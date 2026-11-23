import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { InvoiceRecord } from '../types/gst';
import { FYS, QM } from '../utils/gstFormat';
import { derivePeriodFromMonth } from '../services/registerService';

interface Gstr1ImportModalProps {
  isOpen: boolean;
  isLocked: boolean;
  fy: string;
  existingInvoices: InvoiceRecord[];
  onClose: () => void;
  onImport: (
    newInvoices: InvoiceRecord[],
    replaceMonth?: string,
    replaceFy?: string,
    counts?: { imported: number; duplicate: number; invalid: number }
  ) => void;
  showToast: (msg: string) => void;
}

export interface ParsedGstr1Row {
  rowNum: number;
  rawDate: string;
  parsedDate: string;
  partyName: string;
  gstin: string;
  invNo: string;
  taxableVal: number;
  igstVal: number;
  cgstVal: number;
  sgstVal: number;
  pastedTotalTax: number;
  calculatedTotalTax: number;
  calculatedTotalValue: number;
  taxMismatch: boolean;
  taxMismatchDiff: number;
  isValid: boolean;
  missingFields: string[];
  isDuplicate: boolean;
  rowMonth?: string;
}

// Helper: Parse dates in formats dd-mm-yyyy, dd/mm/yyyy, yyyy-mm-dd, Excel date
export function parseExcelDate(raw: unknown): string {
  if (raw === null || raw === undefined) return '';
  const s = String(raw).trim();
  if (!s) return '';

  // Excel serial number (numeric, e.g. 45000 to 55000)
  if (/^\d{5}$/.test(s)) {
    const serial = parseInt(s, 10);
    const ms = (serial - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d.getTime())) {
      const yyyy = d.getUTCFullYear();
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(d.getUTCDate()).padStart(2, '0');
      return `${yyyy}-${mm}-${dd}`;
    }
  }

  // YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD
  const ymd = s.match(/^(\d{4})[-\/\.](\d{1,2})[-\/\.](\d{1,2})/);
  if (ymd) {
    return `${ymd[1]}-${ymd[2].padStart(2, '0')}-${ymd[3].padStart(2, '0')}`;
  }

  // DD-MM-YYYY or DD/MM/YYYY or DD.MM.YYYY
  const dmy = s.match(/^(\d{1,2})[-\/\.](\d{1,2})[-\/\.](\d{2,4})/);
  if (dmy) {
    const yr = dmy[3].length === 2 ? '20' + dmy[3] : dmy[3];
    return `${yr}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
  }

  // General Date parse fallback
  const dt = new Date(s);
  if (!isNaN(dt.getTime())) {
    const yyyy = dt.getFullYear();
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const dd = String(dt.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  return '';
}

// Helper: Parse currency numbers, strip commas, rupee signs, spaces
export function parseCleanNumber(raw: unknown): number {
  if (raw === null || raw === undefined) return 0;
  if (typeof raw === 'number') {
    return isNaN(raw) ? 0 : raw;
  }
  const s = String(raw).replace(/[₹,\s]/g, '').trim();
  if (!s) return 0;
  const num = parseFloat(s);
  return isNaN(num) ? 0 : num;
}

// Helper: Column normalizer
function normalizeColName(name: string): string {
  return name.toLowerCase().replace(/[^a-z0-9]/g, '');
}

export default function Gstr1ImportModal({
  isOpen,
  isLocked,
  fy,
  existingInvoices,
  onClose,
  onImport,
  showToast
}: Gstr1ImportModalProps) {
  if (!isOpen) return null;

  // 1. Dropdown states
  const [selectedFy, setSelectedFy] = useState<string>((fy && fy !== 'ALL') ? fy : FYS[0]);
  const [selectedQuarter, setSelectedQuarter] = useState<string>('');
  const [selectedMonth, setSelectedMonth] = useState<string>('');

  // 2. Data input states
  const [pastedText, setPastedText] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Quarter-to-month mapping
  const availableMonths = useMemo(() => {
    if (!selectedQuarter) return [];
    return QM[selectedQuarter] || [];
  }, [selectedQuarter]);

  // Handle Quarter Change
  const handleQuarterChange = (q: string) => {
    setSelectedQuarter(q);
    // If current selected month is not in new quarter, reset it
    const months = QM[q] || [];
    if (!months.includes(selectedMonth)) {
      setSelectedMonth('');
    }
  };

  // 3. Parse pasted data
  const { parsedRows, validCount, duplicateCount, invalidCount, warningCount, validImportableRows } = useMemo(() => {
    if (!pastedText.trim()) {
      return {
        parsedRows: [] as ParsedGstr1Row[],
        validCount: 0,
        duplicateCount: 0,
        invalidCount: 0,
        warningCount: 0,
        validImportableRows: [] as ParsedGstr1Row[]
      };
    }

    const lines = pastedText.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    if (lines.length < 2) {
      return {
        parsedRows: [] as ParsedGstr1Row[],
        validCount: 0,
        duplicateCount: 0,
        invalidCount: 0,
        warningCount: 0,
        validImportableRows: [] as ParsedGstr1Row[]
      };
    }

    // Determine delimiter (tab or comma)
    const headerLine = lines[0];
    const isTab = headerLine.includes('\t');
    const delimiter = isTab ? '\t' : ',';

    const headerCols = headerLine.split(delimiter).map((c) => c.trim());
    const normCols = headerCols.map(normalizeColName);

    const isGstHeader = (h: string) => {
      const clean = h.toLowerCase();
      if (clean.includes('cgst') || clean.includes('sgst') || clean.includes('igst') || clean.includes('utgst')) return false;
      return clean.includes('gstin') || clean.includes('gstno') || clean === 'gst' || clean.includes('gstnumber') || clean.includes('tin') || clean.includes('uin');
    };

    const isPartyHeader = (h: string) => {
      const clean = h.toLowerCase();
      if (isGstHeader(clean)) return false;
      return clean.includes('trade') ||
        clean.includes('legalname') ||
        clean.includes('party') ||
        clean.includes('supplier') ||
        clean.includes('transporter') ||
        clean.includes('customer') ||
        clean.includes('buyer') ||
        clean === 'name' ||
        clean.includes('nameof') ||
        clean.includes('particulars');
    };

    // Map column indices by header name or fallback to 9-column order
    let dateIdx = normCols.findIndex((c) => c.includes('date'));
    let gstIdx = normCols.findIndex(isGstHeader);
    let partyIdx = normCols.findIndex(isPartyHeader);
    let invIdx = normCols.findIndex((c) => (c.includes('no') || c.includes('num') || c.includes('inv')) && !c.includes('date') && !isGstHeader(c));
    let taxableIdx = normCols.findIndex((c) => c.includes('taxable') || c.includes('basic'));
    let igstIdx = normCols.findIndex((c) => c.includes('igst') || c.includes('integrated'));
    let cgstIdx = normCols.findIndex((c) => c.includes('cgst') || c.includes('central'));
    let sgstIdx = normCols.findIndex((c) => c.includes('sgst') || c.includes('state') || c.includes('utgst'));
    let totalTaxIdx = normCols.findIndex((c) => c.includes('totaltax') || (c.includes('tax') && !c.includes('taxable') && !c.includes('igst') && !c.includes('cgst') && !c.includes('sgst')));
    let monthIdx = normCols.findIndex((c) => c === 'month' || c.includes('month') || c === 'period');

    // Fallbacks to default 9-column order if not found
    if (dateIdx === -1) dateIdx = 0;
    if (partyIdx === -1) partyIdx = 1;
    if (gstIdx === -1) gstIdx = 2;
    if (invIdx === -1) invIdx = 3;
    if (taxableIdx === -1) taxableIdx = 4;
    if (igstIdx === -1) igstIdx = 5;
    if (cgstIdx === -1) cgstIdx = 6;
    if (sgstIdx === -1) sgstIdx = 7;
    if (totalTaxIdx === -1) totalTaxIdx = 8;

    const isGstRegex = (str: string) => /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/i.test((str || '').trim());

    const rows: ParsedGstr1Row[] = [];
    const seenBatch = new Set<string>();

    for (let i = 1; i < lines.length; i++) {
      const line = lines[i];
      const tokens = line.split(delimiter).map((t) => t.trim());
      if (tokens.every((t) => !t)) continue; // skip blank row

      const rawDate = tokens[dateIdx] || '';
      let rawPartyName = tokens[partyIdx] || '';
      let rawGstin = (tokens[gstIdx] || '').toUpperCase();

      // Intelligent swap / cleanup: Supplier / Party Name should NEVER be a GST number!
      if (isGstRegex(rawPartyName) && !isGstRegex(rawGstin)) {
        const temp = rawGstin;
        rawGstin = rawPartyName;
        rawPartyName = temp;
      } else if (isGstRegex(rawPartyName) && (rawPartyName.trim().toUpperCase() === rawGstin.trim().toUpperCase() || !rawGstin)) {
        rawPartyName = '';
      }

      const partyName = rawPartyName;
      const gstin = rawGstin;
      const invNo = tokens[invIdx] || '';
      const rowMonth = monthIdx !== -1 ? (tokens[monthIdx] || '') : '';
      const taxableVal = parseCleanNumber(tokens[taxableIdx]);
      const igstVal = parseCleanNumber(tokens[igstIdx]);
      const cgstVal = parseCleanNumber(tokens[cgstIdx]);
      const sgstVal = parseCleanNumber(tokens[sgstIdx]);
      const pastedTotalTax = parseCleanNumber(tokens[totalTaxIdx]);

      const parsedDate = parseExcelDate(rawDate);
      const calculatedTotalTax = igstVal + cgstVal + sgstVal;
      const calculatedTotalValue = taxableVal + calculatedTotalTax;

      // Tax mismatch check (> 0.50 difference)
      const rawTaxToken = (tokens[totalTaxIdx] || '').trim();
      const taxMismatch = rawTaxToken !== '' && Math.abs(pastedTotalTax - calculatedTotalTax) > 0.5;
      const taxMismatchDiff = Math.abs(pastedTotalTax - calculatedTotalTax);

      // Validation check
      const missingFields: string[] = [];
      if (!rawDate || !parsedDate) missingFields.push('INVOICE DATE');
      if (!partyName) missingFields.push('PARTY NAME');
      if (!gstin) missingFields.push('GST NO');
      if (!invNo) missingFields.push('INVOICE NO');

      const isValid = missingFields.length === 0;

      // Duplicate check (same INVOICE NO + same GST NO)
      let isDuplicate = false;
      const batchKey = `${invNo.toLowerCase()}_${gstin.toLowerCase()}`;
      if (isValid) {
        const existsInDb = existingInvoices.some(
          (x) =>
            (x.n || '').trim().toLowerCase() === invNo.toLowerCase() &&
            (x.g || '').trim().toUpperCase() === gstin.toUpperCase()
        );
        if (existsInDb || seenBatch.has(batchKey)) {
          isDuplicate = true;
        } else {
          seenBatch.add(batchKey);
        }
      }

      rows.push({
        rowNum: i,
        rawDate,
        parsedDate,
        partyName,
        gstin,
        invNo,
        taxableVal,
        igstVal,
        cgstVal,
        sgstVal,
        pastedTotalTax,
        calculatedTotalTax,
        calculatedTotalValue,
        taxMismatch,
        taxMismatchDiff,
        isValid,
        missingFields,
        isDuplicate,
        rowMonth
      });
    }

    const validRows = rows.filter((r) => r.isValid && !r.isDuplicate);
    const dupCount = rows.filter((r) => r.isValid && r.isDuplicate).length;
    const invCount = rows.filter((r) => !r.isValid).length;
    const warnCount = rows.filter((r) => r.isValid && r.taxMismatch).length;

    return {
      parsedRows: rows,
      validCount: validRows.length,
      duplicateCount: dupCount,
      invalidCount: invCount,
      warningCount: warnCount,
      validImportableRows: validRows
    };
  }, [pastedText, existingInvoices]);

  // Validation: Quarter and Month must be selected
  const isPeriodSelected = Boolean(selectedFy && selectedQuarter && selectedMonth);
  const canImport = isPeriodSelected && validImportableRows.length > 0 && !isLocked;

  // Handle Excel file upload alternative
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      const firstSheet = wb.Sheets[wb.SheetNames[0]];
      const tsv = XLSX.utils.sheet_to_csv(firstSheet, { FS: '\t' });
      setPastedText(tsv);
      showToast('Excel file loaded. Please review preview below.');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Error reading Excel file';
      showToast('Error: ' + msg);
    }
  };

  // Execute Import
  const handleExecuteImport = () => {
    if (!canImport) {
      if (!isPeriodSelected) {
        showToast('Pehle Financial Year, Quarter aur Month select karo.');
      }
      return;
    }

    const QM_LABEL: Record<string, string> = {
      Q1: 'Apr-Jun',
      Q2: 'Jul-Sep',
      Q3: 'Oct-Dec',
      Q4: 'Jan-Mar'
    };
    const qFormatted = QM_LABEL[selectedQuarter]
      ? `${selectedQuarter} (${QM_LABEL[selectedQuarter]})`
      : selectedQuarter;

    // Step 3: Build InvoiceRecord items deriving FY & Q from Month, preserving Month
    const newRecords: InvoiceRecord[] = validImportableRows.map((r) => {
      const rowMonth = r.rowMonth || selectedMonth;
      const period = derivePeriodFromMonth(rowMonth, selectedFy);
      return {
        id: 'inv-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        fy: period.fy,
        q: period.q,
        m: rowMonth || period.m,
        d: r.parsedDate,
        p: r.partyName,
        g: r.gstin,
        n: r.invNo,
        t: Math.round(r.taxableVal * 100),
        i: Math.round(r.igstVal * 100),
        c: Math.round(r.cgstVal * 100),
        s: Math.round(r.sgstVal * 100),
        createdAt: new Date().toISOString()
      };
    });

    onImport(
      newRecords,
      selectedMonth,
      selectedFy,
      {
        imported: newRecords.length,
        duplicate: duplicateCount,
        invalid: invalidCount
      }
    );

    onClose();
  };

  return (
    <div className="mo import-modal-overlay">
      <div className="mb import-modal-box">
        {/* Header */}
        <div className="import-modal-header">
          <div>
            <h3 style={{ margin: 0, color: 'var(--itc)' }}>Import Sales Data (GSTR-1)</h3>
            <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--mute)' }}>
              Excel se 9 columns copy karein aur neeche paste karein.
            </p>
          </div>
          <button
            type="button"
            className="import-close-btn"
            onClick={onClose}
            aria-label="Close"
          >
            ✕
          </button>
        </div>

        {/* 1. Three Required Dropdowns */}
        <div className="import-dropdowns-row">
          <div className="import-dropdown-field">
            <label htmlFor="import_fy">Financial Year *</label>
            <select
              id="import_fy"
              value={selectedFy}
              onChange={(e) => setSelectedFy(e.target.value)}
            >
              {FYS.map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>

          <div className="import-dropdown-field">
            <label htmlFor="import_quarter">Current Quarter *</label>
            <select
              id="import_quarter"
              value={selectedQuarter}
              onChange={(e) => handleQuarterChange(e.target.value)}
            >
              <option value="">-- Select Quarter --</option>
              <option value="Q1">Q1 (Apr–Jun)</option>
              <option value="Q2">Q2 (Jul–Sep)</option>
              <option value="Q3">Q3 (Oct–Dec)</option>
              <option value="Q4">Q4 (Jan–Mar)</option>
            </select>
          </div>

          <div className="import-dropdown-field">
            <label htmlFor="import_month">Select Month *</label>
            <select
              id="import_month"
              value={selectedMonth}
              disabled={!selectedQuarter}
              onChange={(e) => setSelectedMonth(e.target.value)}
            >
              <option value="">
                {!selectedQuarter ? '-- Select Quarter First --' : '-- Select Month --'}
              </option>
              {availableMonths.map((m) => (
                <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Validation Warning Alert if Quarter or Month not selected */}
        {!isPeriodSelected && (
          <div className="import-alert-box alert-warning">
            <span style={{ fontSize: '15px' }}>⚠️</span>
            <span>
              <strong>Pehle Financial Year, Quarter aur Month select karo.</strong> (Ye teeno select
              hona zaroori hai).
            </span>
          </div>
        )}

        {/* 2 & 3. Column Order Guide Strip */}
        <div className="import-columns-guide">
          <div style={{ fontWeight: 600, color: '#350B5E', marginBottom: '4px' }}>
            📋 Required 9 Columns Order (Header row ke saath paste karein):
          </div>
          <div className="import-columns-tags">
            <span>1. INVOICE DATE</span>
            <span>2. SUPPLIER / PARTY NAME</span>
            <span>3. GST NO</span>
            <span>4. INVOICE NO</span>
            <span>5. TAXABLE VALUE</span>
            <span>6. IGST</span>
            <span>7. CGST</span>
            <span>8. SGST</span>
            <span>9. TOTAL TAX</span>
          </div>
          <div style={{ fontSize: '12px', color: '#64748b', marginTop: '6px' }}>
            ✨ <em>MONTH, QUARTER, FINANCIAL YEAR aur TOTAL INVOICE VALUE app khud banayega.</em>
          </div>
        </div>

        {/* Paste Area */}
        <div style={{ marginTop: '12px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
            <label htmlFor="paste_area" style={{ fontWeight: 600, fontSize: '13px' }}>
              Paste Clipboard Data:
            </label>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                type="button"
                className="import-mini-btn"
                onClick={() => fileInputRef.current?.click()}
              >
                📁 Upload Excel File Instead
              </button>
              {pastedText && (
                <button
                  type="button"
                  className="import-mini-btn"
                  onClick={() => setPastedText('')}
                >
                  Clear Paste
                </button>
              )}
            </div>
            <input
              type="file"
              ref={fileInputRef}
              accept=".xlsx,.xls,.csv"
              style={{ display: 'none' }}
              onChange={handleFileUpload}
            />
          </div>

          <textarea
            id="paste_area"
            className="import-textarea"
            rows={3}
            value={pastedText}
            placeholder="Excel se rows copy karke yahan paste karein (Ctrl+V)...&#10;Pehli row column headers honi chahiye."
            onChange={(e) => setPastedText(e.target.value)}
          />
        </div>

        {/* 6. Preview Section (visible only when data is pasted) */}
        {parsedRows.length > 0 && (
          <div className="import-preview-section">
            {/* Preview Status Summary Bar */}
            <div className="import-preview-summary-bar">
                <span className="import-badge import-badge-success">
                  ✓ {validCount} rows ready to import
                </span>
                {duplicateCount > 0 && (
                  <span className="import-badge import-badge-warning">
                    ⚠️ {duplicateCount} duplicate skipped
                  </span>
                )}
                {invalidCount > 0 && (
                  <span className="import-badge import-badge-danger">
                    ✕ {invalidCount} invalid rows (missing fields)
                  </span>
                )}
                {warningCount > 0 && (
                  <span className="import-badge import-badge-info">
                    ℹ️ {warningCount} tax sum mismatch warnings
                  </span>
                )}
              </div>

              {/* Preview Table */}
              <div className="import-preview-table-container">
                <table className="import-preview-table">
                  <thead>
                    <tr>
                      <th style={{ width: '40px' }}>#</th>
                      <th>STATUS</th>
                      <th>FY</th>
                      <th>QUARTER</th>
                      <th>MONTH</th>
                      <th>INVOICE DATE</th>
                      <th>SUPPLIER / PARTY NAME</th>
                      <th>GST NO</th>
                      <th>INVOICE NO</th>
                      <th>TAXABLE (₹)</th>
                      <th>IGST (₹)</th>
                      <th>CGST (₹)</th>
                      <th>SGST (₹)</th>
                      <th>TOTAL TAX</th>
                      <th>CALCULATED TOTAL</th>
                    </tr>
                  </thead>
                  <tbody>
                    {parsedRows.map((r) => {
                      let statusBadge = (
                        <span className="pill-badge pill-valid">Ready</span>
                      );
                      if (!r.isValid) {
                        statusBadge = (
                          <span
                            className="pill-badge pill-invalid"
                            title={`Missing: ${r.missingFields.join(', ')}`}
                          >
                            Row {r.rowNum}: Missing {r.missingFields.join(', ')}
                          </span>
                        );
                      } else if (r.isDuplicate) {
                        statusBadge = (
                          <span className="pill-badge pill-duplicate">Duplicate</span>
                        );
                      } else if (r.taxMismatch) {
                        statusBadge = (
                          <span
                            className="pill-badge pill-tax-mismatch"
                            title={`Pasted ₹${r.pastedTotalTax} ≠ Sum ₹${r.calculatedTotalTax}`}
                          >
                            Tax Mismatch
                          </span>
                        );
                      }

                      return (
                        <tr
                          key={r.rowNum}
                          className={
                            !r.isValid
                              ? 'row-invalid'
                              : r.isDuplicate
                              ? 'row-duplicate'
                              : r.taxMismatch
                              ? 'row-tax-warn'
                              : ''
                          }
                        >
                          <td style={{ textAlign: 'center', color: '#94a3b8' }}>{r.rowNum}</td>
                          <td>{statusBadge}</td>
                          <td style={{ whiteSpace: 'nowrap', textAlign: 'center', fontSize: '11px', color: '#64748b' }}>{selectedFy}</td>
                          <td style={{ textAlign: 'center', fontSize: '11px', color: '#64748b' }}>{selectedQuarter || '---'}</td>
                          <td style={{ textAlign: 'center', fontSize: '11px', color: '#64748b' }}>{selectedMonth || '---'}</td>
                          <td style={{ whiteSpace: 'nowrap' }}>
                            {r.parsedDate || <span style={{ color: '#ef4444' }}>Invalid Date</span>}
                          </td>
                          <td style={{ maxWidth: '200px', textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap' }}>
                            {r.partyName || <span style={{ color: '#ef4444' }}>Missing</span>}
                          </td>
                          <td style={{ fontFamily: 'monospace' }}>
                            {r.gstin || <span style={{ color: '#ef4444' }}>Missing</span>}
                          </td>
                          <td style={{ fontWeight: 600 }}>
                            {r.invNo || <span style={{ color: '#ef4444' }}>Missing</span>}
                          </td>
                          <td style={{ textAlign: 'right' }}>
                            {r.taxableVal.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td style={{ textAlign: 'right' }}>
                            {r.igstVal.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td style={{ textAlign: 'right' }}>
                            {r.cgstVal.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td style={{ textAlign: 'right' }}>
                            {r.sgstVal.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                          <td
                            style={{
                              textAlign: 'right',
                              color: r.taxMismatch ? '#ea580c' : 'inherit',
                              fontWeight: r.taxMismatch ? 700 : 'normal'
                            }}
                          >
                            {r.calculatedTotalTax.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                            {r.taxMismatch && (
                              <small style={{ display: 'block', fontSize: '10px' }}>
                                (Pasted: {r.pastedTotalTax})
                              </small>
                            )}
                          </td>
                          <td style={{ textAlign: 'right', fontWeight: 700, color: 'var(--itc)' }}>
                            ₹{r.calculatedTotalValue.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

        {/* Modal Actions Footer */}
        <div className="import-modal-footer">
          <div style={{ fontSize: '13px', color: '#64748b' }}>
            {validCount > 0 ? (
              <span>
                Ready to import: <strong>{validCount} invoices</strong> into{' '}
                <strong>
                  {selectedMonth || '---'}, {selectedFy}
                </strong>
              </span>
            ) : (
              <span>No valid invoices to import</span>
            )}
          </div>

          <div style={{ display: 'flex', gap: '10px' }}>
            <button
              type="button"
              className="import-btn-cancel"
              onClick={onClose}
            >
              Cancel
            </button>
            <button
              type="button"
              className="import-btn-submit"
              disabled={!canImport}
              onClick={handleExecuteImport}
              title={
                !isPeriodSelected
                  ? 'Pehle Financial Year, Quarter aur Month select karo.'
                  : validCount === 0
                  ? 'No valid rows to import'
                  : ''
              }
            >
              Import Data ({validCount})
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
