import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { RcmRecord } from './RcmRegisterView';
import { FYS, QM } from '../utils/gstFormat';
import { parseExcelDate, parseCleanNumber } from './Gstr1ImportModal';
import { derivePeriodFromMonth } from '../services/registerService';

interface RcmImportModalProps {
  isOpen: boolean;
  isLocked: boolean;
  fy: string;
  existingRecords: RcmRecord[];
  onClose: () => void;
  onImport: (
    newRecords: RcmRecord[],
    replaceMonth?: string,
    replaceFy?: string,
    counts?: { imported: number; duplicate: number; invalid: number }
  ) => void;
  showToast: (msg: string) => void;
}

export interface ParsedRcmRow {
  rowNum: number;
  rawDate: string;
  parsedDate: string;
  transporter: string;
  lrNo: string;
  amount: number;
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

function normalizeColName(name: string): string {
  return name.toLowerCase().replace(/[^a-z0-9]/g, '');
}

export default function RcmImportModal({
  isOpen,
  isLocked,
  fy: initialFy,
  existingRecords,
  onClose,
  onImport,
  showToast
}: RcmImportModalProps) {
  const [selectedFy, setSelectedFy] = useState<string>(initialFy || FYS[0]);
  const [selectedQuarter, setSelectedQuarter] = useState<string>('Q2 (Jul-Sep)');
  const [selectedMonth, setSelectedMonth] = useState<string>('Aug');
  const [pastedText, setPastedText] = useState<string>('');
  const [replaceExisting, setReplaceExisting] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const quarterMonths = useMemo(() => {
    const qKey = selectedQuarter.slice(0, 2) as keyof typeof QM;
    return QM[qKey] || QM.Q2;
  }, [selectedQuarter]);

  const handleQuarterChange = (qVal: string) => {
    setSelectedQuarter(qVal);
    const qKey = qVal.slice(0, 2) as keyof typeof QM;
    const months = QM[qKey] || QM.Q2;
    if (!months.includes(selectedMonth)) {
      setSelectedMonth(months[0]);
    }
  };

  const isPeriodComplete = Boolean(selectedFy && selectedQuarter && selectedMonth);

  // Parse pasted TSV text
  const { parsedRows, validCount, duplicateCount, invalidCount, taxMismatchCount } = useMemo(() => {
    if (!pastedText.trim()) {
      return { parsedRows: [], validCount: 0, duplicateCount: 0, invalidCount: 0, taxMismatchCount: 0 };
    }

    const lines = pastedText.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length === 0) {
      return { parsedRows: [], validCount: 0, duplicateCount: 0, invalidCount: 0, taxMismatchCount: 0 };
    }

    // Row 0 is header
    const headerCells = lines[0].split('\t').map(normalizeColName);

    const isGstHeader = (h: string) => {
      const clean = h.toLowerCase();
      if (clean.includes('cgst') || clean.includes('sgst') || clean.includes('igst') || clean.includes('utgst')) return false;
      return clean.includes('gstin') || clean.includes('gstno') || clean === 'gst' || clean.includes('gstnumber') || clean.includes('tin') || clean.includes('uin');
    };

    let colIndex = {
      date: headerCells.findIndex((h) => h.includes('entrydate') || h === 'date' || h === 'entry' || h.includes('invoicedate')),
      transporter: headerCells.findIndex((h) => !isGstHeader(h) && (h.includes('transportername') || h.includes('transporter') || h.includes('party') || h.includes('supplier') || h === 'name')),
      lrNo: headerCells.findIndex((h) => (h.includes('lrno') || h.includes('lr') || h.includes('lrnumber') || h.includes('invoiceno') || h.includes('invno')) && !isGstHeader(h)),
      amount: headerCells.findIndex((h) => h === 'amount' || h.includes('taxable') || h.includes('basic') || h.includes('totalamount')),
      igst: headerCells.findIndex((h) => h.includes('igst') || h.includes('integrated')),
      cgst: headerCells.findIndex((h) => h.includes('cgst') || h.includes('central')),
      sgst: headerCells.findIndex((h) => h.includes('sgst') || h.includes('state')),
      totalTax: headerCells.findIndex((h) => h.includes('totaltax') || h.includes('taxamount') || h.includes('tax')),
      month: headerCells.findIndex((h) => h === 'month' || h.includes('month') || h === 'period')
    };

    // Fallback if header doesn't match standard names
    if (colIndex.date === -1 && colIndex.transporter === -1 && headerCells.length >= 4) {
      colIndex = {
        date: 0,
        transporter: 1,
        lrNo: 2,
        amount: 3,
        igst: 4,
        cgst: 5,
        sgst: 6,
        totalTax: 7,
        month: headerCells.length > 8 ? 8 : -1
      };
    }

    // Existing set for duplicate checking
    const existingSet = new Set<string>();
    existingRecords.forEach((r) => {
      const key = `${(r.transporter || '').trim().toUpperCase()}_${(r.lrNo || '').trim().toUpperCase()}_${r.d || ''}`;
      existingSet.add(key);
    });

    const seenInPaste = new Set<string>();
    const rows: ParsedRcmRow[] = [];
    let valid = 0;
    let dup = 0;
    let inv = 0;
    let mismatch = 0;

    for (let i = 1; i < lines.length; i++) {
      const cells = lines[i].split('\t');
      if (cells.length === 0 || cells.every((c) => !c.trim())) continue;

      const getCell = (idx: number) => (idx >= 0 && idx < cells.length ? cells[idx].trim() : '');

      const rawDate = getCell(colIndex.date);
      const parsedDate = parseExcelDate(rawDate);
      const transporter = getCell(colIndex.transporter);
      const lrNo = getCell(colIndex.lrNo);
      const rowMonth = getCell(colIndex.month);

      const amount = parseCleanNumber(getCell(colIndex.amount));
      const igstVal = parseCleanNumber(getCell(colIndex.igst));
      const cgstVal = parseCleanNumber(getCell(colIndex.cgst));
      const sgstVal = parseCleanNumber(getCell(colIndex.sgst));
      const pastedTotalTax = parseCleanNumber(getCell(colIndex.totalTax));

      const calculatedTotalTax = igstVal + cgstVal + sgstVal;
      const calculatedTotalValue = amount + calculatedTotalTax;
      const taxDiff = Math.abs(pastedTotalTax - calculatedTotalTax);
      const taxMismatch = pastedTotalTax > 0 && calculatedTotalTax > 0 && taxDiff > 1.5;

      if (taxMismatch) mismatch++;

      const missingFields: string[] = [];
      if (!parsedDate) missingFields.push('ENTRY DATE');
      if (!transporter) missingFields.push('TRANSPORTER NAME');

      const isValid = missingFields.length === 0;

      const key = `${transporter.toUpperCase()}_${lrNo.toUpperCase()}_${parsedDate}`;
      const isDuplicate = existingSet.has(key) || seenInPaste.has(key);
      if (isValid && !isDuplicate) {
        seenInPaste.add(key);
      }

      if (!isValid) inv++;
      else if (isDuplicate) dup++;
      else valid++;

      rows.push({
        rowNum: i,
        rawDate,
        parsedDate,
        transporter,
        lrNo,
        amount,
        igstVal,
        cgstVal,
        sgstVal,
        pastedTotalTax,
        calculatedTotalTax,
        calculatedTotalValue,
        taxMismatch,
        taxMismatchDiff: taxDiff,
        isValid,
        missingFields,
        isDuplicate,
        rowMonth
      });
    }

    return {
      parsedRows: rows,
      validCount: valid,
      duplicateCount: dup,
      invalidCount: inv,
      taxMismatchCount: mismatch
    };
  }, [pastedText, existingRecords]);

  // Excel File upload handler
  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      const firstSheet = wb.Sheets[wb.SheetNames[0]];
      const tsv = XLSX.utils.sheet_to_csv(firstSheet, { FS: '\t' });
      setPastedText(tsv);
      showToast(`Loaded ${file.name} successfully. Review preview below.`);
    } catch {
      showToast('Error reading Excel file. Please try copy-pasting directly.');
    }
  };

  const handleImportClick = () => {
    if (isLocked) {
      showToast('Sheet is locked. Unlock to import.');
      return;
    }
    if (!isPeriodComplete) {
      showToast('Pehle Financial Year, Quarter aur Month select karo.');
      return;
    }
    if (parsedRows.length === 0) {
      showToast('No data to import.');
      return;
    }

    const rowsToImport = parsedRows.filter((r) => r.isValid && (!r.isDuplicate || replaceExisting));

    if (rowsToImport.length === 0) {
      showToast('No valid rows available to import.');
      return;
    }

    const newRecords: RcmRecord[] = rowsToImport.map((r) => {
      const derived = derivePeriodFromMonth(r.rowMonth || selectedMonth, selectedFy);
      return {
        id: 'rcm-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        m: derived.m,
        q: derived.q,
        fy: derived.fy,
        d: r.parsedDate,
        transporter: r.transporter,
        lrNo: r.lrNo || '—',
        amount: Math.round(r.amount * 100),
        igst: Math.round(r.igstVal * 100),
        cgst: Math.round(r.cgstVal * 100),
        sgst: Math.round(r.sgstVal * 100)
      };
    });

    onImport(
      newRecords,
      replaceExisting ? selectedMonth : undefined,
      replaceExisting ? selectedFy : undefined,
      {
        imported: newRecords.length,
        duplicate: duplicateCount,
        invalid: invalidCount
      }
    );
  };

  if (!isOpen) return null;

  return (
    <div className="mo">
      <div className="mb" style={{ maxWidth: '1040px', width: '95vw', maxHeight: '92vh', overflowY: 'auto', padding: '24px' }}>
        <input
          type="file"
          ref={fileInputRef}
          accept=".xlsx,.xls,.csv"
          style={{ display: 'none' }}
          onChange={handleExcelUpload}
        />

        {/* Modal Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '18px' }}>
          <div>
            <h2 style={{ margin: 0, fontSize: '20px', color: '#1e293b', fontWeight: 700 }}>
              Import (RCM) Register Data
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: '13px', color: '#64748b' }}>
              Paste tab-separated data from Excel along with headers or upload a spreadsheet.
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            style={{
              background: 'none',
              border: 'none',
              fontSize: '22px',
              cursor: 'pointer',
              color: '#94a3b8'
            }}
          >
            ✕
          </button>
        </div>

        {/* Top Dropdowns: FY, Quarter, Month */}
        <div
          style={{
            background: '#f8fafc',
            border: '1px solid #e2e8f0',
            borderRadius: '8px',
            padding: '16px',
            marginBottom: '20px',
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
            gap: '16px'
          }}
        >
          <div>
            <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
              Financial Year <span style={{ color: '#ef4444' }}>*</span>
            </label>
            <select
              value={selectedFy}
              onChange={(e) => setSelectedFy(e.target.value)}
              style={{
                width: '100%',
                padding: '9px 12px',
                borderRadius: '6px',
                border: '1px solid #cbd5e1',
                fontSize: '13px',
                background: '#fff'
              }}
            >
              {FYS.map((y) => (
                <option key={y} value={y}>{y}</option>
              ))}
            </select>
          </div>

          <div>
            <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
              Current Quarter <span style={{ color: '#ef4444' }}>*</span>
            </label>
            <select
              value={selectedQuarter}
              onChange={(e) => handleQuarterChange(e.target.value)}
              style={{
                width: '100%',
                padding: '9px 12px',
                borderRadius: '6px',
                border: '1px solid #cbd5e1',
                fontSize: '13px',
                background: '#fff'
              }}
            >
              <option value="Q1 (Apr-Jun)">Q1 (Apr-Jun)</option>
              <option value="Q2 (Jul-Sep)">Q2 (Jul-Sep)</option>
              <option value="Q3 (Oct-Dec)">Q3 (Oct-Dec)</option>
              <option value="Q4 (Jan-Mar)">Q4 (Jan-Mar)</option>
            </select>
          </div>

          <div>
            <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
              Select Month <span style={{ color: '#ef4444' }}>*</span>
            </label>
            <select
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              style={{
                width: '100%',
                padding: '9px 12px',
                borderRadius: '6px',
                border: '1px solid #cbd5e1',
                fontSize: '13px',
                background: '#fff'
              }}
            >
              {quarterMonths.map((m) => (
                <option key={m} value={m}>{m}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Excel Paste Box */}
        <div style={{ marginBottom: '18px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
            <span style={{ fontSize: '13px', fontWeight: 600, color: '#334155' }}>
              Paste Data from Excel (Header row ke SAATH):
            </span>
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              style={{
                padding: '4px 12px',
                borderRadius: '4px',
                border: '1px solid #cbd5e1',
                background: '#fff',
                fontSize: '12px',
                cursor: 'pointer',
                color: '#334155'
              }}
            >
              📁 Choose Excel File instead
            </button>
          </div>

          <div style={{ fontSize: '12px', color: '#b45309', background: '#fef3c7', padding: '8px 12px', borderRadius: '6px', marginBottom: '8px', border: '1px solid #fde68a' }}>
            ⚠️ <b>Important: Please copy & paste the data ALONG WITH the header row.</b>
            <div style={{ marginTop: '2px', color: '#92400e', fontSize: '11px' }}>
              Expected Columns: <code>ENTRY DATE | TRANSPORTER NAME | TRANSPORTER L.R. NO. | AMOUNT | IGST 5% | CGST 2.5% | SGST 2.5% | TOTAL TAX</code>
            </div>
          </div>

          <textarea
            rows={7}
            value={pastedText}
            onChange={(e) => setPastedText(e.target.value)}
            placeholder={`Paste Excel rows here with header row...\nENTRY DATE\tTRANSPORTER NAME\tTRANSPORTER L.R. NO.\tAMOUNT\tIGST 5%\tCGST 2.5%\tSGST 2.5%\tTOTAL TAX\n03/08/2026\tVRL Logistics Limited\tLR-78901\t150000\t0\t3750\t3750\t7500`}
            style={{
              width: '100%',
              padding: '12px',
              fontFamily: 'monospace',
              fontSize: '12px',
              border: '1px solid #cbd5e1',
              borderRadius: '6px',
              boxSizing: 'border-box',
              resize: 'vertical'
            }}
          />
        </div>

        {/* Validation and Preview Counters */}
        {parsedRows.length > 0 && (
          <div style={{ marginBottom: '16px' }}>
            <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center', marginBottom: '12px' }}>
              <span style={{ fontSize: '12px', padding: '4px 10px', background: '#e0f2fe', color: '#0369a1', borderRadius: '4px', fontWeight: 600 }}>
                Total: {parsedRows.length} rows
              </span>
              <span style={{ fontSize: '12px', padding: '4px 10px', background: '#dcfce7', color: '#15803d', borderRadius: '4px', fontWeight: 600 }}>
                Ready to Import: {validCount} rows
              </span>
              {duplicateCount > 0 && (
                <span style={{ fontSize: '12px', padding: '4px 10px', background: '#fef3c7', color: '#b45309', borderRadius: '4px', fontWeight: 600 }}>
                  Duplicates: {duplicateCount} skipped
                </span>
              )}
              {invalidCount > 0 && (
                <span style={{ fontSize: '12px', padding: '4px 10px', background: '#fee2e2', color: '#b91c1c', borderRadius: '4px', fontWeight: 600 }}>
                  Invalid: {invalidCount} rows
                </span>
              )}
              {taxMismatchCount > 0 && (
                <span style={{ fontSize: '12px', padding: '4px 10px', background: '#ffedd5', color: '#c2410c', borderRadius: '4px', fontWeight: 600 }}>
                  ⚠️ Tax Mismatch in {taxMismatchCount} rows
                </span>
              )}
            </div>

            {/* Preview Table */}
            <div style={{ maxHeight: '220px', overflow: 'auto', border: '1px solid #e2e8f0', borderRadius: '6px' }}>
              <table style={{ width: '100%', fontSize: '11px', borderCollapse: 'collapse', textAlign: 'left' }}>
                <thead style={{ background: '#f1f5f9', position: 'sticky', top: 0, zIndex: 1 }}>
                  <tr>
                    <th style={{ padding: '6px 8px', borderBottom: '1px solid #cbd5e1' }}>#</th>
                    <th style={{ padding: '6px 8px', borderBottom: '1px solid #cbd5e1' }}>ENTRY DATE</th>
                    <th style={{ padding: '6px 8px', borderBottom: '1px solid #cbd5e1' }}>TRANSPORTER</th>
                    <th style={{ padding: '6px 8px', borderBottom: '1px solid #cbd5e1' }}>L.R. NO.</th>
                    <th style={{ padding: '6px 8px', borderBottom: '1px solid #cbd5e1' }}>AMOUNT</th>
                    <th style={{ padding: '6px 8px', borderBottom: '1px solid #cbd5e1' }}>TOTAL TAX</th>
                    <th style={{ padding: '6px 8px', borderBottom: '1px solid #cbd5e1' }}>STATUS</th>
                  </tr>
                </thead>
                <tbody>
                  {parsedRows.slice(0, 15).map((r) => {
                    const rowBg = !r.isValid ? '#fef2f2' : r.isDuplicate ? '#fffbeb' : '#ffffff';
                    return (
                      <tr key={r.rowNum} style={{ background: rowBg, borderBottom: '1px solid #f1f5f9' }}>
                        <td style={{ padding: '6px 8px', color: '#64748b' }}>{r.rowNum}</td>
                        <td style={{ padding: '6px 8px' }}>{r.parsedDate || <span style={{ color: '#ef4444' }}>{r.rawDate || 'Empty'}</span>}</td>
                        <td style={{ padding: '6px 8px' }}>{r.transporter || <span style={{ color: '#ef4444' }}>Missing</span>}</td>
                        <td style={{ padding: '6px 8px' }}>{r.lrNo || '—'}</td>
                        <td style={{ padding: '6px 8px' }}>₹{r.amount.toLocaleString('en-IN')}</td>
                        <td style={{ padding: '6px 8px' }}>
                          ₹{r.calculatedTotalTax.toLocaleString('en-IN')}
                          {r.taxMismatch && (
                            <span title={`Pasted: ${r.pastedTotalTax}, Calc: ${r.calculatedTotalTax}`} style={{ color: '#ea580c', marginLeft: '4px', fontWeight: 'bold' }}>
                              ⚠️
                            </span>
                          )}
                        </td>
                        <td style={{ padding: '6px 8px' }}>
                          {!r.isValid ? (
                            <span style={{ color: '#b91c1c', fontWeight: 600 }}>Missing {r.missingFields.join(', ')}</span>
                          ) : r.isDuplicate ? (
                            <span style={{ color: '#b45309', fontWeight: 600 }}>Duplicate</span>
                          ) : (
                            <span style={{ color: '#15803d', fontWeight: 600 }}>Ready</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {parsedRows.length > 15 && (
              <div style={{ fontSize: '11px', color: '#64748b', marginTop: '6px', textAlign: 'right' }}>
                Showing first 15 of {parsedRows.length} rows
              </div>
            )}
          </div>
        )}

        {/* Replace Option */}
        <div style={{ marginBottom: '20px' }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', cursor: 'pointer', color: '#334155' }}>
            <input
              type="checkbox"
              checked={replaceExisting}
              onChange={(e) => setReplaceExisting(e.target.checked)}
            />
            <span>
              Replace existing records for <b>{selectedMonth} ({selectedFy})</b>
            </span>
          </label>
        </div>

        {/* Footer Actions */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px' }}>
          <button
            type="button"
            onClick={onClose}
            style={{
              padding: '8px 18px',
              borderRadius: '6px',
              border: '1px solid #cbd5e1',
              background: '#fff',
              fontSize: '13px',
              cursor: 'pointer',
              color: '#475569'
            }}
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={handleImportClick}
            disabled={validCount === 0 || isLocked}
            style={{
              padding: '8px 22px',
              borderRadius: '6px',
              border: 'none',
              background: validCount > 0 && !isLocked ? '#350B5E' : '#94a3b8',
              color: '#fff',
              fontSize: '13px',
              fontWeight: 600,
              cursor: validCount > 0 && !isLocked ? 'pointer' : 'not-allowed'
            }}
          >
            Import {validCount} Records
          </button>
        </div>
      </div>
    </div>
  );
}
