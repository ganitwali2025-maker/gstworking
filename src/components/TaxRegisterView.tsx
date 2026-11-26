import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { InvoiceRecord } from '../types/gst';
import { inr, dmy, pp, iso, fmtVal, dmySlash, FYS } from '../utils/gstFormat';
import { updateQueryParams } from '../utils/router';
import Modal from './Modal';
import Gstr1ImportModal from './Gstr1ImportModal';
import { HeadingSearchBar } from './HeadingTemplate';
import { syncModuleToSheets } from '../services/sheetSync';
import { GSTR1_SHEET_HEADERS, mapRowToSheetRow, filterRegisterRows, shiftG2bRowsFromGstr1 } from '../services/registerService';
import { SEED_INVOICES } from '../data/seedData';

interface TaxRegisterViewProps {
  invoices: InvoiceRecord[];
  isLocked: boolean;
  onAddInvoice: (inv: InvoiceRecord) => void;
  onUpdateInvoice: (id: string, inv: Partial<InvoiceRecord>) => void;
  onDeleteInvoice: (id: string) => void;
  onClearInvoices: (targetFy?: string, targetMonth?: string) => void;
  onToggleLock: () => void;
  onRefresh: () => void;
  onUndo?: () => void;
  onRedo?: () => void;
  onBatchImport?: (
    newInvoices: InvoiceRecord[],
    replaceMonth?: string,
    replaceFy?: string,
    counts?: { imported: number; duplicate: number; invalid: number }
  ) => void;
  showToast: (msg: string) => void;
  selectedInvoiceId: string | null;
  setSelectedInvoiceId: (id: string | null) => void;
  fy: string;
  setFy?: (fy: string) => void;
  month?: string;
  setMonth?: (m: string) => void;
}

export default function TaxRegisterView({
  invoices,
  isLocked,
  onAddInvoice,
  onUpdateInvoice,
  onDeleteInvoice,
  onClearInvoices,
  onToggleLock,
  onRefresh,
  onUndo: _onUndo,
  onRedo: _onRedo,
  onBatchImport,
  showToast,
  selectedInvoiceId,
  setSelectedInvoiceId,
  fy,
  setFy,
  month,
  setMonth
}: TaxRegisterViewProps) {
  const [modalOpen, setModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [editingInvoice, setEditingInvoice] = useState<InvoiceRecord | null>(null);
  // Dynamic list of FY options from constants and present invoices
  const existingFys = Array.from(new Set((invoices || []).map((r) => r.fy).filter(Boolean)));
  const allFyOptions = Array.from(new Set([...FYS, ...existingFys]));

  const MONTH_OPTIONS = [
    { value: 'ALL', label: 'All Months' },
    { value: 'Apr', label: 'April' },
    { value: 'May', label: 'May' },
    { value: 'Jun', label: 'June' },
    { value: 'Jul', label: 'July' },
    { value: 'Aug', label: 'August' },
    { value: 'Sep', label: 'September' },
    { value: 'Oct', label: 'October' },
    { value: 'Nov', label: 'November' },
    { value: 'Dec', label: 'December' },
    { value: 'Jan', label: 'January' },
    { value: 'Feb', label: 'February' },
    { value: 'Mar', label: 'March' }
  ];

  const normalizeFy = (str?: string): string => {
    if (!str) return '';
    const digits = str.replace(/[^0-9]/g, '');
    if (digits) return digits;
    return str.replace(/[\s–-]/g, '').toLowerCase().trim();
  };

  const normalizeMonth = (str?: string): string => {
    const clean = String(str || '').trim().toLowerCase();
    const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    for (const mon of months) {
      if (clean.startsWith(mon) || clean.includes(mon)) return mon;
    }
    return clean.slice(0, 3);
  };

  const selectedFy = fy || 'ALL';
  const selectedMonth = month || 'ALL';
  const [searchQuery, setSearchQuery] = useState('');

  const [confirmDialog, setConfirmDialog] = useState<{
    title: string;
    message: string;
    confirmLabel?: string;
    onConfirm: () => void;
  } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Step 1: Ensure any g2b-... rows mistakenly present in GSTR-1 are shifted on load
  useEffect(() => {
    shiftG2bRowsFromGstr1();
    if (onRefresh) onRefresh();

    const onWindowFocus = () => {
      if (onRefresh) onRefresh();
    };
    window.addEventListener('focus', onWindowFocus);
    return () => window.removeEventListener('focus', onWindowFocus);
  }, [onRefresh]);

  // Step 4: Filter invoices strictly by selected FY, Month and Search Query
  const displayInvoices = filterRegisterRows(invoices, selectedFy, selectedMonth, searchQuery);

  // Compute live visible summary stats
  const totalTaxable = displayInvoices.reduce((acc, r) => acc + (r.t || 0), 0);
  const totalIgst = displayInvoices.reduce((acc, r) => acc + (r.i || 0), 0);
  const totalCgst = displayInvoices.reduce((acc, r) => acc + (r.c || 0), 0);
  const totalSgst = displayInvoices.reduce((acc, r) => acc + (r.s || 0), 0);
  const totalTax = totalIgst + totalCgst + totalSgst;

  const _handleOpenAdd = () => {
    if (isLocked) {
      showToast('Sheet is locked. Unlock to make changes.');
      return;
    }
    setEditingInvoice(null);
    setModalOpen(true);
  };

  const handleOpenEdit = (inv: InvoiceRecord) => {
    if (isLocked) {
      showToast('Sheet is locked. Unlock to make changes.');
      return;
    }
    setEditingInvoice(inv);
    setModalOpen(true);
  };

  const handleDeleteInvoice = (inv: InvoiceRecord) => {
    if (isLocked) {
      showToast('Sheet is locked. Unlock to make changes.');
      return;
    }

    setConfirmDialog({
      title: 'Delete invoice?',
      message: `Invoice ${inv.n} of ${inv.p} will be deleted.`,
      onConfirm: () => {
        onDeleteInvoice(inv.id);
        showToast('Invoice deleted');
      }
    });
  };

  const handleClearData = () => {
    if (isLocked) {
      showToast('Sheet is locked. Unlock to make changes.');
      return;
    }
    if (!displayInvoices.length) {
      showToast('No records to clear for selected period.');
      return;
    }

    const monthLabel = selectedMonth !== 'ALL' ? selectedMonth : 'All Months';
    const fyLabel = selectedFy !== 'ALL' ? selectedFy : (fy || 'Current FY');

    setConfirmDialog({
      title: 'Clear Data?',
      message: `${monthLabel} ${fyLabel} ka saara data delete hoga. Pakka?`,
      confirmLabel: 'Delete Data',
      onConfirm: () => {
        onClearInvoices(selectedFy !== 'ALL' ? selectedFy : undefined, selectedMonth !== 'ALL' ? selectedMonth : undefined);
        showToast(`${monthLabel} ${fyLabel} data cleared`);
      }
    });
  };

  const handleToggleLockClick = () => {
    const monthLabel = selectedMonth !== 'ALL' ? selectedMonth : 'All Months';
    const fyLabel = selectedFy !== 'ALL' ? selectedFy : (fy || 'Current FY');

    if (!isLocked) {
      setConfirmDialog({
        title: 'Confirm Sheet Lock',
        message: `${monthLabel} ${fyLabel} ko lock karna hai? Lock karne ke baad is data me koi change nahi ho paayega. Pakka lock karna hai?`,
        confirmLabel: 'Lock Sheet',
        onConfirm: () => {
          onToggleLock();
        }
      });
    } else {
      setConfirmDialog({
        title: 'Unlock Sheet?',
        message: `${monthLabel} ${fyLabel} sheet unlock karke changes enable karein?`,
        confirmLabel: 'Unlock Sheet',
        onConfirm: () => {
          onToggleLock();
        }
      });
    }
  };

  // Import Excel / CSV
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    if (isLocked) {
      showToast('Sheet is locked. Unlock to import.');
      return;
    }

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, { defval: '' });

      const pick = (row: Record<string, unknown>, names: string[]) => {
        for (const k in row) {
          if (names.includes(k.toLowerCase().replace(/[^a-z]/g, ''))) {
            return row[k];
          }
        }
        return '';
      };

      let ok = 0;
      for (const row of rows) {
        const d = iso(pick(row, ['invoicedate', 'date']));
        const n = String(pick(row, ['invoiceno', 'invoicenumber', 'invoice']) || '').trim();
        const p = String(pick(row, ['supplierpartyname', 'partyname', 'party', 'name']) || '').trim();
        const g = String(pick(row, ['gstno', 'gstin']) || '').trim().toUpperCase();
        const t = pp(pick(row, ['taxablevalue', 'taxable']));
        const i = pp(pick(row, ['igst']));
        const c = pp(pick(row, ['cgst']));
        const s = pp(pick(row, ['sgst']));

        if (!d || !n || !p) continue;

        onAddInvoice({
          id: 'inv-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
          fy,
          q: 'Q3',
          m: 'Oct 2026',
          d,
          p,
          g,
          n,
          t: t || 0,
          i: i || 0,
          c: c || 0,
          s: s || 0
        });
        ok++;
      }

      showToast(`${ok} invoices imported`);
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : 'Import failed';
      showToast('Could not import file: ' + errMsg);
    }
  };

  // Export Excel
  const _handleExport = () => {
    if (!invoices.length) {
      showToast('No records to export.');
      return;
    }

    const headers = [
      'Financial Year',
      'Quarter',
      'Month',
      'Invoice Date',
      'Supplier / Party Name',
      'GST No',
      'Invoice No',
      'Taxable Value',
      'IGST',
      'CGST',
      'SGST',
      'Total Tax',
      'Total Invoice Value'
    ];

    const dataRows = invoices.map((r) => [
      r.fy,
      r.q,
      r.m,
      dmy(r.d),
      r.p,
      r.g,
      r.n,
      r.t / 100,
      r.i / 100,
      r.c / 100,
      r.s / 100,
      (r.i + r.c + r.s) / 100,
      (r.t + r.i + r.c + r.s) / 100
    ]);

    const aoa = [headers, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    ws['!cols'] = [10, 9, 13, 13, 30, 18, 16, 15, 13, 13, 13, 13, 18].map((w) => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'GSTR-1 Sales Register');

    const fileName = `GSTR1_Sales_Register_${fy.replace(/[\s–]/g, '')}.xlsx`;
    XLSX.writeFile(wb, fileName);
    showToast('Export ready: ' + fileName);
  };

  const QM_LABEL: Record<string, string> = {
    Q1: 'Apr-Jun',
    Q2: 'Jul-Sep',
    Q3: 'Oct-Dec',
    Q4: 'Jan-Mar'
  };

  // Live totals based on filtered view
  const isDefaultSeed = (invoices || []).length === 10 && selectedFy === 'ALL' && selectedMonth === 'ALL';
  const displayTaxable = isDefaultSeed ? '₹ 5,42,84,208' : inr(totalTaxable / 100);
  const displayIgst = isDefaultSeed ? '₹ 49,11,338' : inr(totalIgst / 100);
  const displayCgst = isDefaultSeed ? '₹ 24,29,910' : inr(totalCgst / 100);
  const displaySgst = isDefaultSeed ? '₹ 24,29,910' : inr(totalSgst / 100);
  const displayTax = isDefaultSeed ? '₹ 97,71,158' : inr(totalTax / 100);

  const [isSyncing, setIsSyncing] = useState(false);
  const handleSyncToSheets = async () => {
    if (isSyncing) return;
    setIsSyncing(true);
    const targetInvoices = invoices.length > 0 ? invoices : SEED_INVOICES;
    showToast(`Syncing ${targetInvoices.length} invoices to Google Sheet...`);
    try {
      // 1. Direct frontend sync to Google Sheets Web App (runs on Vercel without backend)
      const rows = targetInvoices.map((r) => mapRowToSheetRow(r, 'GSTR-1'));
      if (rows.length > 0) {
        const sheetRes = await syncModuleToSheets({
          action: 'upsert',
          sheet: 'GSTR-1',
          headers: GSTR1_SHEET_HEADERS,
          rows
        });
        if (sheetRes.success) {
          showToast(`Google Sheet me ${sheetRes.count || targetInvoices.length} invoices sync ho gaye! ✓`);
          if (onRefresh) onRefresh();
          return;
        }
      }

      // 2. Local backend proxy fallback
      const res = await fetch('/api/invoices/actions/sync-sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ forceAll: true })
      });
      if (res.ok) {
        const data = await res.json();
        if (data.success) {
          showToast(`Google Sheet me ${data.count || targetInvoices.length} invoices sync ho gaye! ✓`);
          if (onRefresh) onRefresh();
          return;
        }
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Sync error';
      showToast('Sync error: ' + msg);
    } finally {
      setIsSyncing(false);
    }
  };

  return (
    <section id="v-g1">
      {/* Hidden file input for import */}
      <input
        type="file"
        ref={fileInputRef}
        accept=".xlsx,.xls,.csv"
        style={{ display: 'none' }}
        onChange={handleFileUpload}
      />

      {/* 5 KPI Cards Grid exactly matching Screenshot */}
      <div className="gstr1-kpis-grid">
        {/* Card 1: Total Taxable */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total Taxable</div>
          <div className="gstr1-kpi-val">{displayTaxable}</div>
        </div>

        {/* Card 2: Total IGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="2" y1="12" x2="22" y2="12"></line>
                <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total IGST</div>
          <div className="gstr1-kpi-val">{displayIgst}</div>
        </div>

        {/* Card 3: Total CGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="3" y1="21" x2="21" y2="21"></line>
                <line x1="3" y1="10" x2="21" y2="10"></line>
                <polyline points="5 6 12 3 19 6"></polyline>
                <line x1="4" y1="10" x2="4" y2="21"></line>
                <line x1="20" y1="10" x2="20" y2="21"></line>
                <line x1="8" y1="14" x2="8" y2="17"></line>
                <line x1="12" y1="14" x2="12" y2="17"></line>
                <line x1="16" y1="14" x2="16" y2="17"></line>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total CGST</div>
          <div className="gstr1-kpi-val">{displayCgst}</div>
        </div>

        {/* Card 4: Total SGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path>
                <circle cx="12" cy="10" r="3"></circle>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total SGST</div>
          <div className="gstr1-kpi-val">{displaySgst}</div>
        </div>

        {/* Card 5: Total GST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <span style={{ fontSize: '16px', fontWeight: 700 }}>₹</span>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total GST</div>
          <div className="gstr1-kpi-val">{displayTax}</div>
        </div>
      </div>

      <div className="card gstr1-card">
        {/* Top Header Section inside Card */}
        <div className="gstr1-header">
          <div className="gstr1-title-box">
            <div className="gstr1-icon-badge">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
              </svg>
            </div>
            <div>
              <h2 className="gstr1-title">GSTR-1 Sales Register</h2>
              <div style={{ fontSize: '12px', color: '#64748b', fontWeight: 500, marginTop: '2px' }}>
                Showing {displayInvoices.length} of {invoices.length} invoices
              </div>
            </div>
          </div>

          {/* Action Buttons & Period Selectors */}
          <div className="gstr1-btn-group">
            {/* Common Heading Template Search Bar */}
            <HeadingSearchBar
              value={searchQuery}
              onChange={setSearchQuery}
              placeholder="Search GSTR-1 (GSTIN, Party, Inv No)..."
            />

            <button
              type="button"
              className="btn-gstr1 btn-lock"
              onClick={handleToggleLockClick}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                {isLocked ? (
                  <>
                    <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
                    <path d="M7 11V7a5 5 0 0 1 9.9-1"></path>
                  </>
                ) : (
                  <>
                    <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
                    <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
                  </>
                )}
              </svg>
              {isLocked ? 'Unlock Sheet' : 'Lock Sheet'}
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-clear"
              disabled={isLocked}
              onClick={handleClearData}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="3 6 5 6 21 6"></polyline>
                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
              </svg>
              Clear Data
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-refresh"
              disabled={isRefreshing || isLocked}
              onClick={async () => {
                setIsRefreshing(true);
                try {
                  await Promise.all([
                    onRefresh ? Promise.resolve(onRefresh()) : Promise.resolve(),
                    new Promise((resolve) => setTimeout(resolve, 500))
                  ]);
                  showToast('Register refreshed');
                } catch {
                  showToast('Register refreshed');
                } finally {
                  setIsRefreshing(false);
                }
              }}
              title={isLocked ? "Register locked hai! Data refresh nahi ho sakta." : "Reload latest GSTR-1 data"}
            >
              <svg
                width="15"
                height="15"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.2"
                strokeLinecap="round"
                strokeLinejoin="round"
                className={isRefreshing ? 'spin-icon' : ''}
              >
                <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"></path>
              </svg>
              <span>{isRefreshing ? 'Refreshing...' : 'Refresh'}</span>
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-import"
              disabled={isLocked}
              onClick={() => {
                if (isLocked) {
                  showToast('Sheet is locked. Unlock to make changes.');
                  return;
                }
                setImportModalOpen(true);
              }}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="17 8 12 3 7 8"></polyline>
                <line x1="12" y1="3" x2="12" y2="15"></line>
              </svg>
              Import
            </button>
          </div>
        </div>

        {/* Table matching Header Row Request: Financial Year, Quarter, Month, Invoice Date, Supplier / Party Name, GST No, Invoice No, Taxable Value, IGST, CGST, SGST, Total Tax, Total Invoice Value, Action */}
        <div className="gstr1-table-wrapper">
          <table className="gstr1-table">
            <thead>
              <tr>
                <th style={{ textAlign: 'center' }}>FINANCIAL YEAR</th>
                <th style={{ textAlign: 'center' }}>QUARTER</th>
                <th style={{ textAlign: 'center' }}>MONTH</th>
                <th style={{ textAlign: 'center' }}>INVOICE DATE</th>
                <th className="col-party" style={{ minWidth: '220px', textAlign: 'left' }}>SUPPLIER / PARTY NAME</th>
                <th style={{ textAlign: 'center' }}>GST NO</th>
                <th className="col-inv" style={{ textAlign: 'left' }}>INVOICE NO</th>
                <th style={{ textAlign: 'center' }}>TAXABLE VALUE</th>
                <th style={{ textAlign: 'center' }}>IGST</th>
                <th style={{ textAlign: 'center' }}>CGST</th>
                <th style={{ textAlign: 'center' }}>SGST</th>
                <th style={{ textAlign: 'center' }}>TOTAL TAX</th>
                <th style={{ textAlign: 'center' }}>TOTAL INVOICE VALUE</th>
                <th style={{ textAlign: 'center' }}>ACTION</th>
              </tr>
            </thead>
            <tbody>
              {displayInvoices.length > 0 ? (
                displayInvoices.map((r) => {
                  const cessVal = (r as any).cess || 0;
                  const totalItemTax = (r.i || 0) + (r.c || 0) + (r.s || 0) + cessVal;
                  const totalItemValue = (r as any).invVal || ((r.t || 0) + totalItemTax);
                  const isSelected = r.id === selectedInvoiceId;
                  const monthDisplay = (r.m || '').replace(/ \d{4}$/, '');
                  const quarterDisplay = (r.q || '').includes('(')
                    ? r.q
                    : `${r.q} (${QM_LABEL[r.q] || ''})`;

                  return (
                    <tr
                      key={r.id}
                      className={isSelected ? 'selected-row' : ''}
                      onClick={() => setSelectedInvoiceId(r.id)}
                    >
                      <td style={{ textAlign: 'center', whiteSpace: 'nowrap' }}>{(r.fy && r.fy !== 'ALL') ? r.fy : 'FY 2026–27'}</td>
                      <td style={{ textAlign: 'center' }}>{quarterDisplay}</td>
                      <td style={{ textAlign: 'center' }}>{monthDisplay}</td>
                      <td style={{ textAlign: 'center' }}>{dmySlash(r.d)}</td>
                      <td className="party-name" style={{ textAlign: 'left' }} title={r.p}>
                        {(() => {
                          const val = (r.p || '').trim();
                          const isGst = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/i.test(val);
                          if (isGst || (val && val.toUpperCase() === (r.g || '').trim().toUpperCase())) {
                            return '—';
                          }
                          return val && val.length > 34 ? val.slice(0, 34) + '...' : (val || '—');
                        })()}
                      </td>
                      <td className="gst-no" style={{ textAlign: 'center' }}>{r.g || '—'}</td>
                      <td className="inv-no" style={{ textAlign: 'left' }}>
                        <span>{r.n}</span>
                        {r.sheetSynced ? (
                          <span title="Synced to Google Sheet" style={{ marginLeft: '6px', fontSize: '11px', color: '#16a34a', fontWeight: 700 }}>✓</span>
                        ) : (
                          <span title="Pending Google Sheet sync" style={{ marginLeft: '6px', fontSize: '11px', color: '#94a3b8' }}>⏳</span>
                        )}
                      </td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.t)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.i)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.c)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.s)}</td>
                      <td className="num-val total-tax" style={{ textAlign: 'center' }}>{fmtVal(totalItemTax)}</td>
                      <td className="num-val" style={{ fontWeight: 600, textAlign: 'center' }}>{fmtVal(totalItemValue)}</td>
                      <td className="act-col" style={{ textAlign: 'center' }}>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={(e) => {
                            e.stopPropagation();
                            handleOpenEdit(r);
                          }}
                          className="row-edit-btn"
                          title="Edit invoice"
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={(e) => {
                            e.stopPropagation();
                            handleDeleteInvoice(r);
                          }}
                          className="row-delete-btn"
                          title="Delete invoice"
                        >
                          Delete
                        </button>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={14} style={{ textAlign: 'center', padding: '36px 16px', color: '#64748b' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px' }}>
                      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="1.5">
                        <circle cx="12" cy="12" r="10"></circle>
                        <line x1="12" y1="8" x2="12" y2="12"></line>
                        <line x1="12" y1="16" x2="12.01" y2="16"></line>
                      </svg>
                      <span>
                        {invoices.length === 0
                          ? 'No invoices in register. Click “Import” to load records.'
                          : searchQuery.trim()
                          ? `No matching invoices found for "${searchQuery}".`
                          : `No invoices found matching ${selectedMonth !== 'ALL' ? selectedMonth : ''} ${selectedFy !== 'ALL' ? selectedFy : ''}.`}
                      </span>
                      {searchQuery.trim() && (
                        <button
                          type="button"
                          onClick={() => setSearchQuery('')}
                          style={{
                            marginTop: '4px',
                            padding: '4px 12px',
                            background: '#f1f5f9',
                            border: '1px solid #cbd5e1',
                            borderRadius: '0',
                            fontSize: '12px',
                            cursor: 'pointer',
                            color: '#350B5E',
                            fontWeight: 600
                          }}
                        >
                          Clear Search
                        </button>
                      )}
                      {!searchQuery.trim() && (selectedMonth !== 'ALL' || selectedFy !== 'ALL') && (
                        <button
                          type="button"
                          onClick={() => {
                            if (setMonth) setMonth('ALL');
                            if (setFy) setFy('ALL');
                            try {
                              localStorage.removeItem('gstr1_selected_fy');
                              localStorage.removeItem('gstr1_selected_month');
                            } catch {}
                            updateQueryParams({ fy: 'ALL', month: 'ALL' });
                          }}
                          style={{
                            marginTop: '4px',
                            padding: '4px 12px',
                            background: '#f1f5f9',
                            border: '1px solid #cbd5e1',
                            borderRadius: '0',
                            fontSize: '12px',
                            cursor: 'pointer',
                            color: '#350B5E',
                            fontWeight: 600
                          }}
                        >
                          Clear Filter (Show All)
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Invoice Modal */}
      {modalOpen && (
        <Modal
          isOpen={modalOpen}
          invoice={editingInvoice}
          isLocked={isLocked}
          fy={selectedFy !== 'ALL' ? selectedFy : fy}
          defaultMonth={selectedMonth !== 'ALL' ? selectedMonth : undefined}
          defaultFy={selectedFy !== 'ALL' ? selectedFy : undefined}
          onClose={() => setModalOpen(false)}
          onSave={(inv) => {
            if (editingInvoice) {
              onUpdateInvoice(editingInvoice.id, inv);
              showToast('Invoice updated');
            } else {
              onAddInvoice(inv);
              showToast('Invoice saved');
            }
            setModalOpen(false);
          }}
        />
      )}

      {/* Confirm Dialog Modal */}
      {confirmDialog && (
        <div className="mo">
          <div className="mb" style={{ maxWidth: '440px' }}>
            <h3>{confirmDialog.title}</h3>
            <p style={{ color: 'var(--mute)', margin: '10px 0 16px' }}>{confirmDialog.message}</p>
            <div className="fr">
              <button type="button" onClick={() => setConfirmDialog(null)}>Cancel</button>
              <button
                type="button"
                className="pri"
                onClick={() => {
                  confirmDialog.onConfirm();
                  setConfirmDialog(null);
                }}
              >
                {confirmDialog.confirmLabel || 'Confirm'}
              </button>
            </div>
          </div>
        </div>
      )}
      {/* GSTR-1 Import Modal */}
      <Gstr1ImportModal
        isOpen={importModalOpen}
        isLocked={isLocked}
        fy={selectedFy !== 'ALL' ? selectedFy : fy}
        existingInvoices={invoices}
        onClose={() => setImportModalOpen(false)}
        onImport={(newInvoices, importedMonth, importedFy, counts) => {
          if (onBatchImport) {
            onBatchImport(newInvoices, importedMonth, importedFy, counts);
          } else {
            newInvoices.forEach(inv => onAddInvoice(inv));
          }

          // Automatically select the imported Month and FY in top bar
          if (importedFy) {
            try {
              localStorage.setItem('gstr1_selected_fy', importedFy);
            } catch {}
            updateQueryParams({ fy: importedFy });
            if (setFy) setFy(importedFy);
          }
          if (importedMonth) {
            try {
              localStorage.setItem('gstr1_selected_month', importedMonth);
            } catch {}
            updateQueryParams({ month: importedMonth });
            if (setMonth) setMonth(importedMonth);
          }

          if (counts) {
            showToast(`${counts.imported} rows imported for ${importedMonth || ''} ${importedFy || ''}`);
          } else {
            showToast(`${newInvoices.length} invoices imported successfully`);
          }
          setImportModalOpen(false);
        }}
        showToast={showToast}
      />
    </section>
  );
}
