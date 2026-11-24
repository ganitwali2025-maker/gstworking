import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { getQueryParams, updateQueryParams } from '../utils/router';

// ==========================================
// Types
// ==========================================
export type TabKey =
  | 'books-reco'
  | 'books-itc'
  | 'govt-2b'
  | 'debit-note'
  | 'not-in-book'
  | 'not-in-2b'
  | 'matched'
  | 'not-matched';

export interface TabItem {
  id: TabKey;
  label: string;
  icon: string;
}

export type RecoStatus =
  | 'Matched'
  | 'Not Match'
  | 'Invoice No. Not Match'
  | 'Not in GSTR-2B'
  | 'Not in 2B'
  | 'Not in Purchase Register'
  | 'Not in Book'
  | 'Debit Note';

export type AuditStatus = 'Verified' | 'Query' | 'Pending';

export interface AuditEntry {
  status: AuditStatus;
  note: string;
  updatedAt: string;
}

export interface RecoInvoiceItem {
  id: string;
  status: RecoStatus;
  month2b: string; // e.g. 'Aug-2026' or '-'
  hasDuplicateMonth?: boolean;
  duplicateMonthsList?: string[];
  date: string;
  invNo: string;
  gstin: string;
  supplierName: string;

  // Amounts in Rupees
  taxable: number;
  igst: number;
  cgst: number;
  sgst: number;
  cess?: number;
  total: number;

  // Tab grouping
  tabCategory: 'books-reco' | 'books-itc' | 'govt-2b' | 'debit-note' | 'not-in-book' | 'not-in-2b' | 'matched' | 'not-matched';

  // Audit info
  auditStatus?: AuditStatus;
  auditNote?: string;
  auditUpdatedAt?: string;
}

// 8 Tabs matching existing UI exactly
export const TABS: TabItem[] = [
  { id: 'books-reco', label: 'Books Reco with 2B', icon: 'layers' },
  { id: 'books-itc', label: 'Books ITC', icon: 'file' },
  { id: 'govt-2b', label: 'Government 2B', icon: 'download' },
  { id: 'debit-note', label: 'Debit Note', icon: 'file' },
  { id: 'not-in-book', label: 'Not in Book', icon: 'file-minus' },
  { id: 'not-in-2b', label: 'Not in 2B', icon: 'layers' },
  { id: 'matched', label: 'Match Invoice', icon: 'check' },
  { id: 'not-matched', label: 'Not Match Invoice', icon: 'layers' }
];

// ==========================================
// Default Seed Records (Matches screenshot exactly: 15 total)
// ==========================================
const SEED_RECORDS: RecoInvoiceItem[] = [
  {
    id: 'rec-1',
    status: 'Not in 2B',
    month2b: '-',
    invNo: '26-27/25',
    date: '05/08/2026',
    gstin: '22BVZPK4622B1ZO',
    supplierName: 'Kalam Khan',
    taxable: 778635,
    igst: 0,
    cgst: 70077.15,
    sgst: 70077.15,
    cess: 0,
    total: 918789.30,
    tabCategory: 'not-in-2b'
  },
  {
    id: 'rec-2',
    status: 'Not in 2B',
    month2b: '-',
    invNo: 'SIES/RMIPL/12-2026-27',
    date: '06/08/2026',
    gstin: '09APBPM1406Q1ZO',
    supplierName: 'Sies Operational Services',
    taxable: 3192433.40,
    igst: 574638.01,
    cgst: 0,
    sgst: 0,
    cess: 0,
    total: 3767071.41,
    tabCategory: 'not-in-2b'
  },
  {
    id: 'rec-3',
    status: 'Not in 2B',
    month2b: '-',
    invNo: 'REFRATECH -14',
    date: '12/08/2026',
    gstin: '22AANCR8181R1ZH',
    supplierName: 'Refratech Application Services Pvt Ltd',
    taxable: 335655,
    igst: 0,
    cgst: 30208.95,
    sgst: 30208.95,
    cess: 0,
    total: 396072.90,
    tabCategory: 'not-in-2b'
  },
  {
    id: 'rec-4',
    status: 'Invoice No. Not Match',
    month2b: 'Aug-2026',
    invNo: 'GI/26-27/318',
    date: '13/08/2026',
    gstin: '21AILPG1979R1ZO',
    supplierName: 'Glaxy International',
    taxable: 1800000,
    igst: 324000,
    cgst: 0,
    sgst: 0,
    cess: 0,
    total: 2124000,
    tabCategory: 'not-matched'
  },
  {
    id: 'rec-5',
    status: 'Not in 2B',
    month2b: '-',
    invNo: 'SKE/26-27/1514',
    date: '20/08/2026',
    gstin: '22ALRPJ2590A1ZD',
    supplierName: 'Shri Kamla Enterprises',
    taxable: 450,
    igst: 0,
    cgst: 40.50,
    sgst: 40.50,
    cess: 0,
    total: 531,
    tabCategory: 'not-in-2b'
  },
  {
    id: 'rec-6',
    status: 'Not in 2B',
    month2b: '-',
    invNo: '5300',
    date: '20/08/2026',
    gstin: '27AAZFB9229K1ZK',
    supplierName: 'Barewell Traders Llp',
    taxable: 28654.32,
    igst: 5157.78,
    cgst: 0,
    sgst: 0,
    cess: 0,
    total: 33812.10,
    tabCategory: 'not-in-2b'
  },
  {
    id: 'rec-7',
    status: 'Not in 2B',
    month2b: '-',
    invNo: 'MST12627/00005',
    date: '20/08/2026',
    gstin: '09IPKPS5665Q1ZG',
    supplierName: 'MICROCARE SOLUTION',
    taxable: 354697.20,
    igst: 63845.50,
    cgst: 0,
    sgst: 0,
    cess: 0,
    total: 418542.70,
    tabCategory: 'not-in-2b'
  },
  {
    id: 'rec-8',
    status: 'Invoice No. Not Match',
    month2b: 'Aug-2026',
    invNo: 'KPS/2026/17',
    date: '20/08/2026',
    gstin: '22CANPV4763J1Z4',
    supplierName: 'Kunwargarh Printing Press & Stationers',
    taxable: 300,
    igst: 0,
    cgst: 27,
    sgst: 27,
    cess: 0,
    total: 354,
    tabCategory: 'not-matched'
  },
  // Debit Notes
  {
    id: 'rec-9',
    status: 'Debit Note',
    month2b: '-',
    invNo: 'DN/2026/089',
    date: '15/08/2026',
    gstin: '22AABCB4567M1Z2',
    supplierName: 'Bharat Steel & Alloys Pvt Ltd',
    taxable: -50000,
    igst: 0,
    cgst: -4500,
    sgst: -4500,
    cess: 0,
    total: -59000,
    tabCategory: 'debit-note'
  },
  {
    id: 'rec-10',
    status: 'Debit Note',
    month2b: '-',
    invNo: 'DN-TK-441',
    date: '22/08/2026',
    gstin: '21AABCT1332N1Z8',
    supplierName: 'Tata Steel Downstream Products Ltd',
    taxable: -25000,
    igst: -4500,
    cgst: 0,
    sgst: 0,
    cess: 0,
    total: -29500,
    tabCategory: 'debit-note'
  },
  // Not in Book
  {
    id: 'rec-11',
    status: 'Not in Book',
    month2b: 'Aug-2026',
    invNo: 'INV-2B-901',
    date: '11/08/2026',
    gstin: '22AABCA1122C1Z9',
    supplierName: 'Precision Engineering Tools Co.',
    taxable: 145000,
    igst: 0,
    cgst: 13050,
    sgst: 13050,
    cess: 0,
    total: 171100,
    tabCategory: 'not-in-book'
  },
  {
    id: 'rec-12',
    status: 'Not in Book',
    month2b: 'Aug-2026',
    invNo: 'INV-2B-902',
    date: '19/08/2026',
    gstin: '27AABCM9988D1Z3',
    supplierName: 'Western Fasteners & Hardware Ltd',
    taxable: 88500,
    igst: 15930,
    cgst: 0,
    sgst: 0,
    cess: 0,
    total: 104430,
    tabCategory: 'not-in-book'
  },
  // Matched Invoices
  {
    id: 'rec-13',
    status: 'Matched',
    month2b: 'Aug-2026',
    invNo: 'BS-4421',
    date: '02/08/2026',
    gstin: '22AABCB4567M1Z2',
    supplierName: 'Bharat Steel & Alloys Pvt Ltd',
    taxable: 420000,
    igst: 0,
    cgst: 37800,
    sgst: 37800,
    cess: 0,
    total: 495600,
    tabCategory: 'matched'
  },
  {
    id: 'rec-14',
    status: 'Matched',
    month2b: 'Aug-2026',
    invNo: 'SK-1092',
    date: '05/08/2026',
    gstin: '22AAFCK8812D1Z4',
    supplierName: 'Shree Krishna Logistics & Spares',
    taxable: 280000,
    igst: 0,
    cgst: 25200,
    sgst: 25200,
    cess: 0,
    total: 330400,
    tabCategory: 'matched'
  },
  {
    id: 'rec-15',
    status: 'Matched',
    month2b: 'Aug-2026',
    invNo: 'TS-9901',
    date: '08/08/2026',
    gstin: '21AABCT1332N1Z8',
    supplierName: 'Tata Steel Downstream Products Ltd',
    taxable: 850000,
    igst: 153000,
    cgst: 0,
    sgst: 0,
    cess: 0,
    total: 1003000,
    tabCategory: 'matched'
  }
];

// ==========================================
// Currency & Formatting Helpers
// ==========================================
export const formatIndianAmount = (val: number): string => {
  if (val === undefined || val === null || isNaN(val)) return '0.00';
  if (val === 0) return '0.00';
  const isNeg = val < 0;
  const abs = Math.abs(val);
  const formatted = abs.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  return (isNeg ? '-' : '') + formatted;
};

interface ReconciliationViewProps {
  showToast?: (msg: string) => void;
  fy?: string;
  month?: string;
}

export default function ReconciliationView({
  showToast = () => { },
  fy: _fy = '2026-27'
}: ReconciliationViewProps) {
  // Active Tab state (synced with router query param ?tab=)
  const [activeTab, setActiveTab] = useState<TabKey>(() => {
    try {
      const q = getQueryParams();
      if (q.tab && TABS.some((t) => t.id === q.tab)) {
        return q.tab as TabKey;
      }
    } catch { }
    return 'books-reco';
  });

  // Filters & Search
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [monthFilter, setMonthFilter] = useState<string>('ALL');

  // Modals state
  const [viewingRecord, setViewingRecord] = useState<RecoInvoiceItem | null>(null);
  const [auditingRecord, setAuditingRecord] = useState<RecoInvoiceItem | null>(null);
  const [auditFormStatus, setAuditFormStatus] = useState<AuditStatus>('Verified');
  const [auditFormNote, setAuditFormNote] = useState<string>('');

  // Persistent Audit Store
  const [auditNotes, setAuditNotes] = useState<Record<string, AuditEntry>>(() => {
    try {
      const saved = localStorage.getItem('gst_reco_audit_v1');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  // Base records from localStorage or SEED_RECORDS
  const [records] = useState<RecoInvoiceItem[]>(() => {
    try {
      const saved = localStorage.getItem('gst_reco_records_v1');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch { }
    return SEED_RECORDS;
  });

  // Sync tab from URL query params (e.g. from WorkingStepsView 3 cards)
  useEffect(() => {
    const syncTab = () => {
      try {
        const q = getQueryParams();
        if (q.tab && TABS.some((t) => t.id === q.tab)) {
          setActiveTab(q.tab as TabKey);
        }
      } catch { }
    };
    syncTab();
    window.addEventListener('popstate', syncTab);
    return () => window.removeEventListener('popstate', syncTab);
  }, []);

  const handleTabChange = (tabId: TabKey) => {
    setActiveTab(tabId);
    updateQueryParams({ tab: tabId });
  };

  // Attach audit details to records
  const enrichedRecords = useMemo(() => {
    return records.map((r) => {
      const audit = auditNotes[r.id] || auditNotes[r.invNo];
      return {
        ...r,
        auditStatus: audit?.status || 'Pending',
        auditNote: audit?.note || '',
        auditUpdatedAt: audit?.updatedAt
      };
    });
  }, [records, auditNotes]);

  // Tab count helper for badges
  const getTabCount = (tabId: TabKey) => {
    if (tabId === 'books-reco') return enrichedRecords.length;
    if (tabId === 'books-itc') return enrichedRecords.filter((r) => ['not-in-2b', 'matched', 'not-matched'].includes(r.tabCategory)).length;
    if (tabId === 'govt-2b') return enrichedRecords.filter((r) => ['not-in-book', 'matched', 'not-matched'].includes(r.tabCategory)).length;
    if (tabId === 'debit-note') return enrichedRecords.filter((r) => r.tabCategory === 'debit-note' || r.status === 'Debit Note').length;
    if (tabId === 'not-in-book') return enrichedRecords.filter((r) => r.tabCategory === 'not-in-book' || r.status === 'Not in Book' || r.status === 'Not in Purchase Register').length;
    if (tabId === 'not-in-2b') return enrichedRecords.filter((r) => r.tabCategory === 'not-in-2b' || r.status === 'Not in 2B' || r.status === 'Not in GSTR-2B').length;
    if (tabId === 'matched') return enrichedRecords.filter((r) => r.tabCategory === 'matched' || r.status === 'Matched').length;
    if (tabId === 'not-matched') return enrichedRecords.filter((r) => r.tabCategory === 'not-matched' || r.status === 'Invoice No. Not Match' || r.status === 'Not Match').length;
    return 0;
  };

  // Distinct months for month dropdown filter
  const availableMonths = useMemo(() => {
    const set = new Set<string>();
    enrichedRecords.forEach((r) => {
      if (r.month2b && r.month2b !== '-') set.add(r.month2b);
    });
    return Array.from(set).sort();
  }, [enrichedRecords]);

  // Check if current view is one of the 3 dedicated user sheets
  const isSheet1 = activeTab === 'matched';
  const isSheet2 = activeTab === 'not-in-book';
  const isSheet3 = activeTab === 'not-in-2b';
  const isDedicatedSheet = isSheet1 || isSheet2 || isSheet3;

  // Filter records by Tab, Debit Note rule, Month filter, and Search Query
  const displayedRecords = useMemo(() => {
    let list = enrichedRecords;

    // 1. Tab category filter
    if (activeTab === 'books-reco') {
      list = enrichedRecords;
    } else if (activeTab === 'books-itc') {
      list = enrichedRecords.filter((r) => ['not-in-2b', 'matched', 'not-matched', 'books-reco'].includes(r.tabCategory));
    } else if (activeTab === 'govt-2b') {
      list = enrichedRecords.filter((r) => ['not-in-book', 'matched', 'not-matched'].includes(r.tabCategory));
    } else if (activeTab === 'debit-note') {
      list = enrichedRecords.filter((r) => r.tabCategory === 'debit-note' || r.status === 'Debit Note');
    } else if (activeTab === 'not-in-book') {
      list = enrichedRecords.filter((r) => r.tabCategory === 'not-in-book' || r.status === 'Not in Book' || r.status === 'Not in Purchase Register');
    } else if (activeTab === 'not-in-2b') {
      list = enrichedRecords.filter((r) => r.tabCategory === 'not-in-2b' || r.status === 'Not in 2B' || r.status === 'Not in GSTR-2B');
    } else if (activeTab === 'matched') {
      list = enrichedRecords.filter((r) => r.tabCategory === 'matched' || r.status === 'Matched');
    } else if (activeTab === 'not-matched') {
      list = enrichedRecords.filter((r) => r.tabCategory === 'not-matched' || r.status === 'Invoice No. Not Match' || r.status === 'Not Match');
    }

    // 2. NIYAM: Debit Note (minus amount wali row) in sheets me mat dikhao
    if (isDedicatedSheet) {
      list = list.filter((r) => r.status !== 'Debit Note' && (r.taxable || 0) >= 0 && (r.total || 0) >= 0);
    }

    // 3. Month Filter
    if (monthFilter !== 'ALL') {
      list = list.filter((r) => r.month2b === monthFilter);
    }

    // 4. Search Query Filter (Invoice No, Supplier Name, GSTIN)
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      list = list.filter(
        (r) =>
          (r.invNo || '').toLowerCase().includes(q) ||
          (r.supplierName || '').toLowerCase().includes(q) ||
          (r.gstin || '').toLowerCase().includes(q)
      );
    }

    return list;
  }, [enrichedRecords, activeTab, isDedicatedSheet, monthFilter, searchQuery]);

  // Totals calculation
  const totals = useMemo(() => {
    return displayedRecords.reduce(
      (acc, r) => {
        acc.taxable += r.taxable || 0;
        acc.igst += r.igst || 0;
        acc.cgst += r.cgst || 0;
        acc.sgst += r.sgst || 0;
        acc.cess += r.cess || 0;
        acc.total += r.total || 0;
        return acc;
      },
      { taxable: 0, igst: 0, cgst: 0, sgst: 0, cess: 0, total: 0 }
    );
  }, [displayedRecords]);

  // Handle Save Audit
  const handleSaveAudit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!auditingRecord) return;

    const key = auditingRecord.id;
    const updatedAudit: AuditEntry = {
      status: auditFormStatus,
      note: auditFormNote.trim(),
      updatedAt: new Date().toLocaleDateString('en-IN', {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
      })
    };

    const newNotes = { ...auditNotes, [key]: updatedAudit, [auditingRecord.invNo]: updatedAudit };
    setAuditNotes(newNotes);
    try {
      localStorage.setItem('gst_reco_audit_v1', JSON.stringify(newNotes));
    } catch { }

    setAuditingRecord(null);
    showToast(`Audit note saved for ${auditingRecord.invNo}`);
  };

  // Handle Export to Excel
  const handleExportExcel = () => {
    try {
      let exportRows: any[] = [];

      if (isSheet1) {
        // SHEET 1
        exportRows = displayedRecords.map((r, i) => ({
          'S.NO': i + 1,
          'STATUS': r.status,
          '2B MONTH': r.month2b,
          'DATE': r.date,
          'INVOICE NO.': r.invNo,
          'SUPPLIER GSTIN': r.gstin,
          'SUPPLIER NAME': r.supplierName,
          'TAXABLE': r.taxable,
          'IGST': r.igst,
          'CGST': r.cgst,
          'SGST': r.sgst,
          'TOTAL': r.total,
          'AUDIT STATUS': r.auditStatus || 'Pending',
          'AUDIT NOTE': r.auditNote || ''
        }));
      } else if (isSheet2) {
        // SHEET 2
        exportRows = displayedRecords.map((r, i) => ({
          'S.NO': i + 1,
          'STATUS': r.status,
          '2B MONTH': r.month2b,
          'DATE': r.date,
          'INVOICE NO.': r.invNo,
          'SUPPLIER GSTIN': r.gstin,
          'SUPPLIER NAME': r.supplierName,
          'TAXABLE (2B)': r.taxable,
          'IGST (2B)': r.igst,
          'CGST (2B)': r.cgst,
          'SGST (2B)': r.sgst,
          'TOTAL (2B)': r.total,
          'AUDIT STATUS': r.auditStatus || 'Pending',
          'AUDIT NOTE': r.auditNote || ''
        }));
      } else if (isSheet3) {
        // SHEET 3 (NO 2B Month column)
        exportRows = displayedRecords.map((r, i) => ({
          'S.NO': i + 1,
          'STATUS': r.status,
          'DATE': r.date,
          'INVOICE NO.': r.invNo,
          'SUPPLIER GSTIN': r.gstin,
          'SUPPLIER NAME': r.supplierName,
          'TAXABLE (BOOKS)': r.taxable,
          'IGST (BOOKS)': r.igst,
          'CGST (BOOKS)': r.cgst,
          'SGST (BOOKS)': r.sgst,
          'TOTAL (BOOKS)': r.total,
          'AUDIT STATUS': r.auditStatus || 'Pending',
          'AUDIT NOTE': r.auditNote || ''
        }));
      } else {
        // Default / other tabs
        exportRows = displayedRecords.map((r, i) => ({
          'S.NO': i + 1,
          'STATUS': r.status,
          'DATE': r.date,
          'INVOICE NO.': r.invNo,
          'SUPPLIER GSTIN': r.gstin,
          'SUPPLIER NAME': r.supplierName,
          'TAXABLE': r.taxable,
          'IGST': r.igst,
          'CGST': r.cgst,
          'SGST': r.sgst,
          'CESS': r.cess || 0,
          'TOTAL': r.total
        }));
      }

      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, activeTab);
      const filename = `GST_Reco_${activeTab}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, filename);
      showToast('Excel exported successfully!');
    } catch (e) {
      console.error(e);
      showToast('Failed to export Excel.');
    }
  };

  const handlePrint = () => {
    window.print();
  };

  // Helper for status badge styling
  const getStatusBadge = (status: RecoStatus) => {
    let pillClass = 'pill-matched';
    let label = status;

    if (status === 'Matched') {
      pillClass = 'pill-matched';
    } else if (status === 'Invoice No. Not Match' || status === 'Not Match') {
      pillClass = 'pill-not-match';
    } else if (status === 'Not in 2B' || status === 'Not in GSTR-2B') {
      pillClass = 'pill-not-in-2b';
      label = isSheet3 ? 'Not in GSTR-2B' : status;
    } else if (status === 'Not in Book' || status === 'Not in Purchase Register') {
      pillClass = 'pill-not-in-book';
      label = isSheet2 ? 'Not in Purchase Register' : status;
    } else if (status === 'Debit Note') {
      pillClass = 'pill-debit-note';
    }

    return (
      <span className={`reco-status-pill ${pillClass}`}>
        <span className="reco-status-dot">●</span>
        <span>{label}</span>
      </span>
    );
  };

  return (
    <section id="v-rec" className="reco-page-container">
      <div className="card reco-card">
        {/* ======================================================== */}
        {/* 1. Header (Icon, Title, Search, Month Filter, Actions)   */}
        {/* ======================================================== */}
        <div className="reco-header-card">
          <div className="reco-header-left">
            <div className="reco-icon-badge" style={{ position: 'relative' }}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
              </svg>
              <div className="reco-badge-check">✓</div>
            </div>
            <div>
              <h1 className="reco-title">GST RECO MANAGER</h1>
              <p className="reco-subtitle">
                Automated Reconciliation of Purchase Books with GSTR-2B to Maximize Eligible ITC Claims
              </p>
            </div>
          </div>

          <div className="reco-header-right">
            {/* Month Filter Dropdown */}
            {availableMonths.length > 0 && (
              <select
                value={monthFilter}
                onChange={(e) => setMonthFilter(e.target.value)}
                className="reco-select-field"
                style={{ height: '38px', padding: '0 12px', fontSize: '13px' }}
                title="Filter by 2B Month"
              >
                <option value="ALL">All Months</option>
                {availableMonths.map((m) => (
                  <option key={m} value={m}>{m}</option>
                ))}
              </select>
            )}

            {/* Search Box */}
            <div className="reco-search-wrap" style={{ height: '38px', padding: '0 12px' }}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="11" cy="11" r="8"></circle>
                <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
              </svg>
              <input
                type="text"
                placeholder="Search invoice / GSTIN..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="reco-search-input"
              />
              {searchQuery && (
                <button
                  type="button"
                  className="reco-search-clear"
                  onClick={() => setSearchQuery('')}
                  title="Clear search"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Print Button */}
            <button type="button" className="btn-reco-action" onClick={handlePrint} title="Print Report">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="6 9 6 2 18 2 18 9"></polyline>
                <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path>
                <rect x="6" y="14" width="12" height="8"></rect>
              </svg>
              <span>Print</span>
            </button>

            {/* Excel Button */}
            <button type="button" className="btn-reco-action" onClick={handleExportExcel} title="Export Excel">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
                <line x1="3" y1="9" x2="21" y2="9"></line>
                <line x1="3" y1="15" x2="21" y2="15"></line>
                <line x1="9" y1="3" x2="9" y2="21"></line>
                <line x1="15" y1="3" x2="15" y2="21"></line>
              </svg>
              <span>Excel</span>
            </button>
          </div>
        </div>

        {/* ======================================================== */}
        {/* 2. Tabs Bar (8 Tabs exactly matching screenshot)         */}
        {/* ======================================================== */}
        <div className="reco-tabs-bar" style={{ padding: '12px 20px', background: '#F8FAFC', borderBottom: '1px solid #E2E8F0' }}>
          {TABS.map((tab) => {
            const isActive = activeTab === tab.id;
            const count = getTabCount(tab.id);

            return (
              <button
                key={tab.id}
                type="button"
                className={`reco-tab ${isActive ? 'active' : ''}`}
                onClick={() => handleTabChange(tab.id)}
              >
                {tab.icon === 'layers' && (
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <polygon points="12 2 2 7 12 12 22 7 12 2"></polygon>
                    <polyline points="2 17 12 22 22 17"></polyline>
                    <polyline points="2 12 12 17 22 12"></polyline>
                  </svg>
                )}
                {tab.icon === 'file' && (
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                    <polyline points="14 2 14 8 20 8"></polyline>
                  </svg>
                )}
                {tab.icon === 'download' && (
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="7 10 12 15 17 10"></polyline>
                    <line x1="12" y1="15" x2="12" y2="3"></line>
                  </svg>
                )}
                {tab.icon === 'file-minus' && (
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                    <polyline points="14 2 14 8 20 8"></polyline>
                    <line x1="9" y1="15" x2="15" y2="15"></line>
                  </svg>
                )}
                {tab.icon === 'check' && (
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
                    <polyline points="20 6 9 17 4 12"></polyline>
                  </svg>
                )}
                <span>{tab.label}</span>
                <span className="reco-tab-badge">{count}</span>
              </button>
            );
          })}
        </div>

        {/* ======================================================== */}
        {/* 3. Table with Tab-Specific Column Structures            */}
        {/* ======================================================== */}
        <div className="reco-table-wrapper">
          <table className="reco-table">
            <thead>
              <tr>
                {/* S.No is always column 1 */}
                <th style={{ width: '55px', textAlign: 'center' }}>S.NO</th>
                <th style={{ minWidth: '140px', textAlign: 'left' }}>STATUS</th>

                {/* 2B Month: Shown on Sheet 1, Sheet 2, and 2B tabs, but NOT on Sheet 3 */}
                {!isSheet3 && (
                  <th style={{ minWidth: '100px', textAlign: 'center' }}>2B MONTH</th>
                )}

                <th style={{ minWidth: '95px', textAlign: 'center' }}>DATE</th>
                <th style={{ minWidth: '135px', textAlign: 'left' }}>INVOICE NO.</th>
                <th style={{ minWidth: '150px', textAlign: 'center' }}>SUPPLIER GSTIN</th>
                <th style={{ minWidth: '220px', textAlign: 'left' }}>SUPPLIER NAME</th>

                {/* Tax & Total Columns Header */}
                {isSheet2 ? (
                  <>
                    <th style={{ textAlign: 'right' }}>TAXABLE (2B)</th>
                    <th style={{ textAlign: 'right' }}>IGST (2B)</th>
                    <th style={{ textAlign: 'right' }}>CGST (2B)</th>
                    <th style={{ textAlign: 'right' }}>SGST (2B)</th>
                    <th style={{ textAlign: 'right' }}>TOTAL (2B)</th>
                  </>
                ) : isSheet3 ? (
                  <>
                    <th style={{ textAlign: 'right' }}>TAXABLE (BOOKS)</th>
                    <th style={{ textAlign: 'right' }}>IGST (BOOKS)</th>
                    <th style={{ textAlign: 'right' }}>CGST (BOOKS)</th>
                    <th style={{ textAlign: 'right' }}>SGST (BOOKS)</th>
                    <th style={{ textAlign: 'right' }}>TOTAL (BOOKS)</th>
                  </>
                ) : (
                  <>
                    <th style={{ textAlign: 'right' }}>TAXABLE</th>
                    <th style={{ textAlign: 'right' }}>IGST</th>
                    <th style={{ textAlign: 'right' }}>CGST</th>
                    <th style={{ textAlign: 'right' }}>SGST</th>
                    {!isSheet1 && <th style={{ textAlign: 'right' }}>CESS</th>}
                    <th style={{ textAlign: 'right' }}>TOTAL</th>
                  </>
                )}

                <th style={{ minWidth: '140px', textAlign: 'center' }}>ACTION</th>
              </tr>
            </thead>

            <tbody>
              {displayedRecords.length > 0 ? (
                displayedRecords.map((r, index) => {
                  return (
                    <tr key={r.id}>
                      {/* S.No */}
                      <td style={{ textAlign: 'center', fontWeight: 600, color: '#64748b' }}>
                        {index + 1}
                      </td>

                      {/* Status */}
                      <td style={{ textAlign: 'left' }}>
                        {getStatusBadge(r.status)}
                      </td>

                      {/* 2B Month (Hidden on Sheet 3) */}
                      {!isSheet3 && (
                        <td style={{ textAlign: 'center', fontWeight: 600, color: '#334155' }}>
                          <span>{r.month2b || '-'}</span>
                          {r.hasDuplicateMonth && (
                            <span className="duplicate-warn-tag" title="Found in multiple GSTR-2B months">⚠️</span>
                          )}
                        </td>
                      )}

                      {/* Date */}
                      <td style={{ textAlign: 'center', color: '#475569' }}>
                        {r.date}
                      </td>

                      {/* Invoice No */}
                      <td style={{ textAlign: 'left', fontWeight: 700, color: '#350b5e' }}>
                        <span>{r.invNo}</span>
                        {r.auditStatus && r.auditStatus !== 'Pending' && (
                          <span className={`audit-status-tag ${r.auditStatus === 'Verified' ? 'verified' : 'query'}`}>
                            {r.auditStatus === 'Verified' ? '✓ Verified' : '? Query'}
                          </span>
                        )}
                      </td>

                      {/* Supplier GSTIN */}
                      <td style={{ textAlign: 'center', fontFamily: 'monospace', fontSize: '12px' }}>
                        {r.gstin}
                      </td>

                      {/* Supplier Name */}
                      <td style={{ textAlign: 'left', maxWidth: '240px', overflow: 'hidden', textOverflow: 'ellipsis' }} title={r.supplierName}>
                        {r.supplierName}
                      </td>

                      {/* Taxable */}
                      <td style={{ textAlign: 'right', fontWeight: 600 }}>
                        {formatIndianAmount(r.taxable)}
                      </td>

                      {/* IGST */}
                      <td style={{ textAlign: 'right' }}>
                        {formatIndianAmount(r.igst)}
                      </td>

                      {/* CGST */}
                      <td style={{ textAlign: 'right' }}>
                        {formatIndianAmount(r.cgst)}
                      </td>

                      {/* SGST */}
                      <td style={{ textAlign: 'right' }}>
                        {formatIndianAmount(r.sgst)}
                      </td>

                      {/* Cess (if not Sheet 1, 2, 3) */}
                      {!isSheet1 && !isSheet2 && !isSheet3 && (
                        <td style={{ textAlign: 'right' }}>
                          {formatIndianAmount(r.cess || 0)}
                        </td>
                      )}

                      {/* Total */}
                      <td style={{ textAlign: 'right', fontWeight: 700, color: '#350b5e' }}>
                        {formatIndianAmount(r.total)}
                      </td>

                      {/* Action: View, Audit */}
                      <td style={{ textAlign: 'center', whiteSpace: 'nowrap' }}>
                        <button
                          type="button"
                          className="btn-action-view"
                          onClick={() => setViewingRecord(r)}
                          title="View Invoice Details"
                        >
                          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                            <circle cx="12" cy="12" r="3"></circle>
                          </svg>
                          <span>View</span>
                        </button>
                        <button
                          type="button"
                          className="btn-action-audit"
                          onClick={() => {
                            setAuditingRecord(r);
                            setAuditFormStatus(r.auditStatus === 'Query' ? 'Query' : 'Verified');
                            setAuditFormNote(r.auditNote || '');
                          }}
                          title="Audit Note & Status"
                        >
                          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                          </svg>
                          <span>Audit</span>
                        </button>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={isSheet3 ? 12 : 13} style={{ padding: '36px', textAlign: 'center', color: '#64748b' }}>
                    No invoices found matching the current filters.
                  </td>
                </tr>
              )}
            </tbody>

            {/* Total Row (NIYAM: Teeno sheet ke neeche Total row do) */}
            <tfoot className="reco-table-footer">
              <tr style={{ background: '#F8FAFC', borderTop: '2px solid #CBD5E1', fontWeight: 700 }}>
                {/* Total label spans S.No through Supplier Name */}
                <td
                  colSpan={isSheet3 ? 6 : 7}
                  style={{ textAlign: 'left', padding: '12px 14px', fontSize: '13px', color: '#1e1b4b', fontWeight: 800 }}
                >
                  TOTAL ({displayedRecords.length} INVOICES)
                </td>

                {/* Taxable Total */}
                <td style={{ textAlign: 'right', padding: '12px 14px', color: '#1e1b4b', fontWeight: 800 }}>
                  {formatIndianAmount(totals.taxable)}
                </td>

                {/* IGST Total */}
                <td style={{ textAlign: 'right', padding: '12px 14px', color: '#1e1b4b', fontWeight: 800 }}>
                  {formatIndianAmount(totals.igst)}
                </td>

                {/* CGST Total */}
                <td style={{ textAlign: 'right', padding: '12px 14px', color: '#1e1b4b', fontWeight: 800 }}>
                  {formatIndianAmount(totals.cgst)}
                </td>

                {/* SGST Total */}
                <td style={{ textAlign: 'right', padding: '12px 14px', color: '#1e1b4b', fontWeight: 800 }}>
                  {formatIndianAmount(totals.sgst)}
                </td>

                {/* Cess Total (if not Sheet 1, 2, 3) */}
                {!isSheet1 && !isSheet2 && !isSheet3 && (
                  <td style={{ textAlign: 'right', padding: '12px 14px', color: '#1e1b4b', fontWeight: 800 }}>
                    {formatIndianAmount(totals.cess)}
                  </td>
                )}

                {/* Total Value Total */}
                <td style={{ textAlign: 'right', padding: '12px 14px', color: '#1e1b4b', fontWeight: 800 }}>
                  {formatIndianAmount(totals.total)}
                </td>

                {/* Empty cell under Action column */}
                <td></td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* ======================================================== */}
      {/* 4. View Invoice Details Modal Popup                      */}
      {/* ======================================================== */}
      {viewingRecord && (
        <div className="reco-modal-backdrop" onClick={() => setViewingRecord(null)}>
          <div className="reco-modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="reco-modal-header">
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <h3>Invoice Details: {viewingRecord.invNo}</h3>
                {getStatusBadge(viewingRecord.status)}
              </div>
              <button type="button" className="reco-modal-close" onClick={() => setViewingRecord(null)}>
                ✕
              </button>
            </div>

            <div className="reco-modal-body">
              <div className="reco-modal-grid">
                <div>
                  <label>Invoice Number</label>
                  <b>{viewingRecord.invNo}</b>
                </div>
                <div>
                  <label>Invoice Date</label>
                  <b>{viewingRecord.date}</b>
                </div>
                <div>
                  <label>2B Return Month</label>
                  <b>{viewingRecord.month2b || '-'}</b>
                </div>
                <div>
                  <label>Supplier GSTIN</label>
                  <b style={{ fontFamily: 'monospace' }}>{viewingRecord.gstin}</b>
                </div>
                <div style={{ gridColumn: 'span 2' }}>
                  <label>Supplier / Trade Name</label>
                  <b>{viewingRecord.supplierName}</b>
                </div>
                <div>
                  <label>Taxable Amount</label>
                  <b>₹{formatIndianAmount(viewingRecord.taxable)}</b>
                </div>
                <div>
                  <label>IGST Amount</label>
                  <b>₹{formatIndianAmount(viewingRecord.igst)}</b>
                </div>
                <div>
                  <label>CGST Amount</label>
                  <b>₹{formatIndianAmount(viewingRecord.cgst)}</b>
                </div>
                <div>
                  <label>SGST Amount</label>
                  <b>₹{formatIndianAmount(viewingRecord.sgst)}</b>
                </div>
                <div style={{ gridColumn: 'span 2' }}>
                  <label>Total Invoice Value</label>
                  <b style={{ fontSize: '17px', color: '#350b5e' }}>₹{formatIndianAmount(viewingRecord.total)}</b>
                </div>

                {viewingRecord.auditStatus && viewingRecord.auditStatus !== 'Pending' && (
                  <div style={{ gridColumn: 'span 2', padding: '10px 12px', background: '#F8FAFC', border: '1px solid #E2E8F0' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '4px' }}>
                      <span style={{ fontSize: '11px', fontWeight: 700, color: '#64748b' }}>AUDIT VERIFICATION</span>
                      <span className={`audit-status-tag ${viewingRecord.auditStatus === 'Verified' ? 'verified' : 'query'}`}>
                        {viewingRecord.auditStatus}
                      </span>
                    </div>
                    <div style={{ fontSize: '12.5px', color: '#1e293b' }}>
                      {viewingRecord.auditNote || 'No audit note provided.'}
                    </div>
                    {viewingRecord.auditUpdatedAt && (
                      <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '4px' }}>
                        Last updated: {viewingRecord.auditUpdatedAt}
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>

            <div className="reco-modal-footer">
              <button type="button" className="btn-reco-action" onClick={() => setViewingRecord(null)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* 5. Audit Note & Verification Modal Popup                 */}
      {/* ======================================================== */}
      {auditingRecord && (
        <div className="reco-modal-backdrop" onClick={() => setAuditingRecord(null)}>
          <div className="reco-modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="reco-modal-header">
              <h3>Audit Invoice: {auditingRecord.invNo}</h3>
              <button type="button" className="reco-modal-close" onClick={() => setAuditingRecord(null)}>
                ✕
              </button>
            </div>

            <form onSubmit={handleSaveAudit}>
              <div className="reco-modal-body">
                <div style={{ marginBottom: '14px', padding: '10px 12px', background: '#FAF8FD', border: '1px solid #EBE4F3' }}>
                  <div style={{ fontSize: '13px', fontWeight: 700, color: '#350b5e' }}>{auditingRecord.supplierName}</div>
                  <div style={{ fontSize: '12px', color: '#64748b', marginTop: '2px' }}>
                    GSTIN: {auditingRecord.gstin} | Total: ₹{formatIndianAmount(auditingRecord.total)}
                  </div>
                </div>

                <div className="reco-form-row">
                  <div>
                    <label>Audit Status</label>
                    <select
                      value={auditFormStatus}
                      onChange={(e) => setAuditFormStatus(e.target.value as AuditStatus)}
                      className="reco-form-input"
                    >
                      <option value="Verified">✓ Verified (Eligible ITC Confirmed)</option>
                      <option value="Query">? Query (Discrepancy / Supplier Follow-up)</option>
                      <option value="Pending">Pending Review</option>
                    </select>
                  </div>
                </div>

                <div className="reco-form-row">
                  <div style={{ flex: 1 }}>
                    <label>Audit Note / Observations</label>
                    <textarea
                      value={auditFormNote}
                      onChange={(e) => setAuditFormNote(e.target.value)}
                      placeholder="Enter verification notes, discrepancy remarks, or supplier queries..."
                      rows={3}
                      className="reco-form-input"
                      style={{ resize: 'vertical' }}
                    />
                  </div>
                </div>
              </div>

              <div className="reco-modal-footer">
                <button type="button" className="btn-reco-cancel" onClick={() => setAuditingRecord(null)}>
                  Cancel
                </button>
                <button type="submit" className="btn-reco-action">
                  Save Audit Note
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </section>
  );
}
