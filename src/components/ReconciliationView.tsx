import { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { fmtVal, inr } from '../utils/gstFormat';
import { MonthlyStats } from '../types/gst';
import { getQueryParams, updateQueryParams } from '../utils/router';

export interface RecoRecord {
  id: string;
  status: 'Not in 2B' | 'Invoice No. Not Match' | 'Matched' | 'Not in Book' | 'Debit Note';
  invNo: string;
  date: string;
  gstin: string;
  supplierName: string;
  taxable: number;
  igst: number;
  cgst: number;
  sgst: number;
  cess: number;
  total: number;
  tabCategory: 'books-reco' | 'books-itc' | 'govt-2b' | 'debit-note' | 'not-in-book' | 'not-in-2b' | 'matched' | 'not-matched';
}

const SEED_RECO_RECORDS: RecoRecord[] = [
  {
    id: 'rec-1',
    status: 'Not in 2B',
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

type TabKey =
  | 'books-reco'
  | 'books-itc'
  | 'govt-2b'
  | 'debit-note'
  | 'not-in-book'
  | 'not-in-2b'
  | 'matched'
  | 'not-matched';

interface TabItem {
  id: TabKey;
  label: string;
  icon: string;
}

const TABS: TabItem[] = [
  { id: 'books-reco', label: 'Books Reco with 2B', icon: 'layers' },
  { id: 'books-itc', label: 'Books ITC', icon: 'file' },
  { id: 'govt-2b', label: 'Government 2B', icon: 'download' },
  { id: 'debit-note', label: 'Debit Note', icon: 'file' },
  { id: 'not-in-book', label: 'Not in Book', icon: 'file-minus' },
  { id: 'not-in-2b', label: 'Not in 2B', icon: 'layers' },
  { id: 'matched', label: 'Match Invoice', icon: 'check' },
  { id: 'not-matched', label: 'Not Match Invoice', icon: 'layers' }
];

interface ReconciliationViewProps {
  showToast?: (msg: string) => void;
  fy?: string;
  month?: string;
  monthlyData?: Record<string, MonthlyStats>;
}

export default function ReconciliationView({
  showToast = () => { },
  fy: _fy = '2026-27'
}: ReconciliationViewProps) {
  const [activeTab, setActiveTab] = useState<TabKey>(() => {
    try {
      const q = getQueryParams();
      if (q.tab && TABS.some((t) => t.id === q.tab)) {
        return q.tab as TabKey;
      }
    } catch { }
    return 'books-reco';
  });

  useEffect(() => {
    const syncTab = () => {
      const q = getQueryParams();
      if (q.tab && TABS.some((t) => t.id === q.tab)) {
        setActiveTab(q.tab as TabKey);
      }
    };
    syncTab();
    window.addEventListener('popstate', syncTab);
    return () => window.removeEventListener('popstate', syncTab);
  }, []);
  const [searchQuery, setSearchQuery] = useState('');
  const [records, setRecords] = useState<RecoRecord[]>(() => {
    try {
      const saved = localStorage.getItem('gst_reco_records_v1');
      return saved ? JSON.parse(saved) : SEED_RECO_RECORDS;
    } catch {
      return SEED_RECO_RECORDS;
    }
  });

  const [selectedRecord, setSelectedRecord] = useState<RecoRecord | null>(null);
  const [isEditing, setIsEditing] = useState(false);
  const [viewingRecord, setViewingRecord] = useState<RecoRecord | null>(null);

  // Filter records by Active Tab and Search Query
  const displayedRecords = useMemo(() => {
    let list = records;

    // Filter by Tab
    if (activeTab === 'books-reco') {
      // Shows all records for complete reconciliation
      list = records;
    } else if (activeTab === 'books-itc') {
      list = records.filter((r) => ['not-in-2b', 'matched', 'not-matched', 'books-reco'].includes(r.tabCategory));
    } else if (activeTab === 'govt-2b') {
      list = records.filter((r) => ['not-in-book', 'matched', 'not-matched'].includes(r.tabCategory));
    } else if (activeTab === 'debit-note') {
      list = records.filter((r) => r.tabCategory === 'debit-note' || r.status === 'Debit Note');
    } else if (activeTab === 'not-in-book') {
      list = records.filter((r) => r.tabCategory === 'not-in-book' || r.status === 'Not in Book');
    } else if (activeTab === 'not-in-2b') {
      list = records.filter((r) => r.tabCategory === 'not-in-2b' || r.status === 'Not in 2B');
    } else if (activeTab === 'matched') {
      list = records.filter((r) => r.tabCategory === 'matched' || r.status === 'Matched');
    } else if (activeTab === 'not-matched') {
      list = records.filter((r) => r.tabCategory === 'not-matched' || r.status === 'Invoice No. Not Match');
    }

    // Filter by Search Query
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      list = list.filter(
        (r) =>
          r.invNo.toLowerCase().includes(q) ||
          r.supplierName.toLowerCase().includes(q) ||
          r.gstin.toLowerCase().includes(q)
      );
    }

    return list;
  }, [records, activeTab, searchQuery]);

  // Totals of displayed records
  const totalTaxable = displayedRecords.reduce((acc, r) => acc + (r.taxable || 0), 0);
  const totalIgst = displayedRecords.reduce((acc, r) => acc + (r.igst || 0), 0);
  const totalCgst = displayedRecords.reduce((acc, r) => acc + (r.cgst || 0), 0);
  const totalSgst = displayedRecords.reduce((acc, r) => acc + (r.sgst || 0), 0);
  const totalCess = displayedRecords.reduce((acc, r) => acc + (r.cess || 0), 0);
  const totalInvoiceVal = displayedRecords.reduce((acc, r) => acc + (r.total || 0), 0);

  // Tab count helper
  const getTabCount = (tabId: TabKey) => {
    if (tabId === 'books-reco') return records.length;
    if (tabId === 'books-itc') return records.filter((r) => ['not-in-2b', 'matched', 'not-matched'].includes(r.tabCategory)).length;
    if (tabId === 'govt-2b') return records.filter((r) => ['not-in-book', 'matched', 'not-matched'].includes(r.tabCategory)).length;
    if (tabId === 'debit-note') return records.filter((r) => r.tabCategory === 'debit-note' || r.status === 'Debit Note').length;
    if (tabId === 'not-in-book') return records.filter((r) => r.tabCategory === 'not-in-book' || r.status === 'Not in Book').length;
    if (tabId === 'not-in-2b') return records.filter((r) => r.tabCategory === 'not-in-2b' || r.status === 'Not in 2B').length;
    if (tabId === 'matched') return records.filter((r) => r.tabCategory === 'matched' || r.status === 'Matched').length;
    if (tabId === 'not-matched') return records.filter((r) => r.tabCategory === 'not-matched' || r.status === 'Invoice No. Not Match').length;
    return 0;
  };

  // Export to Excel
  const handleExportExcel = () => {
    try {
      const exportData = displayedRecords.map((r) => ({
        STATUS: r.status,
        'INVOICE NO.': r.invNo,
        DATE: r.date,
        'SUPPLIER GSTIN': r.gstin,
        'SUPPLIER NAME': r.supplierName,
        TAXABLE: r.taxable,
        IGST: r.igst,
        CGST: r.cgst,
        SGST: r.sgst,
        CESS: r.cess,
        TOTAL: r.total
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Reco_Report');
      XLSX.writeFile(wb, `GST_Reco_Manager_${activeTab}.xlsx`);
      showToast('Excel report downloaded successfully!');
    } catch (e) {
      console.error(e);
      showToast('Failed to export Excel.');
    }
  };

  // Print Report
  const handlePrint = () => {
    window.print();
  };

  // Save edited record
  const handleSaveEdit = (updated: RecoRecord) => {
    const updatedList = records.map((r) => (r.id === updated.id ? updated : r));
    setRecords(updatedList);
    try {
      localStorage.setItem('gst_reco_records_v1', JSON.stringify(updatedList));
    } catch { }
    setIsEditing(false);
    setSelectedRecord(null);
    showToast('Record updated successfully!');
  };

  return (
    <section id="v-rec" className="reco-page-container">
      <div className="card reco-card">
        {/* 1. Header Card Matching Image 1 */}
        <div className="reco-header-card">
        <div className="reco-header-left">
          <div className="reco-icon-badge">
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
            <h1 className="reco-title">
              GST RECO <span className="reco-title-sub">MANAGER</span>
            </h1>
            <p className="reco-subtitle">
              Automated Reconciliation of Purchase Books with GSTR-2B to Maximize Eligible ITC Claims
            </p>
          </div>
        </div>

        <div className="reco-header-right">
          {/* Search Input */}
          <div className="reco-search-wrap">
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
          <button type="button" className="btn-reco-action" onClick={handlePrint} title="Print Reconciliation Statement">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <polyline points="6 9 6 2 18 2 18 9"></polyline>
              <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path>
              <rect x="6" y="14" width="12" height="8"></rect>
            </svg>
            <span>Print</span>
          </button>

          {/* Excel Button */}
          <button type="button" className="btn-reco-action" onClick={handleExportExcel} title="Export to Excel">
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

      {/* 2. Reconciliation Tabs Bar Matching Image 2 */}
      <div className="reco-tabs-bar">
        {TABS.map((t) => {
          const isActive = activeTab === t.id;
          const count = getTabCount(t.id);
          return (
            <button
              key={t.id}
              type="button"
              className={`reco-tab ${isActive ? 'active' : ''}`}
              onClick={() => setActiveTab(t.id)}
            >
              {t.icon === 'layers' && (
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2">
                  <polygon points="12 2 2 7 12 12 22 7 12 2"></polygon>
                  <polyline points="2 17 12 22 22 17"></polyline>
                  <polyline points="2 12 12 17 22 12"></polyline>
                </svg>
              )}
              {t.icon === 'file' && (
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                  <polyline points="14 2 14 8 20 8"></polyline>
                </svg>
              )}
              {t.icon === 'download' && (
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                  <polyline points="7 10 12 15 17 10"></polyline>
                  <line x1="12" y1="15" x2="12" y2="3"></line>
                </svg>
              )}
              {t.icon === 'file-minus' && (
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                  <polyline points="14 2 14 8 20 8"></polyline>
                  <line x1="9" y1="15" x2="15" y2="15"></line>
                </svg>
              )}
              {t.icon === 'check' && (
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4">
                  <polyline points="20 6 9 17 4 12"></polyline>
                </svg>
              )}
              <span>{t.label}</span>
              <span className={`reco-tab-badge ${isActive ? 'active' : ''}`}>{count}</span>
            </button>
          );
        })}
      </div>

      {/* 3. Detailed Data Table Matching Image 1 */}
      <div className="reco-table-wrapper">
        <table className="reco-table">
          <thead>
            <tr>
              <th className="col-status" style={{ textAlign: 'left' }}>STATUS</th>
              <th className="col-inv" style={{ textAlign: 'left' }}>INVOICE NO.</th>
              <th style={{ textAlign: 'center' }}>DATE</th>
              <th style={{ textAlign: 'center' }}>SUPPLIER GSTIN</th>
              <th className="col-party" style={{ minWidth: '220px', textAlign: 'left' }}>SUPPLIER NAME</th>
              <th style={{ textAlign: 'center' }}>TAXABLE</th>
              <th style={{ textAlign: 'center' }}>IGST</th>
              <th style={{ textAlign: 'center' }}>CGST</th>
              <th style={{ textAlign: 'center' }}>SGST</th>
              <th style={{ textAlign: 'center' }}>CESS</th>
              <th style={{ textAlign: 'center' }}>TOTAL</th>
              <th style={{ textAlign: 'center' }}>ACTION</th>
            </tr>
          </thead>
          <tbody>
            {displayedRecords.length > 0 ? (
              displayedRecords.map((r) => {
                const isNotIn2b = r.status === 'Not in 2B';
                const isNoMatch = r.status === 'Invoice No. Not Match';
                const isMatched = r.status === 'Matched';
                const isNotInBook = r.status === 'Not in Book';

                return (
                  <tr key={r.id}>
                    {/* STATUS PILL */}
                    <td className="col-status" style={{ textAlign: 'left' }}>
                      <span
                        className={`reco-status-pill ${isNotIn2b
                            ? 'pill-not-in-2b'
                            : isNoMatch
                              ? 'pill-not-match'
                              : isMatched
                                ? 'pill-matched'
                                : isNotInBook
                                  ? 'pill-not-in-book'
                                  : 'pill-debit-note'
                          }`}
                      >
                        <span className="reco-status-dot">●</span> {r.status}
                      </span>
                    </td>

                    {/* INVOICE NO */}
                    <td className="reco-inv-no" style={{ textAlign: 'left' }}>{r.invNo}</td>

                    {/* DATE */}
                    <td style={{ textAlign: 'center' }}>{r.date}</td>

                    {/* SUPPLIER GSTIN */}
                    <td className="reco-gstin" style={{ textAlign: 'center' }}>{r.gstin}</td>

                    {/* SUPPLIER NAME */}
                    <td className="reco-party-name" style={{ textAlign: 'left' }} title={r.supplierName}>
                      {r.supplierName}
                    </td>

                    {/* AMOUNTS */}
                    <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.taxable)}</td>
                    <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.igst)}</td>
                    <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cgst)}</td>
                    <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.sgst)}</td>
                    <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cess)}</td>
                    <td className="num-val reco-total-val" style={{ textAlign: 'center' }}>{fmtVal(r.total)}</td>

                    {/* ACTIONS: EYE & PENCIL */}
                    <td className="reco-actions-col" style={{ textAlign: 'center' }}>
                      <button
                        type="button"
                        className="reco-act-btn reco-btn-view"
                        onClick={() => setViewingRecord(r)}
                        title="View invoice details"
                      >
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                          <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                          <circle cx="12" cy="12" r="3"></circle>
                        </svg>
                      </button>
                      <button
                        type="button"
                        className="reco-act-btn reco-btn-edit"
                        onClick={() => {
                          setSelectedRecord(r);
                          setIsEditing(true);
                        }}
                        title="Edit invoice"
                      >
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                          <path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"></path>
                        </svg>
                      </button>
                    </td>
                  </tr>
                );
              })
            ) : (
              <tr>
                <td colSpan={12} style={{ textAlign: 'center', padding: '42px', color: '#64748b' }}>
                  No invoices found under <b>{TABS.find((t) => t.id === activeTab)?.label}</b>.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      </div>

      {/* 4. View Details Modal */}
      {viewingRecord && (
        <div className="reco-modal-backdrop" onClick={() => setViewingRecord(null)}>
          <div className="reco-modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="reco-modal-header">
              <h3>Invoice Details – {viewingRecord.invNo}</h3>
              <button type="button" className="reco-modal-close" onClick={() => setViewingRecord(null)}>
                ✕
              </button>
            </div>
            <div className="reco-modal-body">
              <div className="reco-modal-grid">
                <div>
                  <label>Status</label>
                  <b>{viewingRecord.status}</b>
                </div>
                <div>
                  <label>Date</label>
                  <b>{viewingRecord.date}</b>
                </div>
                <div>
                  <label>Supplier GSTIN</label>
                  <b>{viewingRecord.gstin}</b>
                </div>
                <div>
                  <label>Supplier Name</label>
                  <b>{viewingRecord.supplierName}</b>
                </div>
                <div>
                  <label>Taxable Value</label>
                  <b>{inr(viewingRecord.taxable)}</b>
                </div>
                <div>
                  <label>Total IGST</label>
                  <b>{inr(viewingRecord.igst)}</b>
                </div>
                <div>
                  <label>CGST</label>
                  <b>{inr(viewingRecord.cgst)}</b>
                </div>
                <div>
                  <label>SGST</label>
                  <b>{inr(viewingRecord.sgst)}</b>
                </div>
                <div>
                  <label>Cess</label>
                  <b>{inr(viewingRecord.cess)}</b>
                </div>
                <div>
                  <label>Total Invoice Value</label>
                  <b style={{ color: '#350b5e', fontSize: '16px' }}>{inr(viewingRecord.total)}</b>
                </div>
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

      {/* 5. Edit Modal */}
      {isEditing && selectedRecord && (
        <div className="reco-modal-backdrop" onClick={() => setIsEditing(false)}>
          <div className="reco-modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="reco-modal-header">
              <h3>Edit Invoice – {selectedRecord.invNo}</h3>
              <button type="button" className="reco-modal-close" onClick={() => setIsEditing(false)}>
                ✕
              </button>
            </div>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                const fd = new FormData(e.currentTarget);
                const taxable = parseFloat(fd.get('taxable') as string) || 0;
                const igst = parseFloat(fd.get('igst') as string) || 0;
                const cgst = parseFloat(fd.get('cgst') as string) || 0;
                const sgst = parseFloat(fd.get('sgst') as string) || 0;
                const cess = parseFloat(fd.get('cess') as string) || 0;
                const total = taxable + igst + cgst + sgst + cess;

                handleSaveEdit({
                  ...selectedRecord,
                  status: (fd.get('status') as RecoRecord['status']) || selectedRecord.status,
                  invNo: (fd.get('invNo') as string) || selectedRecord.invNo,
                  date: (fd.get('date') as string) || selectedRecord.date,
                  gstin: (fd.get('gstin') as string) || selectedRecord.gstin,
                  supplierName: (fd.get('supplierName') as string) || selectedRecord.supplierName,
                  taxable,
                  igst,
                  cgst,
                  sgst,
                  cess,
                  total
                });
              }}
            >
              <div className="reco-modal-body">
                <div className="reco-form-row">
                  <div>
                    <label>Status</label>
                    <select name="status" defaultValue={selectedRecord.status} className="reco-form-input">
                      <option value="Not in 2B">Not in 2B</option>
                      <option value="Invoice No. Not Match">Invoice No. Not Match</option>
                      <option value="Matched">Matched</option>
                      <option value="Not in Book">Not in Book</option>
                      <option value="Debit Note">Debit Note</option>
                    </select>
                  </div>
                  <div>
                    <label>Invoice No.</label>
                    <input name="invNo" defaultValue={selectedRecord.invNo} className="reco-form-input" required />
                  </div>
                </div>

                <div className="reco-form-row">
                  <div>
                    <label>Date (DD/MM/YYYY)</label>
                    <input name="date" defaultValue={selectedRecord.date} className="reco-form-input" required />
                  </div>
                  <div>
                    <label>Supplier GSTIN</label>
                    <input name="gstin" defaultValue={selectedRecord.gstin} className="reco-form-input" required />
                  </div>
                </div>

                <div className="reco-form-row">
                  <div style={{ flex: 1 }}>
                    <label>Supplier Name</label>
                    <input name="supplierName" defaultValue={selectedRecord.supplierName} className="reco-form-input" required />
                  </div>
                </div>

                <div className="reco-form-row">
                  <div>
                    <label>Taxable (₹)</label>
                    <input name="taxable" type="number" step="any" defaultValue={selectedRecord.taxable} className="reco-form-input" required />
                  </div>
                  <div>
                    <label>IGST (₹)</label>
                    <input name="igst" type="number" step="any" defaultValue={selectedRecord.igst} className="reco-form-input" />
                  </div>
                </div>

                <div className="reco-form-row">
                  <div>
                    <label>CGST (₹)</label>
                    <input name="cgst" type="number" step="any" defaultValue={selectedRecord.cgst} className="reco-form-input" />
                  </div>
                  <div>
                    <label>SGST (₹)</label>
                    <input name="sgst" type="number" step="any" defaultValue={selectedRecord.sgst} className="reco-form-input" />
                  </div>
                  <div>
                    <label>Cess (₹)</label>
                    <input name="cess" type="number" step="any" defaultValue={selectedRecord.cess} className="reco-form-input" />
                  </div>
                </div>
              </div>

              <div className="reco-modal-footer">
                <button type="button" className="btn-reco-cancel" onClick={() => setIsEditing(false)}>
                  Cancel
                </button>
                <button type="submit" className="btn-reco-action">
                  Save Changes
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </section>
  );
}
