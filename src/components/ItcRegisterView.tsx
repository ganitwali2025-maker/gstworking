import React, { useState, useRef, useEffect, useMemo } from 'react';
import { fmtVal, dmySlash, inr, pp } from '../utils/gstFormat';
import { updateQueryParams } from '../utils/router';
import { syncModuleToSheets } from '../services/sheetSync';
import ItcImportModal from './ItcImportModal';
import {
  REGISTER_TABS,
  STANDARD_SHEET_HEADERS,
  filterRegisterRows,
  loadRegisterRows,
  mapRowToSheetRow
} from '../services/registerService';

export interface ItcPurchaseRecord {
  id: string;
  m: string;
  q: string;
  fy: string;
  d: string;
  p: string; // Name of Supplier
  g: string; // GST No.
  n: string; // Invoice No
  t: number; // BASIC AS PER BOOK (paise)
  i: number; // Integrated Tax
  c: number; // Central Tax
  s: number; // State Tax
  cess: number; // Cess
  itcInput?: string; // ITC INPUT ('Input', etc.)
  sheetSynced?: boolean;
}

const SEED_ITC_RECORDS: ItcPurchaseRecord[] = [
  {
    id: 'itc-1',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: 'FY 2026–27',
    d: '2026-08-02',
    p: 'Bharat Steel & Alloys Pvt Ltd',
    g: '22AABCB4567M1Z2',
    n: 'BS-4421',
    t: 42000000,
    i: 0,
    c: 3780000,
    s: 3780000,
    cess: 0,
    itcInput: 'Input',
    sheetSynced: true
  },
  {
    id: 'itc-2',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: 'FY 2026–27',
    d: '2026-08-05',
    p: 'Shree Krishna Logistics & Spares',
    g: '22AAFCK8812D1Z4',
    n: 'SK-1092',
    t: 28000000,
    i: 0,
    c: 2520000,
    s: 2520000,
    cess: 0,
    itcInput: 'Input',
    sheetSynced: true
  },
  {
    id: 'itc-3',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: 'FY 2026–27',
    d: '2026-08-08',
    p: 'Tata Steel Downstream Products Ltd',
    g: '21AABCT1332N1Z8',
    n: 'TS-9901',
    t: 185000000,
    i: 33300000,
    c: 0,
    s: 0,
    cess: 0,
    itcInput: 'Input',
    sheetSynced: true
  },
  {
    id: 'itc-4',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: 'FY 2026–27',
    d: '2026-08-14',
    p: 'Hindalco Industries Limited',
    g: '23AAACH1245G1Z5',
    n: 'HND-5519',
    t: 96000000,
    i: 17280000,
    c: 0,
    s: 0,
    cess: 0,
    itcInput: 'Input',
    sheetSynced: true
  }
];

export const PURCHASE_SHEET_NAME = REGISTER_TABS.ITC_PURCHASE;
export const PURCHASE_SHEET_HEADERS = STANDARD_SHEET_HEADERS;

export function mapItcToSheetRow(r: ItcPurchaseRecord): { id: string; values: (string | number)[] } {
  return mapRowToSheetRow(r, REGISTER_TABS.ITC_PURCHASE);
}

interface ItcRegisterViewProps {
  showToast: (msg: string) => void;
  fy?: string;
}

export default function ItcRegisterView({ showToast, fy: _fyProp }: ItcRegisterViewProps) {
  const [records, setRecords] = useState<ItcPurchaseRecord[]>(() => {
    try {
      const saved = localStorage.getItem('itc_register_v1');
      return saved ? JSON.parse(saved) : SEED_ITC_RECORDS;
    } catch {
      return SEED_ITC_RECORDS;
    }
  });

  const [isLocked, setIsLocked] = useState<boolean>(() => {
    try {
      return localStorage.getItem('itc_locked_v1') === '1';
    } catch {
      return false;
    }
  });

  const [editingRecord, setEditingRecord] = useState<ItcPurchaseRecord | null>(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [confirmDialog, setConfirmDialog] = useState<{
    title: string;
    message: string;
    confirmLabel?: string;
    onConfirm: () => void;
  } | null>(null);

  // Month & FY options
  const existingFys = Array.from(new Set(records.map((r) => r.fy).filter(Boolean)));
  const allFyOptions = Array.from(new Set(['FY 2026–27', 'FY 2025–26', ...existingFys]));

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

  // State: selectedFy & selectedMonth with URL / storage persistence
  const [selectedFy, setSelectedFy] = useState<string>(() => {
    if (typeof window !== 'undefined') {
      const q = new URLSearchParams(window.location.search).get('fy')?.trim();
      if (q) {
        if (q.toUpperCase() === 'ALL') return 'ALL';
        const matched = allFyOptions.find(
          (opt) => normalizeFy(opt) === normalizeFy(q) || opt.toLowerCase() === q.toLowerCase()
        );
        if (matched) return matched;
      }
      const saved = localStorage.getItem('purchase_selected_fy')?.trim();
      if (saved) {
        if (saved.toUpperCase() === 'ALL') return 'ALL';
        const matched = allFyOptions.find(
          (opt) => normalizeFy(opt) === normalizeFy(saved) || opt.toLowerCase() === saved.toLowerCase()
        );
        if (matched) return matched;
      }
    }
    return 'ALL';
  });

  const [selectedMonth, setSelectedMonth] = useState<string>(() => {
    if (typeof window !== 'undefined') {
      const q = new URLSearchParams(window.location.search).get('month')?.trim();
      if (q) {
        if (q.toUpperCase() === 'ALL') return 'ALL';
        const matched = MONTH_OPTIONS.find(
          (m) => m.value.toLowerCase() === q.toLowerCase() || m.label.toLowerCase() === q.toLowerCase()
        );
        if (matched) return matched.value;
      }
      const saved = localStorage.getItem('purchase_selected_month')?.trim();
      if (saved) {
        if (saved.toUpperCase() === 'ALL') return 'ALL';
        const matched = MONTH_OPTIONS.find(
          (m) => m.value.toLowerCase() === saved.toLowerCase() || m.label.toLowerCase() === saved.toLowerCase()
        );
        if (matched) return matched.value;
      }
    }
    return 'ALL';
  });

  // Sync selectedFy & selectedMonth with back/forward popstate
  useEffect(() => {
    const handlePopState = () => {
      const params = new URLSearchParams(window.location.search);
      const qFy = params.get('fy')?.trim();
      const qMonth = params.get('month')?.trim();
      if (qFy) {
        if (qFy.toUpperCase() === 'ALL') setSelectedFy('ALL');
        else {
          const matched = allFyOptions.find((opt) => normalizeFy(opt) === normalizeFy(qFy));
          if (matched) setSelectedFy(matched);
        }
      }
      if (qMonth) {
        if (qMonth.toUpperCase() === 'ALL') setSelectedMonth('ALL');
        else {
          const matched = MONTH_OPTIONS.find((m) => m.value.toLowerCase() === qMonth.toLowerCase());
          if (matched) setSelectedMonth(matched.value);
        }
      }
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, [allFyOptions]);

  const persistRecords = (recs: ItcPurchaseRecord[]) => {
    setRecords(recs);
    try {
      localStorage.setItem('itc_register_v1', JSON.stringify(recs));
    } catch { }
  };

  // Step 5: Load data for ITC-Purchase tab on initial mount
  useEffect(() => {
    loadRegisterRows<ItcPurchaseRecord>(REGISTER_TABS.ITC_PURCHASE, SEED_ITC_RECORDS).then((res) => {
      if (res.rows && res.rows.length > 0) {
        setRecords(res.rows);
      }
    });

    const handleSynced = (e: any) => {
      if (e.detail?.tab === REGISTER_TABS.ITC_PURCHASE) {
        try {
          const saved = localStorage.getItem('itc_register_v1');
          if (saved) setRecords(JSON.parse(saved));
        } catch {}
      }
    };
    window.addEventListener('register-synced', handleSynced);
    return () => window.removeEventListener('register-synced', handleSynced);
  }, []);

  const handleRefresh = async () => {
    try {
      const res = await loadRegisterRows<ItcPurchaseRecord>(REGISTER_TABS.ITC_PURCHASE, records);
      if (res.rows) {
        setRecords(res.rows);
      }
      showToast('ITC-Purchase tab data reloaded successfully');
    } catch {
      showToast('ITC-Purchase refreshed');
    }
  };

  // Step 4: Filter display records strictly by selected FY and Month (NEVER by invoice date!)
  const displayRecords = useMemo(() => {
    return filterRegisterRows(records, selectedFy, selectedMonth);
  }, [records, selectedFy, selectedMonth]);

  // KPI calculations based on filtered records
  const totalBasic = displayRecords.reduce((acc, r) => acc + (r.t || 0), 0);
  const totalIgst = displayRecords.reduce((acc, r) => acc + (r.i || 0), 0);
  const totalCgst = displayRecords.reduce((acc, r) => acc + (r.c || 0), 0);
  const totalSgst = displayRecords.reduce((acc, r) => acc + (r.s || 0), 0);
  const totalCess = displayRecords.reduce((acc, r) => acc + (r.cess || 0), 0);
  const totalTax = totalIgst + totalCgst + totalSgst + totalCess;

  const unsyncedCount = records.filter((r) => !r.sheetSynced).length;

  // Actions
  const handleOpenEdit = (rec: ItcPurchaseRecord) => {
    if (isLocked) {
      showToast('Sheet is locked. Unlock to make changes.');
      return;
    }
    setEditingRecord(rec);
    setModalOpen(true);
  };

  const handleSaveRecord = async (updated: ItcPurchaseRecord) => {
    let synced = false;
    try {
      const res = await syncModuleToSheets({
        action: 'upsert',
        sheet: PURCHASE_SHEET_NAME,
        headers: PURCHASE_SHEET_HEADERS,
        rows: [mapItcToSheetRow(updated)]
      });
      if (res.success) synced = true;
      else showToast('Data save ho gaya, par Google Sheet sync fail hua');
    } catch {
      showToast('Data save ho gaya, par Google Sheet sync fail hua');
    }

    // Backend save
    fetch(`/api/registers/${REGISTER_TABS.ITC_PURCHASE}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(updated)
    }).catch(() => {});

    const recordToSave = { ...updated, sheetSynced: synced };
    const exists = records.some((r) => r.id === updated.id);
    const newRecords = exists
      ? records.map((r) => (r.id === updated.id ? recordToSave : r))
      : [recordToSave, ...records];
    persistRecords(newRecords);
    setModalOpen(false);
    showToast('Record updated successfully');
  };

  const handleDelete = (id: string) => {
    if (isLocked) {
      showToast('Sheet is locked.');
      return;
    }

    setConfirmDialog({
      title: 'Delete Record',
      message: 'Kya aap is purchase record ko delete karna chahte hain?',
      confirmLabel: 'Delete',
      onConfirm: async () => {
        persistRecords(records.filter((r) => r.id !== id));
        showToast('Record deleted');

        // Backend delete
        fetch(`/api/registers/${REGISTER_TABS.ITC_PURCHASE}/${id}`, { method: 'DELETE' }).catch(() => {});

        try {
          const res = await syncModuleToSheets({
            action: 'delete',
            sheet: PURCHASE_SHEET_NAME,
            headers: PURCHASE_SHEET_HEADERS,
            ids: [id]
          });
          if (!res.success) {
            showToast('Data save ho gaya, par Google Sheet sync fail hua');
          }
        } catch {
          showToast('Data save ho gaya, par Google Sheet sync fail hua');
        }
      }
    });
  };

  const handleToggleLockClick = () => {
    const periodDesc = `${selectedMonth !== 'ALL' ? selectedMonth : 'All Months'} ${selectedFy !== 'ALL' ? selectedFy : ''}`.trim();

    if (!isLocked) {
      setConfirmDialog({
        title: 'Lock Sheet Confirmation',
        message: `Kya aap ${periodDesc} ke records ko lock karna chahte hain? Lock hone par edit/delete nahi ho sakega.`,
        confirmLabel: 'Lock Sheet',
        onConfirm: async () => {
          setIsLocked(true);
          try { localStorage.setItem('itc_locked_v1', '1'); } catch { }
          showToast('Sheet locked successfully');

          const idsToLock = displayRecords.map((r) => r.id);
          try {
            await syncModuleToSheets({
              action: 'lock',
              sheet: PURCHASE_SHEET_NAME,
              headers: PURCHASE_SHEET_HEADERS,
              ids: idsToLock
            });
          } catch { }
        }
      });
    } else {
      setConfirmDialog({
        title: 'Unlock Sheet Confirmation',
        message: `Kya aap sheet ko unlock karna chahte hain?`,
        confirmLabel: 'Unlock Sheet',
        onConfirm: async () => {
          setIsLocked(false);
          try { localStorage.setItem('itc_locked_v1', '0'); } catch { }
          showToast('Sheet unlocked successfully');

          const idsToUnlock = displayRecords.map((r) => r.id);
          try {
            await syncModuleToSheets({
              action: 'unlock',
              sheet: PURCHASE_SHEET_NAME,
              headers: PURCHASE_SHEET_HEADERS,
              ids: idsToUnlock
            });
          } catch { }
        }
      });
    }
  };

  const handleClearDataClick = () => {
    if (isLocked) {
      showToast('Locked month ka data clear nahi ho sakta.');
      return;
    }

    const periodDesc = `${selectedMonth !== 'ALL' ? selectedMonth : 'All Months'} ${selectedFy !== 'ALL' ? selectedFy : ''}`.trim();

    setConfirmDialog({
      title: 'Clear Data Confirmation',
      message: `${periodDesc} ka saara data delete hoga. Pakka?`,
      confirmLabel: 'Delete All',
      onConfirm: async () => {
        const toDeleteIds = displayRecords.map((r) => r.id);
        const remaining = records.filter((r) => !toDeleteIds.includes(r.id));
        persistRecords(remaining);
        showToast(`${periodDesc} data cleared`);

        // Backend clear
        try {
          const params = new URLSearchParams();
          if (selectedFy !== 'ALL') params.append('fy', selectedFy);
          if (selectedMonth !== 'ALL') params.append('m', selectedMonth);
          fetch(`/api/registers/${REGISTER_TABS.ITC_PURCHASE}/actions/clear?${params.toString()}`, { method: 'DELETE' }).catch(() => {});
        } catch {}

        try {
          const res = await syncModuleToSheets({
            action: 'delete',
            sheet: PURCHASE_SHEET_NAME,
            headers: PURCHASE_SHEET_HEADERS,
            ids: toDeleteIds
          });
          if (!res.success) {
            showToast('Data save ho gaya, par Google Sheet sync fail hua');
          }
        } catch {
          showToast('Data save ho gaya, par Google Sheet sync fail hua');
        }
      }
    });
  };

  const handleSyncToSheets = async () => {
    if (isSyncing) return;
    const unsynced = records.filter((r) => !r.sheetSynced);
    const targetRows = unsynced.length > 0 ? unsynced : records;
    if (targetRows.length === 0) {
      showToast('No records to sync');
      return;
    }

    setIsSyncing(true);
    showToast(`Syncing ${targetRows.length} records to Google Sheet...`);

    try {
      const rows = targetRows.map(mapItcToSheetRow);
      const res = await syncModuleToSheets({
        action: 'upsert',
        sheet: PURCHASE_SHEET_NAME,
        headers: PURCHASE_SHEET_HEADERS,
        rows
      });

      if (res.success) {
        const syncedIds = new Set(targetRows.map((r) => r.id));
        const updated = records.map((r) => (syncedIds.has(r.id) ? { ...r, sheetSynced: true } : r));
        persistRecords(updated);
        showToast(`Successfully synced ${res.count || targetRows.length} records to Google Sheet!`);
      } else {
        showToast('Google Sheet sync fail hua: ' + (res.error || 'Unknown error'));
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Sync error';
      showToast('Sync error: ' + msg);
    } finally {
      setIsSyncing(false);
    }
  };

  const handleBatchImport = async (
    newInvoices: ItcPurchaseRecord[],
    replaceMonth?: string,
    replaceFy?: string,
    counts?: { imported: number; duplicate: number; invalid: number }
  ) => {
    let remaining = records;
    let replacedIds: string[] = [];

    if (replaceMonth && replaceFy) {
      const toReplace = records.filter(
        (r) =>
          normalizeMonth(r.m) === normalizeMonth(replaceMonth) &&
          normalizeFy(r.fy) === normalizeFy(replaceFy)
      );
      replacedIds = toReplace.map((r) => r.id);
      remaining = records.filter((r) => !replacedIds.includes(r.id));
    }

    // Save locally first
    const updatedRecords = [...newInvoices, ...remaining];
    persistRecords(updatedRecords);

    // Automatically select imported month & FY in top bar
    if (replaceFy) {
      setSelectedFy(replaceFy);
      try { localStorage.setItem('purchase_selected_fy', replaceFy); } catch { }
      updateQueryParams({ fy: replaceFy });
    }
    if (replaceMonth) {
      setSelectedMonth(replaceMonth);
      try { localStorage.setItem('purchase_selected_month', replaceMonth); } catch { }
      updateQueryParams({ month: replaceMonth });
    }

    if (counts) {
      showToast(`${counts.imported} imported, ${counts.duplicate} duplicate skipped, ${counts.invalid} invalid skipped`);
    } else {
      showToast(`${newInvoices.length} purchase records imported`);
    }

    setImportModalOpen(false);

    // Background Google Sheets Sync & Backend Sync
    try {
      if (replacedIds.length > 0) {
        await syncModuleToSheets({
          action: 'delete',
          sheet: PURCHASE_SHEET_NAME,
          headers: PURCHASE_SHEET_HEADERS,
          ids: replacedIds
        });
      }

      const rows = newInvoices.map(mapItcToSheetRow);
      const res = await syncModuleToSheets({
        action: 'upsert',
        sheet: PURCHASE_SHEET_NAME,
        headers: PURCHASE_SHEET_HEADERS,
        rows
      });

      // Keep backend server in sync
      fetch(`/api/registers/${REGISTER_TABS.ITC_PURCHASE}/actions/bulk-import`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          items: newInvoices,
          fixedMonth: replaceMonth,
          fixedFy: replaceFy
        })
      }).catch(() => {});

      if (res.success) {
        const syncedIds = new Set(newInvoices.map((r) => r.id));
        const finalRecords = updatedRecords.map((r) => (syncedIds.has(r.id) ? { ...r, sheetSynced: true } : r));
        persistRecords(finalRecords);
      } else {
        showToast('Data save ho gaya, par Google Sheet sync fail hua');
      }
    } catch {
      showToast('Data save ho gaya, par Google Sheet sync fail hua');
    }
  };

  return (
    <section id="v-g1">
      {/* 5 KPI Cards Grid */}
      <div className="gstr1-kpis-grid">
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total Basic</div>
          <div className="gstr1-kpi-val">{inr(totalBasic / 100)}</div>
        </div>

        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="2" y1="12" x2="22" y2="12"></line>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total IGST</div>
          <div className="gstr1-kpi-val">{inr(totalIgst / 100)}</div>
        </div>

        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="3" y1="21" x2="21" y2="21"></line>
                <line x1="3" y1="10" x2="21" y2="10"></line>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total CGST</div>
          <div className="gstr1-kpi-val">{inr(totalCgst / 100)}</div>
        </div>

        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path>
                <circle cx="12" cy="10" r="3"></circle>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total SGST</div>
          <div className="gstr1-kpi-val">{inr(totalSgst / 100)}</div>
        </div>

        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <span style={{ fontSize: '16px', fontWeight: 700 }}>₹</span>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total Tax</div>
          <div className="gstr1-kpi-val">{inr(totalTax / 100)}</div>
        </div>
      </div>

      <div className="card gstr1-card">
        {/* Top Header Card */}
        <div className="gstr1-header">
          <div className="gstr1-title-box">
            <div className="gstr1-icon-badge">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
              </svg>
            </div>
            <div>
              <h2 className="gstr1-title">(ITC) Purchase Register</h2>
              <div className="gstr1-subtitle">
                Showing {displayRecords.length} of {records.length} invoices
              </div>
            </div>
          </div>

          {/* Action Buttons & Dropdowns */}
          <div className="gstr1-btn-group">
            {/* 1. FY Select */}
            <label className="gstr1-select-wrapper" title="Select Financial Year">
              <div className="gstr1-select-icon">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
                  <line x1="16" y1="2" x2="16" y2="6"></line>
                  <line x1="8" y1="2" x2="8" y2="6"></line>
                  <line x1="3" y1="10" x2="21" y2="10"></line>
                </svg>
              </div>
              <span className="gstr1-select-label">FY:</span>
              <select
                className="gstr1-select-field"
                value={selectedFy}
                onChange={(e) => {
                  const val = e.target.value;
                  setSelectedFy(val);
                  try { localStorage.setItem('purchase_selected_fy', val); } catch { }
                  updateQueryParams({ fy: val });
                }}
              >
                <option value="ALL">All FY</option>
                {allFyOptions.map((y) => (
                  <option key={y} value={y}>{y}</option>
                ))}
              </select>
            </label>

            {/* 2. Month Select */}
            <label className="gstr1-select-wrapper" title="Select Month">
              <div className="gstr1-select-icon">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="10"></circle>
                  <polyline points="12 6 12 12 16 14"></polyline>
                </svg>
              </div>
              <span className="gstr1-select-label">Month:</span>
              <select
                className="gstr1-select-field"
                value={selectedMonth}
                onChange={(e) => {
                  const val = e.target.value;
                  setSelectedMonth(val);
                  try { localStorage.setItem('purchase_selected_month', val); } catch { }
                  updateQueryParams({ month: val });
                }}
              >
                {MONTH_OPTIONS.map((m) => (
                  <option key={m.value} value={m.value}>{m.label}</option>
                ))}
              </select>
            </label>

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
              onClick={handleClearDataClick}
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
              onClick={handleRefresh}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"></path>
              </svg>
              Refresh
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-import"
              disabled={isLocked}
              onClick={() => setImportModalOpen(true)}
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

        {/* Table matching requested heading row */}
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
                <th style={{ textAlign: 'center' }}>INVOICE VALUE</th>
                <th style={{ textAlign: 'center' }}>TAXABLE VALUE</th>
                <th style={{ textAlign: 'center' }}>IGST</th>
                <th style={{ textAlign: 'center' }}>CGST</th>
                <th style={{ textAlign: 'center' }}>SGST</th>
                <th style={{ textAlign: 'center' }}>CESS</th>
                <th style={{ textAlign: 'center' }}>TOTAL TAX</th>
                <th style={{ textAlign: 'center' }}>ACTION</th>
              </tr>
            </thead>
            <tbody>
              {displayRecords.length > 0 ? (
                displayRecords.map((r) => {
                  const rowTax = (r.i || 0) + (r.c || 0) + (r.s || 0) + (r.cess || 0);
                  const rowTotalInv = (r.t || 0) + rowTax;

                  return (
                    <tr key={r.id}>
                      <td style={{ textAlign: 'center' }}>{r.fy || '—'}</td>
                      <td style={{ textAlign: 'center' }}>{r.q || '—'}</td>
                      <td style={{ textAlign: 'center' }}>{r.m || '—'}</td>
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
                        {r.n}
                        {r.sheetSynced && (
                          <span style={{ color: '#10b981', marginLeft: '6px', fontSize: '11px' }} title="Synced to Google Sheet">
                            ✓
                          </span>
                        )}
                      </td>
                      <td className="num-val" style={{ textAlign: 'center', fontWeight: 600, color: '#350B5E' }}>{fmtVal(rowTotalInv)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.t)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.i)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.c)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.s)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cess)}</td>
                      <td className="num-val total-tax" style={{ textAlign: 'center' }}>{fmtVal(rowTax)}</td>
                      <td className="act-col" style={{ textAlign: 'center', whiteSpace: 'nowrap' }}>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={() => handleOpenEdit(r)}
                          className="row-edit-btn"
                          title="Edit record"
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={() => handleDelete(r.id)}
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
                  <td colSpan={15} style={{ textAlign: 'center', padding: '36px', color: '#64748b' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px' }}>
                      <span>
                        {records.length === 0
                          ? 'No invoices in register. Click “Import” to load records.'
                          : `No invoices for ${selectedMonth !== 'ALL' ? selectedMonth : ''} ${selectedFy !== 'ALL' ? selectedFy : ''}.`}
                      </span>
                      {(selectedMonth !== 'ALL' || selectedFy !== 'ALL') && (
                        <button
                          type="button"
                          onClick={() => {
                            setSelectedMonth('ALL');
                            setSelectedFy('ALL');
                            try {
                              localStorage.removeItem('purchase_selected_fy');
                              localStorage.removeItem('purchase_selected_month');
                            } catch { }
                            updateQueryParams({ fy: 'ALL', month: 'ALL' });
                          }}
                          style={{
                            padding: '4px 12px',
                            background: '#f1f5f9',
                            border: '1px solid #cbd5e1',
                            borderRadius: '4px',
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

      {/* Edit ITC Modal */}
      {modalOpen && editingRecord && (
        <ItcEditModal
          isOpen={modalOpen}
          record={editingRecord}
          isLocked={isLocked}
          onClose={() => setModalOpen(false)}
          onSave={handleSaveRecord}
        />
      )}

      {/* Import Modal */}
      {importModalOpen && (
        <ItcImportModal
          isOpen={importModalOpen}
          isLocked={isLocked}
          fy={selectedFy !== 'ALL' ? selectedFy : 'FY 2026–27'}
          existingRecords={records}
          onClose={() => setImportModalOpen(false)}
          onImport={handleBatchImport}
          showToast={showToast}
        />
      )}

      {/* Confirmation Dialog */}
      {confirmDialog && (
        <div className="mo">
          <div className="mb" style={{ maxWidth: '420px' }}>
            <h3>{confirmDialog.title}</h3>
            <p style={{ margin: '12px 0 20px', color: '#475569', fontSize: '14px' }}>
              {confirmDialog.message}
            </p>
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
    </section>
  );
}

interface ItcEditModalProps {
  isOpen: boolean;
  record: ItcPurchaseRecord;
  isLocked: boolean;
  onClose: () => void;
  onSave: (record: ItcPurchaseRecord) => void;
}

function ItcEditModal({ isOpen, record, isLocked, onClose, onSave }: ItcEditModalProps) {
  if (!isOpen) return null;

  const [d, setD] = useState(record.d);
  const [p, setP] = useState(record.p);
  const [g, setG] = useState(record.g);
  const [n, setN] = useState(record.n);
  const [t, setT] = useState((record.t / 100).toString());
  const [i, setI] = useState((record.i / 100).toString());
  const [c, setC] = useState((record.c / 100).toString());
  const [s, setS] = useState((record.s / 100).toString());
  const [cess, setCess] = useState((record.cess / 100).toString());
  const [itcInput, setItcInput] = useState(record.itcInput || 'Input');
  const [error, setError] = useState('');

  const tVal = pp(t) || 0;
  const iVal = pp(i) || 0;
  const cVal = pp(c) || 0;
  const sVal = pp(s) || 0;
  const cessVal = pp(cess) || 0;
  const totalTax = iVal + cVal + sVal + cessVal;
  const totalInv = tVal + totalTax;

  const handleSubmit = () => {
    if (isLocked) {
      setError('Sheet is locked.');
      return;
    }
    if (!d) return setError('Enter invoice date.');
    if (!n.trim()) return setError('Enter invoice number.');
    if (!p.trim()) return setError('Enter supplier name.');

    onSave({
      ...record,
      d,
      p: p.trim(),
      g: (g || '').trim().toUpperCase(),
      n: n.trim(),
      t: tVal,
      i: iVal,
      c: cVal,
      s: sVal,
      cess: cessVal,
      itcInput
    });
  };

  return (
    <div className="mo">
      <div className="mb">
        <h3>Edit ITC Purchase Record</h3>
        <div style={{ color: 'var(--mute)', fontSize: '13px', marginBottom: '12px' }}>
          Period: {record.m} | {record.fy}
        </div>

        <div className="fg">
          <div>
            <label>Invoice Date</label>
            <input type="date" value={d} onChange={(e) => setD(e.target.value)} />
          </div>
          <div>
            <label>Invoice No</label>
            <input value={n} onChange={(e) => setN(e.target.value)} />
          </div>
          <div>
            <label>Name of Supplier</label>
            <input value={p} onChange={(e) => setP(e.target.value)} />
          </div>
          <div>
            <label>GSTIN of Supplier</label>
            <input maxLength={15} value={g} onChange={(e) => setG(e.target.value.toUpperCase())} />
          </div>
          <div>
            <label>TAXABLE VALUE (₹)</label>
            <input inputMode="decimal" value={t} onChange={(e) => setT(e.target.value)} />
          </div>
          <div>
            <label>IGST (₹)</label>
            <input inputMode="decimal" value={i} onChange={(e) => setI(e.target.value)} />
          </div>
          <div>
            <label>CGST (₹)</label>
            <input inputMode="decimal" value={c} onChange={(e) => setC(e.target.value)} />
          </div>
          <div>
            <label>SGST (₹)</label>
            <input inputMode="decimal" value={s} onChange={(e) => setS(e.target.value)} />
          </div>
          <div>
            <label>CESS (₹)</label>
            <input inputMode="decimal" value={cess} onChange={(e) => setCess(e.target.value)} />
          </div>
          <div>
            <label>ITC Input</label>
            <select value={itcInput} onChange={(e) => setItcInput(e.target.value)}>
              <option value="Input">Input</option>
              <option value="Capital Goods">Capital Goods</option>
              <option value="Input Service">Input Service</option>
              <option value="Ineligible">Ineligible</option>
            </select>
          </div>
        </div>

        <div style={{ marginTop: '14px', display: 'flex', gap: '24px', flexWrap: 'wrap' }}>
          <span>Total Tax: <b>₹{(totalTax / 100).toLocaleString('en-IN')}</b></span>
          <span>Total Invoice Value: <b>₹{(totalInv / 100).toLocaleString('en-IN')}</b></span>
        </div>

        {error && <div className="fe">{error}</div>}

        <div className="fr">
          <button type="button" onClick={onClose}>Cancel</button>
          <button type="button" className="pri" onClick={handleSubmit}>Save Changes</button>
        </div>
      </div>
    </div>
  );
}
