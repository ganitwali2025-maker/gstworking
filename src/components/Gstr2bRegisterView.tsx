import React, { useState, useEffect, useMemo } from 'react';
import { fmtVal, dmySlash, inr, pp } from '../utils/gstFormat';
import { updateQueryParams } from '../utils/router';
import { syncModuleToSheets } from '../services/sheetSync';
import Gstr2bImportModal from './Gstr2bImportModal';
import { HeadingSearchBar } from './HeadingTemplate';
import {
  REGISTER_TABS,
  STANDARD_SHEET_HEADERS,
  filterRegisterRows,
  loadRegisterRows,
  deleteRegisterRow,
  clearRegisterRows,
  mapRowToSheetRow
} from '../services/registerService';

export interface Gstr2bRecord {
  id: string;
  m: string;
  q: string;
  fy: string;
  d: string; // Date
  gstin: string; // GSTIN of Supplier
  tradeName: string; // Trade / Legal Name
  invNo: string; // Invoice No
  invVal: number; // Invoice Value (in paise)
  taxable: number; // Taxable Value (in paise)
  igst: number; // IGST (in paise)
  cgst: number; // CGST (in paise)
  sgst: number; // SGST (in paise)
  cess: number; // CESS (in paise)
  sheetSynced?: boolean;
}

const SEED_GSTR2B_RECORDS: Gstr2bRecord[] = [
  {
    id: 'g2b-1',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: 'FY 2026–27',
    d: '2026-08-02',
    gstin: '22AABCB4567M1Z2',
    tradeName: 'Bharat Steel & Alloys Pvt Ltd',
    invNo: 'BS-4421',
    invVal: 49560000,
    taxable: 42000000,
    igst: 0,
    cgst: 3780000,
    sgst: 3780000,
    cess: 0,
    sheetSynced: true
  },
  {
    id: 'g2b-2',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: 'FY 2026–27',
    d: '2026-08-05',
    gstin: '22AAFCK8812D1Z4',
    tradeName: 'Shree Krishna Logistics & Spares',
    invNo: 'SK-1092',
    invVal: 33040000,
    taxable: 28000000,
    igst: 0,
    cgst: 2520000,
    sgst: 2520000,
    cess: 0,
    sheetSynced: true
  },
  {
    id: 'g2b-3',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: 'FY 2026–27',
    d: '2026-08-08',
    gstin: '21AABCT1332N1Z8',
    tradeName: 'Tata Steel Downstream Products Ltd',
    invNo: 'TS-9901',
    invVal: 218300000,
    taxable: 185000000,
    igst: 33300000,
    cgst: 0,
    sgst: 0,
    cess: 0,
    sheetSynced: true
  }
];

export const GSTR2B_SHEET_NAME = REGISTER_TABS.GSTR2B;
export const GSTR2B_SHEET_HEADERS = STANDARD_SHEET_HEADERS;

export function mapGstr2bToSheetRow(r: Gstr2bRecord): { id: string; values: (string | number)[] } {
  return mapRowToSheetRow({
    ...r,
    g: r.gstin,
    p: r.tradeName,
    n: r.invNo,
    t: r.taxable,
    i: r.igst,
    c: r.cgst,
    s: r.sgst
  }, REGISTER_TABS.GSTR2B);
}

interface Gstr2bRegisterViewProps {
  showToast?: (msg: string) => void;
  fy?: string;
  setFy?: (fy: string) => void;
  month?: string;
  setMonth?: (month: string) => void;
}

export default function Gstr2bRegisterView({
  showToast,
  fy,
  setFy,
  month,
  setMonth
}: Gstr2bRegisterViewProps) {
  const [records, setRecords] = useState<Gstr2bRecord[]>(() => {
    try {
      const saved = localStorage.getItem('gstr2b_register_v1');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [isLocked, setIsLocked] = useState<boolean>(() => {
    try {
      return localStorage.getItem('gstr2b_locked_v1') === '1';
    } catch {
      return false;
    }
  });

  const [editingRecord, setEditingRecord] = useState<Gstr2bRecord | null>(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [confirmDialog, setConfirmDialog] = useState<{
    title: string;
    message: string;
    confirmLabel?: string;
    onConfirm: () => void;
  } | null>(null);

  const notify = (msg: string) => {
    if (showToast) showToast(msg);
  };

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

  const selectedFy = fy || 'ALL';
  const selectedMonth = month || 'ALL';

  const persistRecords = (recs: Gstr2bRecord[]) => {
    setRecords(recs);
    try {
      localStorage.setItem('gstr2b_register_v1', JSON.stringify(recs));
    } catch {}
  };

  const handleRefresh = async () => {
    try {
      const res = await loadRegisterRows<Gstr2bRecord>(REGISTER_TABS.GSTR2B);
      if (res.rows) {
        setRecords(res.rows);
      }
      if (res.error) {
        notify(res.error);
      } else {
        notify('GSTR-2B tab data reloaded successfully');
      }
    } catch {
      notify('GSTR-2B refreshed');
    }
  };

  // Step 5: Load data for GSTR-2B tab on mount & window focus
  useEffect(() => {
    handleRefresh();

    const onWindowFocus = () => {
      handleRefresh();
    };
    window.addEventListener('focus', onWindowFocus);

    const handleSynced = (e: any) => {
      if (e.detail?.tab === REGISTER_TABS.GSTR2B) {
        try {
          const saved = localStorage.getItem('gstr2b_register_v1');
          if (saved) setRecords(JSON.parse(saved));
        } catch {}
      }
    };
    window.addEventListener('register-synced', handleSynced);
    return () => {
      window.removeEventListener('focus', onWindowFocus);
      window.removeEventListener('register-synced', handleSynced);
    };
  }, []);

  // Step 4: Filter display records strictly by selected FY, Month and Search Query
  const displayRecords = useMemo(() => {
    return filterRegisterRows(records, selectedFy, selectedMonth, searchQuery);
  }, [records, selectedFy, selectedMonth, searchQuery]);

  // KPI calculations based on filtered records
  const totalTaxable = displayRecords.reduce((acc, r) => acc + (r.taxable || 0), 0);
  const totalIgst = displayRecords.reduce((acc, r) => acc + (r.igst || 0), 0);
  const totalCgst = displayRecords.reduce((acc, r) => acc + (r.cgst || 0), 0);
  const totalSgst = displayRecords.reduce((acc, r) => acc + (r.sgst || 0), 0);
  const totalCess = displayRecords.reduce((acc, r) => acc + (r.cess || 0), 0);
  const totalTax = totalIgst + totalCgst + totalSgst + totalCess;

  const unsyncedCount = records.filter((r) => !r.sheetSynced).length;

  // Actions
  const handleOpenEdit = (rec: Gstr2bRecord) => {
    if (isLocked) {
      notify('Sheet is locked. Unlock to make changes.');
      return;
    }
    setEditingRecord(rec);
    setModalOpen(true);
  };

  const handleSaveRecord = async (updated: Gstr2bRecord) => {
    let synced = false;
    try {
      const res = await syncModuleToSheets({
        action: 'upsert',
        sheet: GSTR2B_SHEET_NAME,
        headers: GSTR2B_SHEET_HEADERS,
        rows: [mapGstr2bToSheetRow(updated)]
      });
      if (res.success) synced = true;
      else notify('Data save ho gaya, par Google Sheet sync fail hua');
    } catch {
      notify('Data save ho gaya, par Google Sheet sync fail hua');
    }

    // Backend save
    fetch(`/api/registers/${REGISTER_TABS.GSTR2B}`, {
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
    notify('Record updated successfully');
  };

  const handleDelete = (id: string) => {
    if (isLocked) {
      notify('Sheet is locked.');
      return;
    }

    setConfirmDialog({
      title: 'Delete Record',
      message: 'Kya aap is GSTR-2B record ko delete karna chahte hain?',
      confirmLabel: 'Delete',
      onConfirm: async () => {
        // 1. Google Sheets delete first (Rule 3)
        const res = await deleteRegisterRow(REGISTER_TABS.GSTR2B, id);
        if (!res.success) {
          notify(res.error || 'Google Sheet delete failed. Record not deleted.');
          return;
        }

        // 2. Only on success, remove from app state & localStorage
        persistRecords(records.filter((r) => r.id !== id));
        notify('Record deleted');
      }
    });
  };

  const handleToggleLockClick = () => {
    const periodDesc = `${selectedMonth !== 'ALL' ? selectedMonth : 'All Months'} ${selectedFy !== 'ALL' ? selectedFy : ''}`.trim();

    if (!isLocked) {
      setConfirmDialog({
        title: 'Lock Sheet Confirmation',
        message: `Kya aap ${periodDesc} ke GSTR-2B records ko lock karna chahte hain? Lock hone par edit/delete nahi ho sakega.`,
        confirmLabel: 'Lock Sheet',
        onConfirm: async () => {
          setIsLocked(true);
          try { localStorage.setItem('gstr2b_locked_v1', '1'); } catch {}
          notify('Sheet locked successfully');

          const idsToLock = displayRecords.map((r) => r.id);
          try {
            await syncModuleToSheets({
              action: 'lock',
              sheet: GSTR2B_SHEET_NAME,
              headers: GSTR2B_SHEET_HEADERS,
              ids: idsToLock
            });
          } catch {}
        }
      });
    } else {
      setConfirmDialog({
        title: 'Unlock Sheet Confirmation',
        message: `Kya aap sheet ko unlock karna chahte hain?`,
        confirmLabel: 'Unlock Sheet',
        onConfirm: async () => {
          setIsLocked(false);
          try { localStorage.setItem('gstr2b_locked_v1', '0'); } catch {}
          notify('Sheet unlocked successfully');

          const idsToUnlock = displayRecords.map((r) => r.id);
          try {
            await syncModuleToSheets({
              action: 'unlock',
              sheet: GSTR2B_SHEET_NAME,
              headers: GSTR2B_SHEET_HEADERS,
              ids: idsToUnlock
            });
          } catch {}
        }
      });
    }
  };

  const handleClearDataClick = () => {
    if (isLocked) {
      notify('Locked month ka data clear nahi ho sakta.');
      return;
    }

    const periodDesc = `${selectedMonth !== 'ALL' ? selectedMonth : 'All Months'} ${selectedFy !== 'ALL' ? selectedFy : ''}`.trim();
    const toDeleteIds = displayRecords.map((r) => r.id);
    if (toDeleteIds.length === 0) {
      notify('No records to clear.');
      return;
    }

    setConfirmDialog({
      title: 'Clear Data Confirmation',
      message: `${periodDesc} ka saara data delete hoga. Pakka?`,
      confirmLabel: 'Delete All',
      onConfirm: async () => {
        // 1. Google Sheets batch delete first (Rule 3)
        const res = await clearRegisterRows(REGISTER_TABS.GSTR2B, toDeleteIds, selectedFy, selectedMonth);
        if (!res.success) {
          notify(res.error || 'Google Sheet clear failed. Records not deleted.');
          return;
        }

        // 2. Only on success, remove from app state & localStorage
        const idSet = new Set(toDeleteIds);
        const remaining = records.filter((r) => !idSet.has(r.id));
        persistRecords(remaining);
        notify(`${periodDesc} data cleared`);
      }
    });
  };

  const handleSyncToSheets = async () => {
    if (isSyncing) return;
    const unsynced = records.filter((r) => !r.sheetSynced);
    const targetRows = unsynced.length > 0 ? unsynced : records;
    if (targetRows.length === 0) {
      notify('No records to sync');
      return;
    }

    setIsSyncing(true);
    notify(`Syncing ${targetRows.length} records to Google Sheet...`);

    try {
      const rows = targetRows.map(mapGstr2bToSheetRow);
      const res = await syncModuleToSheets({
        action: 'upsert',
        sheet: GSTR2B_SHEET_NAME,
        headers: GSTR2B_SHEET_HEADERS,
        rows
      });

      if (res.success) {
        const syncedIds = new Set(targetRows.map((r) => r.id));
        const updated = records.map((r) => (syncedIds.has(r.id) ? { ...r, sheetSynced: true } : r));
        persistRecords(updated);
        notify(`Successfully synced ${res.count || targetRows.length} records to Google Sheet!`);
      } else {
        notify('Google Sheet sync fail hua: ' + (res.error || 'Unknown error'));
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Sync error';
      notify('Sync error: ' + msg);
    } finally {
      setIsSyncing(false);
    }
  };

  const handleBatchImport = async (
    newRecords: Gstr2bRecord[],
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
    const updatedRecords = [...newRecords, ...remaining];
    persistRecords(updatedRecords);

    // Automatically select imported month & FY in common heading
    if (replaceFy && setFy) {
      setFy(replaceFy);
    }
    if (replaceMonth && setMonth) {
      setMonth(replaceMonth);
    }

    if (counts) {
      notify(`${counts.imported} imported, ${counts.duplicate} duplicate skipped, ${counts.invalid} invalid skipped`);
    } else {
      notify(`${newRecords.length} GSTR-2B records imported`);
    }

    setImportModalOpen(false);

    // Background Google Sheets Sync & Backend Sync
    try {
      if (replacedIds.length > 0) {
        await syncModuleToSheets({
          action: 'delete',
          sheet: GSTR2B_SHEET_NAME,
          headers: GSTR2B_SHEET_HEADERS,
          ids: replacedIds
        });
      }

      const rows = newRecords.map(mapGstr2bToSheetRow);
      const res = await syncModuleToSheets({
        action: 'upsert',
        sheet: GSTR2B_SHEET_NAME,
        headers: GSTR2B_SHEET_HEADERS,
        rows
      });

      // Keep backend in sync
      fetch(`/api/registers/${REGISTER_TABS.GSTR2B}/actions/bulk-import`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          items: newRecords,
          fixedMonth: replaceMonth,
          fixedFy: replaceFy
        })
      }).catch(() => {});

      if (res.success) {
        const syncedIds = new Set(newRecords.map((r) => r.id));
        const finalRecords = updatedRecords.map((r) => (syncedIds.has(r.id) ? { ...r, sheetSynced: true } : r));
        persistRecords(finalRecords);
      } else {
        notify('Data save ho gaya, par Google Sheet sync fail hua');
      }
    } catch {
      notify('Data save ho gaya, par Google Sheet sync fail hua');
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
          <div className="gstr1-kpi-label">Total Taxable</div>
          <div className="gstr1-kpi-val">{inr(totalTaxable / 100)}</div>
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
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
                <line x1="16" y1="2" x2="16" y2="6"></line>
                <line x1="8" y1="2" x2="8" y2="6"></line>
                <line x1="3" y1="10" x2="21" y2="10"></line>
              </svg>
            </div>
            <div>
              <h2 className="gstr1-title">GSTR-2B Statement Register</h2>
              <div className="gstr1-subtitle">
                Showing {displayRecords.length} of {records.length} invoices
              </div>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="gstr1-btn-group">
            {/* Common Heading Template Search Bar */}
            <HeadingSearchBar
              value={searchQuery}
              onChange={setSearchQuery}
              placeholder="Search 2B (GSTIN, Supplier, Inv No)..."
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
                  const totalTax = (r.igst || 0) + (r.cgst || 0) + (r.sgst || 0) + (r.cess || 0);

                  return (
                    <tr key={r.id}>
                      <td style={{ textAlign: 'center' }}>{(r.fy && r.fy !== 'ALL') ? r.fy : 'FY 2026–27'}</td>
                      <td style={{ textAlign: 'center' }}>{r.q || '—'}</td>
                      <td style={{ textAlign: 'center' }}>{r.m || '—'}</td>
                      <td style={{ textAlign: 'center' }}>{dmySlash(r.d)}</td>
                      <td className="party-name" style={{ textAlign: 'left' }} title={r.tradeName}>
                        {(() => {
                          const val = (r.tradeName || '').trim();
                          const isGst = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/i.test(val);
                          if (isGst || (val && val.toUpperCase() === (r.gstin || '').trim().toUpperCase())) {
                            return '—';
                          }
                          return val && val.length > 34 ? val.slice(0, 34) + '...' : (val || '—');
                        })()}
                      </td>
                      <td className="gst-no" style={{ textAlign: 'center' }}>{r.gstin || '—'}</td>
                      <td className="inv-no" style={{ textAlign: 'left' }}>
                        {r.invNo}
                        {r.sheetSynced && (
                          <span style={{ color: '#10b981', marginLeft: '6px', fontSize: '11px' }} title="Synced to Google Sheet">
                            ✓
                          </span>
                        )}
                      </td>
                      <td className="num-val" style={{ fontWeight: 600, textAlign: 'center' }}>{fmtVal(r.invVal)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.taxable)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.igst)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cgst)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.sgst)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cess)}</td>
                      <td className="num-val total-tax" style={{ textAlign: 'center' }}>{fmtVal(totalTax)}</td>
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
                          title="Delete record"
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
                          : searchQuery.trim()
                          ? `No matching invoices found for "${searchQuery}".`
                          : `No invoices for ${selectedMonth !== 'ALL' ? selectedMonth : ''} ${selectedFy !== 'ALL' ? selectedFy : ''}.`}
                      </span>
                      {searchQuery.trim() && (
                        <button
                          type="button"
                          onClick={() => setSearchQuery('')}
                          style={{
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
                              localStorage.removeItem('gstr2b_selected_fy');
                              localStorage.removeItem('gstr2b_selected_month');
                            } catch {}
                            updateQueryParams({ fy: 'ALL', month: 'ALL' });
                          }}
                          style={{
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

      {/* Edit GSTR-2B Modal */}
      {modalOpen && editingRecord && (
        <Gstr2bEditModal
          isOpen={modalOpen}
          record={editingRecord}
          isLocked={isLocked}
          onClose={() => setModalOpen(false)}
          onSave={handleSaveRecord}
        />
      )}

      {/* Import Modal */}
      {importModalOpen && (
        <Gstr2bImportModal
          isOpen={importModalOpen}
          isLocked={isLocked}
          fy={selectedFy !== 'ALL' ? selectedFy : 'FY 2026–27'}
          existingRecords={records}
          onClose={() => setImportModalOpen(false)}
          onImport={handleBatchImport}
          showToast={notify}
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

interface Gstr2bEditModalProps {
  isOpen: boolean;
  record: Gstr2bRecord;
  isLocked: boolean;
  onClose: () => void;
  onSave: (record: Gstr2bRecord) => void;
}

function Gstr2bEditModal({ isOpen, record, isLocked, onClose, onSave }: Gstr2bEditModalProps) {
  if (!isOpen) return null;

  const [d, setD] = useState(record.d);
  const [gstin, setGstin] = useState(record.gstin);
  const [tradeName, setTradeName] = useState(record.tradeName);
  const [invNo, setInvNo] = useState(record.invNo);
  const [invVal, setInvVal] = useState((record.invVal / 100).toString());
  const [taxable, setTaxable] = useState((record.taxable / 100).toString());
  const [igst, setIgst] = useState((record.igst / 100).toString());
  const [cgst, setCgst] = useState((record.cgst / 100).toString());
  const [sgst, setSgst] = useState((record.sgst / 100).toString());
  const [cess, setCess] = useState((record.cess / 100).toString());
  const [error, setError] = useState('');

  const taxableVal = pp(taxable) || 0;
  const igstVal = pp(igst) || 0;
  const cgstVal = pp(cgst) || 0;
  const sgstVal = pp(sgst) || 0;
  const cessVal = pp(cess) || 0;
  const invValNum = pp(invVal) || (taxableVal + igstVal + cgstVal + sgstVal + cessVal);
  const totalTax = igstVal + cgstVal + sgstVal + cessVal;

  const handleSubmit = () => {
    if (isLocked) {
      setError('Sheet is locked.');
      return;
    }
    if (!d) return setError('Enter invoice date.');
    if (!invNo.trim()) return setError('Enter invoice number.');
    if (!gstin.trim()) return setError('Enter supplier GSTIN.');

    onSave({
      ...record,
      d,
      gstin: gstin.trim().toUpperCase(),
      tradeName: tradeName.trim(),
      invNo: invNo.trim(),
      invVal: invValNum,
      taxable: taxableVal,
      igst: igstVal,
      cgst: cgstVal,
      sgst: sgstVal,
      cess: cessVal
    });
  };

  return (
    <div className="mo">
      <div className="mb">
        <h3>Edit GSTR-2B Record</h3>
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
            <input value={invNo} onChange={(e) => setInvNo(e.target.value)} />
          </div>
          <div>
            <label>GSTIN of Supplier</label>
            <input maxLength={15} value={gstin} onChange={(e) => setGstin(e.target.value.toUpperCase())} />
          </div>
          <div>
            <label>Trade / Legal Name</label>
            <input value={tradeName} onChange={(e) => setTradeName(e.target.value)} />
          </div>
          <div>
            <label>Invoice Value (₹)</label>
            <input inputMode="decimal" value={invVal} onChange={(e) => setInvVal(e.target.value)} />
          </div>
          <div>
            <label>Taxable Value (₹)</label>
            <input inputMode="decimal" value={taxable} onChange={(e) => setTaxable(e.target.value)} />
          </div>
          <div>
            <label>IGST (₹)</label>
            <input inputMode="decimal" value={igst} onChange={(e) => setIgst(e.target.value)} />
          </div>
          <div>
            <label>CGST (₹)</label>
            <input inputMode="decimal" value={cgst} onChange={(e) => setCgst(e.target.value)} />
          </div>
          <div>
            <label>SGST (₹)</label>
            <input inputMode="decimal" value={sgst} onChange={(e) => setSgst(e.target.value)} />
          </div>
          <div>
            <label>CESS (₹)</label>
            <input inputMode="decimal" value={cess} onChange={(e) => setCess(e.target.value)} />
          </div>
        </div>

        <div style={{ marginTop: '14px', display: 'flex', gap: '24px', flexWrap: 'wrap' }}>
          <span>Total Tax: <b>₹{(totalTax / 100).toLocaleString('en-IN')}</b></span>
          <span>Invoice Value: <b>₹{(invValNum / 100).toLocaleString('en-IN')}</b></span>
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
