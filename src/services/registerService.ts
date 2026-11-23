/**
 * registerService.ts - Common logic for all GST registers
 * Implements Step 1, 2, 3, 4, 5, 6:
 * - Single source of truth for tab names and headers
 * - Month imported from file without touching invoice date
 * - Filter strictly on Month column
 * - Common load, save, delete, clear, import functions parameterized by tabName
 */

import { syncModuleToSheets } from './sheetSync';

export type RegisterTabName = 'GSTR-1' | 'ITC-Purchase' | 'RCM' | 'GSTR-2B';

export const REGISTER_TABS: Record<string, RegisterTabName> = {
  GSTR1: 'GSTR-1',
  ITC_PURCHASE: 'ITC-Purchase',
  RCM: 'RCM',
  GSTR2B: 'GSTR-2B'
};

export const STANDARD_SHEET_HEADERS = [
  'ID',
  'Financial Year',
  'Quarter',
  'Month',
  'Invoice Date',
  'Supplier / Party Name',
  'GST No',
  'Invoice No',
  'Invoice Value',
  'Taxable Value',
  'IGST',
  'CGST',
  'SGST',
  'CESS',
  'Total Tax'
];

export const RCM_SHEET_HEADERS = [
  ...STANDARD_SHEET_HEADERS,
  'RCM Category',
  'RCM Tax Rate',
  'RCM Applicable',
  'Eligible ITC'
];

export function getHeadersForTab(tabName: RegisterTabName): string[] {
  if (tabName === 'RCM') return RCM_SHEET_HEADERS;
  return STANDARD_SHEET_HEADERS;
}

export const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export const QUARTER_LABELS: Record<string, string> = {
  Q1: 'Apr-Jun',
  Q2: 'Jul-Sep',
  Q3: 'Oct-Dec',
  Q4: 'Jan-Mar'
};

export function normalizeFy(str?: string): string {
  if (!str) return '';
  const digits = String(str).replace(/[^0-9]/g, '');
  if (digits) return digits;
  return String(str).replace(/^FY\s*/i, '').replace(/[\s–-]/g, '').trim().toLowerCase();
}

export function normalizeMonth(str?: string): string {
  const clean = String(str || '').trim().toLowerCase();
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  for (const mon of months) {
    if (clean.startsWith(mon) || clean.includes(mon)) return mon;
  }
  return clean.slice(0, 3);
}

/**
 * Step 3: Derives FY, Quarter and Month strictly from Month string without looking at Invoice Date
 * Financial Year: April to March.
 * Quarter: Apr-Jun = Q1, Jul-Sep = Q2, Oct-Dec = Q3, Jan-Mar = Q4.
 */
export function derivePeriodFromMonth(
  monthStr: string,
  fallbackFy: string = 'FY 2026–27'
): { m: string; q: string; fy: string } {
  const clean = String(monthStr || '').trim();
  const monKey = normalizeMonth(clean);
  const mi = MONTH_NAMES.findIndex((x) => x.toLowerCase() === monKey);
  const validMi = mi !== -1 ? mi : 3; // default to Apr (3)
  const mLabel = mi !== -1 ? MONTH_NAMES[validMi] : clean || 'Apr';

  // Quarter: Apr-Jun = Q1, Jul-Sep = Q2, Oct-Dec = Q3, Jan-Mar = Q4
  const qCode = validMi >= 3 && validMi <= 5 ? 'Q1' : validMi >= 6 && validMi <= 8 ? 'Q2' : validMi >= 9 && validMi <= 11 ? 'Q3' : 'Q4';
  const qLabel = `${qCode} (${QUARTER_LABELS[qCode] || ''})`;

  // Financial Year: derived from year in month string if available (e.g. "Apr 2026")
  let fy = fallbackFy;
  const yrMatch = clean.match(/\b(20\d\d)\b/);
  if (yrMatch) {
    const yr = parseInt(yrMatch[1], 10);
    const startYr = validMi >= 3 ? yr : yr - 1;
    const endYrShort = String((startYr + 1) % 100).padStart(2, '0');
    fy = `FY ${startYr}–${endYrShort}`;
  }

  return {
    m: mLabel,
    q: qLabel,
    fy
  };
}

/**
 * Step 4: Filter strictly on Month column and FY. NEVER filter on Invoice Date!
 */
export function filterRegisterRows<T extends { m?: string; month?: string; fy?: string; financialYear?: string }>(
  rows: T[],
  selectedFy: string,
  selectedMonth: string
): T[] {
  return (rows || []).filter((r) => {
    // 1. FY Filter
    if (selectedFy && selectedFy !== 'ALL') {
      const rowFy = r.fy || r.financialYear || '';
      if (normalizeFy(rowFy) !== normalizeFy(selectedFy)) {
        return false;
      }
    }

    // 2. Month Filter: strictly on Month column
    if (selectedMonth && selectedMonth !== 'ALL') {
      const targetMon = normalizeMonth(selectedMonth);
      const rowMon = normalizeMonth(r.m || r.month || '');
      if (rowMon !== targetMon) {
        return false;
      }
    }

    return true;
  });
}

function formatDateSlash(dateStr?: string): string {
  if (!dateStr) return '';
  const clean = String(dateStr).trim();
  if (clean.includes('/')) return clean;
  if (clean.includes('-')) {
    const parts = clean.split('-');
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        // YYYY-MM-DD -> DD/MM/YYYY
        return `${parts[2].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[0]}`;
      }
      return `${parts[0].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[2]}`;
    }
  }
  return clean;
}

/**
 * Maps any row item into the exact 15/19 column values for Google Sheet
 */
export function mapRowToSheetRow(
  row: any,
  tabName: RegisterTabName
): { id: string; values: (string | number)[] } {
  const isPaise = Boolean(row.inPaise || (row.t && row.t > 100000) || (row.taxable && row.taxable > 100000));
  const divisor = isPaise ? 100 : 1;

  const tNum = (Number(row.t ?? row.taxable ?? row.taxableValue ?? row.amount) || 0) / divisor;
  const iNum = (Number(row.i ?? row.igst) || 0) / divisor;
  const cNum = (Number(row.c ?? row.cgst) || 0) / divisor;
  const sNum = (Number(row.s ?? row.sgst) || 0) / divisor;
  const cessNum = (Number(row.cess) || 0) / divisor;
  const totalTax = iNum + cNum + sNum + cessNum;
  const invValRaw = row.invVal ?? row.totalInvoiceValue;
  const invVal = invValRaw !== undefined ? (Number(invValRaw) || 0) / divisor : (tNum + totalTax);

  const id = String(row.id || '');
  const fy = String(row.fy || row.financialYear || '');
  const q = String(row.q || row.quarter || '');
  const m = String(row.m || row.month || '');
  const d = formatDateSlash(row.d || row.invoiceDate || '');
  const rawP = String(row.p || row.tradeName || row.partyName || row.supplierName || row.transporter || '');
  const g = String(row.g || row.gstin || row.gstNo || row.supplierGstin || '').toUpperCase();
  const isGstRegex = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/i;
  const p = (isGstRegex.test(rawP.trim()) && (rawP.trim().toUpperCase() === g || !g)) ? '' : rawP;
  const n = String(row.n || row.invNo || row.invoiceNo || row.lrNo || '');

  const standardValues: (string | number)[] = [
    id,
    fy,
    q,
    m,
    d,
    p,
    g,
    n,
    invVal,
    tNum,
    iNum,
    cNum,
    sNum,
    cessNum,
    totalTax
  ];

  if (tabName === 'RCM') {
    const category = row.rcmCategory || row.category || row.expenseCategory || row.head || 'Transport (GTA)';
    const rate = row.rcmTaxRate || row.rate || row.taxRate || 5;
    const rcmApp = row.rcmApplicable ?? 'Yes';
    const eligibleItc = row.eligibleItc !== undefined ? Number(row.eligibleItc) / divisor : totalTax;
    return {
      id,
      values: [...standardValues, category, rate, rcmApp, eligibleItc]
    };
  }

  return {
    id,
    values: standardValues
  };
}

/**
 * Storage key helper
 */
export function getLocalStorageKey(tabName: RegisterTabName): string {
  switch (tabName) {
    case 'GSTR-1':
      return 'gstr1_register_v4';
    case 'ITC-Purchase':
      return 'itc_register_v1';
    case 'RCM':
      return 'rcm_register_v1';
    case 'GSTR-2B':
      return 'gstr2b_register_v1';
  }
}

/**
 * Step 5: Common Load data from backend (or fallback localStorage) for a specific tab
 */
export async function loadRegisterRows<T = any>(
  tabName: RegisterTabName,
  fallbackRows: T[] = []
): Promise<{ rows: T[]; totalCount: number; locked: boolean }> {
  const storageKey = getLocalStorageKey(tabName);
  let localRows: T[] = [];
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      localRows = JSON.parse(saved);
    } else if (tabName === 'ITC-Purchase') {
      const legacy = localStorage.getItem('purchase_register_v1');
      if (legacy) {
        localRows = JSON.parse(legacy);
        try { localStorage.setItem(storageKey, legacy); } catch {}
      }
    }
  } catch {}

  if (localRows.length === 0 && fallbackRows.length > 0) {
    localRows = fallbackRows;
  }

  // Step 1: Ensure no g2b-... rows remain in GSTR-1
  if (tabName === 'GSTR-1') {
    const mixed = localRows.filter((r: any) => String(r.id || '').toLowerCase().startsWith('g2b'));
    if (mixed.length > 0) {
      localRows = localRows.filter((r: any) => !String(r.id || '').toLowerCase().startsWith('g2b'));
      try {
        localStorage.setItem(storageKey, JSON.stringify(localRows));
      } catch {}

      // Add to GSTR-2B local store if not present
      try {
        const key2b = getLocalStorageKey('GSTR-2B');
        const existing2b = JSON.parse(localStorage.getItem(key2b) || '[]');
        const existing2bIds = new Set(existing2b.map((x: any) => x.id));
        const toAdd = mixed.filter((x: any) => !existing2bIds.has(x.id));
        if (toAdd.length > 0) {
          localStorage.setItem(key2b, JSON.stringify([...toAdd, ...existing2b]));
        }
      } catch {}
    }
  }

  // Attempt backend API fetch
  try {
    const res = await fetch(`/api/registers/${tabName}`);
    if (res.ok) {
      const json = await res.json();
      if (json.success && Array.isArray(json.data)) {
        // Smart merge: preserve any local rows not yet in backend data so unsynced data is never lost
        const backendIds = new Set(json.data.map((r: any) => r.id));
        const unsavedLocal = localRows.filter((r: any) => !backendIds.has(r.id));
        const merged = [...json.data, ...unsavedLocal];
        try {
          localStorage.setItem(storageKey, JSON.stringify(merged));
        } catch {}
        return {
          rows: merged as T[],
          totalCount: merged.length,
          locked: Boolean(json.locked)
        };
      }
    }
  } catch {}

  return {
    rows: localRows,
    totalCount: localRows.length,
    locked: false
  };
}

/**
 * Common Save single row to backend and Google Sheet tab
 */
export async function saveRegisterRow(
  tabName: RegisterTabName,
  row: any
): Promise<{ success: boolean; sheetSynced?: boolean; sheetWarning?: string | null }> {
  const mapped = mapRowToSheetRow(row, tabName);

  // Sync with Google Sheet tab
  let sheetSynced = false;
  let sheetWarning: string | null = null;
  try {
    const sheetRes = await syncModuleToSheets({
      action: 'upsert',
      sheet: tabName,
      headers: getHeadersForTab(tabName),
      rows: [mapped]
    });
    if (sheetRes.success) {
      sheetSynced = true;
    } else {
      sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
    }
  } catch {
    sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
  }

  // Also call backend API if available
  try {
    await fetch(`/api/registers/${tabName}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(row)
    });
  } catch {}

  return { success: true, sheetSynced, sheetWarning };
}

/**
 * Common Bulk Import into register tab and Google Sheet tab
 */
export async function bulkImportRegisterRows(
  tabName: RegisterTabName,
  newRows: any[],
  importMonth?: string,
  importFy?: string
): Promise<{ success: boolean; sheetSynced?: boolean; sheetWarning?: string | null; count: number }> {
  if (!Array.isArray(newRows) || newRows.length === 0) {
    return { success: true, count: 0 };
  }

  const mappedRows = newRows.map((r) => mapRowToSheetRow(r, tabName));

  // Sync to Google Sheet tab
  let sheetSynced = false;
  let sheetWarning: string | null = null;
  try {
    const sheetRes = await syncModuleToSheets({
      action: 'upsert',
      sheet: tabName,
      headers: getHeadersForTab(tabName),
      rows: mappedRows
    });
    if (sheetRes.success) {
      sheetSynced = true;
    } else {
      sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
    }
  } catch {
    sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
  }

  // Sync to backend register
  try {
    await fetch(`/api/registers/${tabName}/actions/bulk-import`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        items: newRows,
        fixedMonth: importMonth,
        fixedFy: importFy
      })
    });
  } catch {}

  return { success: true, sheetSynced, sheetWarning, count: newRows.length };
}

/**
 * Common Delete row from register tab and Google Sheet tab
 */
export async function deleteRegisterRow(
  tabName: RegisterTabName,
  id: string
): Promise<{ success: boolean; sheetWarning?: string | null }> {
  let sheetWarning: string | null = null;
  try {
    const sheetRes = await syncModuleToSheets({
      action: 'delete',
      sheet: tabName,
      headers: getHeadersForTab(tabName),
      ids: [id]
    });
    if (!sheetRes.success) {
      sheetWarning = 'Google Sheet me delete sync fail hua.';
    }
  } catch {
    sheetWarning = 'Google Sheet me delete sync fail hua.';
  }

  try {
    await fetch(`/api/registers/${tabName}/${id}`, { method: 'DELETE' });
  } catch {}

  return { success: true, sheetWarning };
}

/**
 * Common Clear register rows for a period from Google Sheet tab
 */
export async function clearRegisterRows(
  tabName: RegisterTabName,
  toDeleteIds: string[],
  fy?: string,
  month?: string
): Promise<{ success: boolean; count: number }> {
  if (toDeleteIds.length > 0) {
    try {
      await syncModuleToSheets({
        action: 'delete',
        sheet: tabName,
        headers: getHeadersForTab(tabName),
        ids: toDeleteIds
      });
    } catch {}
  }

  try {
    const params = new URLSearchParams();
    if (fy && fy !== 'ALL') params.append('fy', fy);
    if (month && month !== 'ALL') params.append('m', month);
    await fetch(`/api/registers/${tabName}/actions/clear?${params.toString()}`, { method: 'DELETE' });
  } catch {}

  return { success: true, count: toDeleteIds.length };
}

/**
 * Step 1: Shifts any g2b-... rows from GSTR-1 to GSTR-2B
 */
export async function shiftG2bRowsFromGstr1(): Promise<{ shiftedCount: number }> {
  try {
    const res = await fetch('/api/sheets/shift-g2b', { method: 'POST' });
    if (res.ok) {
      const json = await res.json();
      return { shiftedCount: json.shiftedCount || 0 };
    }
  } catch {}
  return { shiftedCount: 0 };
}

/**
 * Auto-retries syncing any unsynced records across all 4 modules (GSTR-1, ITC-Purchase, RCM, GSTR-2B)
 * Called when internet connection is restored or periodically in background.
 */
export async function autoRetrySyncAcrossRegisters(): Promise<{ totalSynced: number }> {
  const tabs: RegisterTabName[] = ['GSTR-1', 'ITC-Purchase', 'RCM', 'GSTR-2B'];
  let totalSynced = 0;

  for (const tab of tabs) {
    const storageKey = getLocalStorageKey(tab);
    try {
      const raw = localStorage.getItem(storageKey);
      if (!raw) continue;
      const records = JSON.parse(raw);
      if (!Array.isArray(records) || records.length === 0) continue;

      const unsynced = records.filter((r: any) => r && r.sheetSynced === false);
      if (unsynced.length === 0) continue;

      const mappedRows = unsynced.map((r) => mapRowToSheetRow(r, tab));
      const res = await syncModuleToSheets({
        action: 'upsert',
        sheet: tab,
        headers: getHeadersForTab(tab),
        rows: mappedRows
      });

      if (res.success) {
        totalSynced += unsynced.length;
        const unsyncedIds = new Set(unsynced.map((r: any) => r.id));
        const updatedRecords = records.map((r: any) => (unsyncedIds.has(r.id) ? { ...r, sheetSynced: true } : r));
        localStorage.setItem(storageKey, JSON.stringify(updatedRecords));

        if (typeof window !== 'undefined') {
          window.dispatchEvent(new CustomEvent('register-synced', { detail: { tab, count: unsynced.length } }));
        }
      }
    } catch {}
  }

  return { totalSynced };
}

let isListenerAttached = false;
export function initAutoSyncRetryListener() {
  if (typeof window === 'undefined' || isListenerAttached) return;
  isListenerAttached = true;

  // Retry when browser comes online
  window.addEventListener('online', () => {
    autoRetrySyncAcrossRegisters();
  });

  // Retry periodically every 45 seconds if tab is active
  setInterval(() => {
    if (typeof navigator !== 'undefined' && navigator.onLine) {
      autoRetrySyncAcrossRegisters();
    }
  }, 45000);
}
