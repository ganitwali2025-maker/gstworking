/**
 * registerService.ts - Common logic for all GST registers
 * Implements Step 1, 2, 3, 4, 5, 6:
 * - Single source of truth for tab names and headers
 * - Month imported from file without touching invoice date
 * - Filter strictly on Month column
 * - Common load, save, delete, clear, import functions parameterized by tabName
 */

import { syncModuleToSheets, readModuleFromSheets } from './sheetSync';

export type RegisterTabName = 'GSTR-1' | 'ITC-Purchase' | 'RCM' | 'GSTR-2B';

export const REGISTER_TABS: Record<string, RegisterTabName> = {
  GSTR1: 'GSTR-1',
  ITC_PURCHASE: 'ITC-Purchase',
  RCM: 'RCM',
  GSTR2B: 'GSTR-2B'
};

export const GSTR1_SHEET_HEADERS = [
  'ID',
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

export const ITC_PURCHASE_SHEET_HEADERS = [
  'ID',
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
  'CESS',
  'Total Tax',
  'Total Invoice Value'
];

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
  if (tabName === 'GSTR-1') return GSTR1_SHEET_HEADERS;
  if (tabName === 'ITC-Purchase') return ITC_PURCHASE_SHEET_HEADERS;
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
  let fy = (fallbackFy && fallbackFy !== 'ALL' && fallbackFy.trim()) ? fallbackFy : 'FY 2026–27';
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
  selectedMonth: string,
  searchQuery?: string
): T[] {
  return (rows || []).filter((r: any) => {
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

    // 3. Search Query Filter across GSTIN, Party/Supplier/Transporter Name, Inv/Voucher/LR No, Date, Type/Category
    if (searchQuery && searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      const match =
        String(r.gstin || r.g || r.gstNo || '').toLowerCase().includes(q) ||
        String(r.party || r.tradeName || r.p || r.transporter || r.supplierName || '').toLowerCase().includes(q) ||
        String(r.invNo || r.n || r.lrNo || r.voucherNo || '').toLowerCase().includes(q) ||
        String(r.d || r.date || r.entryDate || '').toLowerCase().includes(q) ||
        String(r.category || r.type || r.invType || r.itcInput || '').toLowerCase().includes(q);

      if (!match) {
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

  if (tabName === 'GSTR-1') {
    // Exact GSTR-1 sequence requested by user and matching sheet headers:
    // ID, FY, Quarter, Month, Invoice Date, Supplier / Party Name, GST No, Invoice No,
    // Taxable Value (tNum), IGST (iNum), CGST (cNum), SGST (sNum), Total Tax (totalTax), Total Invoice Value (invVal)
    return {
      id,
      values: [
        id,
        fy,
        q,
        m,
        d,
        p,
        g,
        n,
        tNum,
        iNum,
        cNum,
        sNum,
        totalTax,
        invVal
      ]
    };
  }

  if (tabName === 'ITC-Purchase') {
    // Purchase sequence: Invoice Value moved after Total Tax to the end
    // ID, FY, Quarter, Month, Invoice Date, Supplier / Party Name, GST No, Invoice No,
    // Taxable Value (tNum), IGST (iNum), CGST (cNum), SGST (sNum), CESS (cessNum), Total Tax (totalTax), Total Invoice Value (invVal)
    return {
      id,
      values: [
        id,
        fy,
        q,
        m,
        d,
        p,
        g,
        n,
        tNum,
        iNum,
        cNum,
        sNum,
        cessNum,
        totalTax,
        invVal
      ]
    };
  }

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
 * Universal mapper from Google Sheet row array to application record
 */
export function mapSheetRowToAppRecord(values: any[], tabName: RegisterTabName): any {
  if (!Array.isArray(values) || values.length === 0) return null;
  const id = String(values[0] || '').trim();
  if (!id) return null;

  const rawFy = String(values[1] || '').trim();
  const fy = (rawFy && rawFy !== 'ALL') ? rawFy : 'FY 2026–27';
  const q = String(values[2] || '').trim();
  const m = String(values[3] || '').trim();
  const rawDate = String(values[4] || '').trim();
  let d = rawDate;
  if (rawDate.includes('T')) {
    d = rawDate.split('T')[0];
  } else if (rawDate.includes('/')) {
    const parts = rawDate.split('/');
    if (parts.length === 3) {
      if (parts[2].length === 4) {
        d = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
      } else {
        d = `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
      }
    }
  }

  const p = String(values[5] || '').trim();
  const g = String(values[6] || '').trim().toUpperCase();
  const isGstRegex = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/i;
  const cleanParty = (isGstRegex.test(p) && (p.toUpperCase() === g || !g)) ? '' : p;
  const n = String(values[7] || '').trim();

  let invValRupees = 0;
  let taxableRupees = 0;
  let igstRupees = 0;
  let cgstRupees = 0;
  let sgstRupees = 0;
  let cessRupees = 0;

  if (tabName === 'GSTR-1') {
    // Exact GSTR-1 sequence:
    // [8]=Taxable Value, [9]=IGST, [10]=CGST, [11]=SGST, [12]=Total Tax, [13]=Total Invoice Value
    taxableRupees = Number(values[8]) || 0;
    igstRupees = Number(values[9]) || 0;
    cgstRupees = Number(values[10]) || 0;
    sgstRupees = Number(values[11]) || 0;
    const totalTaxRupees = Number(values[12]) || (igstRupees + cgstRupees + sgstRupees);
    invValRupees = Number(values[13]) || (taxableRupees + totalTaxRupees);
  } else if (tabName === 'ITC-Purchase') {
    // Exact Purchase sequence: Invoice Value at the end after Total Tax
    // [8]=Taxable Value, [9]=IGST, [10]=CGST, [11]=SGST, [12]=CESS, [13]=Total Tax, [14]=Total Invoice Value
    taxableRupees = Number(values[8]) || 0;
    igstRupees = Number(values[9]) || 0;
    cgstRupees = Number(values[10]) || 0;
    sgstRupees = Number(values[11]) || 0;
    cessRupees = Number(values[12]) || 0;
    const totalTaxRupees = Number(values[13]) || (igstRupees + cgstRupees + sgstRupees + cessRupees);
    invValRupees = Number(values[14]) || (taxableRupees + totalTaxRupees);
  } else {
    invValRupees = Number(values[8]) || 0;
    taxableRupees = Number(values[9]) || 0;
    igstRupees = Number(values[10]) || 0;
    cgstRupees = Number(values[11]) || 0;
    sgstRupees = Number(values[12]) || 0;
    cessRupees = Number(values[13]) || 0;
  }

  const invValPaise = Math.round(invValRupees * 100);
  const taxablePaise = Math.round(taxableRupees * 100);
  const igstPaise = Math.round(igstRupees * 100);
  const cgstPaise = Math.round(cgstRupees * 100);
  const sgstPaise = Math.round(sgstRupees * 100);
  const cessPaise = Math.round(cessRupees * 100);

  const baseRecord: any = {
    id,
    fy,
    q,
    m,
    d,
    p: cleanParty,
    tradeName: cleanParty,
    partyName: cleanParty,
    supplierName: cleanParty,
    transporter: cleanParty,
    g,
    gstin: g,
    gstNo: g,
    n,
    invNo: n,
    lrNo: n,
    invVal: invValPaise,
    totalInvoiceValue: invValPaise,
    taxable: taxablePaise,
    taxableValue: taxablePaise,
    t: taxablePaise,
    amount: taxablePaise,
    igst: igstPaise,
    i: igstPaise,
    cgst: cgstPaise,
    c: cgstPaise,
    sgst: sgstPaise,
    s: sgstPaise,
    cess: cessPaise,
    sheetSynced: true
  };

  if (tabName === 'RCM') {
    baseRecord.rcmCategory = String(values[15] || 'Transport (GTA)');
    baseRecord.rcmTaxRate = Number(values[16]) || 5;
    baseRecord.rcmApplicable = String(values[17] || 'Yes');
    baseRecord.eligibleItc = values[18] !== undefined ? Math.round(Number(values[18]) * 100) : (igstPaise + cgstPaise + sgstPaise);
  }

  return baseRecord;
}

/**
 * Step 5: Common Load data from backend (or fallback Google Sheets / localStorage) for a specific tab
 */
export async function loadRegisterRows<T = any>(
  tabName: RegisterTabName
): Promise<{ rows: T[]; totalCount: number; locked: boolean; error?: string | null }> {
  const storageKey = getLocalStorageKey(tabName);
  let localRows: T[] = [];
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      localRows = JSON.parse(saved);
    }
  } catch {}

  // 1. Attempt backend API fetch (which fetches live from Google Sheets - Source of Truth)
  try {
    const res = await fetch(`/api/registers/${tabName}`);
    if (res.ok) {
      const json = await res.json();
      if (json.success && Array.isArray(json.data)) {
        // Authoritative live data from Google Sheet / server
        // Even if json.data is empty [], we update localStorage to []!
        try {
          localStorage.setItem(storageKey, JSON.stringify(json.data));
        } catch {}
        return {
          rows: json.data as T[],
          totalCount: json.data.length,
          locked: Boolean(json.locked)
        };
      }
    }
  } catch {}

  // 2. Direct fallback to Google Sheets Web App (crucial on Vercel)
  try {
    const sheetRes = await readModuleFromSheets(tabName);
    if (sheetRes.success && Array.isArray(sheetRes.rows)) {
      const mapped = sheetRes.rows.map((r: any) => mapSheetRowToAppRecord(r.values || r, tabName)).filter(Boolean);
      // Sheet read succeeded! If sheet is empty, mapped is [] and localStorage becomes []!
      try {
        localStorage.setItem(storageKey, JSON.stringify(mapped));
      } catch {}
      return {
        rows: mapped as T[],
        totalCount: mapped.length,
        locked: false
      };
    }
  } catch {}

  // 3. ONLY if both live sources failed due to network error, return offline local cache with error flag!
  return {
    rows: localRows,
    totalCount: localRows.length,
    locked: false,
    error: 'Google Sheet read failed. Showing offline cached data.'
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
): Promise<{ success: boolean; error?: string | null }> {
  // 1. Pehle Google Sheet se delete karo
  const sheetRes = await syncModuleToSheets({
    action: 'delete',
    sheet: tabName,
    headers: getHeadersForTab(tabName),
    ids: [id]
  });

  if (!sheetRes.success) {
    return {
      success: false,
      error: sheetRes.error || 'Google Sheet delete failed. Record not deleted.'
    };
  }

  // 2. Google Sheet success milne par hi backend se delete karo
  try {
    const res = await fetch(`/api/registers/${tabName}/${id}`, { method: 'DELETE' });
    if (!res.ok) {
      console.warn(`Backend delete failed for ${id} in ${tabName}`);
    }
  } catch (err) {
    console.warn(`Backend delete error for ${id} in ${tabName}:`, err);
  }

  // 3. localStorage se bhi hatao
  const storageKey = getLocalStorageKey(tabName);
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) {
        const remaining = parsed.filter((r: any) => r.id !== id);
        localStorage.setItem(storageKey, JSON.stringify(remaining));
      }
    }
  } catch {}

  return { success: true };
}

/**
 * Common Clear register rows for a period from Google Sheet tab
 */
export async function clearRegisterRows(
  tabName: RegisterTabName,
  toDeleteIds: string[],
  fy?: string,
  month?: string
): Promise<{ success: boolean; count: number; error?: string | null }> {
  if (!toDeleteIds || toDeleteIds.length === 0) {
    return { success: true, count: 0 };
  }

  // 1. Pehle Google Sheet se ek hi batch request me saare IDs delete karo
  const sheetRes = await syncModuleToSheets({
    action: 'delete',
    sheet: tabName,
    headers: getHeadersForTab(tabName),
    ids: toDeleteIds
  });

  if (!sheetRes.success) {
    return {
      success: false,
      count: 0,
      error: sheetRes.error || 'Google Sheet batch clear failed. Records not deleted.'
    };
  }

  // 2. Google Sheet success milne par hi backend se clear karo
  try {
    const params = new URLSearchParams();
    if (fy && fy !== 'ALL') params.append('fy', fy);
    if (month && month !== 'ALL') params.append('m', month);
    await fetch(`/api/registers/${tabName}/actions/clear?${params.toString()}`, { method: 'DELETE' });
  } catch (err) {
    console.warn(`Backend clear error for ${tabName}:`, err);
  }

  // 3. localStorage se bhi hatao
  const storageKey = getLocalStorageKey(tabName);
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) {
        const idSet = new Set(toDeleteIds);
        const remaining = parsed.filter((r: any) => !idSet.has(r.id));
        localStorage.setItem(storageKey, JSON.stringify(remaining));
      }
    }
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
