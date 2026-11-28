import fs from 'fs';

// GSTR-1 specific 14 columns matching user format
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

// ITC-Purchase specific 15 columns (Invoice Value moved to after Total Tax)
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

// Standard 15 columns required across registers
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

// RCM specific columns appended after standard columns
export const RCM_SHEET_HEADERS = [
  ...STANDARD_SHEET_HEADERS,
  'RCM Category',
  'RCM Tax Rate',
  'RCM Applicable',
  'Eligible ITC'
];

export const REGISTER_TABS = {
  GSTR1: 'GSTR-1',
  ITC_PURCHASE: 'ITC-Purchase',
  RCM: 'RCM',
  GSTR2B: 'GSTR-2B'
};

export function getHeadersForTab(tabName) {
  if (tabName === REGISTER_TABS.GSTR1) {
    return GSTR1_SHEET_HEADERS;
  }
  if (tabName === REGISTER_TABS.ITC_PURCHASE) {
    return ITC_PURCHASE_SHEET_HEADERS;
  }
  if (tabName === REGISTER_TABS.RCM) {
    return RCM_SHEET_HEADERS;
  }
  return STANDARD_SHEET_HEADERS;
}

// Helper: Ensure .env is loaded if present and not already loaded
function ensureEnvLoaded() {
  if (!process.env.APPS_SCRIPT_URL && fs.existsSync('.env')) {
    try {
      if (typeof process.loadEnvFile === 'function') {
        process.loadEnvFile('.env');
      } else {
        const envRaw = fs.readFileSync('.env', 'utf-8');
        envRaw.split('\n').forEach((line) => {
          const trimmed = line.trim();
          if (trimmed && !trimmed.startsWith('#')) {
            const eqIdx = trimmed.indexOf('=');
            if (eqIdx !== -1) {
              const k = trimmed.slice(0, eqIdx).trim();
              const v = trimmed.slice(eqIdx + 1).trim().replace(/^["']|["']$/g, '');
              if (!process.env[k]) process.env[k] = v;
            }
          }
        });
      }
    } catch (_) {}
  }
}

// Default Apps Script credentials (used automatically if .env is missing, e.g. on Vercel)
const DEFAULT_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby-lHw-wHVBXrS0SsuJ0iBeB96Ve_Zd_CGx3KvUJZQ7lvO32QSrkREr1V7xC_E9eK3X/exec';
const DEFAULT_APPS_SCRIPT_SECRET = 'MY_GST_SECRET_KEY_123';

function getConfig() {
  ensureEnvLoaded();
  const url = process.env.APPS_SCRIPT_URL?.trim() || DEFAULT_APPS_SCRIPT_URL;
  const secret = process.env.APPS_SCRIPT_SECRET?.trim() || DEFAULT_APPS_SCRIPT_SECRET;

  return { url, secret };
}

function formatDateSlash(dateStr) {
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
 * Universal mapper: Maps any register record to standard sheet row structure
 */
export function mapRecordToSheetRow(rec, tabName = 'GSTR-1') {
  // If record already has id and values array, pass as is
  if (rec && rec.id && Array.isArray(rec.values)) {
    return rec;
  }

  const isPaise = Boolean(rec.inPaise || (rec.t && rec.t > 100000) || (rec.taxable && rec.taxable > 100000));
  const divisor = isPaise ? 100 : 1;

  const tNum = (Number(rec.t ?? rec.taxable ?? rec.taxableValue ?? rec.amount) || 0) / divisor;
  const iNum = (Number(rec.i ?? rec.igst) || 0) / divisor;
  const cNum = (Number(rec.c ?? rec.cgst) || 0) / divisor;
  const sNum = (Number(rec.s ?? rec.sgst) || 0) / divisor;
  const cessNum = (Number(rec.cess) || 0) / divisor;
  const totalTax = iNum + cNum + sNum + cessNum;
  const invValRaw = rec.invVal ?? rec.totalInvoiceValue;
  const invVal = invValRaw !== undefined ? (Number(invValRaw) || 0) / divisor : (tNum + totalTax);

  const id = String(rec.id || ('rec-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)));
  const fy = String(rec.fy || rec.financialYear || 'FY 2026–27');
  const q = String(rec.q || rec.quarter || '');
  const m = String(rec.m || rec.month || '');
  const d = formatDateSlash(rec.d || rec.invoiceDate || '');
  const rawP = String(rec.p || rec.tradeName || rec.partyName || rec.supplierName || rec.transporter || '');
  const g = String(rec.g || rec.gstin || rec.gstNo || rec.supplierGstin || '').toUpperCase();
  const isGstRegex = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/i;
  const p = (isGstRegex.test(rawP.trim()) && (rawP.trim().toUpperCase() === g || !g)) ? '' : rawP;
  const n = String(rec.n || rec.invNo || rec.invoiceNo || rec.lrNo || '');

  if (tabName === REGISTER_TABS.GSTR1) {
    // Exact GSTR-1 sequence:
    // ID, FY, Quarter, Month, Invoice Date, Supplier / Party Name, GST No, Invoice No,
    // Taxable Value, IGST, CGST, SGST, Total Tax, Total Invoice Value
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

  if (tabName === REGISTER_TABS.ITC_PURCHASE) {
    // Exact Purchase sequence: Invoice Value moved to after Total Tax
    // ID, FY, Quarter, Month, Invoice Date, Supplier / Party Name, GST No, Invoice No,
    // Taxable Value, IGST, CGST, SGST, CESS, Total Tax, Total Invoice Value
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

  const standardValues = [
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

  if (tabName === REGISTER_TABS.RCM) {
    const category = rec.category || rec.expenseCategory || rec.head || 'Transport (GTA)';
    const rate = rec.rate || rec.taxRate || 5;
    const rcmApp = rec.rcmApplicable ?? 'Yes';
    const eligibleItc = (rec.eligibleItc !== undefined ? Number(rec.eligibleItc) / divisor : totalTax);
    return {
      id,
      values: [
        ...standardValues,
        category,
        rate,
        rcmApp,
        eligibleItc
      ]
    };
  }

  return {
    id,
    values: standardValues
  };
}

// Backwards-compatible alias for existing invoice routes
export function mapInvoiceToSheetRow(inv) {
  return mapRecordToSheetRow(inv, 'GSTR-1');
}

/**
 * Core send helper to POST payload to Google Apps Script Web App
 */
async function sendToAppsScript(payload) {
  const config = getConfig();
  if (!config) {
    return { ok: false, skipped: true, error: 'Google Sheets sync not configured' };
  }

  const { url, secret } = config;
  const bodyToSend = {
    secret,
    ...payload
  };

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bodyToSend),
      redirect: 'follow'
    });

    if (!res.ok) {
      return { ok: false, error: `HTTP ${res.status}` };
    }

    const json = await res.json().catch(() => null);
    if (!json) {
      return { ok: false, error: 'Invalid JSON response from Google Sheet Web App' };
    }

    return {
      ok: Boolean(json.ok),
      count: json.count || 0,
      shiftedCount: json.shiftedCount || 0,
      rows: json.rows || [],
      headers: json.headers || [],
      error: json.error || null
    };
  } catch (err) {
    // If local Windows machine or proxy causes TLS leaf cert issue, retry safely
    if (err && (err.cause?.code === 'UNABLE_TO_VERIFY_LEAF_SIGNATURE' || err.code === 'UNABLE_TO_VERIFY_LEAF_SIGNATURE')) {
      try {
        const retryRes = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(bodyToSend),
          redirect: 'follow'
        });
        const retryJson = await retryRes.json().catch(() => null);
        if (retryJson) {
          return {
            ok: Boolean(retryJson.ok),
            count: retryJson.count || 0,
            shiftedCount: retryJson.shiftedCount || 0,
            rows: retryJson.rows || [],
            headers: retryJson.headers || [],
            error: retryJson.error || null
          };
        }
      } catch (_) {}
    }

    const msg = err instanceof Error ? err.message : 'Network error';
    return {
      ok: false,
      error: msg
    };
  }
}

/**
 * Ensures all 4 tabs (GSTR-1, ITC-Purchase, RCM, GSTR-2B) exist with correct headers
 */
export async function ensureAllTabs() {
  const tabs = [
    { name: REGISTER_TABS.GSTR1, headers: GSTR1_SHEET_HEADERS },
    { name: REGISTER_TABS.ITC_PURCHASE, headers: ITC_PURCHASE_SHEET_HEADERS },
    { name: REGISTER_TABS.RCM, headers: RCM_SHEET_HEADERS },
    { name: REGISTER_TABS.GSTR2B, headers: STANDARD_SHEET_HEADERS }
  ];

  const results = {};
  for (const t of tabs) {
    const res = await sendToAppsScript({
      action: 'upsert',
      sheet: t.name,
      headers: t.headers,
      rows: []
    });
    results[t.name] = res.ok;
  }
  return results;
}

/**
 * Upsert generic rows in chunks of 200 to specified sheet tab
 */
export async function upsertGenericRows({ sheet, headers, rows }) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return { ok: true, count: 0 };
  }

  const targetSheet = sheet || REGISTER_TABS.GSTR1;
  const targetHeaders = headers || getHeadersForTab(targetSheet);
  const CHUNK_SIZE = 200;
  let totalCount = 0;
  let allOk = true;
  let lastError = null;

  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    const formattedRows = chunk.map((r) => mapRecordToSheetRow(r, targetSheet));

    const payload = {
      action: 'upsert',
      sheet: targetSheet,
      headers: targetHeaders,
      rows: formattedRows
    };

    const result = await sendToAppsScript(payload);

    if (result.skipped) {
      return result;
    }

    if (!result.ok) {
      allOk = false;
      lastError = result.error || 'Failed to upsert chunk to Google Sheet';
    } else {
      totalCount += result.count || chunk.length;
    }
  }

  return {
    ok: allOk,
    count: totalCount,
    error: lastError
  };
}

/**
 * Backwards-compatible upsertRows (defaults to GSTR-1 if no sheet specified)
 */
export async function upsertRows(invoices, options = {}) {
  const targetSheet = options.sheet || REGISTER_TABS.GSTR1;
  const targetHeaders = options.headers || getHeadersForTab(targetSheet);

  return await upsertGenericRows({
    sheet: targetSheet,
    headers: targetHeaders,
    rows: invoices
  });
}

/**
 * Delete rows by IDs from specified sheet tab
 */
export async function deleteRows(ids, options = {}) {
  if (!Array.isArray(ids) || ids.length === 0) {
    return { ok: true, count: 0 };
  }

  const targetSheet = options.sheet || REGISTER_TABS.GSTR1;
  const payload = {
    action: 'delete',
    sheet: targetSheet,
    ids
  };

  return await sendToAppsScript(payload);
}

/**
 * Lock sheet tab rows
 */
export async function lockSheet(ids, options = {}) {
  const targetSheet = options.sheet || REGISTER_TABS.GSTR1;
  const payload = {
    action: 'lock',
    sheet: targetSheet,
    ids: Array.isArray(ids) ? ids : []
  };

  return await sendToAppsScript(payload);
}

/**
 * Unlock sheet tab rows
 */
export async function unlockSheet(ids, options = {}) {
  const targetSheet = options.sheet || REGISTER_TABS.GSTR1;
  const payload = {
    action: 'unlock',
    sheet: targetSheet,
    ids: Array.isArray(ids) ? ids : []
  };

  return await sendToAppsScript(payload);
}

/**
 * Clear sheet tab rows below header
 */
export async function clearSheet(options = {}) {
  const targetSheet = options.sheet || REGISTER_TABS.GSTR1;
  const payload = {
    action: 'clear',
    sheet: targetSheet
  };

  return await sendToAppsScript(payload);
}

/**
 * Shift any g2b-... rows from GSTR-1 to GSTR-2B tab
 */
export async function shiftG2bFromGstr1(knownG2bRows = []) {
  // If knownG2bRows provided, upsert them to GSTR-2B and delete from GSTR-1
  if (Array.isArray(knownG2bRows) && knownG2bRows.length > 0) {
    const upsertRes = await upsertGenericRows({
      sheet: REGISTER_TABS.GSTR2B,
      rows: knownG2bRows
    });
    const deleteRes = await deleteRows(knownG2bRows.map(r => r.id), {
      sheet: REGISTER_TABS.GSTR1
    });
    return {
      ok: Boolean(upsertRes.ok && deleteRes.ok),
      shiftedCount: knownG2bRows.length
    };
  }

  // Also call shiftG2b action on Apps Script if supported
  const scriptRes = await sendToAppsScript({
    action: 'shiftG2b'
  });
  return scriptRes;
}

/**
 * Read all rows from specified sheet tab
 */
export async function readSheet(options = {}) {
  const targetSheet = options.sheet || REGISTER_TABS.GSTR1;
  const payload = {
    action: 'read',
    sheet: targetSheet
  };
  return await sendToAppsScript(payload);
}

/**
 * Universal mapper from Google Sheet row array to application record
 */
export function mapSheetRowToRecord(values, tabName = 'GSTR-1') {
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
    // DD/MM/YYYY -> YYYY-MM-DD
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

  if (tabName === REGISTER_TABS.GSTR1) {
    taxableRupees = Number(values[8]) || 0;
    igstRupees = Number(values[9]) || 0;
    cgstRupees = Number(values[10]) || 0;
    sgstRupees = Number(values[11]) || 0;
    const totalTaxRupees = Number(values[12]) || (igstRupees + cgstRupees + sgstRupees);
    invValRupees = Number(values[13]) || (taxableRupees + totalTaxRupees);
  } else if (tabName === REGISTER_TABS.ITC_PURCHASE) {
    const val8 = Number(values[8]) || 0;
    const val9 = Number(values[9]) || 0;
    const val10 = Number(values[10]) || 0;
    const val11 = Number(values[11]) || 0;
    const val12 = Number(values[12]) || 0;
    const val13 = Number(values[13]) || 0;

    if (Math.abs(val8) >= Math.abs(val9) && Math.abs(val8 - (val9 + val10 + val11 + val12)) < 5 && (val10 !== 0 || val11 !== 0 || val12 !== 0)) {
      // Standard format: [8] Invoice Value, [9] Taxable Value, [10] IGST, [11] CGST, [12] SGST, [13] Cess
      invValRupees = val8;
      taxableRupees = val9;
      igstRupees = val10;
      cgstRupees = val11;
      sgstRupees = val12;
      cessRupees = val13;
    } else {
      // Reordered format: [8] Taxable Value, [9] IGST, [10] CGST, [11] SGST, [12] Cess, [13] Total Tax, [14] Total Invoice Value
      taxableRupees = val8;
      igstRupees = val9;
      cgstRupees = val10;
      sgstRupees = val11;
      cessRupees = val12;
      const totalTaxRupees = Number(values[13]) || (igstRupees + cgstRupees + sgstRupees + cessRupees);
      invValRupees = Number(values[14]) || (taxableRupees + totalTaxRupees);
    }
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

  const baseRecord = {
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

  if (tabName === REGISTER_TABS.RCM) {
    baseRecord.rcmCategory = String(values[15] || 'Transport (GTA)');
    baseRecord.rcmTaxRate = Number(values[16]) || 5;
    baseRecord.rcmApplicable = String(values[17] || 'Yes');
    baseRecord.eligibleItc = values[18] !== undefined ? Math.round(Number(values[18]) * 100) : (igstPaise + cgstPaise + sgstPaise);
  }

  return baseRecord;
}

export default {
  STANDARD_SHEET_HEADERS,
  RCM_SHEET_HEADERS,
  REGISTER_TABS,
  getHeadersForTab,
  mapRecordToSheetRow,
  mapInvoiceToSheetRow,
  mapSheetRowToRecord,
  ensureAllTabs,
  readSheet,
  upsertRows,
  upsertGenericRows,
  deleteRows,
  lockSheet,
  unlockSheet,
  clearSheet,
  shiftG2bFromGstr1
};
