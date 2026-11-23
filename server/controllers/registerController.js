import { getDb, saveDb } from '../data/store.js';
import googleSheets, { REGISTER_TABS } from '../services/googleSheets.js';

const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const QM_LABEL = {
  Q1: 'Apr-Jun',
  Q2: 'Jul-Sep',
  Q3: 'Oct-Dec',
  Q4: 'Jan-Mar'
};

export const normalizeTabName = (tab) => {
  if (!tab) return REGISTER_TABS.GSTR1;
  const clean = String(tab).trim().toLowerCase();
  if (clean === 'gstr-1' || clean === 'gstr1' || clean === 'sales') return REGISTER_TABS.GSTR1;
  if (clean === 'itc-purchase' || clean === 'purchase' || clean === 'itc') return REGISTER_TABS.ITC_PURCHASE;
  if (clean === 'rcm') return REGISTER_TABS.RCM;
  if (clean === 'gstr-2b' || clean === 'gstr2b' || clean === '2b') return REGISTER_TABS.GSTR2B;
  return tab;
};

export const normalizeFy = (str) => {
  if (!str) return '';
  const digits = String(str).replace(/[^0-9]/g, '');
  if (digits) return digits;
  return String(str).replace(/^FY\s*/i, '').replace(/[\s–-]/g, '').trim().toLowerCase();
};

export const normalizeMonth = (str) => {
  const clean = String(str || '').trim().toLowerCase();
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  for (const mon of months) {
    if (clean.startsWith(mon) || clean.includes(mon)) return mon;
  }
  return clean.slice(0, 3);
};

/**
 * Step 3: Derives FY, Quarter and Month strictly from Month string without looking at Invoice Date
 */
export const derivePeriodFromMonth = (monthStr, fallbackFy = 'FY 2026–27') => {
  const clean = String(monthStr || '').trim();
  const monKey = normalizeMonth(clean);
  const mi = MN.findIndex((x) => x.toLowerCase() === monKey);
  const validMi = mi !== -1 ? mi : 3; // Default to Apr (3) if unspecified
  const mLabel = mi !== -1 ? MN[validMi] : clean || 'Apr';

  // Quarter: Apr-Jun = Q1, Jul-Sep = Q2, Oct-Dec = Q3, Jan-Mar = Q4
  const qCode = validMi >= 3 && validMi <= 5 ? 'Q1' : validMi >= 6 && validMi <= 8 ? 'Q2' : validMi >= 9 && validMi <= 11 ? 'Q3' : 'Q4';
  const qLabel = `${qCode} (${QM_LABEL[qCode] || ''})`;

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
};

/**
 * Get register rows filtered strictly by FY and Month (Step 4 & Step 5)
 * Fetches live data from Google Sheet tab if available, falling back to local DB cache
 */
export const getRegisterRows = async (req, res) => {
  try {
    const tabName = normalizeTabName(req.params.tabName);
    const db = getDb();
    let rows = db.registers?.[tabName] || [];

    // Attempt live fetch from Google Sheet tab (Single Source of Truth)
    try {
      const sheetRes = await googleSheets.readSheet({ sheet: tabName });
      if (sheetRes && sheetRes.ok && Array.isArray(sheetRes.rows)) {
        const liveRows = sheetRes.rows
          .map((r) => googleSheets.mapSheetRowToRecord(r.values || r, tabName))
          .filter(Boolean);

        rows = liveRows;
        if (!db.registers) db.registers = {};
        db.registers[tabName] = liveRows;
        if (tabName === REGISTER_TABS.GSTR1) {
          db.invoices = liveRows;
        }
        saveDb(db);
      }
    } catch (sheetErr) {
      console.warn(`[RegisterController] Live sheet fetch failed for ${tabName}, using local cache:`, sheetErr.message);
    }

    const { fy, m } = req.query;
    let filtered = [...rows];

    // FY Filter
    if (fy && fy !== 'ALL') {
      filtered = filtered.filter((r) => normalizeFy(r.fy || r.financialYear) === normalizeFy(fy));
    }

    // Step 4: Month Filter strictly on Month column, NEVER on Invoice Date!
    if (m && m !== 'ALL') {
      const targetMon = normalizeMonth(m);
      filtered = filtered.filter((r) => normalizeMonth(r.m || r.month) === targetMon);
    }

    res.json({
      success: true,
      tab: tabName,
      totalCount: rows.length,
      filteredCount: filtered.length,
      data: filtered,
      locked: Boolean(db.locked)
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

/**
 * Create a new row in the register tab
 */
export const createRegisterRow = async (req, res) => {
  try {
    const tabName = normalizeTabName(req.params.tabName);
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const item = req.body;
    const rowMonth = item.m || item.month || 'Apr';
    const period = derivePeriodFromMonth(rowMonth, item.fy || item.financialYear || 'FY 2026–27');

    const prefix = tabName === REGISTER_TABS.GSTR1 ? 'inv-' :
                   tabName === REGISTER_TABS.GSTR2B ? 'g2b-' :
                   tabName === REGISTER_TABS.RCM ? 'rcm-' : 'itc-';

    const newRecord = {
      ...item,
      id: item.id || (prefix + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)),
      m: item.m || period.m,
      q: item.q || period.q,
      fy: item.fy || period.fy,
      sheetSynced: false,
      createdAt: new Date().toISOString()
    };

    if (!db.registers[tabName]) db.registers[tabName] = [];
    db.registers[tabName].push(newRecord);
    if (tabName === REGISTER_TABS.GSTR1) {
      db.invoices = db.registers[tabName];
    }

    // Attempt Google Sheets Sync
    let sheetWarning = null;
    try {
      const sheetRes = await googleSheets.upsertGenericRows({
        sheet: tabName,
        rows: [newRecord]
      });
      if (sheetRes.ok) {
        newRecord.sheetSynced = true;
      } else if (!sheetRes.skipped) {
        sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
      }
    } catch (_) {
      sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
    }

    saveDb(db);

    res.status(201).json({
      success: true,
      tab: tabName,
      data: newRecord,
      sheetSynced: newRecord.sheetSynced,
      sheetWarning
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

/**
 * Update a row in the register tab
 */
export const updateRegisterRow = async (req, res) => {
  try {
    const tabName = normalizeTabName(req.params.tabName);
    const { id } = req.params;
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const rows = db.registers?.[tabName] || [];
    const index = rows.findIndex((r) => r.id === id);
    if (index === -1) {
      return res.status(404).json({ success: false, message: 'Record not found.' });
    }

    const updated = {
      ...rows[index],
      ...req.body,
      id,
      sheetSynced: false,
      updatedAt: new Date().toISOString()
    };

    rows[index] = updated;
    if (tabName === REGISTER_TABS.GSTR1) {
      db.invoices = rows;
    }

    // Attempt Google Sheets Sync
    let sheetWarning = null;
    try {
      const sheetRes = await googleSheets.upsertGenericRows({
        sheet: tabName,
        rows: [updated]
      });
      if (sheetRes.ok) {
        updated.sheetSynced = true;
      } else if (!sheetRes.skipped) {
        sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
      }
    } catch (_) {
      sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
    }

    saveDb(db);

    res.json({
      success: true,
      tab: tabName,
      data: updated,
      sheetSynced: updated.sheetSynced,
      sheetWarning
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

/**
 * Delete a row from the register tab
 */
export const deleteRegisterRow = async (req, res) => {
  try {
    const tabName = normalizeTabName(req.params.tabName);
    const { id } = req.params;
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    // 1. Google Sheets delete first (Rule 3)
    const sheetRes = await googleSheets.deleteRows([id], { sheet: tabName });
    if (!sheetRes.ok && !sheetRes.skipped) {
      return res.status(502).json({
        success: false,
        message: 'Google Sheet delete failed: ' + (sheetRes.error || 'Unknown sheet error')
      });
    }

    // 2. Only on sheet success, delete from DB
    const rows = db.registers?.[tabName] || [];
    const filtered = rows.filter((r) => r.id !== id);

    db.registers[tabName] = filtered;
    if (tabName === REGISTER_TABS.GSTR1) {
      db.invoices = filtered;
    }
    saveDb(db);

    res.json({
      success: true,
      tab: tabName,
      message: 'Record deleted successfully.'
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

/**
 * Bulk import into register tab (Step 3: Month comes from file, derives FY & Q, imports into tab)
 */
export const bulkImportRegister = async (req, res) => {
  try {
    const tabName = normalizeTabName(req.params.tabName);
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to import.' });
    }

    const { items, fixedMonth, fixedFy } = req.body;
    if (!Array.isArray(items)) {
      return res.status(400).json({ success: false, message: 'Items must be an array.' });
    }

    const prefix = tabName === REGISTER_TABS.GSTR1 ? 'inv-' :
                   tabName === REGISTER_TABS.GSTR2B ? 'g2b-' :
                   tabName === REGISTER_TABS.RCM ? 'rcm-' : 'itc-';

    const newRecords = [];
    let ok = 0;

    for (const item of items) {
      const rowMonth = item.m || item.month || fixedMonth || 'Apr';
      const period = derivePeriodFromMonth(rowMonth, item.fy || item.financialYear || fixedFy || 'FY 2026–27');

      const rec = {
        ...item,
        id: item.id || (prefix + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)),
        m: item.m || period.m,
        q: item.q || period.q,
        fy: item.fy || period.fy,
        sheetSynced: false,
        createdAt: new Date().toISOString()
      };
      newRecords.push(rec);
      ok++;
    }

    if (!db.registers[tabName]) db.registers[tabName] = [];
    db.registers[tabName].push(...newRecords);
    if (tabName === REGISTER_TABS.GSTR1) {
      db.invoices = db.registers[tabName];
    }

    // Google Sheets Sync
    let sheetWarning = null;
    let sheetSynced = false;
    if (newRecords.length > 0) {
      try {
        const sheetRes = await googleSheets.upsertGenericRows({
          sheet: tabName,
          rows: newRecords
        });
        if (sheetRes.ok) {
          sheetSynced = true;
          newRecords.forEach((r) => { r.sheetSynced = true; });
        } else if (!sheetRes.skipped) {
          sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
        }
      } catch (_) {
        sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
      }
    }

    saveDb(db);

    res.json({
      success: true,
      tab: tabName,
      imported: ok,
      sheetSynced,
      sheetWarning
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

/**
 * Clear period records from register tab
 */
export const clearRegister = async (req, res) => {
  try {
    const tabName = normalizeTabName(req.params.tabName);
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { fy, m } = req.query;
    const rows = db.registers?.[tabName] || [];
    const hasFy = fy && fy !== 'ALL';
    const hasMonth = m && m !== 'ALL';

    const toDeleteIds = [];
    rows.forEach((r) => {
      const fyMatch = !hasFy || normalizeFy(r.fy || r.financialYear) === normalizeFy(fy);
      const monMatch = !hasMonth || normalizeMonth(r.m || r.month) === normalizeMonth(m);
      if (fyMatch && monMatch) {
        toDeleteIds.push(r.id);
      }
    });

    if (toDeleteIds.length > 0) {
      // 1. Google Sheets delete in ONE batch request (Rule 3)
      const sheetRes = await googleSheets.deleteRows(toDeleteIds, { sheet: tabName });
      if (!sheetRes.ok && !sheetRes.skipped) {
        return res.status(502).json({
          success: false,
          message: 'Google Sheet clear failed: ' + (sheetRes.error || 'Unknown sheet error')
        });
      }
    }

    // 2. Only on sheet success, delete from DB
    const toDeleteSet = new Set(toDeleteIds);
    const remaining = rows.filter(r => !toDeleteSet.has(r.id));
    db.registers[tabName] = remaining;
    if (tabName === REGISTER_TABS.GSTR1) {
      db.invoices = remaining;
    }
    saveDb(db);

    res.json({
      success: true,
      tab: tabName,
      count: toDeleteIds.length,
      message: `Cleared ${toDeleteIds.length} records from ${tabName}.`
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

/**
 * Sync unsynced records to Google Sheets
 */
export const syncRegister = async (req, res) => {
  try {
    const tabName = normalizeTabName(req.params.tabName);
    const db = getDb();
    const rows = db.registers?.[tabName] || [];
    const { forceAll } = req.body || {};
    let targetRows = rows.filter((r) => !r.sheetSynced);

    if (forceAll || targetRows.length === 0) {
      targetRows = rows;
    }

    if (targetRows.length === 0) {
      return res.json({
        success: true,
        tab: tabName,
        count: 0,
        message: `No records found in ${tabName} to sync.`
      });
    }

    const result = await googleSheets.upsertGenericRows({
      sheet: tabName,
      rows: targetRows
    });

    if (result.skipped) {
      return res.json({
        success: false,
        skipped: true,
        message: 'Google Sheets sync not configured.'
      });
    }

    if (result.ok) {
      rows.forEach((r) => {
        r.sheetSynced = true;
      });
      saveDb(db);
      return res.json({
        success: true,
        tab: tabName,
        count: targetRows.length,
        message: `${targetRows.length} records synced to Google Sheet ${tabName} successfully.`
      });
    } else {
      return res.status(500).json({
        success: false,
        message: 'Google Sheet sync fail hua: ' + (result.error || 'Unknown error')
      });
    }
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
};
