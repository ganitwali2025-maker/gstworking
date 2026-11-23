import express from 'express';
import googleSheets from '../services/googleSheets.js';
import { getDb, saveDb } from '../data/store.js';

const router = express.Router();

/**
 * Initialize all 4 tabs in Google Sheets (GSTR-1, ITC-Purchase, RCM, GSTR-2B)
 */
router.post('/init-tabs', async (req, res) => {
  try {
    const results = await googleSheets.ensureAllTabs();
    return res.json({ success: true, results });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * Shift any g2b-... rows mistakenly present in GSTR-1 to GSTR-2B
 */
router.post('/shift-g2b', async (req, res) => {
  try {
    const db = getDb();
    const gstr1Rows = db.registers?.['GSTR-1'] || db.invoices || [];
    const g2bRows = gstr1Rows.filter((r) => String(r.id || '').toLowerCase().startsWith('g2b'));

    if (g2bRows.length > 0) {
      // Remove from GSTR-1 in local DB
      db.registers['GSTR-1'] = gstr1Rows.filter((r) => !String(r.id || '').toLowerCase().startsWith('g2b'));
      db.invoices = db.registers['GSTR-1'];

      // Add to GSTR-2B in local DB
      const existing2bIds = new Set((db.registers['GSTR-2B'] || []).map((r) => r.id));
      for (const row of g2bRows) {
        if (!existing2bIds.has(row.id)) {
          db.registers['GSTR-2B'].push(row);
          existing2bIds.add(row.id);
        }
      }
      saveDb(db);
    }

    // Call Google Sheets shift
    const sheetRes = await googleSheets.shiftG2bFromGstr1(g2bRows);

    return res.json({
      success: true,
      shiftedCount: g2bRows.length || sheetRes.shiftedCount || 0,
      sheetResult: sheetRes
    });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * Generic Google Sheets sync endpoint for any register
 */
router.post('/sync', async (req, res) => {
  try {
    const { action, sheet, headers, rows, ids } = req.body;

    if (!sheet) {
      return res.status(400).json({ success: false, error: 'Missing sheet name' });
    }

    if (action === 'read') {
      const result = await googleSheets.readSheet({ sheet });
      const rawRows = result.rows || [];
      const mapped = rawRows.map(r => googleSheets.mapSheetRowToRecord(r.values || r, sheet)).filter(Boolean);
      return res.json({
        success: Boolean(result.ok),
        count: mapped.length,
        rawCount: rawRows.length,
        rows: mapped,
        error: result.error || null
      });
    }

    if (action === 'upsert') {
      const result = await googleSheets.upsertGenericRows({ sheet, headers, rows: rows || [] });
      return res.json({
        success: Boolean(result.ok),
        count: result.count || 0,
        error: result.error || null
      });
    }

    if (action === 'delete') {
      const result = await googleSheets.deleteRows(ids || [], { sheet, headers });
      return res.json({
        success: Boolean(result.ok),
        error: result.error || null
      });
    }

    if (action === 'lock') {
      const result = await googleSheets.lockSheet(ids || [], { sheet, headers });
      return res.json({
        success: Boolean(result.ok),
        error: result.error || null
      });
    }

    if (action === 'unlock') {
      const result = await googleSheets.unlockSheet(ids || [], { sheet, headers });
      return res.json({
        success: Boolean(result.ok),
        error: result.error || null
      });
    }

    if (action === 'clear') {
      const result = await googleSheets.clearSheet({ sheet, headers });
      return res.json({
        success: Boolean(result.ok),
        error: result.error || null
      });
    }

    return res.status(400).json({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * Status and diagnostic endpoint for Google Sheets integration
 */
router.get('/status', async (req, res) => {
  try {
    const config = Boolean(process.env.APPS_SCRIPT_URL?.trim() && process.env.APPS_SCRIPT_SECRET?.trim());
    if (!config) {
      return res.json({
        configured: false,
        message: 'APPS_SCRIPT_URL or APPS_SCRIPT_SECRET not configured in .env'
      });
    }

    const testRes = await googleSheets.ensureAllTabs();
    return res.json({
      configured: true,
      url: process.env.APPS_SCRIPT_URL?.slice(0, 45) + '...',
      testTabs: testRes
    });
  } catch (err) {
    return res.status(500).json({ configured: false, error: err.message });
  }
});

/**
 * Universal Sync All endpoint: Syncs all registers (GSTR-1, ITC-Purchase, RCM, GSTR-2B) to Google Sheets
 */
router.post('/sync-all', async (req, res) => {
  try {
    const db = getDb();
    const payload = req.body || {};

    const gstr1Rows = (payload.gstr1 && Array.isArray(payload.gstr1) && payload.gstr1.length > 0)
      ? payload.gstr1
      : (db.registers?.['GSTR-1'] || db.invoices || []);

    const itcRows = (payload.itc && Array.isArray(payload.itc) && payload.itc.length > 0)
      ? payload.itc
      : (db.registers?.['ITC-Purchase'] || []);

    const rcmRows = (payload.rcm && Array.isArray(payload.rcm) && payload.rcm.length > 0)
      ? payload.rcm
      : (db.registers?.['RCM'] || []);

    const g2bRows = (payload.gstr2b && Array.isArray(payload.gstr2b) && payload.gstr2b.length > 0)
      ? payload.gstr2b
      : (db.registers?.['GSTR-2B'] || []);

    // Save to local db so everything stays unified
    if (!db.registers) db.registers = {};
    if (gstr1Rows.length > 0) {
      db.registers['GSTR-1'] = gstr1Rows;
      db.invoices = gstr1Rows;
    }
    if (itcRows.length > 0) db.registers['ITC-Purchase'] = itcRows;
    if (rcmRows.length > 0) db.registers['RCM'] = rcmRows;
    if (g2bRows.length > 0) db.registers['GSTR-2B'] = g2bRows;
    saveDb(db);

    const results = {};
    if (gstr1Rows.length > 0) {
      results['GSTR-1'] = await googleSheets.upsertGenericRows({
        sheet: googleSheets.REGISTER_TABS.GSTR1,
        rows: gstr1Rows
      });
    }
    if (itcRows.length > 0) {
      results['ITC-Purchase'] = await googleSheets.upsertGenericRows({
        sheet: googleSheets.REGISTER_TABS.ITC_PURCHASE,
        rows: itcRows
      });
    }
    if (rcmRows.length > 0) {
      results['RCM'] = await googleSheets.upsertGenericRows({
        sheet: googleSheets.REGISTER_TABS.RCM,
        rows: rcmRows
      });
    }
    if (g2bRows.length > 0) {
      results['GSTR-2B'] = await googleSheets.upsertGenericRows({
        sheet: googleSheets.REGISTER_TABS.GSTR2B,
        rows: g2bRows
      });
    }

    return res.json({
      success: true,
      message: 'All records synced to Google Sheets successfully',
      counts: {
        gstr1: gstr1Rows.length,
        itc: itcRows.length,
        rcm: rcmRows.length,
        gstr2b: g2bRows.length
      },
      results
    });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

export default router;
