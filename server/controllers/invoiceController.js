import { getDb, saveDb } from '../data/store.js';
import googleSheets from '../services/googleSheets.js';

const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const QM_LABEL = {
  Q1: 'Apr-Jun',
  Q2: 'Jul-Sep',
  Q3: 'Oct-Dec',
  Q4: 'Jan-Mar'
};
const qOf = (mi) => (mi >= 3 && mi <= 5 ? 'Q1' : mi >= 6 && mi <= 8 ? 'Q2' : mi >= 9 && mi <= 11 ? 'Q3' : 'Q4');
const fyOf = (y, mi) => {
  const a = mi >= 3 ? y : y - 1;
  return 'FY ' + a + '–' + String(a + 1).slice(2);
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

export const stampDate = (dateStr, fixedMonth, fixedFy) => {
  const parts = String(dateStr || '').split('-').map(Number);
  const y = parts[0] || 2026;
  const mi = (parts[1] && !isNaN(parts[1])) ? parts[1] - 1 : 7;
  const qCode = qOf(mi);
  const dateFy = fyOf(y, mi);

  const hasFixedFy = Boolean(fixedFy && fixedFy !== 'ALL' && fixedFy.trim());
  const finalFy = hasFixedFy ? fixedFy : dateFy;

  if (fixedMonth && fixedMonth !== 'ALL') {
    const normKey = normalizeMonth(fixedMonth);
    const foundMi = MN.findIndex(x => x.toLowerCase() === normKey);
    const validMi = foundMi !== -1 ? foundMi : mi;
    const finalQ = qOf(validMi);
    return {
      fy: finalFy,
      q: `${finalQ} (${QM_LABEL[finalQ] || ''})`,
      m: MN[validMi]
    };
  }

  return {
    fy: finalFy,
    q: `${qCode} (${QM_LABEL[qCode] || ''})`,
    m: MN[mi]
  };
};

export const getInvoices = async (req, res) => {
  try {
    const db = getDb();
    let list = db.invoices || [];

    // Live fetch from Google Sheets GSTR-1 tab (Google Sheet is Single Source of Truth)
    try {
      const sheetRes = await googleSheets.readSheet({ sheet: 'GSTR-1' });
      if (sheetRes && sheetRes.ok && Array.isArray(sheetRes.rows)) {
        const liveRows = sheetRes.rows
          .map((r) => googleSheets.mapSheetRowToRecord(r.values || r, 'GSTR-1'))
          .filter(Boolean);
        list = liveRows;
        db.invoices = liveRows;
        if (!db.registers) db.registers = {};
        db.registers['GSTR-1'] = liveRows;
        saveDb(db);
      }
    } catch (sheetErr) {
      console.warn('[InvoiceController] Live sheet fetch for GSTR-1 failed, using cache:', sheetErr.message);
    }

    const { fy, q, m } = req.query;

    if (fy && fy !== 'ALL') list = list.filter(r => normalizeFy(r.fy) === normalizeFy(fy));
    if (q && q !== 'ALL') list = list.filter(r => r.q === q);
    if (m && m !== 'ALL') list = list.filter(r => normalizeMonth(r.m) === normalizeMonth(m));

    // Sort by date ascending, then invoice number
    list.sort((a, b) => (a.d < b.d ? -1 : a.d > b.d ? 1 : a.n.localeCompare(b.n)));

    res.json({
      success: true,
      data: list,
      locked: db.locked
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const createInvoice = async (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { d, p, n, g = '', t, i, c, s, fixedMonth, fixedFy, m: bodyM, q: bodyQ, fy: bodyFy } = req.body;

    if (!d) return res.status(400).json({ success: false, message: 'Enter the invoice date.' });
    if (!n) return res.status(400).json({ success: false, message: 'Enter the invoice number.' });
    if (!p) return res.status(400).json({ success: false, message: 'Enter the supplier / party name.' });

    const cleanGst = (g || '').trim().toUpperCase();
    if (cleanGst && !/^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/.test(cleanGst)) {
      return res.status(400).json({ success: false, message: 'Enter a valid 15-character GST No, or leave blank for unregistered.' });
    }

    const tNum = Number(t) || 0;
    const iNum = Number(i) || 0;
    const cNum = Number(c) || 0;
    const sNum = Number(s) || 0;

    if (iNum > 0 && (cNum > 0 || sNum > 0)) {
      return res.status(400).json({ success: false, message: 'An invoice cannot have IGST together with CGST/SGST.' });
    }

    const cleanMonth = (bodyM || fixedMonth)
      ? (MN[MN.findIndex(x => x.toLowerCase() === normalizeMonth(bodyM || fixedMonth))] || bodyM || fixedMonth)
      : null;

    const st = (bodyFy && cleanMonth)
      ? {
          fy: bodyFy,
          q: bodyQ || stampDate(d).q,
          m: cleanMonth
        }
      : stampDate(d, fixedMonth, fixedFy);

    // Check duplicate invoice number in same financial year
    const exists = db.invoices.some(x => normalizeFy(x.fy) === normalizeFy(st.fy) && x.n.toLowerCase() === n.trim().toLowerCase());
    if (exists) {
      return res.status(400).json({ success: false, message: `Invoice No "${n}" already exists in ${st.fy}.` });
    }

    const newRecord = {
      id: 'inv-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      fy: st.fy,
      q: st.q,
      m: st.m,
      d,
      p: p.trim(),
      g: cleanGst,
      n: n.trim(),
      t: tNum,
      i: iNum,
      c: cNum,
      s: sNum,
      sheetSynced: false,
      createdAt: new Date().toISOString()
    };

    // Attempt Google Sheets Sync
    let sheetWarning = null;
    try {
      const sheetRes = await googleSheets.upsertRows([newRecord]);
      if (sheetRes.ok) {
        newRecord.sheetSynced = true;
      } else if (!sheetRes.skipped) {
        sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
      }
    } catch (_) {
      sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
    }

    db.invoices.push(newRecord);
    saveDb(db);

    res.status(201).json({
      success: true,
      data: newRecord,
      sheetSynced: newRecord.sheetSynced,
      sheetWarning,
      warning: (cNum !== sNum && (cNum > 0 || sNum > 0)) ? 'Note: CGST and SGST amounts are different.' : null
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const updateInvoice = async (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { id } = req.params;
    const { d, p, n, g = '', t, i, c, s, fixedMonth, fixedFy, m: bodyM, q: bodyQ, fy: bodyFy } = req.body;

    const index = db.invoices.findIndex(x => x.id === id);
    if (index === -1) {
      return res.status(404).json({ success: false, message: 'Invoice not found.' });
    }

    if (!d || !n || !p) {
      return res.status(400).json({ success: false, message: 'Invoice date, number, and party name are required.' });
    }

    const cleanGst = (g || '').trim().toUpperCase();
    if (cleanGst && !/^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/.test(cleanGst)) {
      return res.status(400).json({ success: false, message: 'Invalid 15-character GSTIN.' });
    }

    const tNum = Number(t) || 0;
    const iNum = Number(i) || 0;
    const cNum = Number(c) || 0;
    const sNum = Number(s) || 0;

    if (iNum > 0 && (cNum > 0 || sNum > 0)) {
      return res.status(400).json({ success: false, message: 'An invoice cannot have IGST together with CGST/SGST.' });
    }

    const cleanMonth = (bodyM || fixedMonth)
      ? (MN[MN.findIndex(x => x.toLowerCase() === normalizeMonth(bodyM || fixedMonth))] || bodyM || fixedMonth)
      : null;

    const st = (bodyFy && cleanMonth)
      ? {
          fy: bodyFy,
          q: bodyQ || db.invoices[index].q,
          m: cleanMonth
        }
      : stampDate(d, fixedMonth, fixedFy);

    // Check duplicate invoice number in same financial year excluding current
    const exists = db.invoices.some(x => x.id !== id && normalizeFy(x.fy) === normalizeFy(st.fy) && x.n.toLowerCase() === n.trim().toLowerCase());
    if (exists) {
      return res.status(400).json({ success: false, message: `Invoice No "${n}" already exists in ${st.fy}.` });
    }

    const updated = {
      ...db.invoices[index],
      fy: st.fy,
      q: st.q,
      m: st.m,
      d,
      p: p.trim(),
      g: cleanGst,
      n: n.trim(),
      t: tNum,
      i: iNum,
      c: cNum,
      s: sNum,
      sheetSynced: false,
      updatedAt: new Date().toISOString()
    };

    // Attempt Google Sheets Sync
    let sheetWarning = null;
    try {
      const sheetRes = await googleSheets.upsertRows([updated]);
      if (sheetRes.ok) {
        updated.sheetSynced = true;
      } else if (!sheetRes.skipped) {
        sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
      }
    } catch (_) {
      sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
    }

    db.invoices[index] = updated;
    saveDb(db);

    res.json({
      success: true,
      data: updated,
      sheetSynced: updated.sheetSynced,
      sheetWarning,
      warning: (cNum !== sNum && (cNum > 0 || sNum > 0)) ? 'Note: CGST and SGST amounts are different.' : null
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const deleteInvoice = async (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { id } = req.params;

    // 1. Google Sheets delete first (Rule 3)
    const sheetRes = await googleSheets.deleteRows([id], { sheet: 'GSTR-1' });
    if (!sheetRes.ok && !sheetRes.skipped) {
      return res.status(502).json({
        success: false,
        message: 'Google Sheet delete failed: ' + (sheetRes.error || 'Unknown sheet error')
      });
    }

    // 2. Only on sheet success, delete from DB
    db.invoices = (db.invoices || []).filter(x => x.id !== id);
    if (db.registers && db.registers['GSTR-1']) {
      db.registers['GSTR-1'] = db.registers['GSTR-1'].filter(x => x.id !== id);
    }
    saveDb(db);

    res.json({ success: true, message: 'Invoice deleted successfully.' });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const clearInvoices = async (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { fy, q, m } = req.query;
    const hasFy = fy && fy !== 'ALL';
    const hasMonth = m && m !== 'ALL';

    const toDeleteIds = [];
    db.invoices.forEach(r => {
      const fyMatch = !hasFy || normalizeFy(r.fy) === normalizeFy(fy);
      const qMatch = !q || q === 'ALL' || r.q === q;
      const monthMatch = !hasMonth || normalizeMonth(r.m) === normalizeMonth(m);
      if (fyMatch && qMatch && monthMatch) {
        toDeleteIds.push(r.id);
      }
    });

    if (toDeleteIds.length > 0) {
      // 1. Google Sheets delete in ONE batch request
      const sheetRes = await googleSheets.deleteRows(toDeleteIds, { sheet: 'GSTR-1' });
      if (!sheetRes.ok && !sheetRes.skipped) {
        return res.status(502).json({
          success: false,
          message: 'Google Sheet clear failed: ' + (sheetRes.error || 'Unknown sheet error')
        });
      }
    }

    // 2. Only on sheet success, delete from DB
    const toDeleteSet = new Set(toDeleteIds);
    db.invoices = db.invoices.filter(r => !toDeleteSet.has(r.id));
    if (db.registers && db.registers['GSTR-1']) {
      db.registers['GSTR-1'] = db.registers['GSTR-1'].filter(r => !toDeleteSet.has(r.id));
    }
    saveDb(db);

    res.json({ success: true, count: toDeleteIds.length, message: `Cleared ${toDeleteIds.length} records.` });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const bulkImport = async (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to import.' });
    }

    const { items, fixedMonth, fixedFy } = req.body;
    if (!Array.isArray(items)) {
      return res.status(400).json({ success: false, message: 'Invalid payload: items must be an array.' });
    }

    let ok = 0;
    let bad = 0;
    const newRecords = [];

    for (const item of items) {
      const { d, p, n, g = '', t = 0, i = 0, c = 0, s = 0, id: existingId, fy: itemFy, q: itemQ, m: itemM } = item;
      if (!d || !n || !p) {
        bad++;
        continue;
      }

      const tNum = Number(t) || 0;
      const iNum = Number(i) || 0;
      const cNum = Number(c) || 0;
      const sNum = Number(s) || 0;

      if (iNum > 0 && (cNum > 0 || sNum > 0)) {
        bad++;
        continue;
      }

      const cleanMonth = (itemM || fixedMonth)
        ? (MN[MN.findIndex(x => x.toLowerCase() === normalizeMonth(itemM || fixedMonth))] || itemM || fixedMonth)
        : null;

      const st = (itemFy && cleanMonth)
        ? { fy: itemFy, q: itemQ || '', m: cleanMonth }
        : stampDate(d, fixedMonth, fixedFy);

      const exists = db.invoices.some(x => normalizeFy(x.fy) === normalizeFy(st.fy) && x.n.toLowerCase() === String(n).trim().toLowerCase());
      if (exists) {
        bad++;
        continue;
      }

      const rec = {
        id: existingId || ('inv-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)),
        fy: st.fy,
        q: st.q,
        m: st.m,
        d,
        p: String(p).trim(),
        g: String(g || '').trim().toUpperCase(),
        n: String(n).trim(),
        t: tNum,
        i: iNum,
        c: cNum,
        s: sNum,
        sheetSynced: false,
        createdAt: new Date().toISOString()
      };

      newRecords.push(rec);
      ok++;
    }

    // Attempt Google Sheets Batch Upsert (chunked in 200s inside service)
    let sheetWarning = null;
    let sheetSynced = false;
    if (newRecords.length > 0) {
      try {
        const sheetRes = await googleSheets.upsertRows(newRecords);
        if (sheetRes.ok) {
          sheetSynced = true;
          newRecords.forEach(r => { r.sheetSynced = true; });
        } else if (!sheetRes.skipped) {
          sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
        }
      } catch (_) {
        sheetWarning = 'Data save ho gaya, par Google Sheet sync fail hua.';
      }
      db.invoices.push(...newRecords);
      saveDb(db);
    }

    res.json({
      success: true,
      imported: ok,
      skipped: bad,
      sheetSynced,
      sheetWarning
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const getLockStatus = (req, res) => {
  const db = getDb();
  res.json({ success: true, locked: !!db.locked });
};

export const toggleLock = async (req, res) => {
  try {
    const db = getDb();
    const { locked } = req.body;
    const newLock = typeof locked === 'boolean' ? locked : !db.locked;
    db.locked = newLock;
    saveDb(db);

    try {
      const invoiceIds = (db.invoices || []).map(x => x.id);
      if (newLock) {
        await googleSheets.lockSheet(invoiceIds);
      } else {
        await googleSheets.unlockSheet(invoiceIds);
      }
    } catch (_) {}

    res.json({
      success: true,
      locked: db.locked,
      message: db.locked ? 'Sheet locked' : 'Sheet unlocked'
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const syncToGoogleSheets = async (req, res) => {
  try {
    const db = getDb();
    const { forceAll } = req.body || {};
    let targetInvoices = (db.invoices || []).filter(r => !r.sheetSynced);

    if (forceAll || targetInvoices.length === 0) {
      // If forceAll requested or all are already marked synced, push all invoices so empty sheets get populated
      targetInvoices = db.invoices || [];
    }

    if (targetInvoices.length === 0) {
      return res.json({
        success: true,
        count: 0,
        message: 'No invoices found in database to sync.'
      });
    }

    const result = await googleSheets.upsertRows(targetInvoices);
    if (result.skipped) {
      return res.json({
        success: false,
        skipped: true,
        message: 'Google Sheets sync is not configured (APPS_SCRIPT_URL/SECRET missing in .env).'
      });
    }

    if (result.ok) {
      db.invoices.forEach(r => {
        r.sheetSynced = true;
      });
      saveDb(db);
      return res.json({
        success: true,
        count: targetInvoices.length,
        message: `${targetInvoices.length} invoices synced to Google Sheet successfully.`
      });
    } else {
      return res.status(500).json({
        success: false,
        message: 'Data save ho gaya, par Google Sheet sync fail hua: ' + (result.error || 'Unknown error')
      });
    }
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
};
