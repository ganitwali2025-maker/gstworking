import { getDb, saveDb } from '../data/store.js';

const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const qOf = (mi) => (mi >= 3 && mi <= 5 ? 'Q1' : mi >= 6 && mi <= 8 ? 'Q2' : mi >= 9 && mi <= 11 ? 'Q3' : 'Q4');
const fyOf = (y, mi) => {
  const a = mi >= 3 ? y : y - 1;
  return 'FY ' + a + '–' + String(a + 1).slice(2);
};

export const stampDate = (dateStr, fixedMonth, fixedFy) => {
  if (fixedMonth && fixedFy) {
    const mi = MN.indexOf(fixedMonth.slice(0, 3));
    return {
      fy: fixedFy,
      q: qOf(mi),
      m: fixedMonth
    };
  }
  const parts = dateStr.split('-').map(Number);
  const y = parts[0];
  const mi = parts[1] - 1;
  return {
    fy: fyOf(y, mi),
    q: qOf(mi),
    m: MN[mi] + ' ' + y
  };
};

export const getInvoices = (req, res) => {
  try {
    const db = getDb();
    const { fy, q, m } = req.query;
    let list = db.invoices || [];

    const normalizeFy = (str) => String(str || '').replace(/^FY\s*/i, '').replace(/–/g, '-').trim();
    if (fy) list = list.filter(r => normalizeFy(r.fy) === normalizeFy(fy));
    if (q) list = list.filter(r => r.q === q);
    if (m) list = list.filter(r => r.m === m);

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

export const createInvoice = (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { d, p, n, g = '', t, i, c, s, fixedMonth, fixedFy } = req.body;

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

    const st = stampDate(d, fixedMonth, fixedFy);

    // Check duplicate invoice number in same financial year
    const exists = db.invoices.some(x => x.fy === st.fy && x.n.toLowerCase() === n.trim().toLowerCase());
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
      createdAt: new Date().toISOString()
    };

    db.invoices.push(newRecord);
    saveDb(db);

    res.status(201).json({
      success: true,
      data: newRecord,
      warning: (cNum !== sNum && (cNum > 0 || sNum > 0)) ? 'Note: CGST and SGST amounts are different.' : null
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const updateInvoice = (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { id } = req.params;
    const { d, p, n, g = '', t, i, c, s, fixedMonth, fixedFy } = req.body;

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

    const st = stampDate(d, fixedMonth, fixedFy);

    // Check duplicate invoice number in same financial year excluding current
    const exists = db.invoices.some(x => x.id !== id && x.fy === st.fy && x.n.toLowerCase() === n.trim().toLowerCase());
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
      updatedAt: new Date().toISOString()
    };

    db.invoices[index] = updated;
    saveDb(db);

    res.json({
      success: true,
      data: updated,
      warning: (cNum !== sNum && (cNum > 0 || sNum > 0)) ? 'Note: CGST and SGST amounts are different.' : null
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const deleteInvoice = (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { id } = req.params;
    const initialLen = db.invoices.length;
    db.invoices = db.invoices.filter(x => x.id !== id);

    if (db.invoices.length === initialLen) {
      return res.status(404).json({ success: false, message: 'Invoice not found.' });
    }

    saveDb(db);
    res.json({ success: true, message: 'Invoice deleted successfully.' });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const clearInvoices = (req, res) => {
  try {
    const db = getDb();
    if (db.locked) {
      return res.status(403).json({ success: false, message: 'Sheet is locked. Unlock to make changes.' });
    }

    const { fy, q, m } = req.query;
    const initialCount = db.invoices.length;

    db.invoices = db.invoices.filter(r => {
      if (fy && r.fy !== fy) return true;
      if (q && r.q !== q) return true;
      if (m && r.m !== m) return true;
      return false;
    });

    const deletedCount = initialCount - db.invoices.length;
    saveDb(db);

    res.json({ success: true, count: deletedCount, message: `Cleared ${deletedCount} records.` });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const bulkImport = (req, res) => {
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

    for (const item of items) {
      const { d, p, n, g = '', t = 0, i = 0, c = 0, s = 0 } = item;
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

      const st = stampDate(d, fixedMonth, fixedFy);
      const exists = db.invoices.some(x => x.fy === st.fy && x.n.toLowerCase() === String(n).trim().toLowerCase());
      if (exists) {
        bad++;
        continue;
      }

      db.invoices.push({
        id: 'inv-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
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
        createdAt: new Date().toISOString()
      });
      ok++;
    }

    saveDb(db);
    res.json({ success: true, imported: ok, skipped: bad });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const getLockStatus = (req, res) => {
  const db = getDb();
  res.json({ success: true, locked: !!db.locked });
};

export const toggleLock = (req, res) => {
  const db = getDb();
  const { locked } = req.body;
  db.locked = typeof locked === 'boolean' ? locked : !db.locked;
  saveDb(db);
  res.json({ success: true, locked: db.locked, message: db.locked ? 'Sheet locked' : 'Sheet unlocked' });
};
