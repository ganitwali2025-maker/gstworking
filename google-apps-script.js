/**
 * Google Apps Script Web App for Unified GST ITC Liability Dashboard
 * Supports: GSTR-1, ITC-Purchase, RCM, GSTR-2B tabs
 *
 * Deployment Instructions:
 * 1. Open your Google Spreadsheet ("GST ITC Liability Dashboard").
 * 2. Click Extensions > Apps Script.
 * 3. Replace all code in Code.gs with this file's code.
 * 4. Click Deploy > Manage Deployments > Edit (pencil icon) > Version: New version > Deploy.
 * 5. Ensure APPS_SCRIPT_SECRET matches your .env file ("MY_GST_SECRET_KEY_123").
 */

const CONFIG = {
  SECRET: 'MY_GST_SECRET_KEY_123',
  TABS: {
    GSTR1: 'GSTR-1',
    ITC_PURCHASE: 'ITC-Purchase',
    RCM: 'RCM',
    GSTR2B: 'GSTR-2B'
  },
  STANDARD_HEADERS: [
    'ID', 'Financial Year', 'Quarter', 'Month', 'Invoice Date',
    'Supplier / Party Name', 'GST No', 'Invoice No', 'Invoice Value',
    'Taxable Value', 'IGST', 'CGST', 'SGST', 'CESS', 'Total Tax'
  ],
  RCM_HEADERS: [
    'ID', 'Financial Year', 'Quarter', 'Month', 'Invoice Date',
    'Supplier / Party Name', 'GST No', 'Invoice No', 'Invoice Value',
    'Taxable Value', 'IGST', 'CGST', 'SGST', 'CESS', 'Total Tax',
    'RCM Category', 'RCM Tax Rate', 'RCM Applicable', 'Eligible ITC'
  ]
};

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    status: 'GST ITC Liability Dashboard Apps Script is active and running.',
    spreadsheetName: ss ? ss.getName() : 'Unknown',
    spreadsheetUrl: ss ? ss.getUrl() : '',
    spreadsheetId: ss ? ss.getId() : '',
    existingTabs: ss ? ss.getSheets().map(s => s.getName()) : [],
    tabs: [CONFIG.TABS.GSTR1, CONFIG.TABS.ITC_PURCHASE, CONFIG.TABS.RCM, CONFIG.TABS.GSTR2B]
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    let payload = {};
    try {
      payload = JSON.parse(e.postData.contents || '{}');
    } catch (err) {
      return jsonResponse({ ok: false, error: 'Malformed JSON payload' });
    }

    if (payload.secret !== CONFIG.SECRET) {
      return jsonResponse({ ok: false, error: 'Unauthorized: Invalid secret key' }, 401);
    }

    const action = String(payload.action || '').trim();
    const sheetName = String(payload.sheet || CONFIG.TABS.GSTR1).trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    switch (action) {
      case 'ping':
      case 'status':
        return jsonResponse({
          ok: true,
          spreadsheetName: ss ? ss.getName() : 'Unknown',
          spreadsheetUrl: ss ? ss.getUrl() : '',
          tabs: ss ? ss.getSheets().map(s => s.getName()) : []
        });

      case 'upsert':
        return handleUpsert(ss, sheetName, payload.headers, payload.rows || []);

      case 'delete':
        return handleDelete(ss, sheetName, payload.ids || []);

      case 'clear':
        return handleClear(ss, sheetName);

      case 'read':
      case 'get':
      case 'load':
        return handleRead(ss, sheetName);

      case 'shiftG2b':
        return handleShiftG2b(ss);

      case 'lock':
        return handleLock(ss, sheetName, true);

      case 'unlock':
        return handleLock(ss, sheetName, false);

      default:
        return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  const defaultHeaders = sheetName === CONFIG.TABS.RCM ? CONFIG.RCM_HEADERS : CONFIG.STANDARD_HEADERS;
  const useHeaders = (headers && headers.length > 0) ? headers : defaultHeaders;

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(useHeaders);
    formatHeaderRow(sheet, useHeaders.length);
  } else {
    // If sheet exists but is empty or missing headers
    if (sheet.getLastRow() < 1) {
      sheet.appendRow(useHeaders);
      formatHeaderRow(sheet, useHeaders.length);
    }
  }
  return sheet;
}

function formatHeaderRow(sheet, colCount) {
  try {
    const range = sheet.getRange(1, 1, 1, colCount);
    range.setFontWeight('bold');
    range.setBackground('#1e293b');
    range.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  } catch (_) {}
}

function handleUpsert(ss, sheetName, headers, rows) {
  if (!rows || rows.length === 0) {
    getOrCreateSheet(ss, sheetName, headers);
    return jsonResponse({ ok: true, count: 0, sheet: sheetName });
  }

  const sheet = getOrCreateSheet(ss, sheetName, headers);
  const lastRow = sheet.getLastRow();
  const idMap = new Map(); // id -> rowIndex (1-based)

  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < existingIds.length; i++) {
      const idVal = String(existingIds[i][0] || '').trim();
      if (idVal) {
        idMap.set(idVal, i + 2);
      }
    }
  }

  const defaultHeaders = sheetName === CONFIG.TABS.RCM ? CONFIG.RCM_HEADERS : CONFIG.STANDARD_HEADERS;
  const colCount = (headers && headers.length > 0) ? headers.length : defaultHeaders.length;

  let upsertCount = 0;
  const appendRows = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    let rowId = '';
    let rowValues = [];

    if (Array.isArray(r.values)) {
      rowId = String(r.id || r.values[0] || '').trim();
      rowValues = r.values;
    } else {
      // Object format
      rowId = String(r.id || '').trim();
      const rawParty = String(r.partyName || r.supplierName || r.tradeName || r.transporter || r.p || '').trim();
      const rawGst = String(r.gstNo || r.gstin || r.g || '').trim().toUpperCase();
      const isGstRegex = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/i;
      const cleanParty = (isGstRegex.test(rawParty) && (rawParty.toUpperCase() === rawGst || !rawGst)) ? '' : rawParty;

      rowValues = [
        rowId,
        r.financialYear || r.fy || '',
        r.quarter || r.q || '',
        r.month || r.m || '',
        r.invoiceDate || r.d || '',
        cleanParty,
        rawGst,
        r.invoiceNo || r.invNo || r.n || '',
        r.totalInvoiceValue || r.invVal || 0,
        r.taxableValue || r.taxable || r.t || r.amount || 0,
        r.igst || r.i || 0,
        r.cgst || r.c || 0,
        r.sgst || r.s || 0,
        r.cess || 0,
        r.totalTax || 0
      ];
      if (sheetName === CONFIG.TABS.RCM) {
        rowValues.push(r.rcmCategory || r.category || 'Transport (GTA)');
        rowValues.push(r.rcmTaxRate || r.taxRate || 5);
        rowValues.push(r.rcmApplicable || 'Yes');
        rowValues.push(r.eligibleItc || 0);
      }
    }

    if (idMap.has(rowId)) {
      // Update existing row
      const targetRow = idMap.get(rowId);
      sheet.getRange(targetRow, 1, 1, Math.min(rowValues.length, colCount)).setValues([rowValues.slice(0, colCount)]);
      upsertCount++;
    } else {
      // Append row
      appendRows.push(rowValues.slice(0, colCount));
      idMap.set(rowId, lastRow + appendRows.length);
      upsertCount++;
    }
  }

  if (appendRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, colCount).setValues(appendRows);
  }
  SpreadsheetApp.flush();

  return jsonResponse({ ok: true, count: upsertCount, sheet: sheetName });
}

function handleDelete(ss, sheetName, ids) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return jsonResponse({ ok: true, count: 0, sheet: sheetName });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1 || !ids || ids.length === 0) {
    return jsonResponse({ ok: true, count: 0, sheet: sheetName });
  }

  const targetIds = new Set(ids.map(x => String(x).trim()));
  const existingValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let deletedCount = 0;

  // Iterate backwards when deleting rows
  for (let i = existingValues.length - 1; i >= 0; i--) {
    const rowId = String(existingValues[i][0] || '').trim();
    if (targetIds.has(rowId)) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }

  return jsonResponse({ ok: true, count: deletedCount, sheet: sheetName });
}

function handleClear(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return jsonResponse({ ok: true, message: 'Sheet does not exist' });
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
  return jsonResponse({ ok: true, message: 'Sheet cleared successfully' });
}

function handleRead(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return jsonResponse({ ok: true, sheet: sheetName, headers: [], rows: [], count: 0 });
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) {
    const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    return jsonResponse({ ok: true, sheet: sheetName, headers, rows: [], count: 0 });
  }

  const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = allData[0];
  const rows = [];

  for (let r = 1; r < allData.length; r++) {
    const row = allData[r];
    const rowId = String(row[0] || '').trim();
    if (rowId) {
      rows.push({
        id: rowId,
        values: row
      });
    }
  }

  return jsonResponse({ ok: true, sheet: sheetName, headers, rows, count: rows.length });
}

/**
 * Step 1: Shifts any row in GSTR-1 starting with g2b- to GSTR-2B tab and deletes it from GSTR-1
 */
function handleShiftG2b(ss) {
  const gstr1 = ss.getSheetByName(CONFIG.TABS.GSTR1);
  if (!gstr1) {
    return jsonResponse({ ok: true, shiftedCount: 0, message: 'GSTR-1 sheet not found' });
  }

  const lastRow = gstr1.getLastRow();
  const lastCol = gstr1.getLastColumn();
  if (lastRow <= 1) {
    return jsonResponse({ ok: true, shiftedCount: 0, message: 'GSTR-1 has no data' });
  }

  const allValues = gstr1.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = allValues[0];
  const g2bSheet = getOrCreateSheet(ss, CONFIG.TABS.GSTR2B, headers);

  const rowsToShift = [];
  const rowsToDeleteIndices = [];

  for (let i = 1; i < allValues.length; i++) {
    const idVal = String(allValues[i][0] || '').trim().toLowerCase();
    if (idVal.startsWith('g2b')) {
      rowsToShift.push(allValues[i]);
      rowsToDeleteIndices.push(i + 1); // 1-based index
    }
  }

  if (rowsToShift.length > 0) {
    // Append to GSTR-2B
    const g2bColCount = Math.min(headers.length, CONFIG.STANDARD_HEADERS.length);
    const cleanedRows = rowsToShift.map(r => r.slice(0, g2bColCount));
    g2bSheet.getRange(g2bSheet.getLastRow() + 1, 1, cleanedRows.length, g2bColCount).setValues(cleanedRows);

    // Delete backwards from GSTR-1
    for (let j = rowsToDeleteIndices.length - 1; j >= 0; j--) {
      gstr1.deleteRow(rowsToDeleteIndices[j]);
    }
  }

  return jsonResponse({
    ok: true,
    shiftedCount: rowsToShift.length,
    message: `${rowsToShift.length} rows shifted from GSTR-1 to GSTR-2B`
  });
}

function handleLock(ss, sheetName, lock) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ ok: true, locked: lock });

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (lock) {
    if (protections.length === 0) {
      const p = sheet.protect().setDescription('Locked by GST Dashboard');
      p.removeEditors(p.getEditors());
    }
  } else {
    for (let i = 0; i < protections.length; i++) {
      protections[i].remove();
    }
  }
  return jsonResponse({ ok: true, locked: lock });
}
