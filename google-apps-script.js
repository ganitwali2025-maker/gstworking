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
  GSTR1_HEADERS: [
    'ID', 'Financial Year', 'Quarter', 'Month', 'Invoice Date',
    'Supplier / Party Name', 'GST No', 'Invoice No',
    'Taxable Value', 'IGST', 'CGST', 'SGST', 'Total Tax', 'Total Invoice Value'
  ],
  ITC_PURCHASE_HEADERS: [
    'ID', 'Financial Year', 'Quarter', 'Month', 'Invoice Date',
    'Supplier / Party Name', 'GST No', 'Invoice No',
    'Taxable Value', 'IGST', 'CGST', 'SGST', 'CESS', 'Total Tax', 'Total Invoice Value'
  ],
  STANDARD_HEADERS: [
    'ID', 'Financial Year', 'Quarter', 'Month', 'Invoice Date',
    'Supplier / Party Name', 'GST No', 'Invoice No', 'Invoice Value',
    'Taxable Value', 'IGST', 'CGST', 'SGST', 'CESS', 'Total Tax'
  ],
  RCM_HEADERS: [
    'ID', 'Financial Year', 'Quarter', 'Month', 'Entry Date',
    'Transporter Name', 'GST No', 'Transporter L.R. No.',
    'Amount', 'IGST 5%', 'CGST 2.5%', 'SGST 2.5%', 'Total Tax',
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
  let defaultHeaders = CONFIG.STANDARD_HEADERS;
  if (sheetName === CONFIG.TABS.GSTR1) {
    defaultHeaders = CONFIG.GSTR1_HEADERS;
  } else if (sheetName === CONFIG.TABS.ITC_PURCHASE) {
    defaultHeaders = CONFIG.ITC_PURCHASE_HEADERS;
  } else if (sheetName === CONFIG.TABS.RCM) {
    defaultHeaders = CONFIG.RCM_HEADERS;
  }
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

  let defaultHeaders = CONFIG.STANDARD_HEADERS;
  if (sheetName === CONFIG.TABS.GSTR1) {
    defaultHeaders = CONFIG.GSTR1_HEADERS;
  } else if (sheetName === CONFIG.TABS.ITC_PURCHASE) {
    defaultHeaders = CONFIG.ITC_PURCHASE_HEADERS;
  } else if (sheetName === CONFIG.TABS.RCM) {
    defaultHeaders = CONFIG.RCM_HEADERS;
  }
  const colCount = (headers && headers.length > 0) ? headers.length : defaultHeaders.length;

  // Ensure GSTR-1 header row is set to exact 14 columns matching user layout
  if (sheetName === CONFIG.TABS.GSTR1) {
    sheet.getRange(1, 1, 1, CONFIG.GSTR1_HEADERS.length).setValues([CONFIG.GSTR1_HEADERS]);
    formatHeaderRow(sheet, CONFIG.GSTR1_HEADERS.length);
    if (sheet.getLastColumn() > 14 && sheet.getLastRow() > 0) {
      sheet.getRange(1, 15, sheet.getLastRow(), sheet.getLastColumn() - 14).clearContent();
    }
  } else if (sheetName === CONFIG.TABS.ITC_PURCHASE) {
    sheet.getRange(1, 1, 1, CONFIG.ITC_PURCHASE_HEADERS.length).setValues([CONFIG.ITC_PURCHASE_HEADERS]);
    formatHeaderRow(sheet, CONFIG.ITC_PURCHASE_HEADERS.length);
  } else if (sheetName === CONFIG.TABS.RCM) {
    sheet.getRange(1, 1, 1, CONFIG.RCM_HEADERS.length).setValues([CONFIG.RCM_HEADERS]);
    formatHeaderRow(sheet, CONFIG.RCM_HEADERS.length);
  }

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

      if (sheetName === CONFIG.TABS.GSTR1) {
        // Col I: Taxable, Col J: IGST, Col K: CGST, Col L: SGST, Col M: Total Tax, Col N: Total Invoice Value
        const tVal = Number(r.taxableValue || r.taxable || r.t || 0) || 0;
        const iVal = Number(r.igst || r.i || 0) || 0;
        const cVal = Number(r.cgst || r.c || 0) || 0;
        const sVal = Number(r.sgst || r.s || 0) || 0;
        const totTax = Number(r.totalTax) || (iVal + cVal + sVal);
        const totInvVal = Number(r.totalInvoiceValue || r.invVal) || (tVal + totTax);

        rowValues = [
          rowId,
          r.financialYear || r.fy || '',
          r.quarter || r.q || '',
          r.month || r.m || '',
          r.invoiceDate || r.entryDate || r.d || '',
          cleanParty,
          rawGst,
          r.invoiceNo || r.invNo || r.n || '',
          tVal,
          iVal,
          cVal,
          sVal,
          totTax,
          totInvVal
        ];
      } else if (sheetName === CONFIG.TABS.ITC_PURCHASE) {
        // Purchase: Col I: Taxable, Col J: IGST, Col K: CGST, Col L: SGST, Col M: CESS, Col N: Total Tax, Col O: Total Invoice Value
        const tVal = Number(r.taxableValue || r.taxable || r.t || 0) || 0;
        const iVal = Number(r.igst || r.i || 0) || 0;
        const cVal = Number(r.cgst || r.c || 0) || 0;
        const sVal = Number(r.sgst || r.s || 0) || 0;
        const cessVal = Number(r.cess) || 0;
        const totTax = Number(r.totalTax) || (iVal + cVal + sVal + cessVal);
        const totInvVal = Number(r.totalInvoiceValue || r.invVal) || (tVal + totTax);

        rowValues = [
          rowId,
          r.financialYear || r.fy || '',
          r.quarter || r.q || '',
          r.month || r.m || '',
          r.invoiceDate || r.entryDate || r.d || '',
          cleanParty,
          rawGst,
          r.invoiceNo || r.invNo || r.n || '',
          tVal,
          iVal,
          cVal,
          sVal,
          cessVal,
          totTax,
          totInvVal
        ];
      } else if (sheetName === CONFIG.TABS.RCM) {
        // RCM: Entry Date, Transporter Name, GST No, Transporter L.R. No., Amount, IGST 5%, CGST 2.5%, SGST 2.5%, Total Tax
        const amt = Number(r.amount || r.taxableValue || r.taxable || r.t || 0) || 0;
        const iVal = Number(r.igst || r.i || 0) || 0;
        const cVal = Number(r.cgst || r.c || 0) || 0;
        const sVal = Number(r.sgst || r.s || 0) || 0;
        const totTax = Number(r.totalTax) || (iVal + cVal + sVal);

        rowValues = [
          rowId,
          r.financialYear || r.fy || '',
          r.quarter || r.q || '',
          r.month || r.m || '',
          r.entryDate || r.invoiceDate || r.d || '',
          cleanParty,
          rawGst,
          r.lrNo || r.invoiceNo || r.invNo || r.n || '',
          amt,
          iVal,
          cVal,
          sVal,
          totTax,
          r.rcmCategory || r.category || 'Transport (GTA)',
          r.rcmTaxRate || r.taxRate || 5,
          r.rcmApplicable || 'Yes',
          r.eligibleItc || totTax
        ];
      } else {
        // GSTR-2B (or fallback standard tab) - 100% UNTOUCHED!
        rowValues = [
          rowId,
          r.financialYear || r.fy || '',
          r.quarter || r.q || '',
          r.month || r.m || '',
          r.invoiceDate || r.entryDate || r.d || '',
          cleanParty,
          rawGst,
          r.invoiceNo || r.invNo || r.lrNo || r.n || '',
          r.totalInvoiceValue || r.invVal || 0,
          r.taxableValue || r.taxable || r.t || r.amount || 0,
          r.igst || r.i || 0,
          r.cgst || r.c || 0,
          r.sgst || r.s || 0,
          r.cess || 0,
          r.totalTax || 0
        ];
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
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0 || !ids || ids.length === 0) {
    return jsonResponse({ ok: true, count: 0, sheet: sheetName });
  }

  const targetIds = new Set(ids.map(x => String(x).trim()));
  // Batch read all data rows at once
  const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const keptRows = [];
  let deletedCount = 0;

  for (let i = 1; i < allData.length; i++) {
    const rowId = String(allData[i][0] || '').trim();
    if (targetIds.has(rowId)) {
      deletedCount++;
    } else {
      keptRows.push(allData[i]);
    }
  }

  // Clear data rows (row 2 downwards) in one call
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  // If there are remaining rows, write them back in one batch call
  if (keptRows.length > 0) {
    sheet.getRange(2, 1, keptRows.length, lastCol).setValues(keptRows);
  }

  SpreadsheetApp.flush();
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
