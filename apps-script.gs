/**
 * PAFTE Verification — Google Apps Script backend
 *
 * Handles three things from the web page:
 *  - doPost (action=create): appends a new response row
 *  - doPost (action=update): updates an existing row, matched by ID
 *  - doGet: returns all responses as JSON, used by "View Responses" and
 *           the duplicate-name/email check on the record form
 *
 * Columns are found BY HEADER NAME, not fixed position — but on top of
 * that, this version is now SELF-HEALING: every request checks row 1 of
 * the Sheet and automatically creates any missing header (most commonly
 * "PAFTE OR No.") in its canonical position, preserving all existing
 * data. This removes the #1 cause of "one field silently isn't saving":
 * a header in the Sheet that doesn't *exactly* match what the script is
 * looking for (extra space, different capitalization, column deleted,
 * etc). You no longer have to get the Sheet's headers exactly right by
 * hand — the script will fix them for you.
 *
 * Canonical header order (also what a brand-new Sheet gets via setup()):
 *   ID | Timestamp | Name (Surname, Given Name, Middle Name) | Email |
 *   PAFTE OR No. | Place of Registration | Status | License No. |
 *   Date of Registration | Last Updated
 *
 * License No. and Date of Registration are only ever populated when
 * Status is "Registered" — they're cleared whenever a row is set back to
 * "Not Registered". PAFTE OR No. is independent of Status and is always
 * saved as entered.
 *
 * TROUBLESHOOTING "PAFTE OR No. isn't showing up":
 *  1. Open Apps Script editor ▸ select "debugHeaders" in the function
 *     dropdown ▸ Run ▸ View ▸ Logs. This prints exactly what your Sheet's
 *     header row currently contains, column by column.
 *  2. Make sure you redeployed after pasting this in: Deploy ▸ Manage
 *     deployments ▸ pencil icon ▸ Version: New version ▸ Deploy. Saving
 *     the script alone does NOT update the live web app.
 *  3. Run the request once (submit a test entry or click "View
 *     Responses ▸ Refresh") — ensureHeaders_ runs automatically on every
 *     request and will insert "PAFTE OR No." into row 1 if missing.
 *
 * Setup instructions are in README.md.
 */

var SHEET_NAME = 'Responses'; // preferred tab name; falls back to the first sheet if not found

var COLS = {
  id: 'ID',
  timestamp: 'Timestamp',
  name: 'Name (Surname, Given Name, Middle Name)',
  email: 'Email',
  paOrNo: 'PAFTE OR No.',
  region: 'Place of Registration',
  status: 'Status',
  licenseNo: 'License No.',
  dateRegistered: 'Date of Registration',
  lastUpdated: 'Last Updated'
};

// Canonical left-to-right order used when creating headers from scratch
// or inserting a missing one.
var HEADER_ORDER = [
  COLS.id, COLS.timestamp, COLS.name, COLS.email, COLS.paOrNo, COLS.region,
  COLS.status, COLS.licenseNo, COLS.dateRegistered, COLS.lastUpdated
];

function getResponseSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  ensureHeaders_(sheet);
  return sheet;
}

/**
 * Makes sure every header in HEADER_ORDER exists somewhere in row 1.
 * - Brand new / empty sheet → writes the full header row.
 * - Existing sheet missing one or more headers (e.g. "PAFTE OR No.")
 *   → inserts each missing column at its canonical position, shifting
 *   existing columns right and preserving all existing data.
 */
function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ORDER);
    sheet.setFrozenRows(1);
    return;
  }

  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });

  HEADER_ORDER.forEach(function (headerName, idx) {
    if (headerRow.indexOf(headerName) !== -1) return; // already present, exact match

    var insertAt = Math.min(idx + 1, sheet.getLastColumn() + 1);
    sheet.insertColumnBefore(insertAt);
    sheet.getRange(1, insertAt).setValue(headerName);
    headerRow.splice(insertAt - 1, 0, headerName); // keep local copy in sync for later checks
  });
}

// Reads row 1 and returns { "Header Text": 0-based column index }
function getHeaderMap(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    var key = String(h).trim();
    if (key) map[key] = i;
  });
  return map;
}

function getCell(rowValues, map, headerName) {
  var idx = map[headerName];
  return idx === undefined ? '' : rowValues[idx];
}

function setCellByHeader(sheet, row, map, headerName, value) {
  var idx = map[headerName];
  if (idx === undefined) return; // shouldn't happen now that ensureHeaders_ runs first, but stay safe
  sheet.getRange(row, idx + 1).setValue(value);
}

function doGet(e) {
  var sheet = getResponseSheet_();
  var map = getHeaderMap(sheet);
  var data = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var id = getCell(r, map, COLS.id);
    if (!id) continue; // skip blank rows

    rows.push({
      id: id,
      timestamp: getCell(r, map, COLS.timestamp),
      name: getCell(r, map, COLS.name),
      email: getCell(r, map, COLS.email),
      region: getCell(r, map, COLS.region),
      status: getCell(r, map, COLS.status),
      licenseNo: getCell(r, map, COLS.licenseNo),
      paOrNo: getCell(r, map, COLS.paOrNo),
      dateRegistered: getCell(r, map, COLS.dateRegistered),
      lastUpdated: getCell(r, map, COLS.lastUpdated)
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = getResponseSheet_();
  var map = getHeaderMap(sheet);
  var action = e.parameter.action || 'create';
  var status = e.parameter.status || 'Not Registered';
  var isRegistered = status === 'Registered';

  if (action === 'update') {
    var id = e.parameter.id;
    var data = sheet.getDataRange().getValues();
    var idCol = map[COLS.id];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        var row = i + 1;
        setCellByHeader(sheet, row, map, COLS.name, e.parameter.name || data[i][map[COLS.name]]);
        setCellByHeader(sheet, row, map, COLS.email, e.parameter.email || data[i][map[COLS.email]]);
        setCellByHeader(sheet, row, map, COLS.region, e.parameter.region || data[i][map[COLS.region]]);
        setCellByHeader(sheet, row, map, COLS.status, status);
        // License No. and Date of Registration only persist while Registered.
        setCellByHeader(sheet, row, map, COLS.licenseNo, isRegistered ? (e.parameter.licenseNo || '') : '');
        // PAFTE OR No. is independent of Status — always saved as entered.
        setCellByHeader(sheet, row, map, COLS.paOrNo, e.parameter.paOrNo || '');
        setCellByHeader(sheet, row, map, COLS.dateRegistered, isRegistered ? (e.parameter.dateRegistered || '') : '');
        setCellByHeader(sheet, row, map, COLS.lastUpdated, new Date());
        break;
      }
    }
  } else {
    var lastCol = sheet.getLastColumn();
    var newRow = new Array(lastCol).fill('');
    var newId = Utilities.getUuid();

    function put(headerName, value) {
      var idx = map[headerName];
      if (idx !== undefined) newRow[idx] = value;
    }

    put(COLS.id, newId);
    put(COLS.timestamp, new Date());
    put(COLS.name, e.parameter.name || '');
    put(COLS.email, e.parameter.email || '');
    put(COLS.region, e.parameter.region || '');
    put(COLS.status, status);
    put(COLS.licenseNo, isRegistered ? (e.parameter.licenseNo || '') : '');
    put(COLS.paOrNo, e.parameter.paOrNo || '');
    put(COLS.dateRegistered, isRegistered ? (e.parameter.dateRegistered || '') : '');

    sheet.appendRow(newRow);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run this once manually (select "setup" in the function dropdown, click Run)
// to create a header row on a brand-new sheet — safe to skip otherwise, since
// ensureHeaders_ now runs automatically on every doGet/doPost and will add
// any missing headers (including "PAFTE OR No.") on its own.
function setup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ORDER);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaders_(sheet);
  }
}

// Diagnostic helper: select "debugHeaders" in the function dropdown, click
// Run, then View ▸ Logs (or Executions) to see exactly what this sheet's
// header row currently contains and whether "PAFTE OR No." is present.
function debugHeaders() {
  var sheet = getResponseSheet_();
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  Logger.log('Sheet tab name: ' + sheet.getName());
  Logger.log('Column count: ' + lastCol);
  headerRow.forEach(function (h, i) {
    var text = String(h);
    var matchesPaOrNo = text.trim() === COLS.paOrNo;
    Logger.log('Column ' + (i + 1) + ': "' + text + '"' + (matchesPaOrNo ? '  ← PAFTE OR No. ✓' : ''));
  });

  if (headerRow.map(function (h) { return String(h).trim(); }).indexOf(COLS.paOrNo) === -1) {
    Logger.log('⚠ "PAFTE OR No." was NOT found. It will be auto-inserted the next time doGet/doPost runs.');
  }
}
