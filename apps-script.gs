/**
 * PAFTE Verification — Google Apps Script backend
 * ─────────────────────────────────────────────────
 * This script turns a Google Sheet into the storage/API layer for the
 * PAFTE Verification portal (index.html). It exposes:
 *   - doGet()  → returns all records as JSON
 *   - doPost() → creates a new record (action=create) or updates an
 *                existing one (action=update)
 *
 * Columns are looked up BY HEADER NAME, not fixed position. This means:
 *   - If you already have a "Responses" sheet from an earlier version of
 *     this script that doesn't have a "PAFTE OR No." column, it will be
 *     inserted automatically (with existing data preserved) the next
 *     time doGet/doPost runs.
 *   - You can reorder columns in the sheet itself without breaking
 *     anything, as long as the header text stays the same.
 *
 * SETUP
 * 1. Create (or open) a Google Sheet that will hold responses.
 * 2. Extensions ▸ Apps Script, delete any boilerplate, paste this file in.
 * 3. Deploy ▸ New deployment ▸ type: "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 4. Copy the resulting /exec URL and paste it into SCRIPT_URL in index.html.
 * 5. IMPORTANT if you had a previous deployment: create a NEW deployment
 *    (or "Manage deployments" ▸ edit ▸ New version) after pasting this
 *    code in, otherwise the web app keeps serving the old script.
 */

const SHEET_NAME = 'Responses'; // change if you want a different tab name

const HEADERS = [
  'ID', 'Name', 'Email', 'Region', 'PAFTE OR No.', 'Status',
  'License No.', 'Date Registered', 'Timestamp', 'Last Updated'
];

// Maps the JSON field names used by index.html to the sheet's header text
const FIELD_TO_HEADER = {
  id: 'ID',
  name: 'Name',
  email: 'Email',
  region: 'Region',
  pafteOrNo: 'PAFTE OR No.',
  status: 'Status',
  licenseNo: 'License No.',
  dateRegistered: 'Date Registered',
  timestamp: 'Timestamp',
  lastUpdated: 'Last Updated'
};

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  ensureHeaders_(sheet);
  return sheet;
}

/**
 * Makes sure the sheet has all HEADERS present. If the sheet is brand
 * new, writes the full header row. If it already has data but is
 * missing a header (e.g. an older sheet without "PAFTE OR No."), that
 * column is inserted in place, preserving all existing rows/values.
 */
function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  HEADERS.forEach((headerName, idx) => {
    if (headerRow.indexOf(headerName) !== -1) return; // already present

    // Insert this missing column at its canonical position (idx+1),
    // clamped to the sheet's current width so we don't leave gaps.
    const insertAt = Math.min(idx + 1, sheet.getLastColumn() + 1);
    sheet.insertColumnBefore(insertAt);
    sheet.getRange(1, insertAt).setValue(headerName);
    headerRow.splice(insertAt - 1, 0, headerName); // keep local copy in sync
  });
}

/** Returns { 'Header Text': columnNumber, ... } for the sheet's current header row */
function getColumnMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headerRow.forEach((h, i) => {
    if (h) map[h] = i + 1; // 1-indexed column number
  });
  return map;
}

/** GET → list all records as JSON */
function doGet(e) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_([]);
  }

  const colMap = getColumnMap_(sheet);
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const records = data
    .filter(r => r.join('').trim() !== '')
    .map(r => {
      const rec = {};
      Object.keys(FIELD_TO_HEADER).forEach(field => {
        const header = FIELD_TO_HEADER[field];
        const col = colMap[header];
        rec[field] = col ? formatValue_(r[col - 1]) : '';
      });
      if (!rec.status) rec.status = 'Not Registered';
      return rec;
    });

  return jsonResponse_(records);
}

/** POST → create or update a record, depending on action */
function doPost(e) {
  const sheet = getSheet_();
  const params = (e && e.parameter) || {};
  const action = params.action;

  if (action === 'create') {
    return handleCreate_(sheet, params);
  }
  if (action === 'update') {
    return handleUpdate_(sheet, params);
  }
  return jsonResponse_({ result: 'error', message: 'Unknown action: ' + action });
}

function handleCreate_(sheet, params) {
  const colMap = getColumnMap_(sheet);
  const id = Utilities.getUuid();
  const now = new Date();

  const values = {
    id: id,
    name: params.name || '',
    email: params.email || '',
    region: params.region || '',
    pafteOrNo: params.pafteOrNo || '',
    status: params.status || 'Not Registered',
    licenseNo: params.licenseNo || '',
    dateRegistered: params.dateRegistered || '',
    timestamp: now,
    lastUpdated: '' // stays blank until the record is edited
  };

  writeRow_(sheet, colMap, sheet.getLastRow() + 1, values);

  return jsonResponse_({ result: 'success', id: id });
}

function handleUpdate_(sheet, params) {
  const colMap = getColumnMap_(sheet);
  const rowIndex = findRowById_(sheet, colMap, params.id);
  if (rowIndex === -1) {
    return jsonResponse_({ result: 'error', message: 'Record not found for id ' + params.id });
  }

  const timestampCol = colMap['Timestamp'];
  const originalTimestamp = timestampCol ? sheet.getRange(rowIndex, timestampCol).getValue() : '';
  const now = new Date();

  const values = {
    name: params.name || '',
    email: params.email || '',
    region: params.region || '',
    pafteOrNo: params.pafteOrNo || '',
    status: params.status || 'Not Registered',
    licenseNo: params.licenseNo || '',
    dateRegistered: params.dateRegistered || '',
    timestamp: originalTimestamp, // never overwrite the original submission time
    lastUpdated: now
  };

  writeRow_(sheet, colMap, rowIndex, values);

  return jsonResponse_({ result: 'success' });
}

/** Writes a set of {field: value} pairs into the given row, one cell at a time, by header name. */
function writeRow_(sheet, colMap, rowIndex, fieldValues) {
  Object.keys(fieldValues).forEach(field => {
    const header = FIELD_TO_HEADER[field];
    const col = colMap[header];
    if (!col) return; // header not present on this sheet — skip rather than error
    sheet.getRange(rowIndex, col).setValue(fieldValues[field]);
  });
}

function findRowById_(sheet, colMap, id) {
  const idCol = colMap['ID'];
  const lastRow = sheet.getLastRow();
  if (!idCol || lastRow < 2) return -1;
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // +2: header row + 1-indexed
  }
  return -1;
}

function formatValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (value.getTime() === 0) return ''; // an "empty" date coerced from a blank cell
    return value.toISOString();
  }
  return value;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
