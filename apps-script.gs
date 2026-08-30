/**
 * PAFTE Verification — Google Apps Script backend
 * ─────────────────────────────────────────────────
 * This script turns a Google Sheet into the storage/API layer for the
 * PAFTE Verification portal (index.html). It exposes:
 *   - doGet()  → returns all records as JSON
 *   - doPost() → creates a new record (action=create) or updates an
 *                existing one (action=update)
 *
 * SETUP
 * 1. Create (or open) a Google Sheet that will hold responses.
 * 2. Extensions ▸ Apps Script, delete any boilerplate, paste this file in.
 * 3. Deploy ▸ New deployment ▸ type: "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 4. Copy the resulting /exec URL and paste it into SCRIPT_URL in index.html.
 * 5. The first request will auto-create a "Responses" tab with headers
 *    if it doesn't already exist — no manual header setup required.
 */

const SHEET_NAME = 'Responses'; // change if you want a different tab name

const HEADERS = [
  'ID', 'Name', 'Email', 'Region', 'PAFTE OR No.', 'Status',
  'License No.', 'Date Registered', 'Timestamp', 'Last Updated'
];

// Column positions (1-indexed) matching HEADERS above
const COL = {
  ID: 1,
  NAME: 2,
  EMAIL: 3,
  REGION: 4,
  PAFTE_OR_NO: 5,
  STATUS: 6,
  LICENSE_NO: 7,
  DATE_REGISTERED: 8,
  TIMESTAMP: 9,
  LAST_UPDATED: 10
};

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** GET → list all records as JSON */
function doGet(e) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_([]);
  }

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  const records = data
    .filter(r => r.join('').trim() !== '')
    .map(r => ({
      id: r[COL.ID - 1],
      name: r[COL.NAME - 1],
      email: r[COL.EMAIL - 1],
      region: r[COL.REGION - 1],
      pafteOrNo: r[COL.PAFTE_OR_NO - 1],
      status: r[COL.STATUS - 1] || 'Not Registered',
      licenseNo: r[COL.LICENSE_NO - 1],
      dateRegistered: formatValue_(r[COL.DATE_REGISTERED - 1]),
      timestamp: formatValue_(r[COL.TIMESTAMP - 1]),
      lastUpdated: formatValue_(r[COL.LAST_UPDATED - 1])
    }));

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
  const id = Utilities.getUuid();
  const now = new Date();

  sheet.appendRow([
    id,
    params.name || '',
    params.email || '',
    params.region || '',
    params.pafteOrNo || '',
    params.status || 'Not Registered',
    params.licenseNo || '',
    params.dateRegistered || '',
    now,
    '' // Last Updated stays blank until the record is edited
  ]);

  return jsonResponse_({ result: 'success', id: id });
}

function handleUpdate_(sheet, params) {
  const rowIndex = findRowById_(sheet, params.id);
  if (rowIndex === -1) {
    return jsonResponse_({ result: 'error', message: 'Record not found for id ' + params.id });
  }

  const originalTimestamp = sheet.getRange(rowIndex, COL.TIMESTAMP).getValue();
  const now = new Date();

  // Write columns Name → Last Updated in one call (keeps ID and original Timestamp untouched)
  sheet.getRange(rowIndex, COL.NAME, 1, HEADERS.length - 1).setValues([[
    params.name || '',
    params.email || '',
    params.region || '',
    params.pafteOrNo || '',
    params.status || 'Not Registered',
    params.licenseNo || '',
    params.dateRegistered || '',
    originalTimestamp,
    now
  ]]);

  return jsonResponse_({ result: 'success' });
}

function findRowById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, COL.ID, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // +2: header row + 1-indexed
  }
  return -1;
}

function formatValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    // Treat an all-zero "empty" date (from a blank cell coerced to Date) as blank
    if (value.getTime() === 0) return '';
    return value.toISOString();
  }
  return value;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
