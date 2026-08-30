/**
 * PAFTE Verification — Google Apps Script backend
 *
 * Handles three things from the web page:
 *  - doPost (action=create): appends a new response row
 *  - doPost (action=update): updates an existing row, matched by ID
 *  - doGet: returns all responses as JSON, used by "View Responses"
 *
 * Sheet columns (created by setup()):
 *  A: ID | B: Timestamp | C: Name | D: Email | E: Region | F: Last Updated
 *
 * Setup instructions are in README.md.
 */

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; // skip blank rows
    rows.push({
      id: data[i][0],
      timestamp: data[i][1],
      name: data[i][2],
      email: data[i][3],
      region: data[i][4],
      lastUpdated: data[i][5]
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var action = e.parameter.action || 'create';

  if (action === 'update') {
    var id = e.parameter.id;
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        var row = i + 1;
        sheet.getRange(row, 3).setValue(e.parameter.name || data[i][2]);
        sheet.getRange(row, 4).setValue(e.parameter.email || data[i][3]);
        sheet.getRange(row, 5).setValue(e.parameter.region || data[i][4]);
        sheet.getRange(row, 6).setValue(new Date());
        break;
      }
    }
  } else {
    var id = Utilities.getUuid();
    sheet.appendRow([
      id,
      new Date(),
      e.parameter.name || '',
      e.parameter.email || '',
      e.parameter.region || '',
      ''
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run this once manually (select "setup" in the function dropdown, click Run)
// to create the header row.
function setup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow(['ID', 'Timestamp', 'Name (Surname, Given Name, Middle Name)', 'Email', 'Place of Registration', 'Last Updated']);
}
