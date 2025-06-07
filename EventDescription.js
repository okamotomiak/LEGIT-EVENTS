// EventDescription.js - Provides setup for the Event Description sheet

/**
 * Creates or resets the "Event Description" sheet with default fields.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Optional spreadsheet to operate on.
 */
function setupEventDescriptionSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Event Description');
  if (!sheet) {
    sheet = ss.insertSheet('Event Description');
    sheet.setTabColor('#f9cb9c');
  } else {
    sheet.clear();
  }

  const headers = ['Field', 'Value'];
  const fields = [
    'Event ID',
    'Event Name',
    'Start Date (And Time)',
    'End Date (And Time)',
    'Single- or Multi-Day?',
    'Location',
    'Theme or Focus',
    'Target Audience',
    'Short Objectives',
    'Description & Messaging',
    'Detailed Description',
    'Key Messages',
    'Attendance Goal (#)',
    'Profit Goal ($)'
  ];
  sheet.getRange(1, 1, 1, 2).setValues([headers])
    .setBackground('#674ea7')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  const data = fields.map(f => [f, '']);
  sheet.getRange(2, 1, data.length, 2).setValues(data);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 250);
  sheet.setFrozenRows(1);
  return sheet;
}
