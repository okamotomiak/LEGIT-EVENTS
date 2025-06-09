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
    'Tagline',
    'Start Date (And Time)',
    'End Date (And Time)',
    'Single- or Multi-Day?',
    'Timezone',
    'Event Type',
    'Location',
    'Venue Address',
    'Virtual Link',
    'Theme or Focus',
    'Target Audience',
    'Categories',
    'Short Objectives',
    'Success Metrics',
    'Description & Messaging',
    'Detailed Description',
    'Key Messages',
    'Attendance Goal (#)',
    'Profit Goal ($)',
    'Special Notes',
    'Event Website'
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

/**
 * Opens the Event Setup dialog for interactive event entry.
 */
function showEventSetupDialog() {
  const html = HtmlService.createHtmlOutputFromFile('EventSetupDialog')
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Event Setup');
}

/**
 * Retrieves existing event details from the Event Description sheet.
 * @return {Object} Event details for the setup dialog.
 */
function getEventDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Event Description');
  if (!sheet) {
    sheet = setupEventDescriptionSheet(ss);
  }

  const getVal = label => {
    const row = _findRow(sheet, label);
    return row ? sheet.getRange(row, 2).getValue() : '';
  };

  // Parse date/time values
  const parseDate = value => {
    if (value instanceof Date) {
      return Utilities.formatDate(value, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    }
    return value ? value.toString() : '';
  };

  const parseTime = value => {
    if (value instanceof Date) {
      return Utilities.formatDate(value, ss.getSpreadsheetTimeZone(), 'HH:mm');
    }
    return '';
  };

  const startValue = getVal('Start Date (And Time)');
  const endValue = getVal('End Date (And Time)');

  return {
    eventName: getVal('Event Name'),
    eventTagline: getVal('Tagline'),
    eventDescription: getVal('Description & Messaging'),
    detailedDescription: getVal('Detailed Description'),
    theme: getVal('Theme or Focus'),
    eventDuration: getVal('Single- or Multi-Day?') || 'single',
    timezone: getVal('Timezone'),
    eventType: getVal('Event Type'),
    startDate: parseDate(startValue),
    startTime: parseTime(startValue),
    endDate: parseDate(endValue),
    endTime: parseTime(endValue),
    venueName: getVal('Location'),
    venueAddress: getVal('Venue Address'),
    virtualLink: getVal('Virtual Link'),
    targetAudience: getVal('Target Audience'),
    categories: getVal('Categories'),
    eventGoals: getVal('Short Objectives'),
    successMetrics: getVal('Success Metrics'),
    keyMessages: getVal('Key Messages'),
    expectedAttendees: getVal('Attendance Goal (#)'),
    profitGoal: getVal('Profit Goal ($)'),
    specialNotes: getVal('Special Notes'),
    eventWebsite: getVal('Event Website')
  };
}

/**
 * Saves event details from the setup dialog back to the sheet.
 * @param {Object} details Object with event data from the dialog.
 */
function saveEventDetails(details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Event Description');
  if (!sheet) {
    sheet = setupEventDescriptionSheet(ss);
  }

  const setVal = (label, value) => {
    let row = _findRow(sheet, label);
    if (!row) {
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1).setValue(label);
    }
    sheet.getRange(row, 2).setValue(value);
  };

  const combineDateTime = (date, time) => {
    if (!date) return '';
    const dtString = date + (time ? ' ' + time : '');
    const parsed = new Date(dtString);
    return isNaN(parsed.getTime()) ? dtString : parsed;
  };

  setVal('Event Name', details.eventName);

  setVal('Tagline', details.eventTagline);
  setVal('Description & Messaging', details.eventDescription);
  setVal('Detailed Description', details.detailedDescription);
  setVal('Theme or Focus', details.theme);
  setVal('Single- or Multi-Day?', details.eventDuration);
  setVal('Timezone', details.timezone);
  setVal('Event Type', details.eventType);
  setVal('Start Date (And Time)', combineDateTime(details.startDate, details.startTime));
  if (details.endDate || details.endTime) {
    setVal('End Date (And Time)', combineDateTime(details.endDate || details.startDate, details.endTime));
  }
  setVal('Location', details.venueName || details.virtualLink || '');
  setVal('Venue Address', details.venueAddress);
  setVal('Virtual Link', details.virtualLink);
  setVal('Target Audience', details.targetAudience);
  setVal('Categories', details.categories);
  setVal('Short Objectives', details.eventGoals);
  setVal('Success Metrics', details.successMetrics);
  setVal('Key Messages', details.keyMessages);
  setVal('Attendance Goal (#)', details.expectedAttendees);
  setVal('Profit Goal ($)', details.profitGoal);
  setVal('Special Notes', details.specialNotes);
  setVal('Event Website', details.eventWebsite);

  return true;
}

