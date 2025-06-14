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

  // Set up headers
  const headers = ['Field', 'Value'];
  const headerRange = sheet.getRange(1, 1, 1, 2).setValues([headers])
    .setBackground('#674ea7')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16);

  // Define field sections with their colors and fields
  const fieldSections = [
    {
      color: '#674ea7', // Purple for basic event info
      fields: ['Event ID', 'Event Name', 'Tagline']
    },
    {
      color: '#93c47d', // Green for dates/timing
      fields: ['Start Date (And Time)', 'End Date (And Time)', 'Single- or Multi-Day?', 'Timezone']
    },
    {
      color: '#f1c232', // Orange/yellow for location
      fields: ['Event Type', 'Location', 'Venue Address', 'Virtual Link']
    },
    {
      color: '#6fa8dc', // Blue for theme/categories
      fields: ['Theme or Focus', 'Categories']
    },
    {
      color: '#c27ba0', // Pink/rose for audience/objectives
      fields: ['Target Audience', 'Short Objectives (How do you want the audience to feel, learn, and do)']
    },
    {
      color: '#6fa8dc', // Blue for success metrics
      fields: ['Success Metrics', 'Attendance Goal (#)', 'Profit Goal ($)']
    },
    {
      color: '#6fa8dc', // Blue for description
      fields: ['Description & Messaging', 'Special Notes']
    },
    {
      color: '#6fa8dc', // Blue for website
      fields: ['Event Website']
    }
  ];

  let currentRow = 2;

  // Create each section
  fieldSections.forEach(section => {
    const sectionData = section.fields.map(field => [field, '']);
    const sectionRange = sheet.getRange(currentRow, 1, section.fields.length, 2);

    // Set the data
    sectionRange.setValues(sectionData);

    // Apply section styling
    sectionRange.setFontSize(12);

    // Style the field names (column A) with the section color
    const fieldRange = sheet.getRange(currentRow, 1, section.fields.length, 1);
    fieldRange.setBackground(section.color)
             .setFontColor('#ffffff')
             .setFontWeight('bold');

    // Style the value column (column B) with light background
    const valueRange = sheet.getRange(currentRow, 2, section.fields.length, 1);
    const lightColor = section.color === '#674ea7' ? '#d9d2e9' : 
                      section.color === '#93c47d' ? '#d9ead3' :
                      section.color === '#f1c232' ? '#fff2cc' :
                      section.color === '#c27ba0' ? '#ead1dc' : '#cfe2f3';

    valueRange.setBackground(lightColor)
             .setWrap(true);

    currentRow += section.fields.length;
  });

  // Set column widths
  sheet.setColumnWidth(1, 300); // Wider for field names
  sheet.setColumnWidth(2, 400); // Wider for values
  
  // Freeze the header row
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
    eventGoals: getVal('Short Objectives (How do you want the audience to feel, learn, and do)'),
    successMetrics: getVal('Success Metrics'),
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
  setVal('Short Objectives (How do you want the audience to feel, learn, and do)', details.eventGoals);
  setVal('Success Metrics', details.successMetrics);
  setVal('Attendance Goal (#)', details.expectedAttendees);
  setVal('Profit Goal ($)', details.profitGoal);
  setVal('Special Notes', details.specialNotes);
  setVal('Event Website', details.eventWebsite);

  return true;
}
