//Schedule.gs

/**
 * Sets up the Schedule sheet with simplified headers, formatting, and sample data.
 * Simplified structure: Time, Duration, Program, Lead/Presenter
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object
 * @param {boolean} addSampleData Whether to add sample data to the sheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The configured Schedule sheet
 */
function setupScheduleSheet(ss, addSampleData = true) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName('Schedule');

  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Schedule');
    sheet.setTabColor('#f1c232'); // Yellow/gold color
  } else {
    // Remove existing filter and clear content if sheet already exists
    const existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    sheet.clear();
  }
  
  // Ensure we have at least 900 rows
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < 900) {
    sheet.insertRowsAfter(currentMaxRows, 900 - currentMaxRows);
  }
  
  // Define simplified headers (4 columns total)
  const headers = ['Time', 'Duration', 'Program', 'Lead/Presenter'];
  
  // Set header values
  const headerRange = sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  headerRange.setFontSize(16);
  
  // Set column widths
  const widths = [100, 100, 300, 200];
  for (let i = 0; i < headers.length; i++) {
    if (i < widths.length) {
      sheet.setColumnWidth(i + 1, widths[i]);
    }
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 900, headers.length).setFontSize(12);
  // Wrap program and lead/presenter for readability
  sheet.getRange(2, 3, 899, 2).setWrap(true);
  
  // Add sample data if requested
  if (addSampleData) {
    const sampleData = [
      ['9:00 AM', '1 hour', 'Opening Session', 'Jane Doe'],
      ['10:30 AM', '1.5 hours', 'Workshop Session', 'John Smith'],
      ['2:00 PM', '45 minutes', 'Panel Discussion', 'Sarah Johnson']
    ];
    
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  }
  
  // Format headers with blue background and white text
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#674ea7') // Brand background
    .setFontColor('#ffffff') // White text
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Set normal white background for the first data row and all other rows
  sheet.getRange(2, 1, 899, headers.length).setBackground(null);
  
  // Apply custom alternating row colors
  for (let i = 2; i <= 900; i += 2) {
    sheet.getRange(i, 1, 1, headers.length).setBackground('#f3f3f3');
  }
  
  // Create filter view for all rows
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(1, 1, 900, headers.length).createFilter();
  
  // Setup time calculation
  setupTimeCalculation(ss);
  
  return sheet;
}

/**
 * Sets up time calculation in the Schedule sheet
 * Time column will be calculated based on Duration column
 */
function setupTimeCalculation(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheet = ss.getSheetByName('Schedule');
  if (!sheet) return;
  
  // Get the last row with data
  const lastRow = Math.max(2, sheet.getLastRow());
  
  // Skip if only header row exists
  if (lastRow <= 1) return;
  
  // Set formula for time calculation in column 1 (Time)
  // Simple approach: first row starts at 9:00 AM, subsequent rows add duration
  for (let row = 2; row <= lastRow; row++) {
    let formula;
    if (row === 2) {
      // First data row starts at 9:00 AM
      formula = '=IF(B2<>"", "9:00 AM", "")';
    } else {
      // Subsequent rows: if previous row was a day separator, start at 9:00 AM
      // Otherwise, add the duration to the previous time
      formula = `=IF(B${row}<>"", 
        IF(ISNUMBER(SEARCH("day",B${row-1})), 
          "9:00 AM", 
          A${row-1} + B${row-1}
        ),
        ""
      )`;
    }
    sheet.getRange(row, 1).setFormula(formula);
  }
}

/**
 * Handles edits specifically for the Schedule sheet
 * Called from the consolidated onEdit function in Core.gs
 * @param {Object} e The edit event object
 */
function handleScheduleEdit(e) {
  // Only proceed if edit is in Schedule sheet
  if (!e || !e.range || e.range.getSheet().getName() !== 'Schedule') return;
  
  // Only proceed if edit is in column B (Duration)
  const col = e.range.getColumn();
  if (col !== 2) return;
  
  // Get the row being edited
  const row = e.range.getRow();
  if (row <= 1) return; // Skip header row
  
  // Get the sheet and values
  const sheet = e.range.getSheet();
  const durationCell = sheet.getRange(row, 2);
  const durationValue = durationCell.getValue();
  
  // Skip if duration is missing
  if (!durationValue) {
    return;
  }
  
  // Check if this is a day separator (contains "day" or "Day")
  if (typeof durationValue === 'string' && 
      (durationValue.toLowerCase().includes('day') || 
       durationValue.toLowerCase().includes('separator'))) {
    // This is a day separator, format it appropriately
    durationCell.setBackground('#e6f3ff');
    durationCell.setFontWeight('bold');
    sheet.getRange(row, 1, 1, 4).setBackground('#e6f3ff');
    return;
  }
  
  // Normal duration entry - recalculate time
  setupTimeCalculation(sheet.getParent());
}

/**
 * Helper function to parse a time string in the format "10:00 AM" or "2:30 PM"
 * @param {string|Date} timeStr - The time string or Date object to parse
 * @return {Date|null} - A Date object set to today with the specified time, or null if parsing fails
 */
function parseTimeString(timeStr) {
  try {
    // Handle Date objects directly
    if (timeStr instanceof Date) {
      return timeStr;
    }
    
    // Make sure we have a string
    timeStr = String(timeStr).trim();
    
    // Regular expression to match time formats with more flexibility
    // This will match "10:00 AM", "2:30PM", "14:00", etc.
    const timeRegex = /^(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)?$/;
    const match = timeStr.match(timeRegex);
    
    if (!match) {
      // Try an alternative format with period separator like "2.30 PM"
      const altTimeRegex = /^(\d{1,2})[\.\s](\d{2})\s*(AM|PM|am|pm)?$/;
      const altMatch = timeStr.match(altTimeRegex);
      
      if (!altMatch) return null;
      
      // Use the alternative match
      let hours = parseInt(altMatch[1], 10);
      const minutes = parseInt(altMatch[2], 10);
      const period = altMatch[3] ? altMatch[3].toUpperCase() : null;
      
      // Convert to 24-hour format if period is specified
      if (period === 'PM' && hours < 12) {
        hours += 12;
      } else if (period === 'AM' && hours === 12) {
        hours = 0;
      }
      
      // Create a Date object for today with the parsed time
      const date = new Date();
      date.setHours(hours, minutes, 0, 0);
      return date;
    }
    
    let hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const period = match[3] ? match[3].toUpperCase() : null;
    
    // If period is specified, convert to 24-hour format
    if (period === 'PM' && hours < 12) {
      hours += 12;
    } else if (period === 'AM' && hours === 12) {
      hours = 0;
    }
    
    // Create a Date object for today with the specified time
    const date = new Date();
    date.setHours(hours, minutes, 0, 0);
    return date;
    
  } catch (error) {
    Logger.log("Error parsing time string: " + error.toString());
    return null;
  }
}

/**
 * Standalone function to update time calculation in Schedule sheet
 * This can be called from the menu to apply time calculation to existing data
 */
function updateScheduleTimes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupTimeCalculation(ss);
  SpreadsheetApp.getUi().alert('Time calculation has been updated in the Schedule sheet.');
}

/**
 * Sets dropdowns for Lead/Presenter in Schedule.
 * Updated to get Lead dropdown from People sheet names.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @param {Object} sheets Cached sheet references
 * @param {Object} rowCounts Cached row counts
 * @param {Object} lists Configuration lists
 * @return {Array} List of updated dropdown fields
 */
function setScheduleDropdowns(ss, sheets, rowCounts, lists) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!sheets) {
    sheets = {
      schedule: ss.getSheetByName('Schedule'),
      people: ss.getSheetByName('People')
    };
  }
  
  const scheduleSheet = sheets.schedule;
  if (!scheduleSheet) return [];
  
  const numRows = rowCounts && rowCounts.schedule 
    ? rowCounts.schedule 
    : Math.max(10, scheduleSheet.getLastRow() - 1);
    
  const updated = [];
  
  // Update Lead/Presenter dropdown using names from the People sheet only
  const peopleSheet = sheets.people;

  if (peopleSheet) {
    const maxRows = peopleSheet.getMaxRows();
    if (maxRows > 1) {
      const peopleRange = peopleSheet.getRange(2, 1, maxRows - 1, 1);
      const leadRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(peopleRange, true)
        .build();
      scheduleSheet.getRange(2, 4, numRows).setDataValidation(leadRule);
      updated.push("Lead/Presenter");
    }
  }
  
  return updated;
}

/**
 * Adds a day separator to the schedule
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @param {string} dayLabel The label for the day (e.g., "Day 2", "Tuesday")
 * @param {number} row The row to insert the separator at
 */
function addDaySeparator(ss, dayLabel, row) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheet = ss.getSheetByName('Schedule');
  if (!sheet) return;
  
  // Insert a new row
  sheet.insertRowBefore(row);
  
  // Set the separator data
  sheet.getRange(row, 1).setValue(''); // Time (will be calculated)
  sheet.getRange(row, 2).setValue(dayLabel); // Duration column used for day label
  sheet.getRange(row, 3).setValue(''); // Program
  sheet.getRange(row, 4).setValue(''); // Lead/Presenter
  
  // Format the separator row
  sheet.getRange(row, 1, 1, 4).setBackground('#e6f3ff');
  sheet.getRange(row, 2).setFontWeight('bold');
  
  // Recalculate times
  setupTimeCalculation(ss);
}

/**
 * Adds menu items for schedule management
 * @param {GoogleAppsScript.UI.Menu} menu The menu to add items to
 */
function addScheduleMenuItems(menu) {
  menu.addSeparator();
  menu.addItem('🕐 Update Time Calculations', 'updateScheduleTimes');
  menu.addItem('📅 Add Day Separator', 'showDaySeparatorDialog');
}

/**
 * Shows a dialog to add a day separator
 */
function showDaySeparatorDialog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Add Day Separator',
    'Enter the day label (e.g., "Day 2", "Tuesday", "Wednesday"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const dayLabel = response.getResponseText().trim();
    if (dayLabel) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Schedule');
      if (sheet) {
        const lastRow = sheet.getLastRow();
        addDaySeparator(ss, dayLabel, lastRow + 1);
        ui.alert('Success', `Day separator "${dayLabel}" has been added to the schedule.`, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Error', 'Please enter a valid day label.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Test function to verify the simplified schedule structure
 * This function can be run to test the new schedule format
 */
function testSimplifiedSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Test 1: Setup the simplified schedule
    const sheet = setupScheduleSheet(ss, false);
    ui.alert('Test 1 Complete', 'Simplified schedule structure created successfully.', ui.ButtonSet.OK);
    
    // Test 2: Add sample data
    const sampleData = [
      ['', '1 hour', 'Opening Session', 'Jane Doe'],
      ['', '45 minutes', 'Workshop Session', 'John Smith'],
      ['', '1.5 hours', 'Panel Discussion', 'Sarah Johnson']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    ui.alert('Test 2 Complete', 'Sample data added successfully.', ui.ButtonSet.OK);
    
    // Test 3: Add day separator
    addDaySeparator(ss, 'Day 2', 6);
    ui.alert('Test 3 Complete', 'Day separator added successfully.', ui.ButtonSet.OK);
    
    // Test 4: Add more data after separator
    const moreData = [
      ['', '30 minutes', 'Break', ''],
      ['', '1 hour', 'Closing Session', 'Mike Wilson']
    ];
    
    sheet.getRange(7, 1, moreData.length, moreData[0].length).setValues(moreData);
    ui.alert('Test 4 Complete', 'Additional data added after separator.', ui.ButtonSet.OK);
    
    // Test 5: Update time calculations
    setupTimeCalculation(ss);
    ui.alert('Test 5 Complete', 'Time calculations updated successfully.', ui.ButtonSet.OK);
    
    ui.alert('All Tests Complete', 'The simplified schedule structure is working correctly!', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Test Failed', 'Error: ' + error.toString(), ui.ButtonSet.OK);
  }
}
