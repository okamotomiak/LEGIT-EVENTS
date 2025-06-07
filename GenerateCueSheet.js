// GenerateCueSheet.js - Generates a professional, print-ready production cue sheet from the Cue Builder sheet.

/**
 * Main function to generate a professional cue sheet.
 * Reads data from "Cue Builder" and creates a new, formatted sheet.
 */
function generateProfessionalCueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const cueData = getCueBuilderData(ss);
    if (!cueData || cueData.length === 0) {
      ui.alert('No Data', 'The "Cue Builder" sheet is empty. Please add cues before generating the sheet.', ui.ButtonSet.OK);
      return;
    }

    const eventInfo = getEventInformation(); // From TaskManagement.js
    const sheetName = `${eventInfo.eventName || 'Event'} - Cue Sheet`;

    // Create or clear the cue sheet
    let cueSheet = ss.getSheetByName(sheetName);
    if (cueSheet) {
      cueSheet.clear();
    } else {
      cueSheet = ss.insertSheet(sheetName);
    }
    cueSheet.setTabColor('#ff9900');

    // Build the professional, print-ready header
    createPrintReadyHeader(cueSheet, eventInfo);

    // Setup the main table headers below the new header
    setupProfessionalCueSheetHeaders(cueSheet);
    
    // Process data and populate the sheet
    populateProfessionalCueSheet(cueSheet, cueData, ss);

    ui.alert('Success', `Professional cue sheet created! Check the "${sheetName}" tab.`, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error generating cue sheet: ' + error.toString());
    ui.alert('Error', 'Failed to generate cue sheet: ' + error.message, ui.ButtonSet.OK);
  }
}
/**
 * NEW: Creates a professional, print-ready header at the top of the cue sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to add the header to.
 * @param {Object} eventInfo The event details object.
 */
/**
 * NEW: Creates a professional, print-ready header at the top of the cue sheet.
 * MODIFIED: Added fields for MC and Producer.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to add the header to.
 * @param {Object} eventInfo The event details object.
 */
function createPrintReadyHeader(sheet, eventInfo) {
  // Set Column Widths first for proper layout
  const widths = [40, 80, 50, 150, 150, 350, 180, 180, 180];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  // --- Main Title ---
  sheet.getRange('A1:I1').merge().setValue(eventInfo.eventName.toUpperCase() + ' - CUE SHEET')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  // --- Event & Crew Info Section ---
  const eventDate = eventInfo.startDate ? eventInfo.startDate.toLocaleDateString() : 'TBD';
  const generatedTime = new Date().toLocaleString();

  // Create the info table with the new rows
  const infoData = [
    ['Event Date:', eventDate, '', 'Show Caller:', ''],
    ['Venue:', eventInfo.location || 'TBD', '', 'Stage Manager:', ''],
    ['MC:', '', '', 'Producer:', ''],
    ['Audio Lead (A1):', '', '', 'Lighting (LD):', '']
  ];
  // The range is now larger to accommodate the new rows
  const infoRange = sheet.getRange('C3:G6');
  infoRange.setValues(infoData)
    .setFontSize(11).setFontWeight('bold');
    
  // Format the info table
  infoRange.setHorizontalAlignment('left');
  sheet.getRange('C3:C6').setHorizontalAlignment('right'); // Labels on the left
  sheet.getRange('F3:F6').setHorizontalAlignment('right'); // Labels on the right
  
  // Add borders to the crew name fields for writing
  sheet.getRange('D3:D6').setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THIN);
  sheet.getRange('G3:G6').setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THIN);
  
  // Add Generated/Version info on a separate line below
  sheet.getRange('A7').setValue(`Generated: ${generatedTime} | Version: 1.0`).setFontSize(8).setFontStyle('italic');
  sheet.setRowHeight(7, 20);
}


/**
 * Retrieves and processes data from the "Cue Builder" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @return {Array} An array of cue objects.
 */
function getCueBuilderData(ss) {
  const cueBuilderSheet = ss.getSheetByName('Cue Builder');
  if (!cueBuilderSheet) throw new Error('"Cue Builder" sheet not found.');

  const data = cueBuilderSheet.getDataRange().getValues();
  if (data.length < 2) return []; // No data beyond header

  const headers = data[0].map(h => h.toString().toLowerCase());
  const cueObjects = data.slice(1).map((row, index) => ({
    scheduleItem: row[headers.indexOf('schedule item')],
    cueTitle: row[headers.indexOf('cue title')],
    lead: row[headers.indexOf('lead / talent')],
    duration: row[headers.indexOf('est. duration (mins)')] || 0,
    mcScript: row[headers.indexOf('mc script / notes')],
    lightingCue: row[headers.indexOf('lighting cue')],
    audioCue: row[headers.indexOf('audio / sound cue')],
    visualsCue: row[headers.indexOf('visuals / screen cue')],
    originalRow: index + 2
  }));

  return cueObjects.filter(cue => cue.scheduleItem && cue.cueTitle); // Only include valid cues
}

/**
 * Sets up the headers for the professional cue sheet.
 * MODIFIED: Removed the "Section" column, reducing the total column count to 9.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to set up.
 */
function setupProfessionalCueSheetHeaders(sheet) {
  const headers = [
    '#', 'Time', 'Dur.', 'Cue Title', 'Lead / Talent', 
    'MC Script', 'Lighting Cue', 'Audio / Sound Cue', 'Visuals / Screen Cue'
  ];
  const headerRange = sheet.getRange(8, 1, 1, headers.length); // Now 9 columns
  headerRange.setValues([headers])
    .setBackground('#1c4587')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setVerticalAlignment('middle');
    
  // Adjusted widths for 9 columns
  const widths = [40, 80, 50, 200, 150, 350, 180, 180, 180];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  sheet.setFrozenRows(8);
}


/**
 * Populates the professional cue sheet with processed data.
 * MODIFIED: Changed the number format for the Time column to exclude seconds.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} cueSheet The destination cue sheet.
 * @param {Array} cueData The array of cue objects from the Cue Builder.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 */
function populateProfessionalCueSheet(cueSheet, cueData, ss) {
  const scheduleSheet = ss.getSheetByName('Schedule');
  if (!scheduleSheet) throw new Error('"Schedule" sheet not found.');

  const scheduleData = scheduleSheet.getDataRange().getValues();
  const scheduleHeaders = scheduleData[0].map(h => h.toString().toLowerCase());
  const scheduleTitleIndex = scheduleHeaders.indexOf('session title');
  const scheduleTimeIndex = scheduleHeaders.indexOf('start time');

  // Create a map for quick lookup of schedule item start times
  const scheduleTimeMap = new Map();
  scheduleData.slice(1).forEach(row => {
    const title = row[scheduleTitleIndex];
    const time = row[scheduleTimeIndex];
    if (title && time) {
      scheduleTimeMap.set(title, time);
    }
  });

  let runningTime = null;
  let currentScheduleItem = '';
  let cueNumber = 1;
  const outputData = [];

  cueData.forEach(cue => {
    // Check if we are starting a new section
    if (cue.scheduleItem !== currentScheduleItem) {
      currentScheduleItem = cue.scheduleItem;
      const scheduleStartTime = scheduleTimeMap.get(currentScheduleItem);
      if (scheduleStartTime instanceof Date) {
        runningTime = new Date(scheduleStartTime.getTime());
      } else {
        runningTime = null;
      }
      
      const separatorRow = ['', '', '', `--- ${currentScheduleItem.toUpperCase()} ---`, '', '', '', '', ''];
      outputData.push(separatorRow);
    }

    const cueStartTime = runningTime ? new Date(runningTime.getTime()) : null;
    const durationMins = parseFloat(cue.duration) || 0;

    outputData.push([
      cueNumber++,
      cueStartTime,
      durationMins > 0 ? `${durationMins}m` : '',
      cue.cueTitle,
      cue.lead,
      cue.mcScript,
      cue.lightingCue,
      cue.audioCue,
      cue.visualsCue
    ]);

    if (runningTime) {
      runningTime.setMinutes(runningTime.getMinutes() + durationMins);
    }
  });

  if (outputData.length > 0) {
    const START_ROW = 9;
    const NUM_COLUMNS = 9;
    const outputRange = cueSheet.getRange(START_ROW, 1, outputData.length, NUM_COLUMNS);
    outputRange.setValues(outputData);
    
    // --- Enhanced Formatting ---
    outputRange.setFontSize(12).setWrap(true).setVerticalAlignment('top');
    cueSheet.getRange(START_ROW, 1, outputData.length, 1).setHorizontalAlignment('center'); // #
    // --- THIS IS THE FIX ---
    // Changed format from 'h:mm:ss AM/PM' to 'h:mm AM/PM'
    cueSheet.getRange(START_ROW, 2, outputData.length, 1).setNumberFormat('h:mm AM/PM').setHorizontalAlignment('center'); // Time
    cueSheet.getRange(START_ROW, 3, outputData.length, 1).setHorizontalAlignment('center'); // Dur.
    
    // Apply borders and formatting more efficiently
    const borderRows = [];
    for (let i = 0; i < outputData.length; i++) {
        const row = i + START_ROW;
        if(outputData[i][3] && outputData[i][3].startsWith('---')){
            const separatorRange = cueSheet.getRange(row, 1, 1, NUM_COLUMNS);
            separatorRange.merge().setBackground('#d9d9d9').setFontWeight('bold').setHorizontalAlignment('center');
        } else {
            borderRows.push(`A${row}:I${row}`);
        }
    }
    if (borderRows.length) {
        cueSheet.getRangeList(borderRows).setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}