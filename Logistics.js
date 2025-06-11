// Logistics.js - Handles the creation and AI-powered generation of the logistics list.

/**
 * Creates and sets up the "Logistics" sheet.
 */
function setupLogisticsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetName = 'Logistics';

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    ui.alert('Sheet Created', 'The "Logistics" sheet has been created.', ui.ButtonSet.OK);
  } else {
    const response = ui.alert('Reset Sheet?', 'This will clear all data. Continue?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
    sheet.clear();
  }
  
  sheet.getRange('A1:F1').merge().setValue('Generated quantities will be based on the attendance goal from the "Event Description" sheet.')
    .setHorizontalAlignment('center')
    .setFontStyle('italic')
    .setFontColor('#666666')
    .setFontSize(12);

  const headers = ['Item', 'Quantity Needed', 'Related Schedule Item', 'Status', 'Assigned To', 'Notes'];
  const headerRange = sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#674ea7')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16);

  const widths = [250, 100, 180, 120, 150, 300];
  widths.forEach((width, i) => {
    sheet.setColumnWidth(i + 1, width);
  });

  const dataRange = sheet.getRange('A3:F');
  dataRange.setVerticalAlignment('top').setWrap(true).setFontSize(12);
  
  const statusOptions = ['Needed', 'Sourced', 'On-site', 'Returned', 'Cancelled'];
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statusOptions, true).build();
  sheet.getRange(3, 4, sheet.getMaxRows() - 2, 1).setDataValidation(statusRule);
  
  updateLogisticsDropdowns(ss);
  sheet.setFrozenRows(2);
}

/**
 * Updates dropdowns in the Logistics sheet.
 */
function updateLogisticsDropdowns(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const peopleSheet = ss.getSheetByName('People');
  const scheduleSheet = ss.getSheetByName('Schedule');
  const logisticsSheet = ss.getSheetByName('Logistics');
  if (!peopleSheet || !logisticsSheet || !scheduleSheet) return;

  const peopleNames = peopleSheet.getRange('A2:A').getValues().flat().filter(String);
  if (peopleNames.length > 0) {
    const assigneeRule = SpreadsheetApp.newDataValidation().requireValueInList(peopleNames, true).build();
    logisticsSheet.getRange(3, 5, logisticsSheet.getMaxRows() - 2, 1).setDataValidation(assigneeRule);
  }
  
  const sessionTitles = scheduleSheet.getRange('E2:E').getValues().flat().filter(String);
  if (sessionTitles.length > 0) {
      const scheduleItemsWithOptions = ["General Event", ...new Set(sessionTitles)];
      const scheduleRule = SpreadsheetApp.newDataValidation().requireValueInList(scheduleItemsWithOptions, true).build();
      logisticsSheet.getRange(3, 3, logisticsSheet.getMaxRows() - 2, 1).setDataValidation(scheduleRule);
  }
}

/**
 * Shows the custom HTML dialog for selecting which logistics to generate.
 */
function showLogisticsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('LogisticsDialog')
      .setWidth(400)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate AI Logistics List');
}

/**
 * Fetches schedule items to populate the dialog.
 * @return {Array<string>} A list of unique session titles.
 */
function getScheduleItemsForDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName('Schedule');
  if (!scheduleSheet) return [];
  const sessionTitles = scheduleSheet.getRange('E2:E').getValues().flat().filter(String);
  return [...new Set(sessionTitles)];
}


/**
 * Main function to generate a logistics list using AI for selected items.
 * MODIFIED: Now includes notes from the Schedule sheet in the AI prompt for better context.
 * @param {Array<string>} selectedItems The list of schedule items selected by the user.
 */
function generateAILogisticsList(selectedItems) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logisticsSheet = ss.getSheetByName('Logistics');
    if (!logisticsSheet) {
        throw new Error('The "Logistics" sheet has not been created yet.');
    }

    const eventInfo = getEventInformation();
    const apiKey = getOpenAIApiKey();

    if (!eventInfo || !apiKey) {
      throw new Error("Missing Event Info or API Key.");
    }

    const attendanceGoal = eventInfo.attendanceGoal || 0;
    logisticsSheet.getRange('A1').setValue(`Quantities below are based on an estimated attendance of ${attendanceGoal} people.`);

    // --- NEW: Get notes for the selected schedule items ---
    const scheduleSheet = ss.getSheetByName('Schedule');
    const scheduleData = scheduleSheet.getDataRange().getValues();
    const scheduleHeaders = scheduleData.shift();
    const titleIndex = scheduleHeaders.indexOf('Session Title');
    const notesIndex = scheduleHeaders.indexOf('Notes');
    
    let scheduleContext = '';
    selectedItems.forEach(itemTitle => {
      if(itemTitle === 'General Event') {
        scheduleContext += `- For the "General Event" (overall needs like registration, signage, etc.)\n`;
      } else {
        const scheduleRow = scheduleData.find(row => row[titleIndex] === itemTitle);
        scheduleContext += `- For "${itemTitle}"`;
        if (scheduleRow && notesIndex !== -1 && scheduleRow[notesIndex]) {
          scheduleContext += `: Note - ${scheduleRow[notesIndex]}\n`;
        } else {
          scheduleContext += `\n`;
        }
      }
    });
    // --------------------------------------------------------

    const prompt = `Based on the event details below, generate a list of logistical items.
    Event Name: ${eventInfo.eventName}
    Tagline: ${eventInfo.eventTagline || 'N/A'}
    The total expected attendance is ${attendanceGoal}. This is a critical number.

    **VERY IMPORTANT INSTRUCTION**: Use the attendance goal to calculate a final number for the 'quantity'. Do NOT return ratios like '1 per 25 attendees'. For an attendance of 100 and a recommendation of 1 per 25, you MUST return the calculated number 4.

    Success Metrics: ${eventInfo.successMetrics || 'N/A'}
    Event Website: ${eventInfo.eventWebsite || 'N/A'}

    Focus on the logistics for these specific parts of the event:
    ${scheduleContext}
    Do NOT include volunteers, staff, or any people. Only list physical items, equipment, or supplies.
    For each item, specify which part it is for using the "relatedScheduleItem" key. You MUST use one of these exact names: ${selectedItems.join(', ')}.

    Return a JSON object with a single key "logistics" containing an array of objects.
    Each object must have THREE keys: "item" (string), "quantity" (string representing a number), and "relatedScheduleItem" (string).

    Example:
    {
      "logistics": [
        { "item": "Guest Check-in Desks", "quantity": "2", "relatedScheduleItem": "General Event" },
        { "item": "3 Volleyball Nets", "quantity": "1", "relatedScheduleItem": "Volleyball Tournament" }
      ]
    }`;

    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: "gpt-4.1-mini",
      messages: [{ "role": "user", "content": prompt }],
      response_format: { "type": "json_object" }
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      throw new Error(`AI API Error (${responseCode}): ${responseText}`);
    }
    
    const parsedResponse = JSON.parse(responseText);
    const contentString = parsedResponse.choices[0].message.content;
    const logisticsData = JSON.parse(contentString);

    const logisticsList = logisticsData.logistics;

    if (!logisticsList || !Array.isArray(logisticsList) || logisticsList.length === 0) {
      throw new Error("The AI returned an empty or invalid logistics list.");
    }

    const rows = logisticsList.map(d => [
        d.item || 'Unnamed Item', 
        d.quantity || '1', 
        d.relatedScheduleItem || 'General Event',
        'Needed', '', ''
    ]);
    logisticsSheet.getRange(logisticsSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    
    return `Success! Added ${rows.length} new items to the Logistics sheet.`;

  } catch (e) {
    Logger.log(e.toString());
    return `Error: ${e.message}`;
  }
}




