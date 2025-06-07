//Budget.gs
/**
 * Sets up or resets the Budget sheet with the Pre-Event Budget structure
 * All data starts at zero and Other/Miscellaneous section corrected
 * Header fields and Attendees Goal removed as requested
 */
function setupBudgetSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = ss.getSheetByName('Budget');
  
  if (!budgetSheet) {
    SpreadsheetApp.getUi().alert('Budget sheet not found. Please create the template first.');
    return;
  }
  
  // STEP 1: Aggressively remove all data validations first
  try {
    const entireSheet = budgetSheet.getDataRange();
    entireSheet.setDataValidation(null);
    
    const unitFeeColumn = budgetSheet.getRange("C:C");
    unitFeeColumn.setDataValidation(null);
  } catch (e) {
    Logger.log("Error clearing validations: " + e.toString());
  }
  
  // STEP 2: Clear existing content
  budgetSheet.clear();
  
  // Set headers (Row 1)
  const headers = ["Category", "Item", "$Unit Fee", "Quantity", "Sub Total $", "Total $"];
  budgetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  budgetSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Now populate the budget structure with CORRECT formulas for each row
  // All header fields and Attendees Goal removed as requested
  // Revenue is now the first row of the table
  const budgetData = [
    ["Revenue", "", "", "", "", ""],                // Row 2 (first row after header)
    ["⮞ Registration Fees", "", "", "", "", ""],    // Row 3
    ["", "Regular Fee", 0, 0, "=C4*D4", ""],        // Row 4
    ["", "Early Bird Discount", 0, 0, "=C5*D5", ""], // Row 5
    ["", "Staff", 0, 0, "=C6*D6", ""],              // Row 6
    ["", "Registration Total", "", "", "", "=SUM(E4:E6)"], // Row 7
    ["", "", "", "", "", ""],                       // Row 8 (blank)
    ["⮞ Donations", "", "", "", "", ""],            // Row 9
    ["", "Org Donations", 0, 0, "=C10*D10", ""],    // Row 10
    ["", "Individual Donations", 0, 0, "=C11*D11", ""], // Row 11
    ["", "Donation Total", "", "", "", "=SUM(E10:E11)"], // Row 12
    ["", "", "", "", "", ""],                       // Row 13 (blank)
    ["⮞ Sales & Other", "", "", "", "", ""],        // Row 14
    ["", "Drinks", 0, 0, "=C15*D15", ""],           // Row 15
    ["", "Vendors", 0, 0, "=C16*D16", ""],          // Row 16
    ["", "Stocks", 0, 0, "=C17*D17", ""],           // Row 17
    ["", "Sales Total", "", "", "", "=SUM(E15:E17)"], // Row 18
    ["", "", "", "", "", ""],                       // Row 19 (blank)
    ["", "Total Revenue", "", "", "", "=SUM(F7,F12,F18)"], // Row 20
    ["", "", "", "", "", ""],                       // Row 21 (blank)
    ["Expenses", "", "", "", "", ""],               // Row 22
    ["⮞ Venue(s) & Production", "", "", "", "", ""], // Row 23
    ["", "Zoom", 0, 0, "=C24*D24", ""],             // Row 24
    ["", "Audio/Visual", 0, 0, "=C25*D25", ""],     // Row 25
    ["", "Venue", 0, 0, "=C26*D26", ""],            // Row 26
    ["", "Translator", 0, 0, "=C27*D27", ""],       // Row 27
    ["", "Venue Total", "", "", "", "=SUM(E24:E27)"], // Row 28
    ["", "", "", "", "", ""],                       // Row 29 (blank)
    ["⮞ Program", "", "", "", "", ""],              // Row 30
    ["", "Production", 0, 0, "=C31*D31", ""],       // Row 31
    ["", "Guest Speakers", 0, 0, "=C32*D32", ""],   // Row 32
    ["", "Performers", 0, 0, "=C33*D33", ""],       // Row 33
    ["", "Prizes", 0, 0, "=C34*D34", ""],           // Row 34
    ["", "Program Total", "", "", "", "=SUM(E31:E34)"], // Row 35
    ["", "", "", "", "", ""],                       // Row 36 (blank)
    ["⮞ Food", "", "", "", "", ""],                 // Row 37
    ["", "Meals", 0, 0, "=C38*D38", ""],            // Row 38
    ["", "Refreshments", 0, 0, "=C39*D39", ""],     // Row 39
    ["", "Food Total", "", "", "", "=SUM(E38:E39)"], // Row 40
    ["", "", "", "", "", ""],                       // Row 41 (blank)
    ["⮞ Lodging", "", "", "", "", ""],              // Row 42
    ["", "Hotel Rooms", 0, 0, "=C43*D43", ""],      // Row 43
    ["", "Lodging Total", "", "", "", "=E43"],      // Row 44
    ["", "", "", "", "", ""],                       // Row 45 (blank)
    ["⮞ Staff", "", "", "", "", ""],                // Row 46
    ["", "Paid Staff", 0, 0, "=C47*D47", ""],       // Row 47
    ["", "Staff Meals", 0, 0, "=C48*D48", ""],      // Row 48
    ["", "Staff Total", "", "", "", "=SUM(E47:E48)"], // Row 49
    ["", "", "", "", "", ""],                       // Row 50 (blank)
    ["⮞ Transportation", "", "", "", "", ""],       // Row 51
    ["", "Flights", 0, 0, "=C52*D52", ""],          // Row 52
    ["", "Bus", 0, 0, "=C53*D53", ""],              // Row 53
    ["", "Rental Car", 0, 0, "=C54*D54", ""],       // Row 54
    ["", "Transportation Total", "", "", "", "=SUM(E52:E54)"], // Row 55
    ["", "", "", "", "", ""],                       // Row 56 (blank)
    ["⮞ Other/Miscellaneous", "", 0, 0, "=C57*D57", ""], // Row 57
    ["", "Total Expenses", "", "", "", "=SUM(F28,F35,F40,F44,F49,F55,E57)"], // Row 58
    ["", "", "", "", "", ""],                       // Row 59 (blank)
    ["Net Balance", "", "", "", "", ""],            // Row 60
    ["", "Net Profit (Loss)", "", "", "", "=F20-F58"], // Row 61
    ["", "% Profit (Loss)", "", "", "", "=F61/F20"], // Row 62
  ];
  
  // Set values
  budgetSheet.getRange(2, 1, budgetData.length, 6).setValues(budgetData);
  
  // Format currency columns
  budgetSheet.getRange(4, 3, 60, 1).setNumberFormat("$#,##0.00");  // $Unit Fee (starting from Row 4)
  budgetSheet.getRange(4, 5, 60, 1).setNumberFormat("$#,##0.00");  // Sub Total (starting from Row 4)
  budgetSheet.getRange(4, 6, 60, 1).setNumberFormat("$#,##0.00");  // Total (starting from Row 4)
  
  // Format percentage (Row 62)
  budgetSheet.getRange("F62").setNumberFormat("0.00%");
  
  // Format category headers
  const categoryRows = [2, 3, 9, 14, 22, 23, 30, 37, 42, 46, 51, 57, 60];
  for (let row of categoryRows) {
    budgetSheet.getRange(row, 1).setFontWeight("bold");
  }
  
  // Format total rows
  const totalRows = [7, 12, 18, 20, 28, 35, 40, 44, 49, 55, 58, 61, 62];
  for (let row of totalRows) {
    budgetSheet.getRange(row, 2, 1, 5).setFontWeight("bold");
  }
  
  // Adjust column widths
  budgetSheet.setColumnWidth(1, 200);  // Category
  budgetSheet.setColumnWidth(2, 200);  // Item
  budgetSheet.setColumnWidth(3, 100);  // $Unit Fee
  budgetSheet.setColumnWidth(4, 100);  // Quantity
  budgetSheet.setColumnWidth(5, 120);  // Sub Total $
  budgetSheet.setColumnWidth(6, 120);  // Total $
  
  // STEP 3: Apply the new formatting as requested
  
  // Highlight specific rows with blue background and white text
  const blueRows = [2, 20, 22, 58, 60]; // Revenue (2), Total Revenue (20), Expenses (22), Total Expenses (58), Net Balance (60)
  for (let row of blueRows) {
    budgetSheet.getRange(row, 1, 1, 6).setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');
  }
  
  // Apply mint green color (#b7e1cd) to rows with "⮞" symbol, etc.
  const mintGreenRows = [
    3, 9, 14, 23, 30, 37, 42, 46, 51, // Rows with "⮞" symbol
    55, // Transportation Total (row 55)
    57, // Other/Miscellaneous (row 57)
    61, 62 // Net Profit (Loss) (row 61), % Profit (Loss) (row 62)
  ];
  
  // Apply the mint green color (#b7e1cd) to the mint green rows
  for (let row of mintGreenRows) {
    budgetSheet.getRange(row, 1, 1, 6).setBackground('#b7e1cd'); // Mint green color
  }
  
  // Light blue sections for user input (using correct row numbers)
  budgetSheet.getRange("C4:D6").setBackground('#d0e0ff');    // Registration (rows 4-6)
  budgetSheet.getRange("C10:D11").setBackground('#d0e0ff');  // Donations (rows 10-11)
  budgetSheet.getRange("C15:D17").setBackground('#d0e0ff');  // Sales (rows 15-17)
  budgetSheet.getRange("C24:D27").setBackground('#d0e0ff');  // Venue (rows 24-27)
  budgetSheet.getRange("C31:D34").setBackground('#d0e0ff');  // Program (rows 31-34)
  budgetSheet.getRange("C38:D39").setBackground('#d0e0ff');  // Food (rows 38-39)
  budgetSheet.getRange("C43:D43").setBackground('#d0e0ff');  // Lodging (row 43)
  budgetSheet.getRange("C47:D48").setBackground('#d0e0ff');  // Staff (rows 47-48)
  budgetSheet.getRange("C52:D54").setBackground('#d0e0ff');  // Transportation (rows 52-54)
  budgetSheet.getRange("C57:D57").setBackground('#d0e0ff');  // Misc (row 57)
  
  // Freeze the header row
  budgetSheet.setFrozenRows(1);
  
  // One final check to ensure no data validations exist in Column C
  try {
    const finalUnitFeeColumn = budgetSheet.getRange("C:C");
    finalUnitFeeColumn.setDataValidation(null);
    
    // Double-check our Unit Fee column currency formatting
    budgetSheet.getRange(2, 3, budgetData.length, 1).setNumberFormat("$#,##0.00");
  } catch (e) {
    Logger.log("Error in final validation cleanup: " + e.toString());
  }
  
  SpreadsheetApp.getUi().alert('Budget sheet has been updated with the Pre-Event Budget structure and formatting!');
}

/**
 * Generates an estimated budget using OpenAI based on event and logistics data.
 */
function generateAIBudget() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Retrieve event information from TaskManagement.js
    const eventInfo = getEventInformation();
    if (!eventInfo) {
      ui.alert('Error', 'Could not retrieve event information.', ui.ButtonSet.OK);
      return;
    }

    // Retrieve API key
    const apiKey = getOpenAIApiKey();
    if (!apiKey) {
      ui.alert('Error', 'OpenAI API key not found. Please add it to the Config sheet.', ui.ButtonSet.OK);
      return;
    }

    // Build list of logistics items marked as "Needed"
    const logisticsSheet = ss.getSheetByName('Logistics');
    const neededItems = [];
    if (logisticsSheet) {
      const allData = logisticsSheet.getDataRange().getValues();
      if (allData.length > 2) {
        const headers = allData[1];
        const itemCol = headers.indexOf('Item');
        const qtyCol = headers.indexOf('Quantity Needed');
        const statusCol = headers.indexOf('Status');

        for (let i = 2; i < allData.length; i++) {
          const row = allData[i];
          const status = row[statusCol];
          if (status && status.toString().toLowerCase() === 'needed') {
            neededItems.push({
              item: row[itemCol],
              quantity: row[qtyCol]
            });
          }
        }
      }
    }

    // Format event dates
    const startDate = formatDate(eventInfo.startDate);
    const endDate = formatDate(eventInfo.endDate);

    // Prepare logistics text for the prompt
    const logisticsText = neededItems.map(it => `- ${it.item} (${it.quantity})`).join('\n');

    // Construct OpenAI prompt with explicit income/expense guidance
    const prompt =
      `Create a detailed budget for the following event that includes both income and expense categories.\n\n` +
      `Event Name: ${eventInfo.eventName}\n` +
      `Dates: ${startDate}${eventInfo.endDate ? ' to ' + endDate : ''}\n` +
      `Location: ${eventInfo.location || 'TBD'}\n` +
      `Attendance Goal: ${eventInfo.attendanceGoal}\n\n` +
      `Needed Logistics Items:\n${logisticsText}\n\n` +
      `List any questions about missing costs or assumptions, such as registration fees per person.\n` +
      `Respond with a JSON object {"budget": [ {"category":"Income or Expense","item":"","unitPrice":0,"quantity":0} ], "questions": [] }`;

    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: 'gpt-4o',
      messages: [{ role: 'user', content: prompt }],
      response_format: { type: 'json_object' }
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`OpenAI API Error (${response.getResponseCode()}): ${response.getContentText()}`);
    }

    const parsed = JSON.parse(response.getContentText());
    const content = JSON.parse(parsed.choices[0].message.content);
    const items = content.budget || [];
    const questions = Array.isArray(content.questions) ? content.questions : [];

    if (!Array.isArray(items) || items.length === 0) {
      throw new Error('No budget data returned from OpenAI.');
    }

    // Reset sheet to default template with formulas and formatting
    setupBudgetSheet();
    const budgetSheet = ss.getSheetByName('Budget');

    // Map of known items to their row numbers in the template
    const rowMap = {
      'regular fee': 4,
      'early bird discount': 5,
      'staff': 6,
      'org donations': 10,
      'individual donations': 11,
      'drinks': 15,
      'vendors': 16,
      'stocks': 17,
      'zoom': 24,
      'audio/visual': 25,
      'venue': 26,
      'translator': 27,
      'production': 31,
      'guest speakers': 32,
      'performers': 33,
      'prizes': 34,
      'meals': 38,
      'refreshments': 39,
      'hotel rooms': 43,
      'paid staff': 47,
      'staff meals': 48,
      'flights': 52,
      'bus': 53,
      'rental car': 54,
      'other/miscellaneous': 57
    };

    // Apply AI values to the matching rows without disturbing formulas
    items.forEach(it => {
      const key = (it.item || '').toString().toLowerCase();
      const row = rowMap[key];
      if (row) {
        budgetSheet.getRange(row, 3).setValue(parseFloat(it.unitPrice) || 0);
        budgetSheet.getRange(row, 4).setValue(parseFloat(it.quantity) || 0);
      } else {
        Logger.log('Unmapped budget item: ' + it.item);
      }
    });

    if (questions.length > 0) {
      ui.alert('AI-generated budget has been added to the Budget sheet.\n\nQuestions:\n' + questions.join('\n'));
    } else {
      ui.alert('AI-generated budget has been added to the Budget sheet.');
    }

  } catch (e) {
    Logger.log(e.toString());
    ui.alert('Error generating AI budget: ' + e.message, ui.ButtonSet.OK);
  }
}
