//Config.gs - Sets up and manages the main configuration sheet.

/**
 * Creates and populates the Config sheet with a clean, organized structure.
 * This can be run once to initialize or to reset the sheet to defaults.
 */
function setupConfigSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  
  // Create the sheet if it doesn't exist
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.setTabColor('#6aa84f'); // Green color
  } else {
    configSheet.clear(); // Clear existing content to ensure a fresh setup
  }
  
  // Set up headers
  const headers = ["Key / Template Name", "Value / Subject", "Body / Notes"];
  configSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold');
  
  // Set column widths
  configSheet.setColumnWidth(1, 220); // Key / Template Name
  configSheet.setColumnWidth(2, 250); // Value / Subject
  configSheet.setColumnWidth(3, 450); // Body / Notes

  // Prepare the configuration data in logical groups
  const configData = [
    // --- Section 1: Dropdown Lists ---
    ['--- DROPDOWN LISTS ---', 'Enter comma-separated values for dropdown menus across the planner.', ''],
    ['People Categories', 'Staff,Volunteer,Speaker,Participant', 'Used in the "Category" dropdown in the People sheet.'],
    ['People Statuses', 'Potential,Invited,Accepted,Registered,Unavailable', 'Used in the "Status" dropdown in the People sheet.'],
    ['Schedule Status Options', 'Tentative,Confirmed,Cancelled', 'Used in the "Status" dropdown in the Schedule sheet.'],
    ['Task Status Options', 'Not Started,In Progress,Blocked,Done,Cancelled', 'Used in the "Status" dropdown in the Task Management sheet.'],
    ['Task Priority Options', 'High,Medium,Low,Critical', 'Used in the "Priority" dropdown in the Task Management sheet.'],
    ['Location List', 'Main Hall,Room 101,Room 102,Outdoor Area', 'Used in the "Location" dropdown in the Schedule sheet.'],
    ['Owners', 'Jane Doe,John Smith,Alex Johnson', 'Used in the "Owner" dropdown in the Task Management sheet.'],
    
    // --- Section 2: System & AI Settings ---
    ['', '', ''], // Spacer
    ['--- SYSTEM & AI SETTINGS ---', '', ''],
    ['OpenAI API Key', '', 'Enter your OpenAI API key here for all AI-powered features.'],
    ['Look-Ahead Days', '1', 'Number of days ahead to look for upcoming session reminders.'],
    ['Reminder Lead Time (days)', '2', 'How many days before a task is due should a reminder be triggered.'],
    
    // --- Section 3: Email Templates ---
    ['', '', ''], // Spacer
    ['--- EMAIL TEMPLATES ---', 'Subject lines go in the middle column, email body in the right.', ''],
    ['InviteTemplate', 'Invitation: {{name}} for [EVENT NAME]', 'Hi {{name}},\n\nYou are invited to [EVENT NAME]!\n\nPlease RSVP by [RSVP Date].\n\nBest regards,\n[Your Name/Org]'],
    ['ReminderTemplate', 'Reminder: [EVENT NAME] is coming up!', 'Hi {{name}},\n\nJust a friendly reminder about the upcoming event: [EVENT NAME] on [Date] at [Time].\n\nBest regards,\n[Your Name/Org]'],
    ['ThankYouTemplate', 'Thank You for Attending [EVENT NAME]!', 'Hi {{name}},\n\nThank you for attending [EVENT NAME]!\n\nWe hope you enjoyed it.\n\nBest regards,\n[Your Name/Org]'],

    // --- Section 4: Generated Form Links ---
    ['', '', ''], // Spacer
    ['--- GENERATED FORM LINKS ---', 'Links to generated forms will appear here automatically.', 'Do not edit this section manually.']
  ];
  
  // Insert the configuration data
  configSheet.getRange(2, 1, configData.length, configData[0].length).setValues(configData);
  
  // --- Formatting ---
  // Highlight section headers
  const sectionHeaderRows = [2, 10, 16, 21];
  sectionHeaderRows.forEach(row => {
    configSheet.getRange(row, 1, 1, 3).setBackground('#d9d9d9').setFontWeight('bold');
  });

  // Highlight the OpenAI API Key row
  configSheet.getRange(11, 1, 1, 3).setBackground('#fff2cc');

  // Format the body column for text wrapping
  const bodyColumn = configSheet.getRange(2, 3, configData.length, 1);
  bodyColumn.setWrap(true).setVerticalAlignment('top');
  
  // Freeze the header row
  configSheet.setFrozenRows(1);
  
  // Alert the user
  SpreadsheetApp.getUi().alert('Config sheet has been cleaned up and reorganized.');
  
  return configSheet;
}
