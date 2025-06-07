// FormTemplates.js - Creates and manages the sheet for dynamic form generation.

/**
 * Creates the "Form Templates" sheet and populates it with default templates.
 * MODIFIED: Added text wrapping to appropriate columns for better readability.
 */
function setupFormTemplatesSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetName = 'Form Templates';
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    ui.alert('Sheet Created', `"Form Templates" sheet has been created and populated with default designs. You can now customize them or add your own.`, ui.ButtonSet.OK);
  } else {
    const response = ui.alert('Reset Sheet?', 'The "Form Templates" sheet already exists. Do you want to clear it and reset it with the default templates?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
    sheet.clear();
  }
  
  // Set up headers
  const headers = [
    'Form Name', 'Form Description', 'Confirmation Message', 
    'Question Title', 'Question Type', 'Options (comma-separated)', 'Is Required?', 'Maps to People Column'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#073763').setFontColor('#ffffff').setFontWeight('bold');

  // Define the default templates
  const templates = [
    // Template 1: Participant Registration
    ['Participant Registration', 'Registration form for our upcoming event. Please fill out all required fields.', 'Thank you for registering! We will be in touch with more details soon.', 'Full Name', 'Text', '', true, 'Name'],
    ['Participant Registration', '', '', 'Email', 'Text', '', true, 'Email'],
    ['Participant Registration', '', '', 'Phone Number', 'Text', '', false, 'Phone'],
    ['Participant Registration', '', '', 'Dietary Restrictions or Preferences', 'Paragraph Text', '', false, 'Notes'],

    // Template 2: Volunteer Signup
    ['Volunteer Signup', 'Join our team to help make this event a success!', 'Thank you for volunteering! We will contact you soon with more details.', 'Full Name', 'Text', '', true, 'Name'],
    ['Volunteer Signup', '', '', 'Email', 'Text', '', true, 'Email'],
    ['Volunteer Signup', '', '', 'Phone Number', 'Text', '', true, 'Phone'],
    ['Volunteer Signup', '', '', 'Availability', 'Checkboxes', 'Setup Day,Event Day - Morning,Event Day - Afternoon,Cleanup Day', true, 'Notes'],
    ['Volunteer Signup', '', '', 'Preferred Role', 'Multiple Choice', 'Registration,Technical Support,Logistics,General Support', false, 'Role/Position'],

    // Template 3: Speaker Information
    ['Speaker Information', 'Please provide your details for our event program.', 'Thank you for submitting your speaker information.', 'Full Name', 'Text', '', true, 'Name'],
    ['Speaker Information', '', '', 'Email', 'Text', '', true, 'Email'],
    ['Speaker Information', '', '', 'Session Title', 'Text', '', true, 'Role/Position'],
    ['Speaker Information', '', '', 'Speaker Bio (50-100 words)', 'Paragraph Text', '', true, 'Notes'],

    // Template 4: Vendor Signup
    ['Vendor Signup', 'Apply to be a vendor at our event.', 'Thank you for your application! We will review it and get back to you shortly.', 'Company Name', 'Text', '', true, 'Name'],
    ['Vendor Signup', '', '', 'Contact Email', 'Text', '', true, 'Email'],
    ['Vendor Signup', '', '', 'Type of Goods/Services', 'Text', '', true, 'Role/Position'],
    ['Vendor Signup', '', '', 'Special Requests (e.g., electricity)', 'Paragraph Text', '', false, 'Notes'],

    // Template 5: Participant Feedback
    ['Participant Feedback', 'Thank you for attending! Please let us know what you thought.', 'Your feedback is valuable to us. Thank you!', 'Name (Optional)', 'Text', '', false, 'Name'],
    ['Participant Feedback', '', '', 'Overall Rating', 'Multiple Choice', 'Excellent,Good,Average,Poor', true, 'Notes'],
    ['Participant Feedback', '', '', 'What did you enjoy most?', 'Paragraph Text', '', false, 'Notes'],
    ['Participant Feedback', '', '', 'How can we improve?', 'Paragraph Text', '', false, 'Notes'],

    // Template 6: Press/Media Pass Application
    ['Press/Media Pass Application', 'Application for a complimentary media pass for our event.', 'Thank you for your interest. We will review your application and be in touch.', 'Full Name', 'Text', '', true, 'Name'],
    ['Press/Media Pass Application', '', '', 'Outlet / Publication', 'Text', '', true, 'Role/Position'],
    ['Press/Media Pass Application', '', '', 'Email', 'Text', '', true, 'Email'],
    ['Press/Media Pass Application', '', '', 'Link to Portfolio/Website', 'Text', '', false, 'Notes']
  ];
  
  // Populate sheet with templates
  sheet.getRange(2, 1, templates.length, headers.length).setValues(templates);
  
  // Formatting
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 250);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 150);
  
  // --- THIS IS THE NEW PART ---
  // Set text wrapping for columns with potentially long text and align to the top
  sheet.getRange('B2:D').setWrap(true).setVerticalAlignment('top'); // Form Desc, Confirm Msg, Question Title
  sheet.getRange('F2:F').setWrap(true).setVerticalAlignment('top'); // Options column
  // -------------------------
  
  // Add dropdowns for Question Type
  const questionTypes = ['Text', 'Paragraph Text', 'Multiple Choice', 'Checkboxes', 'Date', 'Time'];
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(questionTypes).build();
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
  
  // Get the headers from the People sheet to create the mapping dropdown
  const peopleSheet = ss.getSheetByName('People');
  if (!peopleSheet) {
    // If People sheet doesn't exist, create it.
    setupPeopleSheet(ss, false); 
  }
  const peopleHeaders = peopleSheet.getRange(1, 1, 1, peopleSheet.getLastColumn()).getValues().flat();
  const mapRule = SpreadsheetApp.newDataValidation().requireValueInList(peopleHeaders).build();
  sheet.getRange(2, 8, sheet.getMaxRows() - 1, 1).setDataValidation(mapRule);
  
  sheet.setFrozenRows(1);
}
