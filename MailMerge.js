// MailMerge.js - Handles the mail merge and email sending functionality.

/**
 * Shows the custom HTML dialog for sending emails.
 */
function showEmailDialog() {
  const html = HtmlService.createHtmlOutputFromFile('EmailDialog')
      .setWidth(450)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Email to Event Attendees');
}

/**
 * Gets the necessary data to populate the email dialog's dropdowns.
 * This function is called from the client-side HTML.
 * @return {Object} An object containing lists of roles, statuses, and email templates.
 */
function getEmailUIData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const peopleSheet = ss.getSheetByName('People');
    const configSheet = ss.getSheetByName('Config');

    // Get unique roles (categories) from People sheet
    const roles = [...new Set(peopleSheet.getRange('B2:B').getValues().flat().filter(String))];

    // Get unique statuses from People sheet
    const statuses = [...new Set(peopleSheet.getRange('D2:D').getValues().flat().filter(String))];

    // Get email template names from Config sheet
    const templates = configSheet.getRange('A2:A').getValues().flat().filter(val => val.toString().endsWith('Template'));
    
    return {
      roles: ['All Roles', ...roles],
      statuses: ['All Statuses', ...statuses],
      templates: templates
    };
  } catch (e) {
    Logger.log(e);
    return { roles: [], statuses: [], templates: [] };
  }
}

/**
 * Sends emails based on the filters selected in the dialog.
 * MODIFIED: Now correctly replaces the {{name}} placeholder in the subject line.
 * @param {Object} filters An object containing the selected template, role, and status.
 */
function sendEmails(filters) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const peopleSheet = ss.getSheetByName('People');
    const configSheet = ss.getSheetByName('Config');
    
    // Get all people data
    const peopleData = peopleSheet.getDataRange().getValues();
    const peopleHeaders = peopleData.shift(); // Get headers and remove from data
    const nameIndex = peopleHeaders.indexOf('Name');
    const emailIndex = peopleHeaders.indexOf('Email');
    const roleIndex = peopleHeaders.indexOf('Category');
    const statusIndex = peopleHeaders.indexOf('Status');
    
    // Get the email template
    const configData = configSheet.getDataRange().getValues();
    const templateRow = configData.find(row => row[0] === filters.template);
    if (!templateRow) {
      throw new Error(`Template "${filters.template}" not found in Config sheet.`);
    }
    let subjectTemplate = templateRow[1];
    let bodyTemplate = templateRow[2];
    
    // Get event name to replace placeholder
    const eventName = getEventName(); // Assumes getEventName() exists
    subjectTemplate = subjectTemplate.replace(/\[EVENT NAME\]/g, eventName);
    bodyTemplate = bodyTemplate.replace(/\[EVENT NAME\]/g, eventName);

    // Filter people based on selected role and status
    let recipients = peopleData.filter(person => {
      const hasEmail = person[emailIndex];
      const roleMatch = filters.role === 'All Roles' || person[roleIndex] === filters.role;
      const statusMatch = filters.status === 'All Statuses' || person[statusIndex] === filters.status;
      return hasEmail && roleMatch && statusMatch;
    });

    if (recipients.length === 0) {
      return "No recipients found matching your criteria. No emails were sent.";
    }

    // Send emails
    let count = 0;
    recipients.forEach(recipient => {
      const personName = recipient[nameIndex];
      const personEmail = recipient[emailIndex];
      
      // --- THIS IS THE FIX ---
      // Personalize BOTH the subject and the body
      const personalizedSubject = subjectTemplate.replace(/{{name}}/g, personName);
      const personalizedBody = bodyTemplate.replace(/{{name}}/g, personName);
      
      GmailApp.sendEmail(personEmail, personalizedSubject, personalizedBody);
      count++;
    });

    return `Successfully sent ${count} emails using the "${filters.template}" template.`;
  } catch (e) {
    Logger.log(e);
    return `Error: ${e.message}`;
  }
}

