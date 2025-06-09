// MailMerge.js - Handles the mail merge and email sending functionality.

/**
 * Shows the custom HTML dialog for sending emails.
 */
function showEmailDialog() {
  const html = HtmlService.createHtmlOutputFromFile('EmailAIDialog')
      .setWidth(550)
      .setHeight(600);
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

/**
 * Generates an email subject and body using OpenAI based on a prompt.
 * @param {string} prompt The prompt describing the email to generate.
 * @return {{subject:string, body:string}} Generated email content.
 */
function generateEmailWithAI(prompt) {
  const apiKey = getOpenAIApiKey();
  if (!apiKey) {
    throw new Error('OpenAI API key not found.');
  }

  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: 'gpt-4o',
    messages: [{ role: 'user', content: prompt +
      '\nRespond ONLY with JSON {"subject":"","body":"" }' }],
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
    throw new Error('OpenAI API Error: ' + response.getContentText());
  }

  const parsed = JSON.parse(response.getContentText());
  return JSON.parse(parsed.choices[0].message.content);
}

/**
 * Sends emails using either a template name or provided subject/body content.
 * @param {Object} data Options from the dialog.
 * @param {string=} data.template Template name to use.
 * @param {string=} data.subject Subject text if not using a template.
 * @param {string=} data.body Body text if not using a template.
 * @param {string} data.role Filter role.
 * @param {string} data.status Filter status.
 * @return {string} Status message.
 */
function sendEmailsAdvanced(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const peopleSheet = ss.getSheetByName('People');
    const configSheet = ss.getSheetByName('Config');

    const peopleData = peopleSheet.getDataRange().getValues();
    const headers = peopleData.shift();
    const nameIndex = headers.indexOf('Name');
    const emailIndex = headers.indexOf('Email');
    const roleIndex = headers.indexOf('Category');
    const statusIndex = headers.indexOf('Status');

    let subjectTemplate = data.subject;
    let bodyTemplate = data.body;

    if (!subjectTemplate || !bodyTemplate) {
      const configData = configSheet.getDataRange().getValues();
      const row = configData.find(r => r[0] === data.template);
      if (!row) {
        throw new Error('Template "' + data.template + '" not found in Config sheet.');
      }
      subjectTemplate = row[1];
      bodyTemplate = row[2];
    }

    const eventName = getEventName();
    subjectTemplate = subjectTemplate.replace(/\[EVENT NAME\]/g, eventName);
    bodyTemplate = bodyTemplate.replace(/\[EVENT NAME\]/g, eventName);

    const recipients = peopleData.filter(p => {
      const hasEmail = p[emailIndex];
      const roleMatch = data.role === 'All Roles' || p[roleIndex] === data.role;
      const statusMatch = data.status === 'All Statuses' || p[statusIndex] === data.status;
      return hasEmail && roleMatch && statusMatch;
    });

    if (recipients.length === 0) {
      return 'No recipients found matching your criteria. No emails were sent.';
    }

    recipients.forEach(r => {
      const subject = subjectTemplate.replace(/{{name}}/g, r[nameIndex]);
      const body = bodyTemplate.replace(/{{name}}/g, r[nameIndex]);
      GmailApp.sendEmail(r[emailIndex], subject, body);
    });

    return 'Successfully sent ' + recipients.length + ' emails.';
  } catch (e) {
    Logger.log(e.toString());
    return 'Error: ' + e.message;
  }
}

/**
 * Saves a custom email template to the Config sheet.
 * @param {string} name Template name.
 * @param {string} subject Subject line.
 * @param {string} body Email body.
 * @return {string} Status message.
 */
function saveEmailTemplate(name, subject, body) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Config');
    if (!sheet) throw new Error('Config sheet not found.');

    const data = sheet.getDataRange().getValues();
    let rowIndex = data.findIndex(r => r[0] === name);
    if (rowIndex === -1) {
      rowIndex = sheet.getLastRow();
      sheet.appendRow([name, subject, body]);
    } else {
      rowIndex += 1; // convert to 1-indexed
      sheet.getRange(rowIndex, 1, 1, 3).setValues([[name, subject, body]]);
    }

    return 'Template saved.';
  } catch (e) {
    Logger.log(e.toString());
    return 'Error: ' + e.message;
  }
}

