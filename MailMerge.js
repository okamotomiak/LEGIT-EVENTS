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
function sendEmails(filters, subject, body) {
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

    const usingCustom = subject && body;
    let subjectTemplate = subject;
    let bodyTemplate = body;

    if (!subjectTemplate || !bodyTemplate) {
      const configData = configSheet.getDataRange().getValues();
      const templateRow = configData.find(row => row[0] === filters.template);
      if (!templateRow) {
        throw new Error(`Template "${filters.template}" not found in Config sheet.`);
      }
      subjectTemplate = templateRow[1];
      bodyTemplate = templateRow[2];
    }
    
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
    const personalizedSubject = subjectTemplate.replace(/{{name}}/gi, personName);
    const personalizedBody = bodyTemplate.replace(/{{name}}/gi, personName);
      
      GmailApp.sendEmail(personEmail, personalizedSubject, personalizedBody);
      count++;
    });

    if (usingCustom) {
      return `Successfully sent ${count} emails.`;
    }
    return `Successfully sent ${count} emails using the "${filters.template}" template.`;
  } catch (e) {
    Logger.log(e);
    return `Error: ${e.message}`;
  }
}


/**
 * Generates an email subject and body using OpenAI based on event information.
 * The returned strings include a {{Name}} placeholder for personalization.
 * @param {string} prompt Additional instructions for the AI.
 * @return {Object} Object with `subject` and `body` keys.
 */
function generateEmailWithAI(prompt) {
  try {
    const apiKey = getOpenAIApiKey();
    if (!apiKey) throw new Error('OpenAI API key not found.');

    const eventInfo = getEventInformation();
    if (!eventInfo) throw new Error('Event information not available.');

    const start = eventInfo.startDate ? Utilities.formatDate(new Date(eventInfo.startDate), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    const end = eventInfo.endDate ? Utilities.formatDate(new Date(eventInfo.endDate), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    let context = `Event Name: ${eventInfo.eventName}\n` +
                  `Dates: ${start}${eventInfo.durationDays > 1 ? ' to ' + end : ''}\n` +
                  `Location: ${eventInfo.location || 'TBD'}`;
    if (eventInfo.description) context += `\nDescription: ${eventInfo.description}`;

    const fullPrompt = `${prompt}\n\nEVENT INFORMATION:\n${context}\n\n` +
      'Return a JSON object {"subject":"","body":""} ' +
      'and use {{Name}} where the recipient name should appear.';

    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: 'gpt-4.1-mini',
      messages: [{ role: 'user', content: fullPrompt }],
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
    return { subject: content.subject || '', body: content.body || '' };
  } catch (e) {
    Logger.log(e);
    return { subject: '', body: '' };
  }
}

/**
 * Saves or updates an email template in the Config sheet under "EMAIL TEMPLATES".
 * @param {string} name Template name/key.
 * @param {string} subject Subject line text.
 * @param {string} body Body text.
 */
function saveEmailTemplate(name, subject, body) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Config');
    if (!sheet) throw new Error('Config sheet not found.');

    const data = sheet.getDataRange().getValues();
    let emailHeaderRow = -1;
    let nextSectionRow = data.length + 1;
    let existingRow = -1;

    for (let i = 0; i < data.length; i++) {
      const cell = data[i][0];
      if (cell === name) existingRow = i + 1;
      if (cell === '--- EMAIL TEMPLATES ---') {
        emailHeaderRow = i + 1;
        for (let j = i + 1; j < data.length; j++) {
          const val = data[j][0];
          if (val && val.toString().startsWith('---')) {
            nextSectionRow = j + 1;
            break;
          }
        }
      }
    }

    if (existingRow !== -1) {
      sheet.getRange(existingRow, 1, 1, 3).setValues([[name, subject, body]]);
      return;
    }

    if (emailHeaderRow === -1) {
      sheet.appendRow([name, subject, body]);
      return;
    }

    if (nextSectionRow > data.length) {
      sheet.appendRow([name, subject, body]);
    } else {
      sheet.insertRows(nextSectionRow, 1);
      sheet.getRange(nextSectionRow, 1, 1, 3).setValues([[name, subject, body]]);
    }
  } catch (e) {
    Logger.log('Error saving email template: ' + e.toString());
  }
}

