// FormGenerator.js - Dynamically generates Google Forms from templates in a sheet.

/**
 * Shows a dialog to select which forms to generate from the "Form Templates" sheet.
 */
function showFormSelectionDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName('Form Templates');
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert('Error', '"Form Templates" sheet not found. Please create it first using the menu.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Get unique form names from the template sheet
  const formNamesRange = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 1);
  const formNames = [...new Set(formNamesRange.getValues().flat().filter(String))];
  
  if (formNames.length === 0) {
    SpreadsheetApp.getUi().alert('No Templates Found', 'Please define at least one form in the "Form Templates" sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Build the HTML for the dialog with some basic styling
  let html = `
    <style>
      body {
        font-family: 'Inter', sans-serif;
        font-size: 14px;
        line-height: 1.5;
        padding: 20px;
        max-width: 400px;
        margin: 0 auto;
        word-wrap: break-word;
      }
      label { display: block; margin-bottom: 8px; }
      input[type=checkbox] { margin-right: 6px; }
    </style>
    <form>`;
  formNames.forEach(name => {
    html += `<label><input type="checkbox" name="form" value="${name}"> ${name}</label>`;
  });
  html += `<br><input type="button" value="Generate Selected Forms" onclick="google.script.run.withSuccessHandler(google.script.host.close).generateSelectedForms(this.parentNode);"></form>`;

  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Forms to Generate');
}

/**
 * Server-side function to generate forms selected in the dialog.
 * @param {Object} selection The object containing the selected form names.
 */
function generateSelectedForms(selection) {
  const selectedForms = Array.isArray(selection.form) ? selection.form : [selection.form];
  if (!selectedForms || selectedForms[0] == null) {
      SpreadsheetApp.getUi().alert("No forms were selected.");
      return;
  }

  let summaryMessage = 'Form Generation Complete:\n\n';
  selectedForms.forEach(formName => {
    const url = buildFormFromTemplate(formName);
    summaryMessage += url ? `✓ ${formName} created\n` : `✗ ${formName} failed\n`;
  });
  
  summaryMessage += '\nAll form links have been saved to the Config sheet.';
  SpreadsheetApp.getUi().alert('Forms Generated', summaryMessage, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Builds a single Google Form based on a template from the "Form Templates" sheet.
 * @param {string} formName The name of the form template to build.
 * @return {string|null} The URL of the created form or null on failure.
 */
function buildFormFromTemplate(formName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('Form Templates');
    const allTemplates = templateSheet.getDataRange().getValues();
    const headers = allTemplates[0];
    
    // Filter to get only the rows for the specified form template
    const formTemplateRows = allTemplates.slice(1).filter(row => row[0] === formName);
    
    if (formTemplateRows.length === 0) {
      throw new Error(`Template for "${formName}" not found.`);
    }
    
    const templateInfo = formTemplateRows[0];
    const formTitle = `${getEventName()} - ${templateInfo[headers.indexOf('Form Name')]}`;
    const formDescription = templateInfo[headers.indexOf('Form Description')];
    const confirmationMessage = templateInfo[headers.indexOf('Confirmation Message')];

    // Create the form
    const form = FormApp.create(formTitle)
      .setTitle(formTitle)
      .setDescription(formDescription)
      .setConfirmationMessage(confirmationMessage)
      .setAllowResponseEdits(true)
      .setCollectEmail(true);

    // Move the form to an event-specific folder next to the spreadsheet
    const formsFolder = getOrCreateFormsFolder(getEventName(), ss.getId());
    const formFile = DriveApp.getFileById(form.getId());
    formsFolder.addFile(formFile);
    // Remove from root to avoid clutter
    DriveApp.getRootFolder().removeFile(formFile);

    // Add questions based on the template
    formTemplateRows.forEach(row => {
      const title = row[headers.indexOf('Question Title')];
      const type = row[headers.indexOf('Question Type')];
      const options = row[headers.indexOf('Options (comma-separated)')].toString().split(',').map(s => s.trim());
      const isRequired = row[headers.indexOf('Is Required?')] === true;
      
      let item;
      switch (type) {
        case 'Paragraph Text':
          item = form.addParagraphTextItem();
          break;
        case 'Multiple Choice':
          item = form.addMultipleChoiceItem().setChoiceValues(options);
          break;
        case 'Checkboxes':
          item = form.addCheckboxItem().setChoiceValues(options);
          break;
        case 'Date':
          item = form.addDateItem();
          break;
        case 'Time':
          item = form.addTimeItem();
          break;
        default: // 'Text'
          item = form.addTextItem();
          break;
      }
      item.setTitle(title).setRequired(isRequired);
    });
    
    // Set response destination and create trigger
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    ScriptApp.newTrigger('processDynamicFormResponse').forForm(form).onFormSubmit().create();
    
    // Save URL to Config sheet
    saveFormUrl(formName, form.getPublishedUrl());

    return form.getPublishedUrl();
  } catch (error) {
    Logger.log(`Error building form "${formName}": ${error.toString()}`);
    return null;
  }
}

/**
 * A single function to process responses from any dynamically created form.
 * @param {Object} e The form submit event object.
 */
function processDynamicFormResponse(e) {
  try {
    const form = e.source;
    const formResponse = e.response;
    const formTitle = form.getTitle();
    const respondentEmail = formResponse.getRespondentEmail();

    // Extract the template name from the form title
    const templateName = formTitle.split(' - ')[1];
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('Form Templates');
    const allTemplates = templateSheet.getDataRange().getValues();
    const headers = allTemplates[0];
    const formTemplateRows = allTemplates.slice(1).filter(row => row[0] === templateName);
    
    const personData = { 'Email': respondentEmail, 'Notes': '' };

    formResponse.getItemResponses().forEach(itemResponse => {
      const questionTitle = itemResponse.getItem().getTitle();
      const answer = itemResponse.getResponse();
      
      // Find the mapping for this question in the template
      const templateRow = formTemplateRows.find(row => row[headers.indexOf('Question Title')] === questionTitle);
      if (templateRow) {
        const mappedColumn = templateRow[headers.indexOf('Maps to People Column')];
        if (mappedColumn) {
          if (mappedColumn === 'Notes') {
            personData['Notes'] += `${questionTitle}: ${answer}\n`;
          } else {
            personData[mappedColumn] = answer;
          }
        }
      }
    });
    
    addOrUpdatePersonInPeopleSheet(personData); // Assumes this function exists in People.js
  } catch (error) {
    Logger.log(`Error processing dynamic form response: ${error.toString()}`);
  }
}

// Helper functions (should already exist in your project)
function getEventName() {
  // ... (implementation from your existing FormGenerator.js)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventDescSheet = ss.getSheetByName('Event Description');
  if (!eventDescSheet) return "Event";
  const data = eventDescSheet.getDataRange().getValues();
  const eventNameRow = data.find(row => row[0] === 'Event Name');
  return eventNameRow ? eventNameRow[1] : "Event";
}

function saveFormUrl(formName, url) {
 // ... (modified implementation to handle dynamic names)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;
  
  const data = configSheet.getDataRange().getValues();
  const key = formName.toLowerCase();
  let rowFound = -1;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === key) {
      rowFound = i + 1;
      break;
    }
  }

  if (rowFound !== -1) {
    configSheet.getRange(rowFound, 2).setFormula(`=HYPERLINK("${url}","${formName}")`);
  } else {
    // Add new row if key doesn't exist
    configSheet.appendRow([formName, `=HYPERLINK("${url}","${formName}")`]);
  }
}

/**
 * Get or create the folder used to store generated forms.
 * The folder is placed next to the spreadsheet and named "[Event Name] Forms".
 * @param {string} eventName The current event name.
 * @param {string} spreadsheetId The ID of the active spreadsheet.
 * @return {Folder} The Drive folder for the forms.
 */
function getOrCreateFormsFolder(eventName, spreadsheetId) {
  const sheetFile = DriveApp.getFileById(spreadsheetId);
  const parents = sheetFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folderName = `${eventName} Forms`;
  const existing = parentFolder.getFoldersByName(folderName);
  return existing.hasNext() ? existing.next() : parentFolder.createFolder(folderName);
}

function addOrUpdatePersonInPeopleSheet(data){
  // This function should be in your People.js file.
  // This is a placeholder for the logic to add/update a person.
  Logger.log('Adding/Updating Person: ' + JSON.stringify(data));
}