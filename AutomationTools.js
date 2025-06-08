// AutomationTools.js - Creates the AI & Automation Tools info sheet

/**
 * Creates or resets the "AI & Automation Tools" sheet with descriptions of advanced features.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Optional spreadsheet to operate on.
 */
function setupAutomationToolsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let sheet = ss.getSheetByName('AI & Automation Tools');

  if (!sheet) {
    sheet = ss.insertSheet('AI & Automation Tools');
    sheet.setTabColor('#8e7cc3');
  } else {
    const resp = ui.alert('Reset AI & Automation Tools?',
      'This will clear the sheet and re-add the default information.',
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
    sheet.clear();
  }

  const headers = ['Tool or Sheet', 'Purpose'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#674ea7')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  const info = [
    ['AI Generators menu',
     'Create schedules, task lists, logistics, and budgets with one click.'],
    ['Cue Builder & Professional Cue Sheet',
     'Use Production Tools to build cues and generate a professional cue sheet.'],
    ['Form Templates & Mail Merge',
     'Generate Google Forms and send bulk emails from the Communication Tools.'],
    ['Tutorial System', 'Show or hide in-sheet tutorials from the menu.'],
    ['Create New Event Spreadsheet',
     'Duplicate this planner with all scripts and base sheets attached.']
  ];

  sheet.getRange(2, 1, info.length, 2).setValues(info);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 500);
  sheet.setFrozenRows(1);
}
