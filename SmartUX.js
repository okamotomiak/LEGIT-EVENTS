//SmartUX.js - Progressive User Experience System

/**
 * Enhanced onOpen function with progressive menu revelation
 */
// URL to the full user manual
const USER_MANUAL_URL = 'https://docs.google.com/document/d/1w5KCO5O2MiuYDZMATFfLwGqHYrdsvhditDVzRJNmmP8/edit?usp=sharing';

function smartUXOnOpen() {
  const ui = SpreadsheetApp.getUi();
  const userProgress = assessUserProgress();
  
  // Always available items
  let menu = ui.createMenu('Event Planner Pro 🚀')
    .addItem('📖 Help & User Guide', 'showContextualHelp')
    .addItem('🗒️ Quick Event Setup', 'showEventSetupDialog')
    .addItem('📕 User Manual (Google Doc)', 'showUserManual')
    .addSeparator();

  // Progressive menu based on user progress
  if (userProgress.isNewUser) {
    menu = addNewUserMenu(menu, ui);
  } else if (userProgress.isBasicUser) {
    menu = addBasicUserMenu(menu, ui);
  } else if (userProgress.isAdvancedUser) {
    menu = addAdvancedUserMenu(menu, ui);
  } else {
    menu = addExpertUserMenu(menu, ui);
  }

  menu.addToUi();
  
  // Show onboarding for new users
  if (userProgress.isNewUser) {
    Utilities.sleep(1000); // Let menu load first
    checkAndGuideUser();
  }
}

/**
 * Assess user progress to determine menu complexity
 */
function assessUserProgress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check basic sheets completion
  const hasEventDesc = hasContentInSheet('Event Description');
  const hasPeople = hasContentInSheet('People');
  const hasSchedule = hasContentInSheet('Schedule');
  const hasTasks = hasContentInSheet('Task Management');
  
  // Check advanced features
  const hasAPIKey = getOpenAIApiKey() !== null;
  const hasAdvancedSheets = ss.getSheetByName('Budget') || ss.getSheetByName('Logistics');
  const hasGeneratedContent = checkForGeneratedContent();
  
  const basicSheetsComplete = hasEventDesc && hasPeople && hasSchedule && hasTasks;
  
  return {
    isNewUser: !hasEventDesc,
    isBasicUser: hasEventDesc && !basicSheetsComplete,
    isAdvancedUser: basicSheetsComplete && (hasAPIKey || hasAdvancedSheets),
    isExpert: basicSheetsComplete && hasAPIKey && hasGeneratedContent,
    hasAPIKey: hasAPIKey,
    basicSheetsComplete: basicSheetsComplete
  };
}

/**
 * Check if a sheet has meaningful content
 */
function hasContentInSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return false;
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false; // Only header row
  
  // Check if there's actual content (not just empty rows)
  if (sheetName === 'Event Description') {
    // Check if Event Name is filled
    const eventNameRow = findRowByValue(sheet, 'Event Name');
    if (eventNameRow) {
      const eventName = sheet.getRange(eventNameRow, 2).getValue();
      return eventName && eventName.toString().trim() !== '';
    }
  } else {
    // For other sheets, check if there's data beyond headers
    const dataRange = sheet.getRange(2, 1, Math.min(lastRow - 1, 5), 1);
    const values = dataRange.getValues();
    return values.some(row => row[0] && row[0].toString().trim() !== '');
  }
  
  return false;
}

/**
 * Check for AI-generated content
 */
function checkForGeneratedContent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName('Task Management');
  
  if (!taskSheet) return false;
  
  const lastRow = taskSheet.getLastRow();
  if (lastRow < 10) return false; // Generated content usually creates many tasks
  
  return true;
}

/**
 * Menu for brand new users
 */
function addNewUserMenu(menu, ui) {
  return menu
    .addSubMenu(ui.createMenu("🌟 Let's Get Started!")
      .addItem('🚀 2-Minute Setup Wizard', 'startSetupWizard')
      .addItem('📝 Create Event Description', 'setupEventDescriptionSheet')
      .addItem('🗒️ Quick Event Setup', 'showEventSetupDialog')
      .addSeparator()
    )
    .addItem('⚙️ Pro Tools', 'showAdvancedOptionsDialog');
}

/**
 * Menu for users with basic setup
 */
function addBasicUserMenu(menu, ui) {
  return menu
    .addSubMenu(ui.createMenu('📋 Build Your Foundation')
      .addItem('👥 Set Up People', 'setupPeopleSheet')
      .addItem('📅 Set Up Schedule', 'setupScheduleSheet')
      .addItem('✅ Set Up Tasks', 'setupTaskManagementSheet')
      .addItem('📊 View Dashboard', 'setupDashboard'))
    .addSubMenu(ui.createMenu('📈 Ready for More?')
      .addItem('💰 Add Budget Planning', 'setupBudgetSheet')
      .addItem('📦 Add Logistics', 'setupLogisticsSheet')
      .addItem('🤖 Try AI Tools', 'showAIIntroduction'))
    .addItem('⚙️ Pro Tools', 'showAdvancedOptionsDialog');
}

/**
 * Menu for advanced users
 */
function addAdvancedUserMenu(menu, ui) {
  return menu
    .addItem('⚙️ Pro Tools', 'showAdvancedOptionsDialog');
}

/**
 * Full menu for expert users
 */
function addExpertUserMenu(menu, ui) {
  return menu
    .addSubMenu(ui.createMenu('✉️ Communication')
      .addItem('📝 Form Templates', 'setupFormTemplatesSheet')
      .addItem('🔗 Generate Forms', 'showFormSelectionDialog')
      .addItem('📧 Send Emails', 'showEmailDialog'))
    .addSubMenu(ui.createMenu('⚙️ Utilities')
      .addItem('🔧 Configuration', 'showConfigDialog')
      .addItem('📋 New Event Planner', 'createNewEventSpreadsheet'));
}

/**
 * Progressive Help System - Context-sensitive manual
 */
function showContextualHelp() {
  const currentSheet = SpreadsheetApp.getActiveSheet().getName();
  const userProgress = assessUserProgress();
  
  let helpContent = getHelpContentForSheet(currentSheet, userProgress);
  
  const html = HtmlService.createHtmlOutput(createHelpHTML(helpContent))
    .setWidth(700)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, `📖 Help: ${currentSheet}`);
}

/**
 * Generate contextual help content
 */
function getHelpContentForSheet(sheetName, userProgress) {
  const helpContent = {
    'Event Description': {
      title: 'Event Description - Your Foundation',
      phase: userProgress.isNewUser ? 'Getting Started' : 'Foundation Complete',
      content: `
        <h3>🎯 What This Sheet Does</h3>
        <p>This is the foundation of your entire event planner. Everything else builds from the information you put here.</p>
        
        <h3>✅ Essential Fields to Fill Out</h3>
        <ul>
          <li><strong>Event Name:</strong> What are you calling your event?</li>
          <li><strong>Start Date:</strong> When does it begin?</li>
          <li><strong>End Date:</strong> When does it end?</li>
          <li><strong>Location:</strong> Where is it happening? (Even "TBD" is fine)</li>
          <li><strong>Description:</strong> What's the purpose? Who's coming?</li>
          <li><strong>Attendance Goal:</strong> How many people do you expect?</li>
        </ul>
        
        <h3>💡 Pro Tips</h3>
        <ul>
          <li>Write descriptions like you're explaining to a friend</li>
          <li>The more detail you provide, the better AI tools work later</li>
          <li>Don't worry about perfection - you can always edit</li>
        </ul>
        
        ${userProgress.isNewUser ? '<h3>🚀 Next Step</h3><p>After filling this out, go to the People sheet to add your team!</p>' : ''}
      `
    },
    
    'People': {
      title: 'People - Your Event Team',
      phase: userProgress.isBasicUser ? 'Building Your Foundation' : 'Team Management',
      content: `
        <h3>👥 What This Sheet Does</h3>
        <p>Track everyone involved in your event - from core team to attendees.</p>
        
        <h3>📝 How to Add People</h3>
        <ol>
          <li>Click an empty row</li>
          <li>Enter their name</li>
          <li>Choose category: Staff, Volunteer, Speaker, or Participant</li>
          <li>Set their status</li>
          <li>Add contact info if you have it</li>
        </ol>
        
        <h3>✨ Smart Features</h3>
        <ul>
          <li><strong>Auto-Task Creation:</strong> When you mark a Speaker as "Accepted", the system creates a task to collect their bio!</li>
          <li><strong>Dropdowns:</strong> Categories and statuses ensure consistency</li>
          <li><strong>Integration:</strong> Names automatically appear in Schedule lead dropdown</li>
        </ul>
        
        <h3>🎯 Start Simple</h3>
        <p>Add just your core team first (yourself, co-organizer, key helpers). You can always add more people later!</p>
      `
    },
    
    'Schedule': {
      title: 'Schedule - Your Event Timeline',
      phase: 'Planning Your Event Flow',
      content: `
        <h3>📅 What This Sheet Does</h3>
        <p>Plan what happens when during your event.</p>
        
        <h3>⚡ Smart Features</h3>
        <ul>
          <li><strong>Auto Duration:</strong> Enter start/end times, duration calculates automatically</li>
          <li><strong>Status Alerts:</strong> Get notifications when you mark sessions "Confirmed"</li>
          <li><strong>Time Format:</strong> Use "9:00 AM" or "2:30 PM" format</li>
        </ul>
        
        <h3>🚀 Quick Start Method</h3>
        <p>Start with big blocks:</p>
        <ul>
          <li>Arrival/Setup</li>
          <li>Main activity</li>
          <li>Break/Food</li>
          <li>Wrap-up</li>
        </ul>
        
        ${userProgress.hasAPIKey ? '<h3>🤖 AI Option</h3><p>You can also use <strong>AI Generators → Generate Schedule</strong> to create a full timeline automatically!</p>' : ''}
      `
    },
    
    'Task Management': {
      title: 'Task Management - Getting Things Done',
      phase: 'Organizing Your Work',
      content: `
        <h3>✅ What This Sheet Does</h3>
        <p>Track everything that needs to get done for your event.</p>
        
        <h3>📋 Task Organization</h3>
        <ul>
          <li><strong>Categories:</strong> Venue, Marketing, Logistics, Program, Budget, etc.</li>
          <li><strong>Priorities:</strong> Critical, High, Medium, Low</li>
          <li><strong>Due Dates:</strong> When does it need to be done?</li>
          <li><strong>Owners:</strong> Who's responsible?</li>
        </ul>
        
        <h3>🎯 Start Simple</h3>
        <p>Add obvious tasks first:</p>
        <ul>
          <li>Book venue</li>
          <li>Send invitations</li>
          <li>Order food</li>
          <li>Prepare materials</li>
        </ul>
        
        ${userProgress.hasAPIKey ? '<h3>🤖 AI Boost</h3><p>Use <strong>AI Generators → Generate Tasks</strong> to create a comprehensive task list based on your event details!</p>' : ''}
        
        <h3>📊 Track Progress</h3>
        <p>Check the Dashboard to see your completion percentage and stay motivated!</p>
      `
    },
    
    'Dashboard': {
      title: 'Dashboard - Your Event Command Center',
      phase: 'Monitoring Progress',
      content: `
        <h3>📊 What This Sheet Shows</h3>
        <p>Real-time overview of your event planning progress.</p>
        
        <h3>📈 Key Metrics</h3>
        <ul>
          <li><strong>Task Progress:</strong> How many tasks are complete?</li>
          <li><strong>Upcoming Sessions:</strong> What's happening next?</li>
          <li><strong>Status Summary:</strong> Breakdown of all task statuses</li>
          <li><strong>Event Goals:</strong> Attendance and financial targets</li>
        </ul>
        
        <h3>🔄 Keeping It Current</h3>
        <p>Click the <strong>🔄 Refresh</strong> button to update all metrics with the latest information.</p>
        
        <h3>💡 Pro Tip</h3>
        <p>Check this regularly to stay motivated and catch any tasks that might be falling behind!</p>
        <p>You can customize status options and other settings in the <strong>Config</strong> sheet.</p>
      `
    }
  };
  
  return helpContent[sheetName] || {
    title: `${sheetName} Help`,
    phase: 'General Information',
    content: '<p>This sheet helps you manage specific aspects of your event. Check the User Guide for detailed information!</p>'
  };
}

/**
 * Create HTML for help dialog
 */
function createHelpHTML(helpContent) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; line-height: 1.6; }
          h2 { color: #674ea7; margin-bottom: 5px; }
          .phase { color: #666; font-style: italic; margin-bottom: 20px; }
          h3 { color: #333; margin-top: 25px; margin-bottom: 10px; }
          ul, ol { margin-left: 20px; }
          li { margin-bottom: 5px; }
          .next-step { background: #e8f4fd; padding: 15px; border-radius: 5px; margin-top: 20px; }
          .ai-feature { background: #f0f8f0; padding: 10px; border-radius: 5px; margin: 10px 0; }
          strong { color: #674ea7; }
        </style>
      </head>
      <body>
        <h2>${helpContent.title}</h2>
        <div class="phase">Phase: ${helpContent.phase}</div>
        ${helpContent.content}
        
        <div class="next-step">
          <strong>Need more help?</strong> Check the <a href="${USER_MANUAL_URL}" target="_blank">User Guide</a> for step-by-step examples or ask your team for support!
        </div>
      </body>
    </html>
  `;
}


/**
 * Smart Onboarding Flow
 */
function checkAndGuideUser() {
  const userProgress = assessUserProgress();
  
  if (userProgress.isNewUser) {
    showWelcomeWizard();
  }
}

function showWelcomeWizard() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Welcome to Event Planner Pro! 🎉',
    'Looks like you\'re just getting started. Would you like a quick 2-minute setup to get your first event planned?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    startSetupWizard();
  } else {
    ui.alert(
      'No Problem!',
      'You can start the setup wizard anytime from the "Event Planner Pro" menu. Just look for "Let\'s Get Started!"',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 2-Minute Setup Wizard
 */
function startSetupWizard() {
  const ui = SpreadsheetApp.getUi();
  
  // Step 1: Event Name
  const eventNameResponse = ui.prompt(
    'Step 1 of 4: What\'s Your Event?',
    'What are you planning? (e.g., "Sarah\'s Birthday Party", "Annual Team Meeting")',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (eventNameResponse.getSelectedButton() !== ui.Button.OK) return;
  const eventName = eventNameResponse.getResponseText().trim();
  
  if (!eventName) {
    ui.alert('Let\'s try again with an event name!');
    return;
  }
  
  // Step 2: Date
  const dateResponse = ui.prompt(
    'Step 2 of 4: When Is It?',
    'When is your event? (e.g., "December 15, 2024", "Next Saturday")',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  const eventDate = dateResponse.getResponseText().trim();
  
  // Step 3: Location
  const locationResponse = ui.prompt(
    'Step 3 of 4: Where Is It?',
    'Where will it happen? (e.g., "My house", "Conference Room A", "TBD")',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (locationResponse.getSelectedButton() !== ui.Button.OK) return;
  const location = locationResponse.getResponseText().trim();
  
  // Step 4: Description
  const descResponse = ui.prompt(
    'Step 4 of 4: Tell Us More',
    'What\'s this event about? Who\'s coming? (1-2 sentences)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (descResponse.getSelectedButton() !== ui.Button.OK) return;
  const description = descResponse.getResponseText().trim();
  
  // Create the basic setup
  executeWizardSetup(eventName, eventDate, location, description);
  
  // Show completion
  ui.alert(
    'Setup Complete! 🎉',
    `Great! "${eventName}" is now set up in your planner.\n\nNext steps:\n• Add key people in the People sheet\n• Create a simple schedule\n• Add some basic tasks\n\nCheck the Dashboard to see your progress!`,
    ui.ButtonSet.OK
  );
  
  // Navigate to Event Description sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('Event Description');
  if (eventSheet) {
    ss.setActiveSheet(eventSheet);
  }
}

/**
 * Execute the wizard setup
 */
function executeWizardSetup(eventName, eventDate, location, description) {
  // Create Event Description sheet if needed
  setupEventDescriptionSheet();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('Event Description');
  
  if (eventSheet) {
    // Fill in the wizard data
    const updates = [
      ['Event Name', eventName],
      ['Start Date (And Time)', eventDate],
      ['Location', location],
      ['Description & Messaging', description]
    ];
    
    updates.forEach(([field, value]) => {
      const row = findRowByValue(eventSheet, field);
      if (row) {
        eventSheet.getRange(row, 2).setValue(value);
      }
    });
  }
  
  // Create other essential sheets
  setupPeopleSheet(ss, false);
  setupScheduleSheet(ss, false);
  setupTaskManagementSheet(ss, false);
  setupDashboard();
}

/**
 * Show AI introduction for users ready to level up
 */
function showAIIntroduction() {
  const ui = SpreadsheetApp.getUi();
  const hasAPIKey = getOpenAIApiKey() !== null;
  
  if (hasAPIKey) {
    ui.alert(
      'AI Tools Ready! 🤖',
      'You already have AI set up! Try:\n• Generate AI Schedule\n• Generate AI Tasks\n• Generate AI Budget\n\nThese tools analyze your event details and create comprehensive plans automatically.',
      ui.ButtonSet.OK
    );
  } else {
    const response = ui.alert(
      'Unlock AI-Powered Planning! 🤖',
      'AI tools can automatically generate schedules, task lists, and budgets based on your event details.\n\nTo use AI features, you\'ll need a free OpenAI API key. Would you like to set this up now?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ui.alert(
        'Getting Your API Key',
        'Steps:\n1. Go to openai.com and create a free account\n2. Navigate to API Keys section\n3. Create a new key\n4. Come back and use "Set Up AI (API Key)" from the menu\n\nIt takes about 2 minutes and unlocks powerful automation!',
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * Show communication menu for advanced users
 */
function showCommunicationMenu() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Communication Tools',
    'Choose what you\'d like to do:\n\n• Create Forms: Build registration, feedback, or info collection forms\n• Send Emails: Send bulk emails to your team or participants\n• Set Up Templates: Customize email and form templates',
    ui.ButtonSet.OK
  );
}

/**
 * Show advanced options dialog for any user level
 */
function showAdvancedOptionsDialog() {
  const ui = SpreadsheetApp.getUi();
  const userProgress = assessUserProgress();
  
  let message = 'Advanced Features:\n\n';
  
  if (!userProgress.hasAPIKey) {
    message += '🤖 AI Tools (Requires API Key):\n• Auto-generate schedules, tasks, budgets\n\n';
  }
  
  message += '🎬 Production Tools:\n• Professional cue sheets for complex events\n\n';
  message += '✉️ Communication:\n• Auto-generate forms\n• Send bulk emails\n\n';
  message += '⚙️ Utilities:\n• Create new event planners\n• Advanced configuration\n\n';
  message += 'Would you like to explore these features?';
  
  const response = ui.alert('Pro Tools', message, ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    // Temporarily show full menu by updating user to expert level
    showProToolsMenu();
  }
}

function showProToolsMenu() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🚀 Pro Tools')
    .addSubMenu(ui.createMenu('🤖 AI Tools')
      .addItem('🔑 Set Up API Key', 'saveApiKeyToScriptProperties')
      .addItem('📅 Generate Schedule', 'generatePreliminarySchedule')
      .addItem('✅ Generate Tasks', 'generateAITasksWithSchedule')
      .addItem('💰 Generate Budget', 'generateAIBudget')
      .addItem('📦 Generate Logistics', 'showLogisticsDialog'))
    .addSubMenu(ui.createMenu('🎬 Production')
      .addItem('🎯 Cue Builder', 'setupCueBuilderSheet')
      .addItem('📄 Professional Cue Sheet', 'generateProfessionalCueSheet'))
    .addSubMenu(ui.createMenu('✉️ Communication')
      .addItem('📝 Form Templates', 'setupFormTemplatesSheet')
      .addItem('🔗 Generate Forms', 'showFormSelectionDialog')
      .addItem('📧 Send Emails', 'showEmailDialog'))
    .addSubMenu(ui.createMenu('⚙️ Utilities')
      .addItem('🔧 Configuration', 'showConfigDialog')
      .addItem('📋 New Event Planner', 'createNewEventSpreadsheet'))
    .addToUi();
}

/**
 * Display a dialog with a link to the full user manual
 */
function showUserManual() {
  const html = HtmlService.createHtmlOutput(
    `<p><a href="${USER_MANUAL_URL}" target="_blank">Open the User Manual</a></p>`
  ).setWidth(350).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'User Manual');
}
