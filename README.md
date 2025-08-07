# LEGIT Event Planner Pro

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8-blue.svg)](https://developers.google.com/apps-script)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![CLASP](https://img.shields.io/badge/CLASP-Enabled-orange.svg)](https://github.com/google/clasp)

LEGIT Event Planner Pro is a Google Apps Script project for managing events directly inside Google Sheets. It provides tools to generate schedules, task lists, logistics, and budgets with the help of AI services such as OpenAI. The script also includes utilities for sending emails, creating forms, and building professional cue sheets. A built-in help system guides you through these features and links to a full user manual.

## 🚀 Features

- **🤖 AI-Powered Planning** - Automated generation of preliminary schedules, tasks, logistics lists, and budgets using OpenAI
- **📧 Smart Email System** - Role and status filters with AI message generation
- **📋 Form Generation** - Create and manage forms automatically saved to Google Drive
- **👥 People Management** - Comprehensive contact and role management
- **📅 Schedule Management** - Simplified scheduling with Time, Duration, Program, and Lead/Presenter columns
- **💰 Budget Tracking** - AI-powered budget generation with interactive questions
- **📊 Dashboard & Analytics** - Centralized view of all event components
- **⚙️ Configuration Management** - Customizable dropdowns and email templates
- **🎯 Task Management** - Advanced task tracking with categories and assignments
- **📖 Built-in Help System** - Contextual help and comprehensive user manual
- **🔄 Automation Tools** - Pro tools for advanced users

## 📋 Requirements

- **Google Account** - Required for Google Sheets and Apps Script
- **OpenAI API Key** - For AI-powered features (optional but recommended)
- **CLASP** - For development and deployment
- **Node.js** - For local development (optional)

## 🛠️ Quick Start

1. **Open the spreadsheet** - The Quick Start Guide appears automatically
2. **Create Event Description** - Use **📝 Create Event Description** to fill in event basics
3. **Access Help** - Use **📖 Help & User Guide** for context-sensitive tips
4. **Quick Setup** - Use **🗒️ Quick Event Setup** to capture tagline, theme, and key messages

## ⚙️ Setup

### 1. Clone and Install Dependencies
```bash
git clone <repo-url>
cd LEGIT-EVENTS
npm install
```

### 2. Install CLASP
```bash
npm install -g @google/clasp
```

### 3. Authenticate with CLASP
```bash
clasp login
```

### 4. Create Apps Script Project
```bash
clasp create --title "LEGIT Event Planner Pro" --type sheets
```

### 5. Configure API Keys
- Obtain an OpenAI API key from [OpenAI Platform](https://platform.openai.com/)
- Store keys using Apps Script Properties service
- Run **Event Planner Pro → Save API Key to Script Properties** from the sheet menu

### 6. Deploy
```bash
clasp push
```

### 7. Initialize Sheets
- Open the associated Google Sheet and refresh
- Run setup functions from the "Event Planner Pro" menu
- Authorize Google Drive access when prompted

### 8. Create New Planner
Use **Dashboard & Utilities → Create New Event Spreadsheet** to generate a fresh planner

### 9. Enable Automation
Run `createDropdownUpdateTrigger()` in Apps Script editor to enable automatic dropdown updates

## 📅 Schedule Management

### Simplified Schedule Structure
The schedule has been simplified to 4 essential columns:
- **Time** - Automatically calculated from duration entries
- **Duration** - Enter durations like "1h", "45m", "1h 30m"
- **Program** - Session titles and descriptions
- **Lead/Presenter** - Dropdown populated from People sheet

### Multi-Day Events
- Day separators can be added using the **📅 Add Day Separator** menu item
- Separators automatically reset time calculations for the new day
- Format: Enter day labels like "Day 2", "Tuesday", "Wednesday" in the Duration column

### Time Calculation
- Time column automatically calculates based on duration entries
- First session of each day starts at 9:00 AM by default
- Subsequent times are calculated by adding the previous session's duration

## 🤖 AI Integration

### OpenAI Features
- **Schedule Generation** - AI creates preliminary schedules based on event description
- **Task Generation** - Automated task list creation with categories
- **Budget Estimation** - AI-powered budget generation with interactive questions
- **Email Crafting** - Generate professional emails using AI
- **Logistics Planning** - Automated logistics list generation

### API Configuration
```javascript
// Store API key securely
PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', 'your-key-here');
```

## 📧 Email Templates and AI Generation

The **Send Emails** dialog includes:
- **Role and Status Filters** - Target specific groups
- **AI Message Generation** - Craft emails using OpenAI
- **Template Management** - Save and reuse email templates
- **Custom Subjects** - Personalized subject lines

## 📚 Documentation & Help

### Built-in Assistance
1. **📖 Help & User Guide** – Contextual help for active sheet
2. **🗒️ Quick Event Setup** – Fast event configuration dialog
3. **📕 User Manual (Google Doc)** – Complete online manual

### Offline Documentation
- Abbreviated manual available in [`docs/USER_MANUAL.md`](docs/USER_MANUAL.md)
- Full online manual: [Google Doc](https://docs.google.com/document/d/1w5KCO5O2MiuYDZMATFfLwGqHYrdsvhditDVzRJNmmP8/edit?usp=sharing)

## 🏗️ Repository Structure

```
├── Core.js                    # Custom menu and common utilities
├── Config.js                  # Configuration management
├── ScheduleGenerator.js       # AI-powered schedule generation
├── TaskManagement.js          # Advanced task tracking
├── Budget.js                  # Budget management and AI generation
├── Logistics.js               # Logistics planning
├── People.js                  # Contact and role management
├── Dashboard.js               # Central dashboard
├── MailMerge.js              # Email functionality
├── FormGenerator.js           # Form creation utilities
├── AutomationTools.js         # AI & Automation Tools setup
├── SmartUX.js                # User experience enhancements
├── EnhancedTaskManagement.js  # Advanced task features
├── SpeakerTaskCreator.js      # Speaker-specific task creation
├── GenerateCueSheet.js       # Professional cue sheet generation
├── EventDescription.js        # Event description management
├── CueBuilder.js             # Cue building utilities
├── appsscript.json           # Google Apps Script manifest
└── docs/
    └── USER_MANUAL.md        # Offline user manual
```

## 🔧 Troubleshooting

### Common Issues

**"Authorization Required"**
- Ensure you've authorized Google Drive access
- Check that the script has proper OAuth scopes

**"API Key Not Found"**
- Verify API key is stored in Script Properties
- Use **Event Planner Pro → Save API Key to Script Properties**

**"Menu Not Appearing"**
- Refresh the Google Sheet after deployment
- Check that `Core.js` is properly loaded

**"AI Features Not Working"**
- Verify OpenAI API key is valid and has credits
- Check internet connection for API calls

### Debug Mode
Enable debug logging in Apps Script editor:
```javascript
console.log('Debug information');
```

## 🤝 Contributing

We welcome contributions! Please follow these guidelines:

1. **Fork the repository**
2. **Create a feature branch** (`git checkout -b feature/amazing-feature`)
3. **Make your changes** with proper JSDoc comments
4. **Test thoroughly** in a development environment
5. **Submit a pull request**

### Development Guidelines
- Keep functions modular and well-documented
- Add JSDoc comments for all new functions
- Test changes in a separate Apps Script project first
- Follow existing code style and patterns

## 📄 License

This project is released under the MIT License. See [LICENSE](LICENSE) for details.

## 🙏 Acknowledgments

- Built with [Google Apps Script](https://developers.google.com/apps-script)
- AI features powered by [OpenAI](https://openai.com/)
- Deployed using [CLASP](https://github.com/google/clasp)

---

**Need Help?** Check the [User Manual](https://docs.google.com/document/d/1w5KCO5O2MiuYDZMATFfLwGqHYrdsvhditDVzRJNmmP8/edit?usp=sharing) or open an issue on GitHub.
