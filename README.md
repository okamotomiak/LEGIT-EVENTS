# LEGIT Event Planner Pro

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat-square&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=flat-square&logo=javascript&logoColor=black)
![OpenAI](https://img.shields.io/badge/OpenAI-412991?style=flat-square&logo=openai&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=flat-square&logo=google-sheets&logoColor=white)

> A comprehensive Google Apps Script solution for professional event planning directly within Google Sheets, powered by AI automation and intelligent workflows.

## 🎯 Overview

LEGIT Event Planner Pro is a sophisticated Google Apps Script project that transforms Google Sheets into a powerful event management platform. It combines AI-powered automation with practical event planning tools to streamline the entire event lifecycle—from initial planning to execution.

### ✨ Key Highlights

- **🤖 AI-Powered**: Automated generation of schedules, tasks, logistics, and budgets using OpenAI
- **📊 Smart Analytics**: Built-in dashboard with interactive charts and progress tracking
- **📧 Communication Hub**: Advanced email system with role-based filtering and AI message generation
- **🔄 Progressive UX**: Intelligent menu system that adapts to your workflow progress
- **📋 Comprehensive Tools**: Complete suite for managing people, schedules, budgets, and logistics
- **🎨 Professional Output**: Generate polished cue sheets, forms, and documentation

## 🚀 Features

### Core Event Management
- **📝 Event Description Builder**: Structured event information capture and management
- **👥 People Management**: Comprehensive contact management with roles and status tracking
- **📅 Schedule Generator**: AI-powered preliminary schedule creation with time management
- **✅ Task Management**: Intelligent task categorization with automated speaker task creation
- **💰 Budget Planning**: AI-assisted budget generation with detailed cost breakdowns
- **🚚 Logistics Coordination**: Comprehensive logistics planning and tracking

### AI & Automation Tools
- **🧠 Smart Content Generation**: AI-powered creation of schedules, tasks, and budgets
- **📧 Intelligent Email System**: Role and status-based filtering with AI message crafting
- **📋 Dynamic Form Generation**: Automated form creation with organized folder management
- **📊 Live Dashboard**: Real-time charts and analytics for event progress tracking
- **🎵 Professional Cue Sheets**: Automated generation of detailed event cue sheets

### Advanced Features
- **⚙️ Smart Configuration**: Interactive configuration dialogs for customizing workflows
- **🔄 Progressive Menu System**: Context-aware menu that reveals tools as you progress
- **📚 Integrated Help System**: Contextual help and comprehensive user documentation
- **🛠️ Pro Tools**: Quick access to blank sheets and advanced utilities
- **📱 Responsive Design**: Modern HTML dialogs with professional styling

## ⚡ Quick Start

### 1. First Launch
When you first open the spreadsheet, the **Quick Start Guide** appears automatically. You can reopen it anytime from the **Event Planner Pro** menu.

### 2. Basic Setup
1. **📝 Create Event Description** - Fill in your event basics
2. **🗒️ Quick Event Setup** - Capture tagline, theme, key messages, and profit goals
3. **📖 Help & User Guide** - Access context-sensitive tips anytime

### 3. Progressive Workflow
The system intelligently reveals advanced tools as you complete basic setup:
- Complete event description → Unlock schedule tools
- Add people → Enable task management
- Set budget goals → Activate budget generator

## 🛠️ Installation & Setup

### Prerequisites
- Google Account with access to Google Sheets and Google Apps Script
- [Node.js](https://nodejs.org/) (for development)
- [CLASP CLI](https://github.com/google/clasp) for deployment

### 1. Clone and Install Dependencies
```bash
git clone <repo-url>
cd LEGIT-EVENTS
npm install
```

Install CLASP globally:
```bash
npm install -g @google/clasp
```

### 2. Authentication
```bash
clasp login
```

### 3. Create Apps Script Project
```bash
clasp create --title "LEGIT Event Planner Pro" --type sheets
```

### 4. Configure API Keys
1. Obtain an **OpenAI API key** for AI features
2. Store keys securely using Apps Script Properties:
   - Go to Apps Script editor → Project Settings → Script Properties
   - Or use **Event Planner Pro → Save API Key to Script Properties** from the sheet menu

### 5. Deploy
```bash
clasp push
```

After deployment:
1. Open the associated Google Sheet and refresh
2. Authorize Google Drive access when prompted
3. Initialize sheets using the "Event Planner Pro" menu

### 6. Advanced Setup
- **Create New Planner**: Use **Dashboard & Utilities → Create New Event Spreadsheet**
- **Enable Auto-Updates**: Run `createDropdownUpdateTrigger()` in Apps Script editor
- **Learn Advanced Tools**: Create the **AI & Automation Tools** sheet for feature overview

## 📧 Email & AI Features

### Smart Email System
The **Send Emails** dialog includes:
- **Role-based filtering**: Target specific participant types
- **Status filtering**: Email only confirmed attendees
- **AI Generation**: Automatic subject and body creation using OpenAI
- **Template Management**: Save and reuse email templates

### AI Content Generation
- **Schedules**: Generate preliminary event timelines
- **Tasks**: Create comprehensive task lists with categorization
- **Logistics**: Build detailed logistics checklists
- **Budgets**: AI-assisted budget estimation with clarifying questions
- **Email Content**: Craft professional communications automatically

## 📚 Documentation & Help

### Built-in Help System
- **📖 Help & User Guide**: Contextual help for active sheet
- **🗒️ Quick Event Setup**: Fast event configuration dialog
- **📕 User Manual (Google Doc)**: Complete online documentation

### Additional Resources
- **Offline Manual**: [`docs/USER_MANUAL.md`](docs/USER_MANUAL.md)
- **AI Tools Overview**: Auto-generated sheet explaining advanced features
- **Quick Start Guide**: Interactive onboarding experience

## 📁 Project Structure

```
├── Core.js                     # Central functionality and menu system
├── SmartUX.js                  # Progressive UX and intelligent menu management
├── Config.js                   # Configuration management and setup
├── Dashboard.js                # Analytics dashboard with charts
├── ScheduleGenerator.js        # AI-powered schedule generation
├── TaskManagement.js           # Comprehensive task management system
├── EnhancedTaskManagement.js   # Advanced task features
├── Budget.js                   # Budget planning and management
├── Logistics.js                # Logistics coordination tools
├── People.js                   # Contact and people management
├── EventDescription.js         # Event information management
├── MailMerge.js               # Email system and communication tools
├── FormGenerator.js           # Dynamic form creation
├── GenerateCueSheet.js        # Professional cue sheet generation
├── SpeakerTaskCreator.js      # Automated speaker task management
├── AutomationTools.js         # AI & automation tools overview
├── CueBuilder.js              # Cue sheet building utilities
├── formTemplates.js           # Form template definitions
├── appsscript.json            # Google Apps Script manifest
├── docs/
│   └── USER_MANUAL.md         # Offline documentation
└── *.html                     # Dialog interfaces and forms
```

### Key Components

| File | Purpose |
|------|---------|
| `Core.js` | Entry point with `onOpen()` and central utilities |
| `SmartUX.js` | Progressive menu system and UX management |
| `ScheduleGenerator.js` | AI-powered schedule creation (1162 lines) |
| `TaskManagement.js` | Comprehensive task system (1022 lines) |
| `SmartUX.js` | Intelligent user experience management (727 lines) |

## 🎨 User Interface

### Modern Dialog System
- **Responsive HTML dialogs** with professional styling
- **Context-aware interfaces** that adapt to current data
- **Progressive disclosure** of advanced features
- **Integrated help** within each dialog

### Smart Menu System
The menu system intelligently adapts to your progress:
- **Basic Tools**: Always available for core functionality
- **Advanced Features**: Revealed as you complete prerequisites
- **Pro Tools**: Unlocked for power users and complex events

## 🔧 Development

### Code Organization
- **Modular Architecture**: Each feature in its own file for maintainability
- **JSDoc Documentation**: Comprehensive inline documentation
- **Progressive Enhancement**: Features build on each other logically
- **Error Handling**: Robust error management throughout

### Contributing Guidelines
1. **Keep functions modular** and single-purpose
2. **Document with JSDoc comments** for all public functions
3. **Follow existing code style** and naming conventions
4. **Test thoroughly** before submitting pull requests
5. **Update documentation** for any new features

### Development Workflow
```bash
# Make changes to .js files
# Push to Apps Script
clasp push

# Pull latest from Apps Script (if editing online)
clasp pull
```

## 🚀 Advanced Usage

### Custom Automation
- **Trigger Setup**: Automatic dropdown updates and data synchronization
- **API Integration**: Extend with additional AI services or external APIs
- **Custom Forms**: Generate specialized forms for unique event types
- **Bulk Operations**: Process multiple events or participants simultaneously

### Power User Features
- **Blank Sheet Generator**: Quickly add Budget, Logistics, or custom sheets
- **Advanced Filtering**: Complex queries for people and task management
- **Bulk Email Operations**: Send targeted communications to large groups
- **Custom Cue Sheets**: Professional event production documentation

## 🤝 Contributing

We welcome contributions! Here's how to get started:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/amazing-feature`
3. **Make your changes** following our coding standards
4. **Add tests** if applicable
5. **Update documentation** for any new features
6. **Submit a pull request**

### Development Standards
- Use JSDoc for all function documentation
- Follow existing naming conventions
- Keep functions modular and focused
- Test with real Google Sheets data
- Ensure mobile-friendly dialog interfaces

## 📄 License

This project is released under the **MIT License**. See [LICENSE](LICENSE) for details.

## 🙏 Acknowledgments

- **OpenAI** for AI content generation capabilities
- **Google Apps Script** platform for serverless execution
- **Google Workspace** for integrated productivity tools
- **CLASP** for development workflow automation

---

**Need Help?** Check our [User Manual](docs/USER_MANUAL.md) or open an issue on GitHub.

**Professional Event Planning Made Simple** ✨
