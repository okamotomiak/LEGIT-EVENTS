# LEGIT Event Planner Pro

LEGIT Event Planner Pro is a Google Apps Script project for managing events directly inside Google Sheets. It provides tools to generate schedules, task lists, logistics, and budgets with the help of AI services such as OpenAI. The script also includes utilities for sending emails, creating forms, and building professional cue sheets.

## Features

- Automated menu with one-click access to event planning tools.
- AI-powered generation of preliminary schedules, tasks, logistics lists, and budgets.
- Form and email generators to streamline communication. Generated forms are saved in a "[Event Name] Forms" folder next to the spreadsheet.
- Tools for managing people, schedules, and cues.
- Dedicated **AI & Automation Tools** sheet to explain advanced menu options.
- Modular code organized by feature for easier maintenance.

## Setup

1. **Clone and install dependencies**
   ```bash
   git clone <repo-url>
   cd LEGIT-EVENTS
   npm install
   ```
   The project uses [CLASP](https://github.com/google/clasp) for deployment. Make sure it is installed globally:
   ```bash
   npm install -g @google/clasp
   ```

2. **Authenticate with CLASP**
   ```bash
   clasp login
   ```

3. **Create a new Apps Script project**
   ```bash
   clasp create --title "LEGIT Event Planner Pro" --type sheets
   ```
   This will link the local files with your new script project.

4. **Configure API Keys**
   - Obtain an OpenAI API key and any other credentials required by your workflow.
   - Store sensitive keys using the Apps Script Properties service rather than directly in the spreadsheet. In the Apps Script editor, go to `Project Settings` ➜ `Script Properties` or run **Event Planner Pro → Save API Key to Script Properties** from the sheet menu to add your keys.

5. **Deploy**
   ```bash
   clasp push
   ```
   After pushing, open the associated Google Sheet and refresh to load the custom menu.

6. **Initialize Sheets**
   Run the setup functions from the "Event Planner Pro" menu to create the necessary sheets (Config, Schedule, Logistics, Budget, etc.). These provide templates for your event data and settings.

7. **Create a New Planner**
   Use **Dashboard & Utilities → Create New Event Spreadsheet** to generate a fresh planner. The new file includes this script project and only the base sheets (Dashboard, Event Description, People, Schedule, Task Management, and Config).

8. **Learn Advanced Tools**
   Run **Dashboard & Utilities → Create/Reset AI & Automation Tools Sheet** for a quick overview of optional automation features like cue sheets and form generators.

## Repository Structure

- `Core.js` – Creates the custom menu and houses common utilities.
- `Config.js` – Handles setup and management of configuration data.
- `ScheduleGenerator.js` – Generates preliminary schedules using AI services.
- `Logistics.js`, `Budget.js`, `TaskManagement.js` – Sheets and logic for logistics, budgeting, and tasks.
- `AutomationTools.js` – Sets up the AI & Automation Tools overview sheet.
- `MailMerge.js`, `FormGenerator.js` – Communication helpers for emailing and form creation.
- `appsscript.json` – Google Apps Script project manifest.

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests. Please keep functions modular and document any new code with JSDoc comments.

## License

This project is released under the MIT License. See [LICENSE](LICENSE) for details.
