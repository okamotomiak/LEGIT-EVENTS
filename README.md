# LEGIT Event Planner Pro

LEGIT Event Planner Pro is a Google Apps Script project for managing events directly inside Google Sheets. It provides tools to generate schedules, task lists, logistics, and budgets with the help of AI services such as OpenAI. The script also includes utilities for sending emails, creating forms, and building professional cue sheets. A built-in help system guides you through these features and links to a full user manual.

## Features

- Automated menu with one-click access to event planning tools.
- AI-powered generation of preliminary schedules, tasks, logistics lists, and budgets.
- Form and email generators to streamline communication. Generated forms are saved in a "[Event Name] Forms" folder next to the spreadsheet.
- New email dialog with role and status filters, plus an **Generate with AI** option to craft messages automatically.
- Tools for managing people, schedules, and cues.
- Interactive configuration dialog for customizing dropdown lists and email templates.
- Dedicated **AI & Automation Tools** sheet to explain advanced menu options.
- Modular code organized by feature for easier maintenance.
- Helpful onboarding with a **🚀 2-Minute Setup Wizard** and contextual help menu.
- One-click access to a **📕 User Manual (Google Doc)** and an offline copy in `docs/USER_MANUAL.md`.

## Quick Start

1. After installing, open the Google Sheet and choose **🚀 2-Minute Setup Wizard** from the **Event Planner Pro** menu.
2. Follow the on-screen prompts to generate your core sheets and sample data.
3. Access **📖 Help & User Guide** at any time for context-sensitive tips.
4. Explore **Show Me Around (Tutorial)** to add tutorial columns explaining each sheet. You can remove them later via **Remove Tutorial Overlays** in the same menu.

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
   After pushing, open the associated Google Sheet and refresh to load the custom menu. When first running the script you may be asked to authorize
   Google Drive access so that new spreadsheets and generated forms can be created.

6. **Initialize Sheets**
   Run the setup functions from the "Event Planner Pro" menu to create the necessary sheets (Config, Schedule, Logistics, Budget, etc.). These provide templates for your event data and settings.

7. **Create a New Planner**
   Use **Dashboard & Utilities → Create New Event Spreadsheet** to generate a fresh planner. The new file includes this script project and only the base sheets (Dashboard, Event Description, People, Schedule, Task Management, and Config).

8. **Learn Advanced Tools**
   Run **Dashboard & Utilities → Create/Reset AI & Automation Tools Sheet** for a quick overview of optional automation features like cue sheets and form generators. The sheet explains what each advanced menu item does and provides links to setup dialogs.

9. **Enable Automatic Dropdown Updates**
   In the Apps Script editor run `createDropdownUpdateTrigger()` once. This sets
   up a daily trigger that refreshes dropdown lists across all sheets.

## Email Templates and AI Generation

The **Send Emails** dialog now lets you filter recipients by role and status and includes a **Generate with AI** button. When used, OpenAI crafts a subject line and body using your event description. You can tweak the generated text and click **Save Template** to add it to the `Config` sheet for later use. This feature relies on the OpenAI API key saved in your script properties.

## Documentation & Help

The **Event Planner Pro** menu provides built-in assistance:

1. **📖 Help & User Guide** – shows contextual help for the active sheet.
2. **🗒️ Quick Event Setup** – opens a dialog for fast event configuration.
3. **📕 User Manual (Google Doc)** – opens the full online manual.

An abbreviated offline manual is available in [`docs/USER_MANUAL.md`](docs/USER_MANUAL.md).

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
