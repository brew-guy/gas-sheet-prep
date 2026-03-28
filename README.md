# SheetPrep - Google Sheets Data File Importer

This project is a custom Data File Importer designed for Google Sheets, built using Google Apps Script (GAS) and HTML Service. It provides an intuitive UI (sidebar and dialog) to import, trim, and clean CSV, XLSX, and XLS data before writing it to a sheet.

## Features

- **Client-Side Parsing:** Uses PapaParse (for CSV) and SheetJS (for Excel) to parse files locally in the browser, preventing server timeouts and maintaining privacy.
- **Multiple File Types:** Supports CSV and multi-sheet Excel files.
- **Data Trimming:** Skip rows from the top/bottom and specify advanced crop options to skip columns from the left/right before importing.
- **Find & Replace:** Add powerful replacement rules to clean up text data dynamically during the import preview phase.
- **Saved Flows:** Your import configurations (cropping and replacement rules) are automatically saved to a hidden `_SavedFlows` sheet against the destination sheet name. Whenever you import to that target sheet again, your previous rules are instantly restored.
- **Import Methods:** Create a new sheet, or append/overwrite an existing sheet. Also features header mismatch detection so you don't accidentally write bad data over columns.
- **Constraints Handling:** Automatically handles Google Sheets' 50,000-character-per-cell limits.

## Project Structure

The project has been refactored into modular components for easier maintainability:

- `src/server/`: Backend GAS code (e.g., `Menu.js`, `SheetUtils.js`). These handle the Google Sheets API interactions (getting headers, committing imports, saving rules).
- `src/ui/`: Frontend HTML dialog components (e.g., `ImporterModal.html`).
- `src/ui/js/`: Frontend logic (e.g., `ModalLogic.html`) containing the necessary `<script>` tags for UI-server communication.
- `src/ui/styles/`: Frontend styling (e.g., `ModalStyles.html`) containing CSS for the modal and UI elements.

## Local Development & Deployment

Code synchronisation is managed by `clasp` (Command Line Apps Script Projects).

### Prerequisites

1. Ensure you have Node.js and NPM installed.
2. Install `clasp` globally:

   ```bash
   npm install -g @google/clasp
   ```

3. Enable the Google Apps Script API in your [Google User Settings](https://script.google.com/home/usersettings).

### Setup

1. Clone this repository.
2. Authenticate `clasp` with your Google account:

   ```bash
   clasp login
   ```

3. Connect the repository to your existing Apps Script project. Ensure that `.clasp.json` is correctly set with the `scriptId`.

   ```bash
   # Automatically inferred to .clasp.json if setup correctly
   clasp pull
   ```

### Pushing Code

When you make changes locally, push them to Google Apps Script:

```bash
clasp push
```

## Using as a Plug-and-Play Library

This codebase is designed using **Bruce McPherson's Universal Proxy Pattern** (from Desktop Liberation) to solve `google.script.run` encapsulation issues. This means you can deploy this project as a reusable GAS Library and attach it to any existing spreadsheet without polluting its global namespace!

1. Deploy this script as a **Library** from the GAS Editor (Deploy > New Deployment > Library).
2. Note the generated **Script ID**.
3. In your target spreadsheet (the consumer), add the Library using the Script ID and set the identifier to `DataImporter`.
4. Paste the following "boilerplate" directly into the target spreadsheet's code:

```javascript
/* =========================================
   🔌 DATA IMPORTER LIBRARY LINK
   ========================================= */
// Add the menu initialized by the library
function onOpen(e) {
  DataImporter.initImporterMenu(SpreadsheetApp.getUi());
}

// Bruce McPherson's Universal Library Router
// This routes all client-side UI calls back to the library's internal functions
function runRouter(funcName, ...args) {
  return DataImporter[funcName].apply(this, args);
}
```

That's it! Because the UI uses `runRouter` to communicate, you never need to define wrapper functions for individual backend methods. Your dashboard sheet stays clean, and you can update the importer library centrally.

## How It Works

1. On opening Google Sheets, an **Importer** menu is created.
2. Selecting **Open Importer UI** opens the `ImporterModal.html` using the GAS HtmlService.
3. The UI runs entirely in the client environment, processing the selected file, rendering a data preview, and allowing rule adjustments.
4. When you click **Import Data**, the UI communicates context (trim limits, replace values, parsed data) back to the server.
5. `SheetUtils.js` processes the data structure constraints and pushes the finalized 2D array matrix cleanly to the target Sheet.
