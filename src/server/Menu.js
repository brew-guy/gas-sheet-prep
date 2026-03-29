/**
 * Creates a custom menu in Google Sheets when the document opens (Standalone Mode).
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e 
 */
function onOpen(e) {
  initImporterMenu(SpreadsheetApp.getUi());
}

/**
 * Main initialization wrapper for library consumers.
 * @param {GoogleAppsScript.Base.Ui} ui 
 * @param {string} [namespace] - Optional namespace (Library identifier) to prefix menu calls.
 */
function initImporterMenu(ui, namespace) {
  ui = ui || SpreadsheetApp.getUi();
  const prefix = namespace ? namespace + '.' : '';
  ui.createMenu('🔌 Importer')
    .addItem('Open Importer UI', prefix + 'openImporterModal')
    .addToUi();
}

/**
 * Universal Library Router (McPherson Proxy Pattern).
 * Receives calls from the UI and routes them to the correct library function.
 * @param {string} funcName 
 * @param  {...any} args 
 */
function runRouter(funcName, ...args) {
  return this[funcName].apply(this, args);
}

/**
 * Opens the Data File Importer modal dialog.
 */
function openImporterModal() {
  const htmlTemplate = HtmlService.createTemplateFromFile('src/ui/ImporterModal');
  const htmlOutput = htmlTemplate.evaluate()
    .setWidth(1000)
    .setHeight(650)
    .setTitle('Data File Importer');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Data File Importer');
}

/**
 * Helper strictly required by HtmlService to include other files.
 * @param {string} filename 
 * @returns {string} 
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
