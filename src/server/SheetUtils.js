/**
 * Utility functions for interacting with Google Sheets in the Importer UI.
 */

const RESERVED_SHEETS = ['_CleaningRules', '_SavedFlows'];

/**
 * Returns a list of sheet names in the active spreadsheet, excluding reserved sheets.
 */
function getAvailableSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets
    .map(sheet => sheet.getName())
    .filter(name => !name.startsWith('_'));
}

/**
 * Fetches the top row of an existing sheet to compare against parsed file headers.
 */
function getSheetHeaders(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return [];
  
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return headers.map(h => String(h).trim());
}

/**
 * Fetches the saved flow rules for a specific destination sheet.
 */
function getSavedFlowRules(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  
  const flowSheet = ss.getSheetByName('_SavedFlows');
  if (!flowSheet) return null;
  
  const lastRow = flowSheet.getLastRow();
  const lastCol = flowSheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 6) return null;
  
  const data = flowSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === sheetName) {
      let replacements = [];
      try {
        replacements = data[i][5] ? JSON.parse(data[i][5]) : [];
      } catch (_) {
        replacements = [];
      }
      return {
        skipTop: Number(data[i][1]) || 0,
        skipBottom: Number(data[i][2]) || 0,
        skipLeft: Number(data[i][3]) || 0,
        skipRight: Number(data[i][4]) || 0,
        replacements
      };
    }
  }
  
  return null;
}

/**
 * Saves or updates flow rules for a specific destination sheet.
 */
function saveFlowRules(sheetName, skipTop, skipBottom, skipLeft, skipRight, replacements) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let flowSheet = ss.getSheetByName('_SavedFlows');
  
  if (!flowSheet) {
    flowSheet = ss.insertSheet('_SavedFlows');
    flowSheet.hideSheet();
    flowSheet.appendRow(['Destination Sheet', 'Skip Top', 'Skip Bottom', 'Skip Left', 'Skip Right', 'Replacements JSON', 'Last Updated']);
    // Format headers
    flowSheet.getRange("1:1").setFontWeight("bold");
    flowSheet.setFrozenRows(1);
  }
  
  const replacementsStr = JSON.stringify(replacements || []);
  const timestamp = new Date().toISOString();
  
  const lastRow = flowSheet.getLastRow();
  const lastCol = Math.max(flowSheet.getLastColumn(), 7);
  let rowIndex = -1;
  
  if (lastRow > 1) {
    const data = flowSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === sheetName) {
        rowIndex = i + 2; // 1-indexed + header row
        break;
      }
    }
  }
  
  if (rowIndex !== -1) {
    flowSheet.getRange(rowIndex, 2, 1, 6).setValues([[skipTop, skipBottom, skipLeft, skipRight, replacementsStr, timestamp]]);
  } else {
    flowSheet.appendRow([sheetName, skipTop, skipBottom, skipLeft, skipRight, replacementsStr, timestamp]);
  }
  
  return { success: true };
}

/**
 * Commits the imported data to the specified sheet.
 */
function commitImport(data, targetSheetName, isNewSheet, method) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Could not access active spreadsheet.");

    if (targetSheetName.startsWith('_')) {
      throw new Error(`Sheet name "${targetSheetName}" is reserved and cannot be used as an import target.`);
    }
    
    let sheet;
    if (isNewSheet) {
      sheet = ss.getSheetByName(targetSheetName);
      if (sheet) {
        throw new Error(`Sheet "${targetSheetName}" already exists. Please choose a different name or set "Existing Sheet".`);
      }
      sheet = ss.insertSheet(targetSheetName);
    } else {
      sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) {
        throw new Error(`Target sheet "${targetSheetName}" not found.`);
      }
      if (method === 'overwrite') {
        sheet.clear();
      }
    }
    
    if (data && data.length > 0) {
      const rows = data.length;
      const cols = data[0].length;
      let startRow = 1;
      
      if (!isNewSheet && method === 'append') {
        const lastRow = sheet.getLastRow();
        startRow = lastRow > 0 ? lastRow + 1 : 1;
      }
      
      // Normalize row lengths to match `cols` and enforce 50,000 char limits
      for (let r = 0; r < data.length; r++) {
        // Prevent array mismatch errors (e.g. PapaParse returning empty lines as arrays of length 1)
        if (data[r].length < cols) {
          const diff = cols - data[r].length;
          data[r].push(...Array(diff).fill(""));
        } else if (data[r].length > cols) {
          data[r] = data[r].slice(0, cols);
        }

        for (let c = 0; c < cols; c++) {
          let cellVal = data[r][c];
          if (typeof cellVal === 'string' && cellVal.length > 50000) {
            data[r][c] = cellVal.substring(0, 50000);
          }
        }
      }

      const range = sheet.getRange(startRow, 1, rows, cols);
      range.setValues(data);
    }
    
    return { success: true, message: `Successfully imported ${data.length} rows to "${targetSheetName}".` };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
