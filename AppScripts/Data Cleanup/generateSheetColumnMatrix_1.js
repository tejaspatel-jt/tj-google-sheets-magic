/**
 * This script provides utility functions for Google Sheets:
 * 
 * 1. Freezes the first row in all sheets of the active spreadsheet.
 * 2. Generates a matrix of all columns from valid sheets, excluding those with "❌" in their name or specified in an exclusion list.
 *    - The output is placed in a new sheet named "All_Sheet_Columns".
 *    - The script collects all unique headers from the sheets and aligns the data accordingly.
 *    - It skips empty sheets and those with names containing "CombinedData" or "Old".
 *    - The first column of the output contains the sheet names.
 *    - from 2nd row onwards, it lists the headers of each sheet.
 * 
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Cleanup Tools 🚀')
    .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    .addItem('📋 Get Columns of All Sheet excluding ❌', 'generateSheetColumnMatrix')
    .addToUi();
}

// 1. Freeze first row in all sheets
function freezeFirstRowInAllSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    sheets[i].setFrozenRows(1);
  }
}

function generateSheetColumnMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ✅ SHOW START TOAST
  ss.toast('Extracting column headers from all sheets...', '⚠️ Attention ⚠️', -1);

  // Define excluded sheet names explicitly
  const excludedSheets = ["CombinedData", "Analytics", "Filter_CombinedData", "All_Sheet_Columns"];

  // Filter valid sheets
  const validSheets = ss.getSheets().filter(sheet => {
    const name = sheet.getName();
    return !name.includes("❌") &&
           !excludedSheets.includes(name) &&
           !name.toLowerCase().includes("missing") &&
           !name.toLowerCase().includes("combineddata") &&
           !name.toLowerCase().includes("old");
  });

  // Create or clear new output sheet
  let outputSheet = ss.getSheetByName("All_Sheet_Columns");
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet("All_Sheet_Columns");
  }

  // Start from first row
  let currentRow = 1;

  validSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return; // skip empty sheets

    // Read header row (first row)
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Write: first cell - sheet name, followed by header columns
    const rowData = [sheetName, ...headers];
    outputSheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);

    currentRow++;
  });

  outputSheet.setFrozenRows(1);

  // ✅ SHOW COMPLETE TOAST
  ss.toast('Sheet column scan - COMPLETED ✅', '✅ Success ✅', -1);

  ui.alert("✅ All sheet columns extracted successfully to 'All_Sheet_Columns' sheet!");
}
