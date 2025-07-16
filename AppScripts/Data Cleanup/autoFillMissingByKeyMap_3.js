/**
 * This script provides utility functions for Google Sheets:
 * 
 * 1. This script is designed to auto-fill missing values in a specified column based on a key column.
 *      - It uses a mapping of existing key-value pairs to fill in missing values.
 *      - The target sheet, key column, and value column are configurable.
 *      - It highlights the filled cells with a specified color.
 *      - It provides user feedback through toasts and alerts.
 * 
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Cleanup Tools 🚀')
    // .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    // .addItem('🧹 Clean & Normalize ICP Lead Data', 'cleanAndAggregateLeads')
    .addItem('🔁 Autofill missing by key', 'autoFillMissingByKeyMap')
    .addToUi();
}

function autoFillMissingByKeyMap() {
  const config = {
    targetSheetName: 'Lead_CleanedData',
    keyColumnName: 'Company City',
    valueColumnName: 'Company Country',
    highlightColor: '#f4cccc'
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Auto-filling missing values based on key match...', '⚠️ Attention ⚠️', -1);
  const sheet = ss.getSheetByName(config.targetSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert(`❗ Sheet '${config.targetSheetName}' not found.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyColIndex = headers.indexOf(config.keyColumnName);
  const valueColIndex = headers.indexOf(config.valueColumnName);

  if (keyColIndex === -1 || valueColIndex === -1) {
    ui.alert(`❗ One or both columns not found:\n- ${config.keyColumnName}\n- ${config.valueColumnName}`);
    return;
  }

  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
  const data = dataRange.getValues();

  // STEP 1: Build lookup from existing key → value
  const keyToValueMap = {};
  data.forEach(row => {
    const key = row[keyColIndex]?.toString().trim();
    const value = row[valueColIndex]?.toString().trim();
    if (key && value && !keyToValueMap[key]) {
      keyToValueMap[key] = value;
    }
  });

  // STEP 2: Fill missing values
  const filledRows = [];
  data.forEach((row, i) => {
    const key = row[keyColIndex]?.toString().trim();
    const value = row[valueColIndex];
    if (key && (!value || value.toString().trim() === '') && keyToValueMap[key]) {
      row[valueColIndex] = keyToValueMap[key];
      filledRows.push(i + 2); // row index to highlight (add 2 for header + 1-based index)
    }
  });

  // ✅ Update the modified cells
  dataRange.setValues(data);

  if (filledRows.length > 0) {
    // Highlight cells that were auto-filled
    const highlightRanges = filledRows.map(r => sheet.getRange(r, valueColIndex + 1));
    sheet.getRangeList(highlightRanges.map(r => r.getA1Notation()))
      .setBackground(config.highlightColor);
  }

  // ✅ Toast & Alert
  ss.toast('Auto-fill completed ✅', '✅ Success ✅', -1);
  ui.alert(`✅ Auto-filled ${filledRows.length} missing '${config.valueColumnName}' values based on '${config.keyColumnName}' match.`)
}
