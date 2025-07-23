/**
 * This script provides utility functions for Google Sheets:
 * 
 * 1. Freezes the first row in all sheets of the active spreadsheet.
 * 
 * 2. Cleans and aggregates leads data based on specified headers.
 *    - It collects data from all valid sheets, excluding those with "❌" in their name or specified in an exclusion list.
 *    - It normalizes the data according to a predefined set of headers.
 *    - region mapping is applied based on company country.
 *    - The output is placed in a new sheet named "Lead_CleanedData".
 * 
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Cleanup Tools 🚀')
    .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    .addItem('🧹 Clean & Normalize ICP Lead Data', 'cleanAndNormalize_ICP_Lead_Data')
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

// 2. Clean and aggregate leads data based on specified headers
function cleanAndNormalize_ICP_Lead_Data() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Cleaning & Normalizing Lead Data...', '⚠️ Attention ⚠️', -1);

  const ui = SpreadsheetApp.getUi();
  const excludedNames = ['CombinedData', 'Analytics', 'Filter_CombinedData', 'All_Sheet_Columns', 'Lead_CleanedData'];

  const desiredHeaders = [
    'Email',
    'First Name',
    'Last Name',
    'Title',
    'Seniority',
    'Departments',
    'Person Linkedin Url',
    'Company Name',
    '# Employees',
    'Industry',
    'Company Country',
    'Company State',
    'Company City',
    'Region'
  ];

  // Country ➝ Region Mapping
  const regionMap = {
    'India': 'APAC',
    'Singapore': 'APAC',
    'Japan': 'APAC',
    'Germany': 'EMEA',
    'France': 'EMEA',
    'United Kingdom': 'EMEA',
    'USA': 'North America',
    'United States': 'North America',
    'Brazil': 'LATAM',
    'Mexico': 'LATAM',
    'Canada': 'North America',
    'Australia': 'Oceania'
    // ➕ Add more as you need
  };

  // Collect data
  const allSheets = ss.getSheets().filter(s =>
    !excludedNames.includes(s.getName()) &&
    !s.getName().includes('❌')
  );

  const combinedData = [];

  allSheets.forEach(sheet => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

    // Build index map to locate fields
    const indexMap = {};
    headers.forEach((header, i) => {
      if (!header) return;
      const cleanedHeader = header.toString().trim().toLowerCase();
      if (cleanedHeader === 'job title') header = 'Title';
      indexMap[header.toLowerCase()] = i;
    });

    data.forEach(row => {
      if (!row.length) return;
      const newRow = [];

      desiredHeaders.forEach(col => {
        let val = '';
        const colKey = col.toLowerCase();
        if (col === 'Company Name') {
          val = row[indexMap['company']] || row[indexMap['company name']] || '';
        } else if (col === 'Region') {
          const country = row[indexMap['company country']] || '';
          val = regionMap[country.trim()] || '';
        } else {
          const headerAlias = indexMap[colKey];
          val = typeof headerAlias !== 'undefined' ? row[headerAlias] : '';
        }
        newRow.push(val);
      });

      // Include only if email exists
      if (newRow[0] && typeof newRow[0] === 'string') {
        combinedData.push(newRow);
      }
    });
  });

  if (combinedData.length === 0) {
    ss.toast('Cleaning Skipped! No valid emails found.', '❗ Empty Output', -1);
    ui.alert("❗ No valid records with 'Email' found.");
    return;
  }

  // Write to new cleaned sheet
  let cleanedSheet = ss.getSheetByName('Lead_CleanedData');
  if (cleanedSheet) {
    cleanedSheet.clearContents();
  } else {
    cleanedSheet = ss.insertSheet('Lead_CleanedData');
  }

  cleanedSheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
  cleanedSheet.getRange(2, 1, combinedData.length, desiredHeaders.length).setValues(combinedData);
  cleanedSheet.setFrozenRows(1);

  ss.toast('Lead Cleanup - COMPLETED ✅', '✅ Success ✅', -1);
  ui.alert(`✅ ${combinedData.length} unique rows successfully extracted and cleaned in 'Lead_CleanedData'`);
}

