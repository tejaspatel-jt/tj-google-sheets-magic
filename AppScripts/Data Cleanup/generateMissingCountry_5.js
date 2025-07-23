/**
 * This script provides utility functions for Google Sheets:
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Cleanup Tools 🚀')
    // .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    .addItem('⁉️ Find Missing Country - from State & City ⁉️', 'generateMissingCountryReport')
    .addToUi();
}

/**
 * This function generates a report of missing 'Company Country' in the 'Lead_CleanedData' sheet.
 *    - It checks for missing country based on the presence of 'Company City' and 'Company State'.
 *    - It creates a new sheet named 'MissingCountry' or clears it if it already exists.
 *    - It populates the sheet with rows where 'Company Country' is missing.
 *    - Rows with both city and state missing are ignored.
 *    - Rows where state exists but city is missing can be excluded via configuration.
 *    - Includes rows where city is present but state is missing, or both city and state are present.
 *    - Configuration options:
 *          - `targetSheetName`: Name of the source sheet to check for missing country.
 *          - `outputSheetName`: Name of the output sheet where results will be written.
 *          - `columns`: Object defining the column names for city, state, and country.
 *          - `excludeStateMissingCity`: Boolean to exclude rows where state is present but city is missing.
 * 
 * @returns 
 */
function generate_Missing_Country_By_City_And_State() {
  const config = {
    targetSheetName: 'Lead_CleanedData',
    outputSheetName: 'MissingCountry',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
    },
    excludeStateMissingCity: true // true to ignore rows where state present but city missing
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheet = ss.getSheetByName(config.targetSheetName);
  if (!sheet) {
    ui.alert(`❗ Sheet '${config.targetSheetName}' not found.`);
    return;
  }

  // Toast for process started
  ss.toast('Generating Missing Country report...', '⚠️ Attention ⚠️', -1);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const cityColIdx = headers.indexOf(config.columns.city);
  const stateColIdx = headers.indexOf(config.columns.state);
  const countryColIdx = headers.indexOf(config.columns.country);

  if (cityColIdx === -1 || stateColIdx === -1 || countryColIdx === -1) {
    ui.alert(`❗ One or more required columns not found: ${Object.values(config.columns).join(', ')}`);
    return;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  const missingRows = [];

  data.forEach(row => {
    const city = (row[cityColIdx] || '').toString().trim();
    const state = (row[stateColIdx] || '').toString().trim();
    const country = (row[countryColIdx] || '').toString().trim();

    if (!country) { // country missing
      // Ignore rows where state AND city both missing
      if (!state && !city) return;

      // If config set to ignore state present but city missing rows
      if (config.excludeStateMissingCity && state && !city) return;

      // In all other cases (city present, or both state+city), include
      missingRows.push([city, state, '']); // country is empty by definition
    }
  });

  // Prepare output sheet
  let outputSheet = ss.getSheetByName(config.outputSheetName);
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet(config.outputSheetName);
  }

  // Write headers and data
  outputSheet.getRange(1, 1, 1, 3).setValues([[config.columns.city, config.columns.state, config.columns.country]]);
  if (missingRows.length > 0) {
    outputSheet.getRange(2, 1, missingRows.length, 3).setValues(missingRows);
  }

  outputSheet.setFrozenRows(1);

  // Completion toast & alert
  ss.toast('Missing Country report generated ✅', '✅ Success ✅', -1);
  ui.alert(`✅ ${missingRows.length} rows found with missing 'Country'.\nOutput is in '${config.outputSheetName}' sheet.`);
}
