/**
 * This script provides utility functions for Google Sheets:
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Cleanup Tools 🚀')
    // .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    .addItem('➕ Fill Missing Country from CityStateCountryMapping sheet ➕', 'fillMissingCountryFromMapping')
    .addToUi();
}

/**
 * Normalizes geographic data in the 'Lead_CleanedData' sheet by:
 * - Fixing swapped placements between City, State, and Country.
 * - Filling missing State, Country, and Region from the 'CityStateCountryRegionMapping' master sheet.
 *
 * 🔄 Dynamically corrects:
 * - City ↔ Country
 * - City ↔ State
 * - State ↔ Country
 * - All 3 if mixed
 *
 * 🚫 Skips update if all geo values are missing
 * ✅ Logs how many rows were updated and how many values were swapped.
 */
function fixCityStateCountrySwapAndFillFromMapping() {
  const config = {
    leadSheetName: 'Lead_CleanedData',
    mappingSheetName: 'CityStateCountryRegionMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region',
    },
    highlightFilledCells: true,         // Highlight cells corrected with this color
    highlightColor: '#f4cccc',           // Light red highlight (change if you want)
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Fetch sheets
  const leadSheet = ss.getSheetByName(config.leadSheetName);
  const mappingSheet = ss.getSheetByName(config.mappingSheetName);

  // Guard: Sheets must exist
  if (!leadSheet || !mappingSheet) {
    ui.alert(`❗ Missing required sheet(s): '${config.leadSheetName}' or '${config.mappingSheetName}'. Aborting.`);
    return;
  }

  ss.toast('Geo normalization started...', '⏳ Processing city/state/country/region', -1);

  // Get header rows for both sheets
  const leadHeaders = leadSheet.getRange(1, 1, 1, leadSheet.getLastColumn()).getValues()[0];
  const mapHeaders = mappingSheet.getRange(1, 1, 1, mappingSheet.getLastColumn()).getValues()[0];

  // Map lead sheet column indexes
  const leadCityIdx = leadHeaders.indexOf(config.columns.city);
  const leadStateIdx = leadHeaders.indexOf(config.columns.state);
  const leadCountryIdx = leadHeaders.indexOf(config.columns.country);
  const leadRegionIdx = leadHeaders.indexOf(config.columns.region);

  // Map mapping sheet column indexes
  const mapCityIdx = mapHeaders.indexOf(config.columns.city);
  const mapStateIdx = mapHeaders.indexOf(config.columns.state);
  const mapCountryIdx = mapHeaders.indexOf(config.columns.country);
  const mapRegionIdx = mapHeaders.indexOf(config.columns.region);

  // Guard: Columns must be found
  if ([leadCityIdx, leadStateIdx, leadCountryIdx, leadRegionIdx,
       mapCityIdx, mapStateIdx, mapCountryIdx, mapRegionIdx].some(i => i === -1)) {
    ui.alert(`❗ One or more required columns missing in either '${config.leadSheetName}' or '${config.mappingSheetName}'.`);
    return;
  }

  // Load mapping data into sets and a cityMap for lookups
  const citySet = new Set();
  const stateSet = new Set();
  const countrySet = new Set();
  const cityMap = new Map();

  const mapData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, mappingSheet.getLastColumn()).getValues();

  mapData.forEach(row => {
    const city = (row[mapCityIdx] || '').toString().trim();
    const state = (row[mapStateIdx] || '').toString().trim();
    const country = (row[mapCountryIdx] || '').toString().trim();
    const region = (row[mapRegionIdx] || '').toString().trim();

    if (city) citySet.add(city.toLowerCase());
    if (state) stateSet.add(state.toLowerCase());
    if (country) countrySet.add(country.toLowerCase());

    if (city && !cityMap.has(city.toLowerCase())) {
      cityMap.set(city.toLowerCase(), { city, state, country, region });
    }
  });

  // Fetch lead data rows
  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();

  // Track counts and changed cell positions for highlighting
  let fixCount = 0;
  let swapped = 0;
  const cellChanges = [];

  // Process each lead row
  for (let i = 0; i < leadData.length; i++) {
    let rawCity = (leadData[i][leadCityIdx] || '').toString().trim();
    let rawState = (leadData[i][leadStateIdx] || '').toString().trim();
    let rawCountry = (leadData[i][leadCountryIdx] || '').toString().trim();
    let rawRegion = (leadData[i][leadRegionIdx] || '').toString().trim();

    let cityVal = rawCity;
    let stateVal = rawState;
    let countryVal = rawCountry;
    let changed = false;

    // Normalize keys for lookup
    const rawCityKey = rawCity.toLowerCase();
    const rawStateKey = rawState.toLowerCase();
    const rawCountryKey = rawCountry.toLowerCase();

    // Detect actual role of each value by presence in sets
    const detected = {
      city: citySet.has(rawCityKey) ? 'city' :
            stateSet.has(rawCityKey) ? 'state' :
            countrySet.has(rawCityKey) ? 'country' : null,
      state: citySet.has(rawStateKey) ? 'city' :
             stateSet.has(rawStateKey) ? 'state' :
             countrySet.has(rawStateKey) ? 'country' : null,
      country: citySet.has(rawCountryKey) ? 'city' :
               stateSet.has(rawCountryKey) ? 'state' :
               countrySet.has(rawCountryKey) ? 'country' : null,
    };

    const values = { city: rawCity, state: rawState, country: rawCountry };
    const roles = {}; // actualRole => currentField
    for (const field in detected) {
      const actualType = detected[field];
      if (actualType) roles[actualType] = field;
    }

    // If at least two fields detected, reassign values correctly
    if (Object.keys(roles).length >= 2) {
      cityVal = values[roles.city] || '';
      stateVal = values[roles.state] || '';
      countryVal = values[roles.country] || '';

      if (leadData[i][leadCityIdx] !== cityVal) {
        leadData[i][leadCityIdx] = cityVal;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadCityIdx + 1 });
        changed = true;
      }
      if (leadData[i][leadStateIdx] !== stateVal) {
        leadData[i][leadStateIdx] = stateVal;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadStateIdx + 1 });
        changed = true;
      }
      if (leadData[i][leadCountryIdx] !== countryVal) {
        leadData[i][leadCountryIdx] = countryVal;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadCountryIdx + 1 });
        changed = true;
      }
      swapped++;
    }

    // Fill missing info from the master mapping (by city)
    const cityKey = cityVal.toLowerCase();
    if (cityMap.has(cityKey)) {
      const mapEntry = cityMap.get(cityKey);

      if (mapEntry.state && leadData[i][leadStateIdx] !== mapEntry.state) {
        leadData[i][leadStateIdx] = mapEntry.state;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadStateIdx + 1 });
        changed = true;
      }
      if (mapEntry.country && leadData[i][leadCountryIdx] !== mapEntry.country) {
        leadData[i][leadCountryIdx] = mapEntry.country;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadCountryIdx + 1 });
        changed = true;
      }
      if (mapEntry.region && leadData[i][leadRegionIdx] !== mapEntry.region) {
        leadData[i][leadRegionIdx] = mapEntry.region;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadRegionIdx + 1 });
        changed = true;
      }
      if (mapEntry.city && leadData[i][leadCityIdx] !== mapEntry.city) {
        leadData[i][leadCityIdx] = mapEntry.city;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadCityIdx + 1 });
        changed = true;
      }
    }

    if (changed) fixCount++;
  }

  // Write back corrected data to sheet
  leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);

  // Highlight fixed cells if configured
  if (config.highlightFilledCells && cellChanges.length > 0) {
    cellChanges.forEach(change => {
      leadSheet.getRange(change.row, change.col).setBackground(config.highlightColor);
    });
  }

  // Final UI notifications
  ui.alert(`✅ Geo normalization done!\nRows changed: ${fixCount}\nSwaps applied: ${swapped}`);
  ss.toast(`Geo normalization complete for ${fixCount} rows`, 'Success', 4);
}