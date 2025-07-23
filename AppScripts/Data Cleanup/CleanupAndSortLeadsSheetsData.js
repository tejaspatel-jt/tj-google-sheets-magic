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
    .addItem('1️⃣ Freeze First Row in All Sheets ✅', 'freezeFirstRowInAllSheets')
    .addItem('2️⃣ 📋 Get Columns of All Sheet excluding ❌ and Specified ✅', 'generateSheetColumnMatrix')
    .addItem('3️⃣ 🧹 Clean & Normalize ICP Lead Data ✅', 'cleanAndNormalize_ICP_Lead_Data')
    .addItem('4️⃣ ➕➕ Fill City State Country Region Very Independently and Flexibly Roubust 💫➕➕✅', 'fillCityStateCountrySwapAndFillFromMapping_Master_Independently_Flexibly_Robust')
    .addItem('5️⃣ ➕➕ 🧠 Categorize Scattered Job Titles into Sorted Designations ✅', 'categorizeScatteredTitlesIntoDesignation')
    .addItem('6️⃣ ➕➕ Normalize Regions in Region Column ➕➕ ✅', 'normalizeRegionsInSheet')
    .addItem('➕➕ Fix CityStateCountry SwapAndFill From Mapping CityStateCountryRegionMapping', 'fixCityStateCountrySwapAndFillFromMapping')
    .addItem('➕➕ Generate Master Mapping Sheet from Existing Mapping and Newly Created Missing Geo Mapping ➕➕✅', 'generateMasterMappingFromExistingAndNewMapping')
    .addItem('➕➕ Fill City State Country Region Flexibly ➕➕✅', 'fillCityStateCountrySwapAndFillFromMapping_Master_Flexibly')
    .addItem('➕➕ Fill City State Country Region Very Independently and Flexibly ➕➕✅', 'fillCityStateCountrySwapAndFillFromMapping_Master_Independently_Flexibly')
    .addItem('🌍 Extract City-State-Country-Region Lookup', 'extractGeoMappings')
    .addItem('🔁 Autofill missing by key', 'autoFillMissingByKeyMap')
    .addItem('❓ Find Missing Details - Duo', 'generateMissingDetailsReport')
    .addItem('❓ Find Missing Details - Triplets', 'generateMissingDetailsAdvancedReport')
    .addItem('⁉️ Find Missing Country - from State & City ⁉️', 'generateMissingCountryReport')
    .addItem('➕ Fill Missing Country from CityStateCountryMapping sheet ➕', 'fillMissingCountryFromMapping')
    .addItem('➕ Generate City-State-Country-Region Master Mapping ➕', 'generateMasterMapping')
    .addItem('➕➕ Fill Missing Country&Region from CityStateCountryRegionMapping sheet ➕➕ ✅', 'fillMissingCountryRegionFromMapping')
    .addItem('➕➕ correct City Country Mapping ✅', 'correctCityCountryMapping')
    .addItem('➕➕ Correct citystatecountryregion in LeadCleanedData From right CityStateCountryRegionMapping sheet ❌', 'fixLeadCleanedDataFromMapping')
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

// 2. Generate a matrix of column headers from all valid sheets to clean up and sort Headers
/**
 * Scans all valid Google Sheets (excluding defined or partially matched ones)
 * and extracts column headers (first row) into a central sheet 'All_Sheet_Columns'.
 *
 * 🧩 Useful for auditing headers across your workspace.
 *
 * 🛠️ Configuration:
 * - `excludedSheets`: Exact names to skip.
 * - `excludedNameIncludes`: Substring matches to exclude sheet names dynamically.
 * - `outputSheetName`: The target sheet where output is written.
 *
 * ✅ Includes: Toasts, Alerts, Configurable filtering.
 */
function generateSheetColumnMatrix() {
  const config = {

    // Define excluded sheet names explicitly
    excludedSheets: [
      'CombinedData',
      'Analytics',
      'Filter_CombinedData',
      'All_Sheet_Columns'
    ],
    excludedNameIncludes: [
      '❌',
      'missing',
      'combineddata',
      'old'
    ],
    outputSheetName: 'All_Sheet_Columns'
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ✅ START TOAST
  ss.toast('Extracting column headers from sheets...', '⏳ Scanning', -1);

  // Filter valid sheets
  const validSheets = ss.getSheets().filter(sheet => {
    const name = sheet.getName();
    const nameLower = name.toLowerCase();

    const isExcludedExact = config.excludedSheets.some(ex =>
      ex.toLowerCase() === nameLower
    );

    const isExcludedPartial = config.excludedNameIncludes.some(keyword =>
      nameLower.includes(keyword)
    );

    return !isExcludedExact && !isExcludedPartial;
  });

  // Prepare output sheet
  let outputSheet = ss.getSheetByName(config.outputSheetName);
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet(config.outputSheetName);
  }

  // Start from first row
  let rowIndex = 1;

  validSheets.forEach(sheet => {
    const name = sheet.getName();
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return; // Skip empty

    // Read header row (first row)
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Write: first cell - sheet name, followed by header columns
    const row = [name, ...headers];

    outputSheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    rowIndex++;
  });

  // outputSheet.setFrozenRows(1);

  // 🔒 Freeze first column
  sheet.setFrozenColumns(1);

  // ✅ COMPLETE TOAST + ALERT
  ss.toast('Header scan completed ✅', '✅ Done', 3);
  ui.alert(
    `✅ Sheet header scan completed.\n\n📄 Output Sheet: '${config.outputSheetName}'\n📊 Sheets processed: ${validSheets.length}`
  );
}


// 3. Clean and aggregate leads data based on specified headers
/**
 * Cleans and normalizes raw leads data across multiple source sheets into a single
 * standardized format and outputs it to a clean tab called "Lead_CleanedData".
 *
 * 🔍 What it does:
 * - Loops through all source sheets (excluding specified ones)
 * - Reads data from every sheet with flexible column detection
 * - Respects and builds rows based on `desiredHeaders` column structure
 * - Keeps only unique rows based on Email
 * - Ensures final output in a new sheet ('Lead_CleanedData')
 * - Optionally logs skipped records with missing/duplicate info
 *
 * 💡 Configuration:
 * - `excludedSheets`: List of sheet names to skip
 * - `desiredHeaders`: Final columns along with their accepted aliases in source sheets
 * - `outputSheetName`: Final output tab ("Lead_CleanedData")
 * - `logOptions`: Controls tooltip tracking and skipped rows logging
 *
 * ✅ Includes toasts, alerts, and optional skipped row tracking
 */
function cleanAndNormalize_ICP_Lead_Data() {
  const config = {
    excludedSheets: ['CombinedData', 'Analytics', 'Filter_CombinedData', 'All_Sheet_Columns', 'Lead_CleanedData'],
    outputSheetName: 'Lead_CleanedData',

    // Define desired headers with aliases for flexible matching column names
    desiredHeaders: [
      { output: 'Email', aliases: [] },
      { output: 'First Name', aliases: ['First Name'] },
      { output: 'Last Name', aliases: ['Last Name'] },
      { output: 'Title', aliases: ['Job Title'] },
      { output: 'Seniority', aliases: ['Seniority'] },
      { output: 'Departments', aliases: ['Departments'] },
      { output: 'Person Linkedin Url', aliases: ['Person Linkedin Url'] },
      { output: 'Company Name', aliases: ['Company'] },
      { output: '# Employees', aliases: ['# Employees', 'Employees'] },
      { output: 'Size of Company', aliases: ['Company number of employees'] },
      { output: 'Industry', aliases: ['Company industry', 'Company main industry'] },
      { output: 'Company City', aliases: ['Company City', 'City'] },
      { output: 'Company State', aliases: ['Company State', 'State'] },
      { output: 'Company Country', aliases: ['Company Country', 'Country'] },
      { output: 'Region', aliases: ['Region'] } // Keep for placeholder if Region exists
    ],

    logOptions: {
      wantUnprocessedDataInfo: false,               // 🔄 Master toggle to track skipped/broken rows
      includeIssueColumn: true,                    // ❔ Add 'Issues' column showing what's missing
      exportSkippedRows: true,                     // ⚠️ Push skipped rows to separate tab
      skippedSheetName: 'Skipped_Lead_Rows',       // 📄 Tab name for skipped rows
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ✅ Toast: Start
  ss.toast('Cleaning & Normalizing ICP Lead Data...', '⚙️ Processing...', -1);

  // 🔍 Filter valid sheets (excluding system or configured excluded ones)
  const allSheets = ss.getSheets().filter(s =>
    !config.excludedSheets.includes(s.getName()) &&
    !s.getName().includes('❌')
  );

  const emailSet = new Set();    // To track unique emails
  const combinedData = [];       // Final data result
  const skippedData = [];        // Info for skipped/invalid rows

  allSheets.forEach(sheet => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
   
    const totalRows = sheet.getLastRow();

    // If No data row or only headers, skip processing - HANDLE EMPTY SHEET
    if (totalRows <= 1) return;

    const data = sheet.getRange(2, 1, totalRows - 1, headers.length).getValues();

    const indexMap = {};

    // 🔁 Build indexMap: Map each desired header to the column index in current sheet
    config.desiredHeaders.forEach(headerGroup => {
      const outputKey = headerGroup.output.toLowerCase();

      // First: try exact match
      let foundIdx = headers.findIndex(
        h => h && h.toString().trim().toLowerCase() === outputKey
      );

      // Fallback: search through aliases (if defined)
      if (foundIdx === -1 && Array.isArray(headerGroup.aliases)) {
        for (const alias of headerGroup.aliases) {
          foundIdx = headers.findIndex(
            h => h && h.toString().trim().toLowerCase() === alias.toLowerCase()
          );
          if (foundIdx !== -1) break;
        }
      }

      if (foundIdx !== -1) {
        indexMap[outputKey] = foundIdx;
      }
    });

    // 🔁 Process each row of data
    data.forEach(row => {
      if (!row || row.length === 0) return; // completely empty row
    
      const outputRow = [];
      const missingHeaders = [];
      let issueText = '';
      let skipReason = '';
    
      // 📤 Build structured row from mapped columns
      config.desiredHeaders.forEach(headerGroup => {
        const key = headerGroup.output.toLowerCase();
        const idx = indexMap[key];
        const value = typeof idx !== 'undefined' ? row[idx] : '';
      
        // Optionally collect missing headers for issue tracking
        if (config.logOptions.wantUnprocessedDataInfo && value === '') {
          missingHeaders.push(headerGroup.output);
        }
      
        outputRow.push(value);
      });
    
      // ✅ Add 'Issues' column (if enabled), but do NOT use it to skip row
      if (config.logOptions.wantUnprocessedDataInfo && config.logOptions.includeIssueColumn) {
        issueText = missingHeaders.length > 0 ? `Missing: ${missingHeaders.join(', ')}` : '';
        outputRow.push(issueText);
      }
    
      // 📌 Email Validations (the only place we SKIP)
      const rawEmail = outputRow[0];
      const email = rawEmail && typeof rawEmail === 'string' ? rawEmail.trim().toLowerCase() : '';
    
      if (!email || emailSet.has(email)) {
        if (config.logOptions.wantUnprocessedDataInfo && config.logOptions.exportSkippedRows) {
          skipReason = !email ? '❌ Missing Email' : '⚠️ Duplicate Email';
        
          const skippedRow = [...outputRow];
        
          // Include issueText if column was already added
          if (config.logOptions.includeIssueColumn) {
            skippedRow.push(issueText);
          }
        
          skippedRow.push(skipReason);
          skippedData.push(skippedRow);
        }
        return; // Skip pushing to main data
      }
    
      // ✅ Add row to result
      emailSet.add(email);
      combinedData.push(outputRow);
    });

  });

  // ⛔ If no usable rows found
  if (combinedData.length === 0) {
    ss.toast('No valid emails found – Skipped.', '⚠️ Empty Output', -1);
    ui.alert("❗ No valid unique records with 'Email' found.");
    return;
  }

  // 📝 Export skipped rows (if enabled via config)
  if (
    config.logOptions.wantUnprocessedDataInfo &&
    config.logOptions.exportSkippedRows &&
    skippedData.length
  ) {
    writeSkippedLeadRows(skippedData, config);
  }

  // 📄 Create or clear output sheet and insert cleaned data
  let outSheet = ss.getSheetByName(config.outputSheetName);
  if (outSheet) {
    outSheet.clearContents();
  } else {
    outSheet = ss.insertSheet(config.outputSheetName);
  }

  // 🧾 Set headers and data rows
  const headers = config.desiredHeaders.map(h => h.output);

  // ✅ Add 'Issues' column if tracking enabled
  if (config.logOptions.wantUnprocessedDataInfo && config.logOptions.includeIssueColumn) headers.push('Issues');

  outSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ✅ Ensure each row in combinedData has exactly the same number of columns as the headers
  const finalColCount = headers.length;
  const correctedData = combinedData.map(row => {
    const fixed = [...row];
    while (fixed.length < finalColCount) fixed.push('');
    if (fixed.length > finalColCount) fixed.length = finalColCount;
    return fixed;
  });

  // outSheet.getRange(2, 1, combinedData.length, headers.length).setValues(combinedData);
  outSheet.getRange(2, 1, correctedData.length, finalColCount).setValues(correctedData);
  outSheet.setFrozenRows(1);

  // ✅ Final toast + Alert
  ss.toast(`Cleaned ${combinedData.length} unique leads ✅`, '✅ Cleaning Complete', 3);
  ui.alert(`✅ Cleaning complete!\n${combinedData.length} unique ICP lead rows saved in '${config.outputSheetName}'`);

  /**
   * Logs skipped rows (with reasons) into a separate sheet for review
   * Runs only when: config.logOptions.wantUnprocessedDataInfo === true
   * Ensures no mismatch between headers and row values.
   */
  function writeSkippedLeadRows(data, config) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = config.logOptions.skippedSheetName;

    ss.toast(`Writing skipped rows to '${sheetName}'...`, '📝 Writing Skipped Rows', -1);

    let skippedSheet = ss.getSheetByName(sheetName);
    if (skippedSheet) {
      skippedSheet.clearContents();
    } else {
      skippedSheet = ss.insertSheet(sheetName);
    }

    // Create headers based on config
    const headers = config.desiredHeaders.map(h => h.output);
    if (config.logOptions.includeIssueColumn) headers.push('Issues');
    headers.push('Skip Reason');

    // ✅ Normalize each row in skippedData to match header count
    const finalColCount = headers.length;
    const correctedSkippedData = data.map(row => {
      const fixed = [...row];
      while (fixed.length < finalColCount) fixed.push('');
      if (fixed.length > finalColCount) fixed.length = finalColCount;
      return fixed;
    });

    // ✍️ Write header + data
    skippedSheet.getRange(1, 1, 1, finalColCount).setValues([headers]);
    skippedSheet.getRange(2, 1, correctedSkippedData.length, finalColCount).setValues(correctedSkippedData);
    skippedSheet.setFrozenRows(1);

    ss.toast(`Skipped rows written to '${sheetName}' ✅`, '✅ Skipped Rows Written', 3);
  }

}

// 4. Fix City-State-Country-Region Mapping is Wrong by SwapAndFill
/**
 * 📍 Normalizes geographic data in 'Lead_CleanedData'
 * - Fixes swapped City, State, Country values (↔ any 2 or all 3)
 * - Fills missing State, Country, Region from 'CityStateCountryRegionMapping'
 * - 📝 Optionally logs unmatched/geographically incomplete rows in 'Missing_GeoMapping'
 * - 🎨 Optionally highlights corrected cells
 * - 
 * 
 * 🔄 Dynamically corrects:
 * - City ↔ Country
 * - City ↔ State
 * - State ↔ Country
 * - All 3 if mixed
 * 
 * ⚙️ Config:
 *   leadSheetName          - name of lead data sheet
 *   mappingSheetName       - name of master geo mapping sheet
 *   columns                - column names to map
 *   highlightFilledCells   - toggle cell highlight on corrections
 *   highlightColor         - highlight color hex (e.g., red, yellow)
 *   trackMissingGeoRows    - enable creating 'Missing_GeoMapping' sheet with unmatched city data
 *   trackMissingGeoRowsSheetName - name of sheet to track missing geo rows
 *   clearMissingGeoMappingBeforeAppend - clear existing 'Missing_GeoMapping' before appending new data
 *
 * 🚫 Skips update if all geo values are empty
 * ✅ Logs how many rows updated & how many swaps were applied
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
    highlightColor: '#f4cccc',          // Light red highlight (change if you want)
    trackMissingGeoRows: true,          // Collect partially missing rows in 'Missing_GeoMapping'
    trackMissingGeoRowsSheetName: 'Missing_GeoMapping',
    clearMissingGeoMappingBeforeAppend: true   // 🚨 Set to true to clear the sheet and write fresh
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
    ui.alert(`❗ One or more required geo columns not found in the sheet headers.`);
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

    const key = city.toLowerCase();
    if (city && !cityMap.has(key)) {
      cityMap.set(key, { city, state, country, region });
    }
  });

  // Fetch lead data rows
  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();

  // Track counts and changed cell positions for highlighting
  let fixCount = 0;
  let swapped = 0;
  const cellChanges = [];

  // For tracking new missing geo rows and existing missing records avoid duplicates
  const missingRecordsSet = new Set();
  const newMissingRows = [];

  if (config.trackMissingGeoRows) {
    const missSheet = ss.getSheetByName(config.trackMissingGeoRowsSheetName);
    if (missSheet && !config.clearMissingGeoMappingBeforeAppend) {
      const missData = missSheet.getDataRange().getValues();
      for (let r = 1; r < missData.length; r++) {
        const row = missData[r];
        const key = [
          (row[0] || '').toString().trim().toLowerCase(),
          (row[1] || '').toString().trim().toLowerCase(),
          (row[2] || '').toString().trim().toLowerCase(),
          (row[3] || '').toString().trim().toLowerCase()
        ].join('||');
        missingRecordsSet.add(key);
      }
    }
  }

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

    // Collect unmatched rows for 'Missing_GeoMapping' (optional)
    if (config.trackMissingGeoRows) {
      if ((rawCity || rawState || rawCountry) && (!rawCountry || !rawRegion)) {
        const missingKey = `${rawCityKey}||${rawStateKey}||${rawCountryKey}||${rawRegion.toLowerCase()}`;
        if (!missingRecordsSet.has(missingKey)) {
          const note = [
            !rawCountry && "Missing Country",
            !rawRegion && "Missing Region",
          ].filter(Boolean).join(', ');
          newMissingRows.push([rawCity, rawState, rawCountry, rawRegion, note]);
          missingRecordsSet.add(missingKey);
        }
      }
    }
  }

  // Write back corrected data to sheet
  leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);

  // Highlight fixed cells if configured
  if (config.highlightFilledCells && cellChanges.length > 0) {
    cellChanges.forEach(change => {
      leadSheet.getRange(change.row, change.col).setBackground(config.highlightColor);
    });
  }

  // Write to Missing_GeoMapping sheet if enabled
  if (config.trackMissingGeoRows && newMissingRows.length > 0) {
    let missingSheet = ss.getSheetByName(config.trackMissingGeoRowsSheetName);
    if (!missingSheet) {
      missingSheet = ss.insertSheet(config.trackMissingGeoRowsSheetName);
    } else if (config.clearMissingGeoMappingBeforeAppend) {
      missingSheet.clear();
    }

    // Write headers if missing or after clearing
    if (missingSheet.getLastRow() === 0) {
      missingSheet.appendRow([
        config.columns.city,
        config.columns.state,
        config.columns.country,
        config.columns.region,
        "Notes"
      ]);
    }
    const lastRow = missingSheet.getLastRow();
    missingSheet.getRange(lastRow + 1, 1, newMissingRows.length, 5).setValues(newMissingRows);
  }

  // Final UI notifications
  ui.alert(`✅ Geo normalization done!\nRows changed: ${fixCount}\nSwaps applied: ${swapped}`);
  ss.toast(`Geo normalization complete for ${fixCount} rows`, 'Success', -1);
}

// 5. Generate Master Mapping Sheet from Existing Mapping and Newly Created Missing Geo Mapping
/**
 * Combines 'CityStateCountryRegionMapping' and 'Missing_GeoMapping'
 * into a cleaned-up and deduplicated sheet called 'Master_Mapping'.
 *
 * Rules:
 * ✅ Include all from master mapping as-is
 * ✅ Add only unique rows from Missing_GeoMapping
 *    - Skip if full city/state/country/region combo already exists
 *    - If city/state are blank, only add if country+region not already present
 */
function generateMasterMappingFromExistingAndNewMapping() {
  const config = {
    mappingSheetName: 'CityStateCountryRegionMapping',
    missingSheetName: 'Missing_GeoMapping',
    outputSheetName: 'Master_Mapping',
    columns: ['Company City', 'Company State', 'Company Country', 'Region'],
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mappingSheet = ss.getSheetByName(config.mappingSheetName);
  const missingSheet = ss.getSheetByName(config.missingSheetName);

  if (!mappingSheet || !missingSheet) {
    SpreadsheetApp.getUi().alert('❗ Required input sheets not found.');
    return;
  }

  // Get data from master mapping sheet
  const mappingData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, config.columns.length).getValues();
  const combined = [...mappingData]; // Start with all original rows

  // Prepare lookup sets to detect duplicates
  const fullComboSet = new Set(); // city|state|country|region
  const countryRegionSet = new Set(); // country|region (for partials)

  mappingData.forEach(row => {
    const key = row.map(v => (v || '').toString().toLowerCase().trim()).join('|');
    const countryRegionKey = [row[2], row[3]].map(v => (v || '').toString().toLowerCase().trim()).join('|');

    fullComboSet.add(key);
    countryRegionSet.add(countryRegionKey);
  });

  // Get data from Missing_GeoMapping
  const missingData = missingSheet.getRange(2, 1, missingSheet.getLastRow() - 1, config.columns.length).getValues();

  missingData.forEach(row => {
    const cleaned = row.map(v => (v || '').toString().trim());
    const city = cleaned[0];
    const state = cleaned[1];
    const country = cleaned[2];
    const region = cleaned[3];

    const fullKey = cleaned.map(v => v.toLowerCase()).join('|');
    const countryRegionKey = [country.toLowerCase(), region.toLowerCase()].join('|');

    const isCityBlank = !city;
    const isStateBlank = !state;

    if (fullComboSet.has(fullKey)) return; // full duplicate → skip

    if (isCityBlank && isStateBlank) {
      if (!country || !region) return; // insufficient data → skip
      if (countryRegionSet.has(countryRegionKey)) return; // country+region duplicate → skip
    }

    // ✅ Append unique record
    combined.push(cleaned);
    fullComboSet.add(fullKey);
    if (country && region) countryRegionSet.add(countryRegionKey);
  });

  // Create/clear output sheet
  let outputSheet = ss.getSheetByName(config.outputSheetName);
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet(config.outputSheetName);
  }

  // Write header + data
  outputSheet.appendRow(config.columns);
  if (combined.length > 0) {
    outputSheet.getRange(2, 1, combined.length, config.columns.length).setValues(combined);
  }

  SpreadsheetApp.getUi().alert(`✅ Master_Mapping created.\nRows: ${combined.length}`);
  ss.toast('Master_Mapping generation complete ✅', 'Done', 4);
}

// 6. This Will Fill City, State, Country, Region from Mapping Independently if Other Data Missing Like City,State But Country Exists
/**
 * 
 * With Country→Region Fallback
 * 
 */
function fillCityStateCountrySwapAndFillFromMapping_Master_Flexibly() {
  const config = {
    leadSheetName: 'Lead_CleanedData',
    mappingSheetName: 'Master_GeoMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region',
    },
    highlightFilledCells: true,         // Highlight cells corrected with this color
    highlightColor: '#ff9195',          // Light red highlight (change if you want)
    trackMissingGeoRows: true,          // Collect partially missing rows in 'Missing_GeoMapping'
    trackMissingGeoRowsSheetName: 'Missing_GeoMapping',
    clearMissingGeoMappingBeforeAppend: true   // 🚨 Set to true to clear the sheet and write fresh
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
    ui.alert(`❗ One or more required geo columns not found in the sheet headers.`);
    return;
  }

  // Load mapping data into sets and a cityMap for lookups
  const citySet = new Set();
  const stateSet = new Set();
  const countrySet = new Set();
  const cityMap = new Map();
  const countryToRegionMap = new Map(); // 👉 NEW: Country → Region fallback map

  const mapData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, mappingSheet.getLastColumn()).getValues();

  mapData.forEach(row => {
    const city = (row[mapCityIdx] || '').toString().trim();
    const state = (row[mapStateIdx] || '').toString().trim();
    const country = (row[mapCountryIdx] || '').toString().trim();
    const region = (row[mapRegionIdx] || '').toString().trim();

    if (city) citySet.add(city.toLowerCase());
    if (state) stateSet.add(state.toLowerCase());
    if (country) countrySet.add(country.toLowerCase());

    const key = city.toLowerCase();
    if (city && !cityMap.has(key)) {
      cityMap.set(key, { city, state, country, region });
    }

    // 👉 MAP country to region if region exists
    const countryKey = country.toLowerCase();
    if (country && region && !countryToRegionMap.has(countryKey)) {
      countryToRegionMap.set(countryKey, region);
    }
  });

  // Fetch lead data rows
  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();

  // Track counts and changed cell positions for highlighting
  let fixCount = 0;
  let swapped = 0;
  const cellChanges = [];

  // For tracking new missing geo rows and existing missing records avoid duplicates
  const missingRecordsSet = new Set();
  const newMissingRows = [];

  if (config.trackMissingGeoRows) {
    const missSheet = ss.getSheetByName(config.trackMissingGeoRowsSheetName);
    if (missSheet && !config.clearMissingGeoMappingBeforeAppend) {
      const missData = missSheet.getDataRange().getValues();
      for (let r = 1; r < missData.length; r++) {
        const row = missData[r];
        const key = [
          (row[0] || '').toString().trim().toLowerCase(),
          (row[1] || '').toString().trim().toLowerCase(),
          (row[2] || '').toString().trim().toLowerCase(),
          (row[3] || '').toString().trim().toLowerCase()
        ].join('||');
        missingRecordsSet.add(key);
      }
    }
  }

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

    // ✅ NEW: Country → Region fallback if city/state empty & region missing
    if (!cityVal && !stateVal && !rawRegion && rawCountry) {
      const countryKey = rawCountry.toLowerCase();
      const resolvedRegion = countryToRegionMap.get(countryKey);
      if (resolvedRegion) {
        leadData[i][leadRegionIdx] = resolvedRegion;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadRegionIdx + 1 });
        changed = true;
      }
    }

    if (changed) fixCount++;

    // Collect unmatched rows for 'Missing_GeoMapping' (optional)
    if (config.trackMissingGeoRows) {
      if ((rawCity || rawState || rawCountry) && (!rawCountry || !rawRegion)) {
        const missingKey = `${rawCityKey}||${rawStateKey}||${rawCountryKey}||${rawRegion.toLowerCase()}`;
        if (!missingRecordsSet.has(missingKey)) {
          const note = [
            !rawCountry && "Missing Country",
            !rawRegion && "Missing Region",
          ].filter(Boolean).join(', ');
          newMissingRows.push([rawCity, rawState, rawCountry, rawRegion, note]);
          missingRecordsSet.add(missingKey);
        }
      }
    }
  }

  // Write back corrected data to sheet
  leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);

  // Highlight fixed cells if configured
  if (config.highlightFilledCells && cellChanges.length > 0) {
    cellChanges.forEach(change => {
      leadSheet.getRange(change.row, change.col).setBackground(config.highlightColor);
    });
  }

  // Write to Missing_GeoMapping sheet if enabled
  if (config.trackMissingGeoRows && newMissingRows.length > 0) {
    let missingSheet = ss.getSheetByName(config.trackMissingGeoRowsSheetName);
    if (!missingSheet) {
      missingSheet = ss.insertSheet(config.trackMissingGeoRowsSheetName);
    } else if (config.clearMissingGeoMappingBeforeAppend) {
      missingSheet.clear();
    }

    // Write headers if missing or after clearing
    if (missingSheet.getLastRow() === 0) {
      missingSheet.appendRow([
        config.columns.city,
        config.columns.state,
        config.columns.country,
        config.columns.region,
        "Notes"
      ]);
    }
    const lastRow = missingSheet.getLastRow();
    missingSheet.getRange(lastRow + 1, 1, newMissingRows.length, 5).setValues(newMissingRows);
  }

  // Final UI notifications
  ui.alert(`✅ Geo normalization done!\nRows changed: ${fixCount}\nSwaps applied: ${swapped}`);
  ss.toast(`Geo normalization complete for ${fixCount} rows`, 'Success', 4);
}

// 7. This Will Fill City, State, Country, Region from Mapping Independently if Other Data Missing Like City,State But Country Exists
/**
 * 
 * With "Country → Region" Fallback
 * With "State + Country → Region" Fallback
 * Adding Data in Missing Sheet before the Mapping - WRONG
 * 
 */
function fillCityStateCountrySwapAndFillFromMapping_Master_Independently_Flexibly() { // 7.
  const config = {
    leadSheetName: 'Lead_CleanedData',
    mappingSheetName: 'Master_GeoMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region',
    },
    highlightFilledCells: true,         // Highlight cells corrected with this color
    highlightColor: '#ff9195',          // Not So Light red highlight (change if you want)
    trackMissingGeoRows: true,          // Collect partially missing rows in 'Missing_GeoMapping'
    trackMissingGeoRowsSheetName: 'Missing_GeoMapping',
    clearMissingGeoMappingBeforeAppend: true   // 🚨 Set to true to clear the sheet and write fresh
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
    ui.alert(`❗ One or more required geo columns not found in the sheet headers.`);
    return;
  }

  // Load mapping data into sets and a cityMap for lookups
  const citySet = new Set();
  const stateSet = new Set();
  const countrySet = new Set();
  const cityMap = new Map();
  const countryToRegionMap = new Map();           // 👉 NEW: Country → Region fallback map
  const stateCountryToRegionMap = new Map();      // 👉 NEW: State + Country to region fallback map

  const mapData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, mappingSheet.getLastColumn()).getValues();

  mapData.forEach(row => {
    const city = (row[mapCityIdx] || '').toString().trim();
    const state = (row[mapStateIdx] || '').toString().trim();
    const country = (row[mapCountryIdx] || '').toString().trim();
    const region = (row[mapRegionIdx] || '').toString().trim();

    if (city) citySet.add(city.toLowerCase());
    if (state) stateSet.add(state.toLowerCase());
    if (country) countrySet.add(country.toLowerCase());

    const key = city.toLowerCase();
    if (city && !cityMap.has(key)) {
      cityMap.set(key, { city, state, country, region });
    }

    // 👉 NEW: Map country → region
    if (country && region && !countryToRegionMap.has(country.toLowerCase())) {
      countryToRegionMap.set(country.toLowerCase(), region);
    }

    // 👉 NEW: Map state+country → region
    if (state && country && region) {
      const scKey = `${state.toLowerCase()}|${country.toLowerCase()}`;
      if (!stateCountryToRegionMap.has(scKey)) {
        stateCountryToRegionMap.set(scKey, region);
      }
    }
  });

  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();
  let fixCount = 0;
  let swapped = 0;
  const cellChanges = [];

  const missingRecordsSet = new Set();
  const newMissingRows = [];

  if (config.trackMissingGeoRows) {
    const missSheet = ss.getSheetByName(config.trackMissingGeoRowsSheetName);
    if (missSheet && !config.clearMissingGeoMappingBeforeAppend) {
      const missData = missSheet.getDataRange().getValues();
      for (let r = 1; r < missData.length; r++) {
        const row = missData[r];
        const key = [
          (row[0] || '').toString().trim().toLowerCase(),
          (row[1] || '').toString().trim().toLowerCase(),
          (row[2] || '').toString().trim().toLowerCase(),
          (row[3] || '').toString().trim().toLowerCase()
        ].join('||');
        missingRecordsSet.add(key);
      }
    }
  }

  for (let i = 0; i < leadData.length; i++) {
    let rawCity = (leadData[i][leadCityIdx] || '').toString().trim();
    let rawState = (leadData[i][leadStateIdx] || '').toString().trim();
    let rawCountry = (leadData[i][leadCountryIdx] || '').toString().trim();
    let rawRegion = (leadData[i][leadRegionIdx] || '').toString().trim();

    let cityVal = rawCity;
    let stateVal = rawState;
    let countryVal = rawCountry;
    let changed = false;

    const rawCityKey = rawCity.toLowerCase();
    const rawStateKey = rawState.toLowerCase();
    const rawCountryKey = rawCountry.toLowerCase();

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
    const roles = {};
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

    // ✅ NEW: "Country → Region" fallback: always try to fill region if missing and country present
    if (!rawRegion && rawCountry) {
      const countryKey = rawCountry.toLowerCase();
      const resolvedRegion = countryToRegionMap.get(countryKey);
      if (resolvedRegion) {
        leadData[i][leadRegionIdx] = resolvedRegion;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadRegionIdx + 1 });
        changed = true;
      }
    }

    // ✅ NEW "State + Country → Region" fallback (City & State both missing, but Only Country present)
    if (!rawRegion && rawState && rawCountry) {
      const scKey = `${rawState.toLowerCase()}|${rawCountry.toLowerCase()}`;
      const resolvedRegion = stateCountryToRegionMap.get(scKey);
      if (resolvedRegion) {
        leadData[i][leadRegionIdx] = resolvedRegion;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadRegionIdx + 1 });
        changed = true;
      }
    }

    if (changed) fixCount++;

    // Collect unmatched rows for 'Missing_GeoMapping' (optional) - This is filling before the final check - WRONG
    // if (config.trackMissingGeoRows) {
    //   // if ((rawCity || rawState || rawCountry) && (!rawCountry || !rawRegion)) {
    //   // ✅ Only collect if city or state is present, but either country or region is missing
    //   if ((rawCity || rawState) && (!rawCountry || !rawRegion)) {
    //     const missingKey = `${rawCityKey}||${rawStateKey}||${rawCountryKey}||${rawRegion.toLowerCase()}`;
    //     if (!missingRecordsSet.has(missingKey)) {
    //       const note = [
    //         !rawCountry && "Missing Country",
    //         !rawRegion && "Missing Region",
    //       ].filter(Boolean).join(', ');
    //       newMissingRows.push([rawCity, rawState, rawCountry, rawRegion, note]);
    //       missingRecordsSet.add(missingKey);
    //     }
    //   }
    // }

    // ✅ Collect unmatched rows for 'Missing_GeoMapping' (AFTER all corrections are applied)
    if (config.trackMissingGeoRows) {
      const finalCity = (leadData[i][leadCityIdx] || '').toString().trim();
      const finalState = (leadData[i][leadStateIdx] || '').toString().trim();
      const finalCountry = (leadData[i][leadCountryIdx] || '').toString().trim();
      const finalRegion = (leadData[i][leadRegionIdx] || '').toString().trim();
    
      const finalCityKey = finalCity.toLowerCase();
      const finalStateKey = finalState.toLowerCase();
      const finalCountryKey = finalCountry.toLowerCase();
      const finalRegionKey = finalRegion.toLowerCase();
    
      // ✅ Only collect if city or state is present, but either country or region is missing
      if ((finalCity || finalState) && (!finalCountry || !finalRegion)) {
        const missingKey = `${finalCityKey}||${finalStateKey}||${finalCountryKey}||${finalRegionKey}`;
        if (!missingRecordsSet.has(missingKey)) {
          const note = [
            !finalCountry && "Missing Country",
            !finalRegion && "Missing Region",
          ].filter(Boolean).join(', ');
          newMissingRows.push([finalCity, finalState, finalCountry, finalRegion, note]);
          missingRecordsSet.add(missingKey);
        }
      }
    }

  }

  // Write back corrected data to sheet
  leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);

  // Highlight fixed cells if configured
  if (config.highlightFilledCells && cellChanges.length > 0) {
    cellChanges.forEach(change => {
      leadSheet.getRange(change.row, change.col).setBackground(config.highlightColor);
    });
  }

  // Write to Missing_GeoMapping sheet if enabled
  if (config.trackMissingGeoRows && newMissingRows.length > 0) {
    let missingSheet = ss.getSheetByName(config.trackMissingGeoRowsSheetName);
    if (!missingSheet) {
      missingSheet = ss.insertSheet(config.trackMissingGeoRowsSheetName);
    } else if (config.clearMissingGeoMappingBeforeAppend) {
      missingSheet.clear();
    }

    // Write headers if missing or after clearing
    if (missingSheet.getLastRow() === 0) {
      missingSheet.appendRow([
        config.columns.city,
        config.columns.state,
        config.columns.country,
        config.columns.region,
        "Notes"
      ]);
    }
    const lastRow = missingSheet.getLastRow();
    missingSheet.getRange(lastRow + 1, 1, newMissingRows.length, 5).setValues(newMissingRows);
  }

  // Final UI notifications
  ui.alert(`✅ Geo normalization done!\nRows changed: ${fixCount}\nSwaps applied: ${swapped}`);
  ss.toast(`Geo normalization complete for ${fixCount} rows`, 'Success', 4);
}

// 8. Fix City-State-Country-Region Mapping if Wrong by SwapAndFill
/**
 * 📍 Normalizes geographic data in 'Lead_CleanedData'
 * 
 * 🔄 Fixes swapped values between City, State, Country, and also handles Region
 * 🧩 Dynamically detects and swaps any 2 or all 3 mixed geo fields
 * 🧠 Independently fills missing State, Country, and Region:
 *    - Uses "City → State, Country, Region" mapping
 *    - Falls back with "Country → Region" mapping
 *    - Falls back with "State + Country → Region" mapping
 * 🎯 Only swaps if at least 2 fields are actually misplaced, includes Region in validation
 * 📝 (Optionally) logs unmatched/geographically incomplete rows in 'Missing_GeoMapping'
 * 🎨 (Optionally) Highlights corrected cells using configured color
 * 🔥 (Optionally) Avoids duplicate missing entries and clears log sheet before appending
 * 
 * 🔄 Dynamically corrects:
 * - City ↔ Country
 * - City ↔ State
 * - State ↔ Country
 * - All 3 if mixed with Region included
 * 
 * ⚙️ Config:
 *   leadSheetName                 - name of lead data sheet
 *   mappingSheetName              - name of master geo mapping sheet
 *   columns                       - column names to map
 *   highlightFilledCells          - toggle cell highlight on corrections
 *   highlightColor                - highlight color hex (e.g., red, yellow)
 *   trackMissingGeoRows           - enable creating 'Missing_GeoMapping' sheet with unmatched geo data
 *   trackMissingGeoRowsSheetName  - name of sheet to track missing geo rows
 *   clearMissingGeoMappingBeforeAppend - clear existing 'Missing_GeoMapping' before appending new data
 *   logging                       - logging configuration for missing rows, missing sheet name, and clearing behavior
 *
 * 🚫 Skips update if all geo values are empty
 * ✅ Logs how many rows updated, swaps applied, and missing rows logged
 */
function fillCityStateCountrySwapAndFillFromMapping_Master_Independently_Flexibly_Robust() { // 8.
  const config = {
    leadSheetName: 'Lead_CleanedData',
    mappingSheetName: 'Master_GeoMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region',
    },
    highlightFilledCells: true,         // Highlight cells corrected with this color
    highlightColor: '#ff9195',          // Not So Light red highlight (change if you want)
    logging : {
        trackMissingGeoRows: true,          // Collect partially missing rows in 'Missing_GeoMapping'
        trackMissingGeoRowsSheetName: 'Missing_GeoMapping',
        clearMissingGeoMappingBeforeAppend: true   // 🚨 Set to true to clear the sheet and write fresh
    }
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
    ui.alert(`❗ One or more required geo columns not found in the sheet headers.`);
    return;
  }

  // Load mapping data into sets and a cityMap for lookups
  const citySet = new Set();
  const stateSet = new Set();
  const countrySet = new Set();
  const cityMap = new Map();
  const countryToRegionMap = new Map();           // 👉 NEW: Country → Region fallback map
  const stateCountryToRegionMap = new Map();      // 👉 NEW: State + Country to region fallback map

  const mapData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, mappingSheet.getLastColumn()).getValues();

  mapData.forEach(row => {
    const city = (row[mapCityIdx] || '').toString().trim();
    const state = (row[mapStateIdx] || '').toString().trim();
    const country = (row[mapCountryIdx] || '').toString().trim();
    const region = (row[mapRegionIdx] || '').toString().trim();

    if (city) citySet.add(city.toLowerCase());
    if (state) stateSet.add(state.toLowerCase());
    if (country) countrySet.add(country.toLowerCase());

    const key = city.toLowerCase();
    if (city && !cityMap.has(key)) {
      cityMap.set(key, { city, state, country, region });
    }

    // 👉 NEW: Map country → region
    if (country && region && !countryToRegionMap.has(country.toLowerCase())) {
      countryToRegionMap.set(country.toLowerCase(), region);
    }

    // 👉 NEW: Map state+country → region
    if (state && country && region) {
      const scKey = `${state.toLowerCase()}|${country.toLowerCase()}`;
      if (!stateCountryToRegionMap.has(scKey)) {
        stateCountryToRegionMap.set(scKey, region);
      }
    }
  });

  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();
  let fixCount = 0;
  let swapped = 0;
  const cellChanges = [];

  const missingRecordsSet = new Set();
  const newMissingRows = [];

  if (config.logging.trackMissingGeoRows) {
    const missSheet = ss.getSheetByName(config.logging.trackMissingGeoRowsSheetName);
    if (missSheet && !config.logging.clearMissingGeoMappingBeforeAppend) {
      const missData = missSheet.getDataRange().getValues();
      for (let r = 1; r < missData.length; r++) {
        const row = missData[r];
        const key = [
          (row[0] || '').toString().trim().toLowerCase(),
          (row[1] || '').toString().trim().toLowerCase(),
          (row[2] || '').toString().trim().toLowerCase(),
          (row[3] || '').toString().trim().toLowerCase()
        ].join('||');
        missingRecordsSet.add(key);
      }
    }
  }

  for (let i = 0; i < leadData.length; i++) {
    let rawCity = (leadData[i][leadCityIdx] || '').toString().trim();
    let rawState = (leadData[i][leadStateIdx] || '').toString().trim();
    let rawCountry = (leadData[i][leadCountryIdx] || '').toString().trim();
    let rawRegion = (leadData[i][leadRegionIdx] || '').toString().trim();

    let cityVal = rawCity;
    let stateVal = rawState;
    let countryVal = rawCountry;
    let changed = false;

    const rawCityKey = rawCity.toLowerCase();
    const rawStateKey = rawState.toLowerCase();
    const rawCountryKey = rawCountry.toLowerCase();

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
    const roles = {};
    for (const field in detected) {
      const actualType = detected[field];
      if (actualType) roles[actualType] = field;
    }

    // If at least two fields detected, reassign values correctly
    // if (Object.keys(roles).length >= 2) {
    //   cityVal = values[roles.city] || '';
    //   stateVal = values[roles.state] || '';
    //   countryVal = values[roles.country] || '';

    //   if (leadData[i][leadCityIdx] !== cityVal) {
    //     leadData[i][leadCityIdx] = cityVal;
    //     if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadCityIdx + 1 });
    //     changed = true;
    //   }
    //   if (leadData[i][leadStateIdx] !== stateVal) {
    //     leadData[i][leadStateIdx] = stateVal;
    //     if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadStateIdx + 1 });
    //     changed = true;
    //   }
    //   if (leadData[i][leadCountryIdx] !== countryVal) {
    //     leadData[i][leadCountryIdx] = countryVal;
    //     if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadCountryIdx + 1 });
    //     changed = true;
    //   }
    //   swapped++;
    // }

    // ✅ Only swap if at least 2 values are misplaced among the 4 fields
    const actualCity = values[roles.city] || '';
    const actualState = values[roles.state] || '';
    const actualCountry = values[roles.country] || '';
    const actualRegion = values[roles.region] || '';

    // Count how many are in wrong place
    let misplacedCount = 0;
    if (roles.city && leadData[i][leadCityIdx] !== actualCity) misplacedCount++;
    if (roles.state && leadData[i][leadStateIdx] !== actualState) misplacedCount++;
    if (roles.country && leadData[i][leadCountryIdx] !== actualCountry) misplacedCount++;
    if (roles.region && leadData[i][leadRegionIdx] !== actualRegion) misplacedCount++;

    if (misplacedCount >= 2) {
      if (roles.city && leadData[i][leadCityIdx] !== actualCity) {
        leadData[i][leadCityIdx] = actualCity;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadCityIdx + 1 });
        changed = true;
      }
      if (roles.state && leadData[i][leadStateIdx] !== actualState) {
        leadData[i][leadStateIdx] = actualState;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadStateIdx + 1 });
        changed = true;
      }
      if (roles.country && leadData[i][leadCountryIdx] !== actualCountry) {
        leadData[i][leadCountryIdx] = actualCountry;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadCountryIdx + 1 });
        changed = true;
      }
      if (roles.region && leadData[i][leadRegionIdx] !== actualRegion) {
        leadData[i][leadRegionIdx] = actualRegion;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadRegionIdx + 1 });
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

    // ✅ NEW: "Country → Region" fallback: always try to fill region if missing and country present
    if (!rawRegion && rawCountry) {
      const countryKey = rawCountry.toLowerCase();
      const resolvedRegion = countryToRegionMap.get(countryKey);
      if (resolvedRegion) {
        leadData[i][leadRegionIdx] = resolvedRegion;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadRegionIdx + 1 });
        changed = true;
      }
    }

    // ✅ NEW "State + Country → Region" fallback (City & State both missing, but Only Country present)
    if (!rawRegion && rawState && rawCountry) {
      const scKey = `${rawState.toLowerCase()}|${rawCountry.toLowerCase()}`;
      const resolvedRegion = stateCountryToRegionMap.get(scKey);
      if (resolvedRegion) {
        leadData[i][leadRegionIdx] = resolvedRegion;
        if (config.highlightFilledCells) cellChanges.push({ row: i + 2, col: leadRegionIdx + 1 });
        changed = true;
      }
    }

    if (changed) fixCount++;

    // ✅ Collect unmatched rows for 'Missing_GeoMapping' (AFTER all corrections are applied)
    if (config.logging.trackMissingGeoRows) {
      const finalCity = (leadData[i][leadCityIdx] || '').toString().trim();
      const finalState = (leadData[i][leadStateIdx] || '').toString().trim();
      const finalCountry = (leadData[i][leadCountryIdx] || '').toString().trim();
      const finalRegion = (leadData[i][leadRegionIdx] || '').toString().trim();
    
      const finalCityKey = finalCity.toLowerCase();
      const finalStateKey = finalState.toLowerCase();
      const finalCountryKey = finalCountry.toLowerCase();
      const finalRegionKey = finalRegion.toLowerCase();
    
      // ✅ Only collect if city or state is present, but either country or region is missing
      if ((finalCity || finalState) && (!finalCountry || !finalRegion)) {
        const missingKey = `${finalCityKey}||${finalStateKey}||${finalCountryKey}||${finalRegionKey}`;
        if (!missingRecordsSet.has(missingKey)) {
          const note = [
            !finalCountry && "Missing Country",
            !finalRegion && "Missing Region",
          ].filter(Boolean).join(', ');
          newMissingRows.push([finalCity, finalState, finalCountry, finalRegion, note]);
          missingRecordsSet.add(missingKey);
        }
      }
    }

  }

  // Write back corrected data to sheet
  leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);

  // Highlight fixed cells if configured
  if (config.highlightFilledCells && cellChanges.length > 0) {
    cellChanges.forEach(change => {
      leadSheet.getRange(change.row, change.col).setBackground(config.highlightColor);
    });
  }

  // Write to Missing_GeoMapping sheet if enabled
  if (config.logging.trackMissingGeoRows && newMissingRows.length > 0) {
    let missingSheet = ss.getSheetByName(config.logging.trackMissingGeoRowsSheetName);
    if (!missingSheet) {
      missingSheet = ss.insertSheet(config.logging.trackMissingGeoRowsSheetName);
    } else if (config.logging.clearMissingGeoMappingBeforeAppend) {
      missingSheet.clear();
    }

    // Write headers if missing or after clearing
    if (missingSheet.getLastRow() === 0) {
      missingSheet.appendRow([
        config.columns.city,
        config.columns.state,
        config.columns.country,
        config.columns.region,
        "Notes"
      ]);
    }
    const lastRow = missingSheet.getLastRow();
    missingSheet.getRange(lastRow + 1, 1, newMissingRows.length, 5).setValues(newMissingRows);
  }

  // Final UI notifications
  ui.alert(`✅ Geo normalization done!
    📌 Rows changed (corrected): ${fixCount}
    🔁 Swaps applied: ${swapped}`
    + (config.logging.trackMissingGeoRows ? `\n 🕳️ Missing rows logged: ${newMissingRows.length}` : '')
  );

  ss.toast(`Geo normalization complete for ${fixCount} rows`, '✅Success✅', -1);
}


function extractGeoMappings() {
  const config = {
    sourceSheet: 'Lead_CleanedData',
    outputSheet: 'Geo_LookupData',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country'
    },
    regionMap: {
      // North America
      'United States':         'North America',
      'USA':                   'North America',
      'Canada':                'North America',
      'Mexico':                'North America',

      // Europe
      'United Kingdom':        'Europe',
      'UK':                    'Europe',
      'Germany':               'Europe',
      'France':                'Europe',
      'Netherlands':           'Europe',
      'Italy':                 'Europe',
      'Spain':                 'Europe',
      'Switzerland':           'Europe',
      'Sweden':                'Europe',
      'Denmark':               'Europe',
      'Finland':               'Europe',
      'Norway':                'Europe',
      'Poland':                'Europe',
      'Czech Republic':        'Europe',
      'Hungary':               'Europe',
      'Romania':               'Europe',

      // Southeast Asia
      'Singapore':             'Southeast Asia',
      'Indonesia':             'Southeast Asia',
      'Thailand':              'Southeast Asia',
      'Vietnam':               'Southeast Asia',
      'Malaysia':              'Southeast Asia',
      'Philippines':           'Southeast Asia',

      // Rest of APAC
      'India':                 'Asia-Pacific',
      'Australia':             'Asia-Pacific',
      'Japan':                 'Asia-Pacific',
      'South Korea':           'Asia-Pacific',
      'China':                 'Asia-Pacific',
      'New Zealand':           'Asia-Pacific',

      // Middle East & MENA
      'United Arab Emirates':  'Middle East',
      'UAE':                   'Middle East',
      'Saudi Arabia':          'Middle East',
      'Qatar':                 'Middle East',
      'Turkey':                'Middle East',
      'Israel':                'Middle East',
      'Egypt':                 'MENA',
      'Morocco':               'MENA',
      'Algeria':               'MENA',
      'Tunisia':               'MENA',

      // Latin America (LATAM)
      'Brazil':                'Latin America',
      'Argentina':             'Latin America',
      'Chile':                 'Latin America',
      'Colombia':              'Latin America',
      'Peru':                  'Latin America',
      'Mexico':                'Latin America',

      // Africa
      'South Africa':          'Africa',
      'Nigeria':               'Africa'
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(config.sourceSheet);

  if (!sheet) {
    ui.alert(`❗ Source sheet '${config.sourceSheet}' not found.`);
    return;
  }

  ss.toast('Extracting Geo Lookup Data...', '🏙️🌎', -1);

  // Read headers & all data
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cityIdx = headers.indexOf(config.columns.city);
  const stateIdx = headers.indexOf(config.columns.state);
  const countryIdx = headers.indexOf(config.columns.country);

  if (cityIdx === -1 || stateIdx === -1 || countryIdx === -1) {
    ui.alert(`❗ Required columns not found: ${Object.values(config.columns).join(', ')}`);
    return;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  // Use a Set for unique city|state|country keys.
  const mappingSet = new Set();
  const outputRows = [];

  data.forEach(row => {
    const city    = (row[cityIdx]    || '').toString().trim();
    const state   = (row[stateIdx]   || '').toString().trim();
    const country = (row[countryIdx] || '').toString().trim();
    if (!city) return; // Only map where city is present
    // Compose unique key
    const uniqueKey = `${city}|${state}|${country}`;
    if (mappingSet.has(uniqueKey)) return;
    mappingSet.add(uniqueKey);

    // Region
    let region = config.regionMap[country] || 'Other';
    outputRows.push([city, state, country, region]);
  });

  // Write output
  let outputSheet = ss.getSheetByName(config.outputSheet);
  if (outputSheet) outputSheet.clearContents();
  else outputSheet = ss.insertSheet(config.outputSheet);

  // Write Headers
  outputSheet.getRange(1, 1, 1, 4).setValues([
    [config.columns.city, config.columns.state, config.columns.country, 'Region']
  ]);
  // Write Data
  if (outputRows.length) {
    outputSheet.getRange(2, 1, outputRows.length, 4).setValues(outputRows);
  }
  outputSheet.setFrozenRows(1);

  // Highlight missing country
  const toHighlight = [];
  for (let r = 0; r < outputRows.length; r++) {
    if (!outputRows[r][2]) toHighlight.push(r + 2); // Row index in sheet
  }
  if (toHighlight.length > 0) {
    const rangeList = toHighlight.map(r => `A${r}:D${r}`);
    outputSheet.getRangeList(rangeList).setBackground('#fff2cc');
  }

  ss.toast('Geo Lookup Extraction - COMPLETED ✅', '✅ Success ✅', -1);
  ui.alert(
    `✅ Geo_LookupData ready with ${outputRows.length} unique rows.\n` +
    `❗ ${toHighlight.length} rows have missing country.`
  );
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

function generateMissingDetailsReport() {
  const config = {
    targetSheetName: 'Lead_CleanedData',

    // Specify pairs of columns to check missing data
    // Each entry: { keyCol: 'City', valueCol: 'Country' }
    // It will find keys with missing values and values with missing keys
    columnPairs: [
      { keyCol: 'Company City', valueCol: 'Company Country' },
      // Add more pairs if needed, e.g.,
      // { keyCol: 'Some Column', valueCol: 'Another Column' }
    ],

    missingSheetName: 'MissingDetails',
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(config.targetSheetName);

  if (!sheet) {
    ui.alert(`❗ Sheet '${config.targetSheetName}' not found.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Helper to get column index or -1
  function getColIndex(colName) {
    return headers.indexOf(colName);
  }

  // Object to hold missing data results
  // Format example:
  // {
  //   'Company City - missing Company Country': ['City1', 'City2'],
  //   'Company Country - missing Company City': ['Country1'],
  //   ...
  // }
  const missingData = {};

  // Process each configured pair
  config.columnPairs.forEach(pair => {
    const keyIdx = getColIndex(pair.keyCol);
    const valIdx = getColIndex(pair.valueCol);

    if (keyIdx === -1 || valIdx === -1) {
      ui.alert(`❗ Column '${pair.keyCol}' or '${pair.valueCol}' not found on '${config.targetSheetName}'.`);
      return;
    }

    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length);
    const data = dataRange.getValues();

    // Sets to keep unique missing keys and missing values
    const missingValuesForKeys = new Set(); // Keys with missing Values
    const missingKeysForValues = new Set(); // Values with missing Keys

    // Build reverse mapping: value → keys present
    const valueToKeys = {};

    // First pass: build value → key(s) map
    data.forEach(row => {
      const key = (row[keyIdx] || '').toString().trim();
      const val = (row[valIdx] || '').toString().trim();

      if (val) {
        if (!valueToKeys[val]) valueToKeys[val] = new Set();
        if (key) valueToKeys[val].add(key);
      }
    });

    // Second pass: find missing values for keys and missing keys for values
    data.forEach(row => {
      const key = (row[keyIdx] || '').toString().trim();
      const val = (row[valIdx] || '').toString().trim();

      if (key && !val) {
        missingValuesForKeys.add(key);
      }

      if (val && (!key || key === '')) {
        // Check if value associated with no keys
        // But if valueToKeys[val] contains empty string key?
        missingKeysForValues.add(val);
      }
    });

    // Store results as arrays
    missingData[`${pair.keyCol} - missing ${pair.valueCol}`] = Array.from(missingValuesForKeys).sort();
    missingData[`${pair.valueCol} - missing ${pair.keyCol}`] = Array.from(missingKeysForValues).sort();
  });

  // Prepare output rows: 1st row with headers = keys of missingData
  // Rows = max length among all arrays
  const keys = Object.keys(missingData);
  const maxLength = keys.reduce((max, k) => Math.max(max, missingData[k].length), 0);

  const output = [];

  // Header row
  output.push(keys);

  // Data rows
  for (let i = 0; i < maxLength; i++) {
    const row = keys.map(k => missingData[k][i] || '');
    output.push(row);
  }

  // Write to "MissingDetails" sheet
  let missingSheet = ss.getSheetByName(config.missingSheetName);
  if (missingSheet) {
    missingSheet.clearContents();
  } else {
    missingSheet = ss.insertSheet(config.missingSheetName);
  }

  missingSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  missingSheet.setFrozenRows(1);

  ui.alert(`✅ Missing details report created at '${config.missingSheetName}'.\nColumns checked: ${config.columnPairs.map(p => `'${p.keyCol}' & '${p.valueCol}'`).join(', ')}`);
}

function generateMissingDetailsAdvancedReport() {
  const config = {
    targetSheetName: 'Lead_CleanedData',

    // Configure pairs or triplets to check
    // For triples: keyCol + valueCol + valueCol2
    // For pairs: just keyCol + valueCol (omit valueCol2)
    columnSets: [
      { keyCol: 'Company State', valueCol: 'Company City', valueCol2: 'Company Country' },
      // Example for pairs:
      // { keyCol: 'Department', valueCol: 'Seniority' }
    ],

    missingSheetName: 'MissingDetails',
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(config.targetSheetName);

  if (!sheet) {
    ui.alert(`❗ Sheet '${config.targetSheetName}' not found.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  function getColIndex(colName) { return headers.indexOf(colName); }

  // Will hold missing and conflict data per columnSet
  // Format:
  // {
  //   'Company State + Company City -> missing Company Country': [...],
  //   'Company City + Company Country -> missing Company State': [...],
  //   'Conflicting mappings (Company City)': ['City1', 'City2'],
  //   ...
  // }
  const missingData = {};
  const conflictData = {};

  const maxRows = sheet.getLastRow() - 1;
  if (maxRows < 1) {
    ui.alert('❗ No data rows found.');
    return;
  }

  // Process each configured column set
  config.columnSets.forEach(set => {
    const keyColIdx = getColIndex(set.keyCol);
    const valColIdx = getColIndex(set.valueCol);
    const valCol2Idx = set.valueCol2 ? getColIndex(set.valueCol2) : -1;

    // Validate columns
    if (keyColIdx === -1 || valColIdx === -1 || (set.valueCol2 && valCol2Idx === -1)) {
      ui.alert(`❗ One or more columns not found: ${set.keyCol}, ${set.valueCol}, ${set.valueCol2 || ''}`);
      return;
    }

    // Read all data for this sheet
    const data = sheet.getRange(2, 1, maxRows, headers.length).getValues();

    // Maps for detecting missing and conflicts
    // Map key: For triples: keyVal + '||' + valVal ; For pairs: keyVal only
    const valPresenceMap = new Map();   // Maps key → Set of dependent values found (to detect conflicts)
    const val2PresenceMap = set.valueCol2 ? new Map() : null;

    // Sets to store missing keys or values
    const missingValForKeys = new Set();
    const missingVal2ForKeyValPairs = new Set();
    const missingKeyForValues = new Set();
    const missingKeyForVal2 = new Set();

    // Build maps and detect missingness
    data.forEach(row => {
      const keyVal = row[keyColIdx]?.toString().trim();
      const valVal = row[valColIdx]?.toString().trim();
      const val2Val = valCol2Idx !== -1 ? row[valCol2Idx]?.toString().trim() : null;

      // Detect missing values for keys
      if (keyVal && !valVal) missingValForKeys.add(keyVal);

      // For triples: detect missing val2 for key+val pairs
      if (set.valueCol2 && keyVal && valVal && !val2Val) {
        missingVal2ForKeyValPairs.add(keyVal + '||' + valVal);
      }

      // For missing keys given a value: if valVal exists but keyVal missing
      if (!keyVal && valVal) missingKeyForValues.add(valVal);

      // For triples: if val2 exists but key or val missing, mark missing keys
      if (set.valueCol2 && val2Val && (!keyVal || !valVal)) {
        missingKeyForVal2.add(val2Val);
      }

      // Populate presence map for conflicts:
      if (keyVal) {
        if (!valPresenceMap.has(keyVal)) valPresenceMap.set(keyVal, new Set());
        if (valVal) valPresenceMap.get(keyVal).add(valVal);
      }
    });

    // Detect conflicts in valPresenceMap: key with >1 associated values
    const conflictsForKey = [...valPresenceMap.entries()]
      .filter(([k, vSet]) => vSet.size > 1)
      .map(([k]) => k);

    // Store all results
    missingData[`${set.keyCol} - missing ${set.valueCol}`] = Array.from(missingValForKeys).sort();
    if (set.valueCol2) {
      missingData[`${set.keyCol} + ${set.valueCol} - missing ${set.valueCol2}`] = Array.from(missingVal2ForKeyValPairs).sort();
    }
    missingData[`${set.valueCol} - missing ${set.keyCol}`] = Array.from(missingKeyForValues).sort();

    if (set.valueCol2) {
      missingData[`${set.valueCol2} - missing ${set.keyCol} or ${set.valueCol}`] = Array.from(missingKeyForVal2).sort();
    }

    if (conflictsForKey.length > 0) {
      conflictData[`Conflicting mappings (${set.keyCol})`] = conflictsForKey.sort();
    }

  });

  // Prepare output structure: Combine missingData and conflictData
  const combinedKeys = [...Object.keys(missingData), ...Object.keys(conflictData)];

  if (combinedKeys.length === 0) {
    ui.alert('✅ No missing or conflicting data found.');
    return;
  }

  // Find longest list
  const maxLen = combinedKeys.reduce((max, k) => {
    const arr = missingData[k] || conflictData[k] || [];
    return Math.max(max, arr.length);
  }, 0);

  // Build output rows
  const output = [];
  output.push(combinedKeys);
  for (let i = 0; i < maxLen; i++) {
    const row = combinedKeys.map(k => {
      const arr = missingData[k] || conflictData[k] || [];
      return arr[i] || '';
    });
    output.push(row);
  }

  // Write to MissingDetails sheet
  let missingSheet = ss.getSheetByName(config.missingSheetName);
  if (!missingSheet) {
    missingSheet = ss.insertSheet(config.missingSheetName);
  } else {
    missingSheet.clearContents();
  }

  missingSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  missingSheet.setFrozenRows(1);

  ui.alert(`✅ Missing details & conflicts reported in '${config.missingSheetName}'.`);
  ss.toast('Missing details report - COMPLETED ✅', '✅ Success ✅', -1);
}

function generateMissingCountryReport() {
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

  // Use a Map to ensure uniqueness on combined key (city|state)
  const uniqueRowsMap = new Map();

  data.forEach(row => {
    const city = (row[cityColIdx] || '').toString().trim();
    const state = (row[stateColIdx] || '').toString().trim();
    const country = (row[countryColIdx] || '').toString().trim();

    if (!country) { // country missing
      // Ignore rows where state AND city both missing
      if (!state && !city) return;

      // If config set to ignore state present but city missing rows
      if (config.excludeStateMissingCity && state && !city) return;

      // Compose unique key with city & state (both)
      const uniqueKey = city + '|' + state;

      if (!uniqueRowsMap.has(uniqueKey)) {
        uniqueRowsMap.set(uniqueKey, [city, state, '']); // country empty
      }
    }
  });

  const missingRows = Array.from(uniqueRowsMap.values());

  let outputSheet = ss.getSheetByName(config.outputSheetName);
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet(config.outputSheetName);
  }

  outputSheet.getRange(1, 1, 1, 3).setValues([[config.columns.city, config.columns.state, config.columns.country]]);
  if (missingRows.length > 0) {
    outputSheet.getRange(2, 1, missingRows.length, 3).setValues(missingRows);
  }

  outputSheet.setFrozenRows(1);

  ss.toast('Missing Country report generated ✅', '✅ Success ✅', -1);
  ui.alert(`✅ ${missingRows.length} unique rows found with missing 'Country'.\nOutput is in '${config.outputSheetName}' sheet.`);
}

function fillMissingCountryFromMapping() {
  const config = {
    leadSheetName: 'Lead_CleanedData',
    mappingSheetName: 'CityStateCountryMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
    },
    highlightFilledCells: true,     // Set false to disable highlighting
    highlightColor: '#d9ead3'       // Light green
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get sheets and check columns
  const leadSheet = ss.getSheetByName(config.leadSheetName);
  const mapSheet = ss.getSheetByName(config.mappingSheetName);

  if (!leadSheet) {
    ui.alert(`❗ Lead data sheet '${config.leadSheetName}' not found.`);
    return;
  }
  if (!mapSheet) {
    ui.alert(`❗ Mapping sheet '${config.mappingSheetName}' not found.`);
    return;
  }

  ss.toast('Filling missing Company Country using CityStateCountryMapping...', '⚠️ Attention ⚠️', -1);

  const leadHeaders = leadSheet.getRange(1, 1, 1, leadSheet.getLastColumn()).getValues()[0];
  const mapHeaders = mapSheet.getRange(1, 1, 1, mapSheet.getLastColumn()).getValues()[0];

  const leadCityIdx    = leadHeaders.indexOf(config.columns.city);
  const leadStateIdx   = leadHeaders.indexOf(config.columns.state);
  const leadCountryIdx = leadHeaders.indexOf(config.columns.country);

  const mapCityIdx     = mapHeaders.indexOf(config.columns.city);
  const mapStateIdx    = mapHeaders.indexOf(config.columns.state);
  const mapCountryIdx  = mapHeaders.indexOf(config.columns.country);

  if (
    leadCityIdx === -1 || leadStateIdx === -1 || leadCountryIdx === -1 ||
    mapCityIdx  === -1 || mapStateIdx  === -1 || mapCountryIdx  === -1
  ) {
    ui.alert(`❗ One or more required columns NOT found. Ensure consistent column names!`);
    return;
  }

  // Build city+state map and city-only map for fallback
  const mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, mapSheet.getLastColumn()).getValues();
  const cityStateMap = new Map(); // key: "city|state"
  const cityOnlyMap = new Map();  // key: "city|"

  mapData.forEach(row => {
    const city    = (row[mapCityIdx] || '').toString().trim().toLowerCase();
    const state   = (row[mapStateIdx] || '').toString().trim().toLowerCase();
    const country = (row[mapCountryIdx] || '').toString().trim();
    if (!city || !country) return; // skip incomplete mappings
    const keyCombo = city + '|' + state;
    cityStateMap.set(keyCombo, country);
    if (!state) cityOnlyMap.set(city + '|', country);
  });

  // Lead data
  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();

  // Track which need filling and which get filled
  let fillCount = 0;
  let couldNotFill = 0;
  let highlightRows = [];

  for (let i = 0; i < leadData.length; i++) {
    const row = leadData[i];
    const city    = (row[leadCityIdx] || '').toString().trim().toLowerCase();
    const state   = (row[leadStateIdx] || '').toString().trim().toLowerCase();
    let   country = (row[leadCountryIdx] || '').toString().trim();

    if (!country && city) {
      // Try city+state key first
      let found = false;
      const keyCombo = city + '|' + state;
      if (cityStateMap.has(keyCombo)) {
        row[leadCountryIdx] = cityStateMap.get(keyCombo);
        fillCount++;
        highlightRows.push(i + 2);
        found = true;
      } else if (cityOnlyMap.has(city + '|')) {
        row[leadCountryIdx] = cityOnlyMap.get(city + '|');
        fillCount++;
        highlightRows.push(i + 2);
        found = true;
      }
      if (!found) couldNotFill++;
    }
  }

  // Update sheet if needed
  if (fillCount > 0) {
    leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);
    if (config.highlightFilledCells) {
      const leadCountryColA1 = leadHeaders[leadCountryIdx];
      const rangeNotations = highlightRows.map(r => leadSheet.getRange(r, leadCountryIdx + 1).getA1Notation());
      leadSheet.getRangeList(rangeNotations).setBackground(config.highlightColor);
    }
  }

  ss.toast(`Filled ${fillCount} missing Company Country cells`, '✅ Success ✅', -1);

  ui.alert(
    `✓ ${fillCount} rows updated with missing Company Country from mapping.\n` +
    (
      couldNotFill > 0
        ? `⚠️ ${couldNotFill} rows NOT filled (no matching city/state mapping found).`
        : `All possible rows were filled.`
    )
  );
}

function generateMasterMapping_duplicateRecords() {
  const config = {
    sheets: [
      {
        name: '❌GeneratedForMissingRegion',
        columns: {
          city: 'Company City',
          state: 'Company State',
          country: 'Company Country',
          region: 'Region'
        }
      },
      {
        name: 'Geo_LookupData',
        columns: {
          city: 'Company City',
          state: 'Company State',
          country: 'Company Country',
          region: 'Region'
        }
      },
      {
        name: 'CityStateCountryMapping',
        columns: {
          city: 'Company City',
          state: 'Company State',
          country: 'Company Country',
          region: null // no region column
        }
      }
    ],
    outputSheet: 'CityStateCountryRegionMapping',
    outputColumns: ['Company City', 'Company State', 'Company Country', 'Region'],
    unknownRegionPlaceholder: 'Other'
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Helper: build row key
  function rowKey(city, state, country) {
    return `${city.toLowerCase().trim()}|${state.toLowerCase().trim()}|${country.toLowerCase().trim()}`;
  }

  // Priority store: highest = best (region present && not "Other"), next = region present, then fallback mapping, then blank
  const masterMap = new Map();

  for (const tab of config.sheets) {
    const sheet = ss.getSheetByName(tab.name);
    if (!sheet) continue;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cityIdx = headers.indexOf(tab.columns.city);
    const stateIdx = headers.indexOf(tab.columns.state);
    const countryIdx = headers.indexOf(tab.columns.country);
    const regionIdx = tab.columns.region ? headers.indexOf(tab.columns.region) : -1;

    if (cityIdx === -1 || stateIdx === -1 || countryIdx === -1) continue;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    for (const row of data) {
      const city = (row[cityIdx] || '').toString().trim();
      const state = (row[stateIdx] || '').toString().trim();
      const country = (row[countryIdx] || '').toString().trim();
      if (!(city && country)) continue; // city/country are minimum required

      let region = regionIdx !== -1 ? (row[regionIdx] || '').toString().trim() : '';
      if (!region) region = config.unknownRegionPlaceholder;

      // Pick the "best" mapping for deduplication:
      const key = rowKey(city, state, country);

      // Decide: Higher priority if region is not blank/Other
      const existing = masterMap.get(key);
      const isBetter = () => {
        if (!existing) return true;
        if (region && region !== config.unknownRegionPlaceholder) {
          if (!(existing.region && existing.region !== config.unknownRegionPlaceholder)) return true;
        }
        return false;
      };

      if (isBetter()) {
        masterMap.set(key, { city, state, country, region });
      }
    }
  }

  // Prepare data and sort by City and State for easier review
  const output = Array.from(masterMap.values())
    .sort((a, b) => 
      a.city.localeCompare(b.city, undefined, {sensitivity: 'base'}) ||
      a.state.localeCompare(b.state, undefined, {sensitivity: 'base'})
    )
    .map(obj => [obj.city, obj.state, obj.country, obj.region]);

  let outSheet = ss.getSheetByName(config.outputSheet);
  if (outSheet) outSheet.clearContents();
  else outSheet = ss.insertSheet(config.outputSheet);
  
  outSheet.getRange(1, 1, 1, config.outputColumns.length).setValues([config.outputColumns]);
  if (output.length) outSheet.getRange(2, 1, output.length, config.outputColumns.length).setValues(output);
  outSheet.setFrozenRows(1);

  ui.alert(`✅ MasterMapping sheet created with ${output.length} unique mappings.`);
}

/**
 * Combines unique city-state-country-region mappings from
 * 'Geo_LookupData', 'CityStateCountryMapping', and 'CityStateCountryRegionMapping'
 * into a single deduplicated 'MasterMapping' sheet with columns:
 * Company City | Company State | Company Country | Region
 *
 * - Prefers records with non-empty State over empty State for same City+Country.
 * - Gives priority to mappings with a valid Region (not blank/'Other').
 * - No duplicate rows by (Company City, Company State, Company Country).
 * - Output is sorted alphabetically by City, then State.
 */
function generateMasterMapping() {
  const config = {
    sheets: [
      {
        name: 'CityStateCountryRegionMapping',
        columns: {
          city: 'Company City',
          state: 'Company State',
          country: 'Company Country',
          region: 'Region'
        }
      },
      {
        name: 'Geo_LookupData',
        columns: {
          city: 'Company City',
          state: 'Company State',
          country: 'Company Country',
          region: 'Region'
        }
      },
      {
        name: 'CityStateCountryMapping',
        columns: {
          city: 'Company City',
          state: 'Company State',
          country: 'Company Country',
          region: null // no region column
        }
      }
    ],
    outputSheet: 'MasterMapping',
    outputColumns: ['Company City', 'Company State', 'Company Country', 'Region'],
    unknownRegionPlaceholder: 'Other'
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Helper: build row key ignoring state to allow overwrite by more detailed state info
  // Key is City + Country only for initial lookup to detect duplicates
  function rowKey(city, country) {
    return `${city.toLowerCase().trim()}|${country.toLowerCase().trim()}`;
  }

  const masterMap = new Map();

  for (const tab of config.sheets) {
    const sheet = ss.getSheetByName(tab.name);
    if (!sheet) continue;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cityIdx = headers.indexOf(tab.columns.city);
    const stateIdx = headers.indexOf(tab.columns.state);
    const countryIdx = headers.indexOf(tab.columns.country);
    const regionIdx = tab.columns.region ? headers.indexOf(tab.columns.region) : -1;

    if (cityIdx === -1 || stateIdx === -1 || countryIdx === -1) continue;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    for (const row of data) {
      const city = (row[cityIdx] || '').toString().trim();
      const state = (row[stateIdx] || '').toString().trim();
      const country = (row[countryIdx] || '').toString().trim();
      if (!(city && country)) continue; // city & country required

      let region = regionIdx !== -1 ? (row[regionIdx] || '').toString().trim() : '';
      if (!region) region = config.unknownRegionPlaceholder;

      const key = rowKey(city, country);
      const existing = masterMap.get(key);

      function hasMorePriority(newEntry, existingEntry) {
        // 1. Prefer record with more complete state
        if (newEntry.state && !existingEntry.state) return true;
        if (!newEntry.state && existingEntry.state) return false;

        // 2. Prefer region different than 'Other'
        if (newEntry.region !== config.unknownRegionPlaceholder && existingEntry.region === config.unknownRegionPlaceholder) return true;
        if (existingEntry.region !== config.unknownRegionPlaceholder && newEntry.region === config.unknownRegionPlaceholder) return false;

        // 3. Otherwise, keep existing (or treat equal)
        return false;
      }

      if (!existing) {
        masterMap.set(key, { city, state, country, region });
      } else if (hasMorePriority({ city, state, country, region }, existing)) {
        masterMap.set(key, { city, state, country, region });
      }
      // else keep existing
    }
  }

  // Sort by city then state
  const output = Array.from(masterMap.values())
    .sort((a, b) => a.city.localeCompare(b.city, undefined, { sensitivity: 'base' }) || a.state.localeCompare(b.state, undefined, { sensitivity: 'base' }))
    .map(r => [r.city, r.state, r.country, r.region]);

  let outSheet = ss.getSheetByName(config.outputSheet);
  if (outSheet) outSheet.clearContents();
  else outSheet = ss.insertSheet(config.outputSheet);

  outSheet.getRange(1, 1, 1, config.outputColumns.length).setValues([config.outputColumns]);
  if (output.length) outSheet.getRange(2, 1, output.length, config.outputColumns.length).setValues(output);
  outSheet.setFrozenRows(1);

  ui.alert(`✅ MasterMapping sheet created with ${output.length} unique mappings, optimized for more complete state info.`);
}

/**
 * Fills missing 'Company Country' and 'Region' values in 'Lead_CleanedData'
 * by using mappings from 'CityStateCountryRegionMapping'.
 *
 * - Matches by 'City + State' first, then by 'City' alone if State is blank.
 * - Updates 'Company Country' and 'Region' columns in 'Lead_CleanedData'.
 * - Optionally highlights filled cells (configurable).
 * - Shows alert on completion including counts of missing countries and regions.
 */
function fillMissingCountryRegionFromMapping() {
  const config = {
    leadSheetName: 'Lead_CleanedData',
    mappingSheetName: 'CityStateCountryRegionMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region'
    },
    highlightFilledCells: true,
    highlightColor: '#d9ead3'
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const leadSheet = ss.getSheetByName(config.leadSheetName);
  const mapSheet = ss.getSheetByName(config.mappingSheetName);

  if (!leadSheet) {
    ui.alert(`❗ Sheet '${config.leadSheetName}' not found.`);
    return;
  }
  if (!mapSheet) {
    ui.alert(`❗ Sheet '${config.mappingSheetName}' not found.`);
    return;
  }

  ss.toast('Filling missing Country/Region from mapping...', '⚠️ Working ⚠️', -1);

  const leadHeaders = leadSheet.getRange(1, 1, 1, leadSheet.getLastColumn()).getValues()[0];
  const mapHeaders  = mapSheet.getRange(1, 1, 1, mapSheet.getLastColumn()).getValues()[0];

  const leadCityIdx    = leadHeaders.indexOf(config.columns.city);
  const leadStateIdx   = leadHeaders.indexOf(config.columns.state);
  const leadCountryIdx = leadHeaders.indexOf(config.columns.country);
  const leadRegionIdx  = leadHeaders.indexOf(config.columns.region);

  const mapCityIdx     = mapHeaders.indexOf(config.columns.city);
  const mapStateIdx    = mapHeaders.indexOf(config.columns.state);
  const mapCountryIdx  = mapHeaders.indexOf(config.columns.country);
  const mapRegionIdx   = mapHeaders.indexOf(config.columns.region);

  if (
    [leadCityIdx, leadStateIdx, leadCountryIdx, leadRegionIdx,
     mapCityIdx, mapStateIdx, mapCountryIdx, mapRegionIdx].some(idx => idx === -1)
  ) {
    ui.alert(`❗ One or more required columns NOT found. Please check column headers.`);
    return;
  }

  // Build Mapping Dicts: (city|state|country), (city|country), (country)
  const mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, mapSheet.getLastColumn()).getValues();
  const cityStateCountryMap = new Map(); // key: city|state|country --> { country, region }
  const cityCountryMap = new Map();      // key: city|country --> { country, region }
  const countryMap = new Map();          // key: country --> region

  mapData.forEach(row => {
    const city    = (row[mapCityIdx]    || '').toString().trim().toLowerCase();
    const state   = (row[mapStateIdx]   || '').toString().trim().toLowerCase();
    const country = (row[mapCountryIdx] || '').toString().trim().toLowerCase();
    const region  = (row[mapRegionIdx]  || '').toString().trim();

    if (city && country) {
      cityCountryMap.set(city + '|' + country, { country: row[mapCountryIdx], region });
    }
    if (city && state && country) {
      cityStateCountryMap.set(city + '|' + state + '|' + country, { country: row[mapCountryIdx], region });
    }
    if (country && region) {
      countryMap.set(country, region);
    }
  });

  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();

  let fillCount = 0, missingCountryCount = 0, missingRegionCount = 0;
  let highlightRowsCountry = [], highlightRowsRegion = [];

  for (let i = 0; i < leadData.length; i++) {
    let row = leadData[i];

    const city    = (row[leadCityIdx]    || '').toString().trim().toLowerCase();
    const state   = (row[leadStateIdx]   || '').toString().trim().toLowerCase();
    let   country = (row[leadCountryIdx] || '').toString().trim();
    let   region  = (row[leadRegionIdx]  || '').toString().trim();

    let originalCountry = country, originalRegion = region;

    // Priority 1: match by city+state+country (most precise)
    let mapEntry = null;
    if (city && state && country) {
      mapEntry = cityStateCountryMap.get(city + '|' + state + '|' + country.toLowerCase());
    }
    // Priority 2: match by city+country only (if no state in lead)
    if (!mapEntry && city && country) {
      mapEntry = cityCountryMap.get(city + '|' + country.toLowerCase());
    }
    // Priority 3: region only by country (if region is missing)
    let regFallback = (!region && country && countryMap.has(country.toLowerCase()))
      ? countryMap.get(country.toLowerCase())
      : null;

    // Fill as needed by rules
    if ((!country || !region) && (mapEntry || regFallback)) {
      // Fill country if missing and mapping provides it
      if (!country && mapEntry && mapEntry.country) {
        row[leadCountryIdx] = mapEntry.country;
        country = mapEntry.country;
        fillCount++;
        highlightRowsCountry.push(i + 2);
      }
      // Fill region if missing
      if (!region) {
        if (mapEntry && mapEntry.region) {
          row[leadRegionIdx] = mapEntry.region;
          region = mapEntry.region;
          fillCount++;
          highlightRowsRegion.push(i + 2);
        } else if (regFallback && typeof regFallback === 'string') {
          row[leadRegionIdx] = regFallback;
          region = regFallback;
          fillCount++;
          highlightRowsRegion.push(i + 2);
        }
      }
    }

    // Counts for missing fields after filling
    if (!country) missingCountryCount++;
    if (!region)  missingRegionCount++;
  }

  // Write back if updated
  if (fillCount > 0) {
    leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);
    if (config.highlightFilledCells) {
      let highlights = [];
      if (highlightRowsCountry.length)
        highlights = highlights.concat(
          highlightRowsCountry.map(r => leadSheet.getRange(r, leadCountryIdx + 1).getA1Notation())
        );
      if (highlightRowsRegion.length)
        highlights = highlights.concat(
          highlightRowsRegion.map(r => leadSheet.getRange(r, leadRegionIdx + 1).getA1Notation())
        );
      if (highlights.length)
        leadSheet.getRangeList(highlights).setBackground(config.highlightColor);
    }
  }

  ss.toast(`Filled ${fillCount} missing Country/Region.`, '✅ Success ✅', -1);
  ui.alert(
    `✅ ${fillCount} cells updated from mapping.\n` +
    `❗ Rows still missing Company Country: ${missingCountryCount}\n` +
    `❗ Rows still missing Region: ${missingRegionCount}`
  );
}


function fillMissingCountryRegionFromMapping_Issue() {
  const config = {
    leadSheetName: 'Lead_CleanedData',
    mappingSheetName: 'CityStateCountryRegionMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region'
    },
    highlightFilledCells: true,     // Set false to disable highlighting
    highlightColor: '#d9ead3'       // Light green
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get sheets & validate
  const leadSheet = ss.getSheetByName(config.leadSheetName);
  const mapSheet = ss.getSheetByName(config.mappingSheetName);

  if (!leadSheet) {
    ui.alert(`❗ Lead data sheet '${config.leadSheetName}' not found.`);
    return;
  }
  if (!mapSheet) {
    ui.alert(`❗ Mapping sheet '${config.mappingSheetName}' not found.`);
    return;
  }

  ss.toast('Filling missing Company Country and Region using CityStateCountryRegionMapping...', '⚠️ Attention ⚠️', -1);

  // Headers
  const leadHeaders = leadSheet.getRange(1, 1, 1, leadSheet.getLastColumn()).getValues()[0];
  const mapHeaders = mapSheet.getRange(1, 1, 1, mapSheet.getLastColumn()).getValues()[0];

  const leadCityIdx = leadHeaders.indexOf(config.columns.city);
  const leadStateIdx = leadHeaders.indexOf(config.columns.state);
  const leadCountryIdx = leadHeaders.indexOf(config.columns.country);
  const leadRegionIdx = leadHeaders.indexOf(config.columns.region);

  const mapCityIdx = mapHeaders.indexOf(config.columns.city);
  const mapStateIdx = mapHeaders.indexOf(config.columns.state);
  const mapCountryIdx = mapHeaders.indexOf(config.columns.country);
  const mapRegionIdx = mapHeaders.indexOf(config.columns.region);

  if (
    leadCityIdx === -1 || leadStateIdx === -1 || leadCountryIdx === -1 || leadRegionIdx === -1 ||
    mapCityIdx === -1 || mapStateIdx === -1 || mapCountryIdx === -1 || mapRegionIdx === -1
  ) {
    ui.alert(`❗ One or more required columns NOT found in either lead or mapping sheet. Please verify column names.`);
    return;
  }

  // Build mapping dictionaries
  const mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, mapSheet.getLastColumn()).getValues();
  const cityStateMap = new Map(); // key: city|state  => {country, region}
  const cityOnlyMap = new Map();  // key: city| => {country, region}

  mapData.forEach(row => {
    const city = (row[mapCityIdx] || '').toString().trim().toLowerCase();
    const state = (row[mapStateIdx] || '').toString().trim().toLowerCase();
    const country = (row[mapCountryIdx] || '').toString().trim();
    const region = (row[mapRegionIdx] || '').toString().trim();
    if (!city || !country) return; // skip incomplete
    const keyCombo = city + '|' + state;
    cityStateMap.set(keyCombo, { country, region });
    if (!state) cityOnlyMap.set(city + '|', { country, region });
  });

  // Lead data
  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();

  // Counters & rows for highlighting
  let fillCount = 0;
  let missingCountryCount = 0;
  let missingRegionCount = 0;
  let highlightRows = [];

  for (let i = 0; i < leadData.length; i++) {
    const row = leadData[i];
    const city = (row[leadCityIdx] || '').toString().trim().toLowerCase();
    const state = (row[leadStateIdx] || '').toString().trim().toLowerCase();
    let country = (row[leadCountryIdx] || '').toString().trim();
    let region = (row[leadRegionIdx] || '').toString().trim();

    let found = false;
    if (city) {
      // Try city+state
      const keyCombo = city + '|' + state;
      if (cityStateMap.has(keyCombo)) {
        const mapped = cityStateMap.get(keyCombo);
        if (!country && mapped.country) {
          row[leadCountryIdx] = mapped.country;
          country = mapped.country;
          fillCount++;
          found = true;
        }
        if (!region && mapped.region) {
          row[leadRegionIdx] = mapped.region;
          region = mapped.region;
          fillCount++;
          found = true;
        }
      } else if (cityOnlyMap.has(city + '|')) {
        // Fallback to city only mapping
        const mapped = cityOnlyMap.get(city + '|');
        if (!country && mapped.country) {
          row[leadCountryIdx] = mapped.country;
          country = mapped.country;
          fillCount++;
          found = true;
        }
        if (!region && mapped.region) {
          row[leadRegionIdx] = mapped.region;
          region = mapped.region;
          fillCount++;
          found = true;
        }
      }
    }

    // Track missing counts
    if (!country) missingCountryCount++;
    if (!region) missingRegionCount++;

    if (found) highlightRows.push(i + 2);
  }

  // Update sheet if any change
  if (fillCount > 0) {
    leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);

    if (config.highlightFilledCells) {
      const rangeNotations = highlightRows.map(r => leadSheet.getRange(r, leadCountryIdx + 1).getA1Notation())
        .concat(highlightRows.map(r => leadSheet.getRange(r, leadRegionIdx + 1).getA1Notation()));
      leadSheet.getRangeList(rangeNotations).setBackground(config.highlightColor);
    }
  }

  ss.toast(`Filled ${fillCount} missing Company Country/Region cells`, '✅ Success ✅', -1);

  ui.alert(
    `✅ ${fillCount} cells updated with missing values from mapping.\n` +
    `❗ Rows with missing Company Country: ${missingCountryCount}\n` +
    `❗ Rows with missing Region: ${missingRegionCount}`
  );
}


/** ✅
 * Fixes incorrect Country values in 'CityStateCountryRegionMapping' based on corrected entries
 * in 'RightCityCountryMapping'. Outputs a deduplicated sheet with unique cities into
 * 'CityStateCountryRegionMapping_TRUE'.
 * 
 * - Matches cities from RightCityCountryMapping and overrides Country values in base sheet if needed.
 * - Keeps full City, State, corrected Country, and Region.
 * - Skips duplicate cities (based on first/best valid correction).
 */
function correctCityCountryMapping() {
  const config = {
    baseSheet: 'CityStateCountryRegionMapping',
    correctionSheet: 'RightCityCountryMapping',
    outputSheet: 'CityStateCountryRegionMapping_TRUE',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region',
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const baseSheet = ss.getSheetByName(config.baseSheet);
  const fixSheet = ss.getSheetByName(config.correctionSheet);
  if (!baseSheet || !fixSheet) {
    ui.alert(`❗One or more sheets not found: '${config.baseSheet}' or '${config.correctionSheet}'`);
    return;
  }

  const baseHeaders = baseSheet.getRange(1, 1, 1, baseSheet.getLastColumn()).getValues()[0];
  const fixHeaders  = fixSheet.getRange(1, 1, 1, fixSheet.getLastColumn()).getValues()[0];

  const baseCityIdx    = baseHeaders.indexOf(config.columns.city);
  const baseStateIdx   = baseHeaders.indexOf(config.columns.state);
  const baseCountryIdx = baseHeaders.indexOf(config.columns.country);
  const baseRegionIdx  = baseHeaders.indexOf(config.columns.region);

  const fixCityIdx     = fixHeaders.indexOf(config.columns.city);
  const fixCountryIdx  = fixHeaders.indexOf(config.columns.country);

  if ([baseCityIdx, baseStateIdx, baseCountryIdx, baseRegionIdx, fixCityIdx, fixCountryIdx].some(idx => idx === -1)) {
    ui.alert('❗ One or more required columns are missing in source or correction sheet. Please review column headers.');
    return;
  }

  const baseData = baseSheet.getRange(2, 1, baseSheet.getLastRow() - 1, baseSheet.getLastColumn()).getValues();
  const fixData = fixSheet.getRange(2, 1, fixSheet.getLastRow() - 1, fixSheet.getLastColumn()).getValues();

  // Create correction lookup: { city.toLowerCase() : correct_country }
  const fixMap = new Map();
  fixData.forEach(row => {
    const city = (row[fixCityIdx] || '').toString().trim().toLowerCase();
    const correctCountry = (row[fixCountryIdx] || '').toString().trim();
    if (city && correctCountry) {
      fixMap.set(city, correctCountry);
    }
  });

  // Deduplication and correction application
  const cityMap = new Map(); // key: city → full cleaned row

  baseData.forEach(row => {
    const city = (row[baseCityIdx] || '').toString().trim();
    if (!city) return;

    const cityKey = city.toLowerCase();
    const originalState = (row[baseStateIdx] || '').toString().trim();
    let originalCountry = (row[baseCountryIdx] || '').toString().trim();
    const region = (row[baseRegionIdx] || '').toString().trim();

    // Apply correction if present
    if (fixMap.has(cityKey)) {
      originalCountry = fixMap.get(cityKey);
    }

    // Deduplicate: prefer first-found or more complete entry
    if (!cityMap.has(cityKey)) {
      cityMap.set(cityKey, [city, originalState, originalCountry, region]);
    } else {
      const existing = cityMap.get(cityKey);
      // Preference logic: keep entry with more filled values
      const newFilled =
        [city, originalState, originalCountry, region].filter(x => !!x).length;
      const oldFilled = existing.filter(x => !!x).length;

      if (newFilled > oldFilled) {
        cityMap.set(cityKey, [city, originalState, originalCountry, region]);
      }
    }
  });

  // Output result
  const output = Array.from(cityMap.values());

  // Write output to new sheet
  let outSheet = ss.getSheetByName(config.outputSheet);
  if (outSheet) outSheet.clearContents();
  else outSheet = ss.insertSheet(config.outputSheet);

  outSheet.getRange(1, 1, 1, 4).setValues([
    [config.columns.city, config.columns.state, config.columns.country, config.columns.region]
  ]);
  if (output.length > 0) {
    outSheet.getRange(2, 1, output.length, 4).setValues(output);
  }
  outSheet.setFrozenRows(1);

  ss.toast(`✅ Fixed City-Country mappings applied.\n📄 Final sheet: ${config.outputSheet}`, '✅ Success ✅', -1);
  ui.alert(`✅ Mapping corrections completed.\n${output.length} unique cities exported to '${config.outputSheet}'.`);
}


/**
 * Cleans and aligns every row in 'Lead_CleanedData' to canonical mapping from 'CityStateCountryRegionMapping'.
 * - Fixes any city/country mixups in columns.
 * - For each city found in mapping, writes back correct State, Country, Region.
 * - Leaves other columns unchanged.
 */
function fixLeadCleanedDataFromMapping_WRONG() {
  const config = {
    leadSheetName: 'Lead_CleanedData',
    masterMappingSheet: 'CityStateCountryRegionMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region'
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const leadSheet = ss.getSheetByName(config.leadSheetName);
  const mappingSheet = ss.getSheetByName(config.masterMappingSheet);
  if (!leadSheet || !mappingSheet) {
    ui.alert(`❗ One or both sheets not found: "${config.leadSheetName}", "${config.masterMappingSheet}"`);
    return;
  }

  const leadHeaders = leadSheet.getRange(1,1,1,leadSheet.getLastColumn()).getValues()[0];
  const mappingHeaders = mappingSheet.getRange(1,1,1,mappingSheet.getLastColumn()).getValues()[0];

  const leadCityIdx = leadHeaders.indexOf(config.columns.city);
  const leadStateIdx = leadHeaders.indexOf(config.columns.state);
  const leadCountryIdx = leadHeaders.indexOf(config.columns.country);
  const leadRegionIdx = leadHeaders.indexOf(config.columns.region);

  const mapCityIdx = mappingHeaders.indexOf(config.columns.city);
  const mapStateIdx = mappingHeaders.indexOf(config.columns.state);
  const mapCountryIdx = mappingHeaders.indexOf(config.columns.country);
  const mapRegionIdx = mappingHeaders.indexOf(config.columns.region);

  if (
    [leadCityIdx, leadStateIdx, leadCountryIdx, leadRegionIdx,
     mapCityIdx, mapStateIdx, mapCountryIdx, mapRegionIdx].some(idx => idx === -1)
  ) {
    ui.alert(`❗ One or more columns not found in 'Lead_CleanedData' or 'CityStateCountryRegionMapping'.`);
    return;
  }

  // --- Build canonical (city -> {state, country, region}) from mapping, keep most complete row per city
  const mappingData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow()-1, mappingSheet.getLastColumn()).getValues();
  const cityMap = new Map();
  mappingData.forEach(row => {
    const city = (row[mapCityIdx] || '').toString().trim();
    const state = (row[mapStateIdx] || '').toString().trim();
    const country = (row[mapCountryIdx] || '').toString().trim();
    const region = (row[mapRegionIdx] || '').toString().trim();
    if (!city) return;
    const key = city.toLowerCase();

    // Prefer row with most complete mapping
    const score = [city, state, country, region].filter(v => v).length;
    let replace = true;
    if (cityMap.has(key)) {
      const existing = cityMap.get(key);
      const existingScore = [existing.city, existing.state, existing.country, existing.region].filter(v => v).length;
      if (existingScore >= score) replace = false;
    }
    if (replace) {
      cityMap.set(key, {city, state, country, region});
    }
  });

  // --- Scan "Lead_CleanedData", fix swapped columns and values as needed
  const leadData = leadSheet.getRange(2,1,leadSheet.getLastRow()-1,leadSheet.getLastColumn()).getValues();
  let fixCount = 0;
  let cityFieldFixCount = 0;
  let countryFieldFixCount = 0;

  for (let i=0; i < leadData.length; i++) {
    let city = (leadData[i][leadCityIdx] || '').toString().trim();
    let state = (leadData[i][leadStateIdx] || '').toString().trim();
    let country = (leadData[i][leadCountryIdx] || '').toString().trim();
    let region = (leadData[i][leadRegionIdx] || '').toString().trim();

    let originalCity = city, originalCountry = country;

    // Detect and fix swaps (city in country column, country in city column)
    // If a city column value matches a known country, and the country column value matches a known city, swap them.
    if (
      country &&
      city &&
      cityMap.has(country.toLowerCase()) && // what's in the 'country' field is actually a city!
      !cityMap.has(city.toLowerCase())      // 'city' field's value isn't an actual city
    ) {
      // Swap detected
      let tmp = city;
      city = country;
      country = tmp;
      leadData[i][leadCityIdx] = city;
      leadData[i][leadCountryIdx] = country;
      cityFieldFixCount++;
      countryFieldFixCount++;
    }

    // Now, use mapping for this city
    const key = city.toLowerCase();
    if (city && cityMap.has(key)) {
      const m = cityMap.get(key);

      // If state is wrong or missing, fix
      if (m.state && state !== m.state) {
        leadData[i][leadStateIdx] = m.state;
        state = m.state; // update for further checks
        fixCount++;
      }
      // If country is wrong or missing, fix
      if (m.country && country !== m.country) {
        leadData[i][leadCountryIdx] = m.country;
        country = m.country;
        fixCount++;
      }
      // If region is wrong or missing, fix
      if (m.region && region !== m.region) {
        leadData[i][leadRegionIdx] = m.region;
        region = m.region;
        fixCount++;
      }
    }
  }

  // Write data back to sheet
  leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);

  ui.alert(
    `✅ Lead_CleanedData corrected using CityStateCountryRegionMapping.\n` +
    `Rows fixed: ${fixCount}\n` +
    `Col swaps fixed: [City→Country: ${cityFieldFixCount}, Country→City: ${countryFieldFixCount}]`
  );

  ss.toast('Lead_CleanedData now matches canonical mapping.','✅ Mapping sync complete!',-1);
}


/**
 * Fixes misaligned City/State/Country columns in 'Lead_CleanedData' using reference data from 'CityStateCountryRegionMapping'.
 * - Identifies values in wrong columns (e.g., city is in country column)
 * - Swaps them to correct columns if unambiguous.
 * - Fills out missing or incorrect fields with authoritative values.
 * - Produces finalized rows for Company City, State, Country, Region.
 */
function fixCityCountrySwapAndFillFromMapping() {
  const config = {
    leadSheetName: 'Lead_CleanedData',
    mappingSheetName: 'CityStateCountryRegionMapping',
    columns: {
      city: 'Company City',
      state: 'Company State',
      country: 'Company Country',
      region: 'Region'
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const leadSheet = ss.getSheetByName(config.leadSheetName);
  const mappingSheet = ss.getSheetByName(config.mappingSheetName);
  if (!leadSheet || !mappingSheet) {
    ui.alert(`❗ Missing required sheets: ${config.leadSheetName} or ${config.mappingSheetName}`);
    return;
  }

  // Get column indexes
  const leadHeaders = leadSheet.getRange(1, 1, 1, leadSheet.getLastColumn()).getValues()[0];
  const leadCityIdx = leadHeaders.indexOf(config.columns.city);
  const leadStateIdx = leadHeaders.indexOf(config.columns.state);
  const leadCountryIdx = leadHeaders.indexOf(config.columns.country);
  const leadRegionIdx = leadHeaders.indexOf(config.columns.region);

  const mapHeaders = mappingSheet.getRange(1, 1, 1, mappingSheet.getLastColumn()).getValues()[0];
  const mapCityIdx = mapHeaders.indexOf(config.columns.city);
  const mapStateIdx = mapHeaders.indexOf(config.columns.state);
  const mapCountryIdx = mapHeaders.indexOf(config.columns.country);
  const mapRegionIdx  = mapHeaders.indexOf(config.columns.region);

  if ([leadCityIdx, leadStateIdx, leadCountryIdx, leadRegionIdx,
       mapCityIdx, mapStateIdx, mapCountryIdx, mapRegionIdx].some(i => i === -1)) {
    ui.alert(`❗ One or more required columns not found in either sheet.`);
    return;
  }

  // Build sets and city-to-full-record map
  const citySet = new Set();
  const countrySet = new Set();
  const stateSet = new Set();
  const cityMap = new Map(); // canonical: city.toLowerCase() => {city, state, country, region}

  const mapData = mappingSheet.getRange(2,1, mappingSheet.getLastRow()-1, mappingSheet.getLastColumn()).getValues();
  mapData.forEach(row => {
    const city    = (row[mapCityIdx] || '').toString().trim();
    const state   = (row[mapStateIdx] || '').toString().trim();
    const country = (row[mapCountryIdx] || '').toString().trim();
    const region  = (row[mapRegionIdx] || '').toString().trim();

    if (city) citySet.add(city.toLowerCase());
    if (state) stateSet.add(state.toLowerCase());
    if (country) countrySet.add(country.toLowerCase());

    if (city) {
      const key = city.toLowerCase();
      if (!cityMap.has(key)) {
        cityMap.set(key, {city, state, country, region});
      }
    }
  });

  // Process lead data
  const leadData = leadSheet.getRange(2, 1, leadSheet.getLastRow() - 1, leadSheet.getLastColumn()).getValues();

  let fixCount = 0;
  let swapped = 0;

  for (let i = 0; i < leadData.length; i++) {
    let rawCity = (leadData[i][leadCityIdx] || '').toString().trim();
    let rawState = (leadData[i][leadStateIdx] || '').toString().trim();
    let rawCountry = (leadData[i][leadCountryIdx] || '').toString().trim();
    let rawRegion = (leadData[i][leadRegionIdx] || '').toString().trim();

    let cityVal = rawCity;
    let stateVal = rawState;
    let countryVal = rawCountry;
    let changed = false;

    // Detect if these values are potentially in the wrong place
    const valMap = {
      [rawCity.toLowerCase()]: 'city',
      [rawState.toLowerCase()]: 'state',
      [rawCountry.toLowerCase()]: 'country'
    };

    const seen = {
      city: citySet.has(rawCity.toLowerCase()),
      state: stateSet.has(rawState.toLowerCase()),
      country: countrySet.has(rawCountry.toLowerCase())
    };

    // Swap if they appear in wrong places (and only one is right)
    const cityIsActuallyCountry = countrySet.has(rawCity.toLowerCase()) && !citySet.has(rawCity.toLowerCase());
    const countryIsActuallyCity = citySet.has(rawCountry.toLowerCase()) && !countrySet.has(rawCountry.toLowerCase());

    if (cityIsActuallyCountry && countryIsActuallyCity) {
      // Swap city & country
      [cityVal, countryVal] = [countryVal, cityVal];
      leadData[i][leadCityIdx] = cityVal;
      leadData[i][leadCountryIdx] = countryVal;
      swapped++;
      changed = true;
    }

    // Now use city map to override everything if city match exists
    const cityKey = cityVal.toLowerCase();
    if (cityMap.has(cityKey)) {
      const mapped = cityMap.get(cityKey);

      if (mapped.state && mapped.state !== stateVal) {
        leadData[i][leadStateIdx] = mapped.state;
        changed = true;
      }
      if (mapped.country && mapped.country !== countryVal) {
        leadData[i][leadCountryIdx] = mapped.country;
        changed = true;
      }
      if (mapped.region && mapped.region !== rawRegion) {
        leadData[i][leadRegionIdx] = mapped.region;
        changed = true;
      }

      // Also overwrite city column just in case it's capitalized differently
      if (mapped.city !== cityVal) {
        leadData[i][leadCityIdx] = mapped.city;
        changed = true;
      }
    }

    if (changed) fixCount++;
  }

  // Write updated data
  leadSheet.getRange(2, 1, leadData.length, leadSheet.getLastColumn()).setValues(leadData);

  ui.alert(`✅ Clean-up complete.\n${fixCount} rows updated.\n🪄 ${swapped} rows had city & country swapped.`);
  ss.toast('Lead_CleanedData normalized with master mapping ✅', 'Done', -1);
}

/**
 * Categorizes job titles into predefined standard roles and inserts a 'Designation' column in the configured sheet.
 * 
 * Roles:
 *   1. Founder / CEO
 *   2. CTO / CIO / IT Head
 *   3 .VP / Director of Engineering
 *   4. QA Manager / QA Lead
 *   5. Product Owner / Manager
 *   6. Project Manager
 * Configuration:
 * - `sheetName`: The name of the source sheet.
 * - `titleColumnName`: The column containing job titles.
 * - `designationColumnName`: The column where categorization will be placed or created.
 */
function categorizeScatteredTitlesIntoDesignation() {
  const config = {
    sheetName: 'Lead_CleanedData',
    titleColumnName: 'Title',
    designationColumnName: 'Designation',
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.sheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert(`❗ Sheet '${config.sheetName}' not found.`);
    return;
  }

  ss.toast('Categorizing job titles...', '⏳ Please wait', -1);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let titleColIdx = headers.indexOf(config.titleColumnName);
  let designationColIdx = headers.indexOf(config.designationColumnName);

  if (titleColIdx === -1) {
    ui.alert(`❗ Column '${config.titleColumnName}' not found.`);
    return;
  }

  // Insert Designation column if it does not exist
  if (designationColIdx === -1) {
    sheet.insertColumnAfter(titleColIdx + 1);
    sheet.getRange(1, titleColIdx + 2).setValue(config.designationColumnName);
    designationColIdx = titleColIdx + 1;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert("⚠️ No data to process.");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // Categorize function using regex patterns based on your Gemini logic
  function categorize(title) {
  if (!title) return 'Other';
  const t = title.toLowerCase();

  if (/founder|ceo|chief executive officer|president|owner|co[- ]?founder|co[- ]?ceo|algemeen directeur|geschäftsführer|md and co - founder|managing director \(cto\)|chairman/i.test(t)) {
    return 'Founder / CEO';
  }
  if (/cto|chief technology officer|chief technical officer|cheif technology officer|cio|chief information officer|it head|head of it|it director|chief digital officer|chief digital & information officer|chief technology & data officer|chief technology & innovation|chief technology & operations officer|chief information security officer|ciso|head of technology|group cio|group cto|corporate information technology manager|dsi \/ cto|information chief & technology architect|it manager|global head of it|interim cio|interim cto|partner & cto|chief ai architect|chief artificial intelligence officer|chief technology lead/i.test(t)) {
    return 'CTO / CIO / IT Head';
  }
  if (/vp engineering|vice president engineering|director of engineering|engineering director|head of engineering|engineering manager|associate engineering manager|lead engineer|lead developer|technical director|software engineering manager|hardware engineering manager|cloud engineering manager|data engineering manager|applied engineering manager|architect , customer first product success and quality|solutions architect|associate director, engineering|group engineering manager|global head of engineering|principal engineer|sr. manager, engineering|staff software engineer|software architect|assistant manager engineering|vp of engineering|vp of software engineering/i.test(t)) {
    return 'VP / Director of Engineering';
  }
  if (/qa manager|qa lead|quality assurance manager|quality assurance lead|quality lead|director - quality assurance|head of quality assurance|quality engineer|quality control|associate director - quality assurance|assistant director - quality assurance|lead quality assurance engineer|software quality assurance|test manager|quality assurance business head|quality and devops|assurance quality operations|head - purchase quality assurance|it director of quality assurance|practice head , quality assurance|principal quality assurance analyst|principal quality assurance engineer|director engineering quality assurance|manager - assurance quality services/i.test(t)) {
    return 'QA Manager / QA Lead';
  }
  if (/product manager|product owner|head of product|chief product officer|cpo|vp product|director of product|associate product manager|assistant product manager|product lead|product strategist|product specialist|agile product manager|ai product owner|brand product manager|digital product manager|clinical product leader|chief products officer|chief product & engineering officer|chief product & strategy officer|chief product & tech officer|chief product & technology officer|chief product and business development officer|chief product and marketing officer|chief product and operations officer|chief product and technology officer|chief product manager|chief product owner|global product manager|principal product manager|principal product owner|product development manager|product marketing manager|product management|product management lead|produktmanager|^product$/i.test(t)) {
    return 'Product Owner / Manager';
  }
  if (/project manager|program manager|agile program manager|agile project manager|assistant director - project manager|assistant director , project manager|associate director , project manager|business program manager|project coordinator|pmo manager|global product \/ project manager|global program manager|head of program & project delivery|head of project management office|information technology project manager|it project manager|senior project manager/i.test(t)) {
    return 'Project Manager';
  }
  return 'Other';
}


  // Categorize titles
  for (let i = 0; i < data.length; i++) {
    const titleValue = data[i][titleColIdx];
    const role = categorize(titleValue);
    data[i][designationColIdx] = role;
  }

  // Write back results
  sheet.getRange(2, 1, data.length, sheet.getLastColumn()).setValues(data);

  ss.toast('Categorization complete ✅', 'Done', -1);
  ui.alert(`✅ Job titles successfully categorized for ${data.length} rows in '${config.sheetName}'.`);
}

/**
 * Normalizes various region names in 'Lead_CleanedData' sheet to a fixed set of 8 regions.
 * 
 * This script looks at the 'Region' column and:
 * - Replaces all equivalent/variant names (e.g., 'APAC', 'Oceania') with a standardized name (e.g., 'Asia-Pacific').
 * - Updates the values directly in-place.
 * - Shows a completion toast and summary alert.
 * 
 * 🛠️ Configuration:
 * - `sheetName`: Source sheet where region data exists.
 * - `regionColumnName`: Name of the column that holds region values.
 * - `regionMap`: A mapping of lowercased variant names → standardized region labels (Final 8).
 */
function normalizeRegionsInSheet() {
  const config = {
    sheetName: 'Lead_CleanedData',
    regionColumnName: 'Region',
    regionMap: {
      // Variants → Standard 8-region mapping
      'africa':             'Africa',
      'apac':               'Asia-Pacific',
      'asia-pacific':       'Asia-Pacific',
      'oceania':            'Asia-Pacific',
      'emea':               'Europe',
      'europe':             'Europe',
      'latam':              'Latin America',
      'latin america':      'Latin America',
      'mena':               'MENA',
      'middle east':        'Middle East',
      'north america':      'North America',
      'southeast asia':     'Southeast Asia',
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(config.sheetName);

  if (!sheet) {
    ui.alert(`❗ Sheet '${config.sheetName}' not found.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const regionIdx = headers.indexOf(config.regionColumnName);

  if (regionIdx === -1) {
    ui.alert(`❗ Column '${config.regionColumnName}' not found in '${config.sheetName}'.`);
    return;
  }

  ss.toast('Normalizing Region values...', '🔄 Processing', -1);

  const totalRows = sheet.getLastRow() - 1;
  if (totalRows <= 0) {
    ui.alert(`⚠️ No data rows found in '${config.sheetName}'.`);
    return;
  }

  const dataRange = sheet.getRange(2, 1, totalRows, sheet.getLastColumn());
  const data = dataRange.getValues();

  let updatedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const regionRaw = (data[i][regionIdx] || '').toString().trim();
    const regionKey = regionRaw.toLowerCase();

    if (config.regionMap[regionKey]) {
      const normalized = config.regionMap[regionKey];
      if (regionRaw !== normalized) {
        data[i][regionIdx] = normalized;
        updatedCount++;
      }
    }
  }

  // Write back if any changes applied
  if (updatedCount > 0) {
    dataRange.setValues(data);
  }

  ss.toast(
    `✅ Region normalization complete. ${updatedCount} updated.`,
    '✅ Success',
    5
  );

  ui.alert(
    `✅ Region values normalized in '${config.sheetName}'.\n\n` +
    `🔁 Total rows updated: ${updatedCount}\n` +
    `🎯 Final Regions:\n- Africa\n- Asia-Pacific\n- Europe\n- Latin America\n- MENA\n- Middle East\n- North America\n- Southeast Asia`
  );
}

