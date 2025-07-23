/**
 * This script provides utility functions for Google Sheets:
 * 
 * 1. This script is designed to generate a report of missing details in the 'Lead_CleanedData' sheet.
 *    - It checks pairs or triplets of columns for missing data.
 *    - The report is generated in a new or cleared sheet named 'MissingDetails'.
 *    - It can handle pairs of columns (e.g., City and Country) or triplets (e.g., State, City, and Country).
 *    -   
 * 
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Cleanup Tools 🚀')
    // .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    .addItem('❓ Find Missing Details - Triplets ⁉️', 'generateMissingDetailsAdvancedReport')
    .addToUi();
}

/**
 * Generates a detailed report of missing details in the 'Lead_CleanedData' sheet.
 * 
 *    - It checks pairs or triplets of columns for missing data and conflicts.
 *    - The report is generated in a new or cleared sheet named 'MissingDetails'.
 *    - Handles both pairs (e.g., City and Country) and triplets (e.g., State, City, and Country).
 *    - It can also detect conflicts where multiple values exist for a single key.
 * 
 *    - Configuration is done at the top of the script for easy modification.
 *        - like which columns to check, and whether to include pairs or triplets.
 *        - like `targetSheetName` for defining the source sheet name.
 *        - like `columnSets` for defining which columns to check.
 *        - like `missingSheetName` for defining the output sheet name.
 */
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

