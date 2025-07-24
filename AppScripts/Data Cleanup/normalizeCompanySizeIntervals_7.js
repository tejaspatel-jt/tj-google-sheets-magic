function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Cleanup Tools 🚀')

    // CLEAN-UP PART
    // .addItem('1️⃣ Freeze First Row in All Sheets ✅', 'freezeFirstRowInAllSheets')
    .addItem('8️⃣ ➕➕ 📏 Normalize Company Size Intervals', 'normalize_CompanySize_Intervals')
    .addToUi();
    
}

/**
 * 📏 Normalize Company Size Intervals with Output Column Configuration 🏢➡️📊
 * 
 * 🧠 This function cleans and standardizes company size data by mapping numeric values
 * from either '# Employees' (preferred if non-empty) or 'Size of Company' columns
 * into defined size intervals such as '0-10', '11-50', etc.
 * 
 * - It extracts the first number from each source cell (ignoring commas and plus).
 * - Matches this number to configured intervals.
 * - Writes the normalized label into a configurable output column.
 * - If the output column does not exist, it creates the column immediately 
 *   after the 'Size of Company' column (or '# Employees' if 'Size of Company' is missing).
 * - If output column is the same as 'Size of Company', then it overwrites in place.
 * - Skips empty or invalid values by labeling them as 'NA'.
 * 
 * 🔁 Safe to rerun multiple times; doesn't alter unchanged cells unnecessarily.
 * 📊 Displays progress with toast and final UI alert.
 */
function normalize_CompanySize_Intervals() {
  const config = {
    sheetName: 'Leads_MasterData', // Set your actual master sheet name here
    columns: {
      employees: '# Employees',
      sizeOfCompany: 'Size of Company',
      outputColumn: 'Size of Company' // Can be same as sizeOfCompany to overwrite, Yes We have kept the same
    },
    intervals: [
      { label: '0 - 10', min: 0, max: 10 },
      { label: '11 - 50', min: 11, max: 50 },
      { label: '51 - 200', min: 51, max: 200 },
      { label: '201 - 500', min: 201, max: 500 },
      { label: '501 - 1000', min: 501, max: 1000 },
      { label: '1001 - 2000', min: 1001, max: 2000 },
      { label: '2001 - 3000', min: 2001, max: 3000 },
      { label: '3001 - 4000', min: 3001, max: 4000 },
      { label: '4001 - 5000', min: 4001, max: 5000 },
      { label: '5001 - 10000', min: 5001, max: 10000 },
      { label: '10001 - 20000', min: 10001, max: 20000 },
      { label: '20001 - 30000', min: 20001, max: 30000 },
      { label: '30001 - 40000', min: 30001, max: 40000 },
      { label: '50001+', min: 50001, max: Number.MAX_SAFE_INTEGER }
    ]
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) {
    ui.alert(`❗ Sheet "${config.sheetName}" not found.`);
    return;
  }

  // ✅ Toast: Operation start
  ss.toast('🔄 Starting company size normalization...', 'Processing', -1);

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];

  // 🔍 Locate relevant columns
  const employeesCol = headers.indexOf(config.columns.employees);
  const sizeCol = headers.indexOf(config.columns.sizeOfCompany);
  const outputColName = config.columns.outputColumn.trim();
  let outputCol = headers.indexOf(outputColName);

  if (employeesCol < 0 || sizeCol < 0) {
    ui.alert(`❗ Required columns missing. Ensure both '${config.columns.employees}' and '${config.columns.sizeOfCompany}' exist.`);
    return;
  }

  // 🔄 Handle output column creation or clearing existing data if present
  if (outputCol === -1) {
    // 🔧 Prefer insertion after 'Size of Company', fallback to '# Employees'
    const insertAfterCol = sizeCol !== -1 ? sizeCol : employeesCol;
    sheet.insertColumnAfter(insertAfterCol + 1);
    outputCol = insertAfterCol + 1;
    headers.splice(outputCol, 0, outputColName);
    values[0] = headers; // Reflect updated header in memory
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]); // Write new header to sheet
  } else if (outputCol !== sizeCol && outputColName !== config.columns.sizeOfCompany) {
    // 🧹 Clear existing output column content (except header)
    sheet.getRange(2, outputCol + 1, sheet.getLastRow() - 1, 1).clearContent();
  } else if (outputColName === config.columns.sizeOfCompany) {
    // ✍️ Overwrite existing 'Size of Company'
    outputCol = sizeCol;
  }

  // 🔍 Extracts the first number from a string, ignoring commas and plus signs (cleaned)
  function extractFirstNumber(str) {
    if (str === null || str === undefined || String(str).trim() === '') return NaN;
    const cleaned = String(str).replace(/[,+]/g, '').trim();
    const match = cleaned.match(/^\d+/);
    return match ? parseInt(match[0], 10) : NaN;
  }

  // 🧠 Returns the interval label for a given number or 'NA' if invalid
  function getIntervalLabelFromNumber(num) {
    if (isNaN(num) || num < 0) return 'NA';
    const interval = config.intervals.find(interval => num >= interval.min && num <= interval.max);
    return interval ? interval.label : 'NA';
  }

  let updateCount = 0;
  const outputData = [];

  // 🔁 Process all data rows
  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    // 🔎 Pick '# Employees' if valid else 'Size of Company'
    const sourceVal = (employeesCol !== -1 && row[employeesCol] !== null && String(row[employeesCol]).trim() !== '') ?
                      row[employeesCol] :
                      (sizeCol !== -1 && row[sizeCol] !== null && String(row[sizeCol]).trim() !== '') ?
                      row[sizeCol] : '';

    let label = 'NA';

    // 📊 If there's a value, normalize it
    if (sourceVal !== '') {
      const num = extractFirstNumber(sourceVal);
      label = getIntervalLabelFromNumber(num);
    }

    outputData.push([label]);

    // ✍️ Track changes to avoid unnecessary writes
    if (outputCol !== -1 && row[outputCol] !== label) {
      updateCount++;
    }
  }

  // 📝 Write all normalized labels back in one batch to the output column
  if (outputCol !== -1 && outputData.length > 0) {
    sheet.getRange(2, outputCol + 1, outputData.length, 1).setValues(outputData);
  } else if (outputCol !== -1) {
    // 🧹 Clear column if no data rows but output column exists
    sheet.getRange(2, outputCol + 1, sheet.getLastRow() - 1, 1).clearContent();
  }

  // ✅ Toast: Operation complete
  ss.toast(`✅ Company size normalization complete! Rows updated: ${updateCount}`, '✅ Done', 5);

  // 🛎 Final UI alert with multiline message and OK button
  ui.alert(
    "✅ Success ✅",
    `Company size intervals normalization completed successfully.\n
    Rows updated: ${updateCount}`,
    ui.ButtonSet.OK
  );
}