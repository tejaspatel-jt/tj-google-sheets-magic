/**
 * This script provides utility functions for Google Sheets:
 * 
 * 1. Freezes the first row in all sheets of the active spreadsheet.
 * 
 * 2. Combines data from all sheets, excluding those with "❌" in their name or specified in an exclusion list.
 *    - The combined data is placed in a new sheet named "CombinedData".
 *    - The script collects all unique headers from the sheets and aligns the data accordingly.
 *    - It skips completely empty rows and fills in missing columns with blanks.
 * 
 * 3. Automatically fills the "Company Country" column based on the "Company City" column.
 *    - If a city has a corresponding country, it fills in the country for rows where it is missing.
 *    - It highlights the filled 'Company Country' cells with a light red background.
 * 
 * 4. Creates a "Filter_CombinedData" sheet with vertical headers and checkboxes for filtering.
 * 
 * 5. Show Count of combined rows and the time taken to complete the operation.
 * 
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Analytics Tools 🚀')
    .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    .addSeparator()
    .addItem('Combine Data (Exclude ❌ and Specified Sheets)', 'mergeSheetsByHeaders')
    .addSeparator()
    .addItem('Auto-Fill Company Country', 'autoFillCompanyCountryMenu')
    .addItem('Create Filter_CombinedData Sheet', 'createFilterCombinedDataSheetMenu')
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

// 2. Combine data from all sheets except those in the exclusion list or with "❌" in their name
function mergeSheetsByHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Please wait while Combining data from all the sheets.. \n YOU WILL BE NOTIFIED ONCE DONE !', '⚠️ Attention ⚠️', -1);

  // Start timing
  var startTime = new Date();

  var allSheets = ss.getSheets();
  var allHeaders = [];
  var dataRows = [];
  
  // === CONFIGURABLE EXCLUSIONS ===
  var excludeSheetNames = ["CombinedData", "Analytics", "Filter_CombinedData"]; // <-- Add any other sheet names to exclude here

  // Filter out sheets with "❌" in the name or in the exclusion list
  var validSheets = allSheets.filter(function(sheet) {
    var name = sheet.getName();
    var lowerName = name.toLowerCase();
    return (
      name.indexOf("❌") === -1 &&
      excludeSheetNames.indexOf(name) === -1 &&
      lowerName.indexOf("combineddata") === -1 &&
      lowerName.indexOf("old") === -1
    );
  });
  
  // 1. Collect all unique headers
  validSheets.forEach(function(sheet) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(function(header) {
      if (header && allHeaders.indexOf(header) === -1) {
        allHeaders.push(header);
      }
    });
  });
  
  // 2. Collect all data, mapping to the master header order, and skip empty rows
  validSheets.forEach(function(sheet) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, Math.max(0, sheet.getLastRow()-1), sheet.getLastColumn()).getValues();
    var headerMap = {};
    headers.forEach(function(header, idx) {
      headerMap[header] = idx;
    });
    data.forEach(function(row) {
      // Skip completely empty rows
      if (row.every(function(cell) { return cell === "" || cell === null; })) return;
      var newRow = [];
      allHeaders.forEach(function(masterHeader) {
        if (headerMap.hasOwnProperty(masterHeader)) {
          newRow.push(row[headerMap[masterHeader]]);
        } else {
          newRow.push(""); // Fill blanks for missing columns
        }
      });
      dataRows.push(newRow);
    });
  });
  
  // 3. Output to the "CombinedData" sheet
  var outSheetName = "CombinedData";
  var outSheet = ss.getSheetByName(outSheetName);
  if (!outSheet) {
    outSheet = ss.insertSheet(outSheetName);
  } else {
    outSheet.clearContents();
  }

  // Freeze the first row
  outSheet.setFrozenRows(1);
  
  outSheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
  if (dataRows.length > 0) {
    outSheet.getRange(2, 1, dataRows.length, allHeaders.length).setValues(dataRows);
  }

  // End timing
  var endTime = new Date();
  var durationMs = endTime - startTime;
  var seconds = Math.floor((durationMs / 1000) % 60);
  var minutes = Math.floor((durationMs / (1000 * 60)) % 60);
  var durationStr = '';
  if (minutes > 0) {
    durationStr += minutes + ' min' + (minutes > 1 ? 's' : '');
    if (seconds > 0) durationStr += ' ';
  }
  if (seconds > 0 || minutes === 0) {
    durationStr += seconds + ' sec' + (seconds !== 1 ? 's' : '');
  }

  // Count of combined rows (excluding header)
  var combinedRowsCount = dataRows.length;

  // Show alert when done, with duration and row count
  ui.alert('Data combined successfully! ✅\n\nTime taken: ' + durationStr + '\nRows combined: ' + combinedRowsCount);
}

// 3. Auto-fill Company Country based on Company City mapping (menu wrapper)
function autoFillCompanyCountryMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CombinedData");
  if (!sheet) {
    SpreadsheetApp.getUi().alert('CombinedData sheet not found!');
    return;
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  autoFillCompanyCountry(sheet, headers);
  SpreadsheetApp.getUi().alert('Auto-fill Company Country completed!');
}

// 3b. Core logic (can be reused elsewhere)
function autoFillCompanyCountry(sheet, headers) {
  var cityCol = headers.indexOf("Company City");
  var countryCol = headers.indexOf("Company Country");
  if (cityCol === -1 || countryCol === -1) return; // If either column doesn't exist, skip

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // No data

  var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  // Build map: city => country (from rows where both are present)
  var cityToCountry = {};
  data.forEach(function(row) {
    var city = row[cityCol];
    var country = row[countryCol];
    if (city && country && !(city in cityToCountry)) {
      cityToCountry[city] = country;
    }
  });

  // Fill missing country where city matches and country is empty
  var toFill = [];
  data.forEach(function(row, i) {
    var city = row[cityCol];
    var country = row[countryCol];
    if (city && (!country || country === "")) {
      if (cityToCountry[city]) {
        row[countryCol] = cityToCountry[city];
        toFill.push(i + 2); // Row number in sheet (offset by header)
      }
    }
  });

  // Write back updated data
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  // Highlight filled cells with "light red 3" (#f4cccc)
  if (toFill.length > 0) {
    var range = sheet.getRangeList(
      toFill.map(function(r) {
        return sheet.getRange(r, countryCol + 1).getA1Notation();
      })
    );
    range.setBackground('#f4cccc');
  }
}

// 4. Create Filter_CombinedData sheet with vertical headers and checkboxes (menu wrapper)
function createFilterCombinedDataSheetMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var combinedSheet = ss.getSheetByName("CombinedData");
  if (!combinedSheet) {
    SpreadsheetApp.getUi().alert('CombinedData sheet not found!');
    return;
  }
  var headers = combinedSheet.getRange(1, 1, 1, combinedSheet.getLastColumn()).getValues()[0];
  createFilterCombinedDataSheet(ss, headers);
  SpreadsheetApp.getUi().alert('Filter_CombinedData sheet created!');
}

// 4b. Core logic (can be reused elsewhere)
function createFilterCombinedDataSheet(ss, headers) {
  var sheetName = "Filter_CombinedData";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }
  for (var i = 0; i < headers.length; i++) {
    sheet.getRange(i + 1, 1).setValue(headers[i]);
    sheet.getRange(i + 1, 2).insertCheckboxes();
  }
}
