/**
 * This script provides utility functions for Google Sheets:
 * 
 * 1. Freezes the first row in all sheets of the active spreadsheet.
 * 2. Combines data from all sheets, excluding those with "❌" in their name or specified in an exclusion list.
 *    - The combined data is placed in a new sheet named "CombinedData".
 *    - The script collects all unique headers from the sheets and aligns the data accordingly.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Analytics Tools 🚀')
    .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    .addItem('Combine Data (Exclude ❌ and Specified Sheets)', 'mergeSheetsByHeaders')
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
  //ss.toast('Combining data, please wait...', 'Info', 5);
  ss.toast('Please wait while Combining data from all the sheets.. \n YOU WILL BE NOTIFIED ONCE DONE !', '⚠️ Attention ⚠️', -1);

  // Start timing
  var startTime = new Date();

  var allSheets = ss.getSheets();
  var allHeaders = [];
  var dataRows = [];
  
  // === CONFIGURABLE EXCLUSIONS ===
  var excludeSheetNames = ["CombinedData", "Analytics"]; // <-- Add any other sheet names to exclude here

  // Filter out sheets with "❌" in the name or in the exclusion list
  var validSheets = allSheets.filter(function(sheet) {
    var name = sheet.getName();
    return name.indexOf("❌") === -1 && excludeSheetNames.indexOf(name) === -1;
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
  
  // 2. Collect all data, mapping to the master header order
  validSheets.forEach(function(sheet) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, Math.max(0, sheet.getLastRow()-1), sheet.getLastColumn()).getValues();
    var headerMap = {};
    headers.forEach(function(header, idx) {
      headerMap[header] = idx;
    });
    data.forEach(function(row) {
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

  // Show alert when done, with duration
  ui.alert('Data combined successfully! ✅\n\nTime taken: ' + durationStr);

}