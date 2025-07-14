/**
 * This script provides 4 separate utility functions for Google Sheets:
 *      1.  Freeze First Row in All Sheets
 *      2.  Combine Data in Batches (Exclude ❌ and Specified Sheets)
 *      3.  Auto-Fill Company Country
 *      4.  Create Filter_CombinedData Sheet
 * 
 * DETAILED DESCRIPTION:
 * 
 * 1. Freezes the first row in all sheets of the active spreadsheet.
 *    - The function iterates through all sheets and sets the first row as frozen.
 * 
 * 2. Combines data from all sheets, excluding those with "❌" in their name or specified in an exclusion list.
 *    - The script processes sheets in batches (default batch size is 3) to avoid timeouts.
 *    - The combined data is placed in a new sheet named "CombinedData".
 *    - The script collects all unique headers from the sheets and aligns the data accordingly.
 *    - It skips completely empty rows and fills in missing columns with blanks.
 *    - It maintains a progress state using script properties to handle large datasets without losing progress.
 *    - shows alert after each batch is processed, indicating how many sheets were processed and how many remain. 
 * 
 * 
 * 3. Automatically fills the "Company Country" column based on the "Company City" column.
 *    - If a city has a corresponding country, it fills in the country for rows where it is missing.
 *    - It highlights the filled 'Company Country' cells with a light red background.
 * 
 * 4. Creates a "Filter_CombinedData" sheet with vertical headers and checkboxes for filtering.
 *    - The headers are copied from the "CombinedData" sheet and placed vertically.
 *    - Each header has a corresponding checkbox for filtering purposes.
 * 
 * 5. Show Count of combined rows and the time taken to complete the operation.
 * 
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas Analytics Tools 🚀')
    .addItem('Freeze First Row in All Sheets', 'freezeFirstRowInAllSheets')
    .addItem('Combine Data in Batches (Exclude ❌ and Specified Sheets)', 'batchCombineSheets')
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

// 2. BATCH COMBINE
function batchCombineSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Batch combining data... You will be notified once this batch is done!', '⚠️ Attention ⚠️', -1);

  var allSheets = ss.getSheets();
  var excludeSheetNames = ["CombinedData", "Analytics", "Filter_CombinedData"];
  
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

  var scriptProperties = PropertiesService.getScriptProperties();
  var allHeaders = JSON.parse(scriptProperties.getProperty('allHeaders') || '[]');
  var lastProcessedIndex = Number(scriptProperties.getProperty('lastProcessedIndex')) || 0;
  var batchSize = 3; // Process 3 sheets per run (change as needed)
  var end = Math.min(lastProcessedIndex + batchSize, validSheets.length);
  var sheetsToProcess = validSheets.slice(lastProcessedIndex, end);

  // 1. Collect all unique headers (if first batch)
  if (lastProcessedIndex === 0) {
    sheetsToProcess.forEach(function(sheet) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      headers.forEach(function(header) {
        if (header && allHeaders.indexOf(header) === -1) {
          allHeaders.push(header);
        }
      });
    });
    // Also scan remaining sheets for headers (to avoid missing columns)
    for (var i = end; i < validSheets.length; i++) {
      var headers = validSheets[i].getRange(1, 1, 1, validSheets[i].getLastColumn()).getValues()[0];
      headers.forEach(function(header) {
        if (header && allHeaders.indexOf(header) === -1) {
          allHeaders.push(header);
        }
      });
    }
    scriptProperties.setProperty('allHeaders', JSON.stringify(allHeaders));
    // Clear or create CombinedData sheet and write headers
    var outSheet = ss.getSheetByName("CombinedData");
    if (!outSheet) outSheet = ss.insertSheet("CombinedData");
    else outSheet.clearContents();
    outSheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);

    // Freeze the first row
    outSheet.setFrozenRows(1);

    scriptProperties.setProperty('combinedRowsCount', 0);
    scriptProperties.setProperty('startTime', new Date().getTime().toString());
  } else {
    // Read headers from property
    allHeaders = JSON.parse(scriptProperties.getProperty('allHeaders') || '[]');
  }

  // 2. Collect and append data for this batch
  var dataRows = [];
  sheetsToProcess.forEach(function(sheet) {
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

  // Append to CombinedData (robust: always create if missing)
  var outSheet = ss.getSheetByName("CombinedData");
  if (!outSheet) {
    outSheet = ss.insertSheet("CombinedData");
    outSheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
    outSheet.setFrozenRows(1);
  }
  var existingRows = outSheet.getLastRow();
  if (dataRows.length > 0) {
    outSheet.getRange(existingRows + 1, 1, dataRows.length, allHeaders.length).setValues(dataRows);
  }

  // Update progress
  scriptProperties.setProperty('lastProcessedIndex', end);
  var totalCombinedRows = Number(scriptProperties.getProperty('combinedRowsCount')) + dataRows.length;
  scriptProperties.setProperty('combinedRowsCount', totalCombinedRows);

  // Check if done
  if (end >= validSheets.length) {
    // Calculate duration
    var startTime = Number(scriptProperties.getProperty('startTime'));
    var endTime = new Date().getTime();
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
    // Clean up properties
    scriptProperties.deleteProperty('lastProcessedIndex');
    scriptProperties.deleteProperty('allHeaders');
    scriptProperties.deleteProperty('startTime');
    scriptProperties.deleteProperty('combinedRowsCount');
    ui.alert('All Batches Processed ✔️\n Data in All Sheets combined successfully! ✅\n\nTime taken: ' + durationStr + '\nRows combined: ' + totalCombinedRows);

  } else {

    // Build processed/unprocessed sheet name lists
    var processedSheets = validSheets.slice(0, end).map(function (sheet) {
      return sheet.getName() + " ✅";
    });
    var unprocessedSheets = validSheets.slice(end).map(function (sheet) {
      return sheet.getName() + " ⭕";
    });

    var processedMsg = processedSheets.length > 0
      ? "Processed sheets:\n" + processedSheets.join('\n')
      : "No sheets processed yet.";

    var unprocessedMsg = unprocessedSheets.length > 0
      ? "\n\nUnprocessed sheets:\n" + unprocessedSheets.join('\n')
      : "";

    ui.alert(
      '👉 Processed sheets : ' + lastProcessedIndex + ' to ' + (end - 1) + '.\n' +
      '👉 Run "Combine Data in Batches" again to continue ⚠️.\n\n' +
      '👉 Rows combined so far: ' + totalCombinedRows + '\n\n' +
      processedMsg + unprocessedMsg
    );
  }
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
