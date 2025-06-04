/**
 * Automatically updates IDs in Column A for any sheet ending with "-Expenses"
 * after any change (edit, row insert, row delete, paste, etc.).
 * Also adds a menu "Expense Tools" for manual regeneration.
 * 
 * For This you have to add Trigger in the script editor:
 * 1. Open the script editor in Google Sheets.
 * 2. Click on the clock icon (Triggers).
 * 3. Click on "+ Add Trigger".
 * 4. Choose `onChange` as the function to run.
 * 5. Set the event type to "On change".
 * 6. Save the trigger.
 * 
 * This script assumes that the first row is a header and starts processing from the second row.
 */

// Auto-triggered on any change in the spreadsheet
function onChange(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet.getName().endsWith("-Expenses")) return;
  regenerateExpenseIDs(sheet);
}

// Main function to regenerate IDs
function regenerateExpenseIDs(sheet) {
  var prefix = sheet.getName().substring(0, 2).toUpperCase();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // No data

  var colBValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Column B, excluding header
  var ids = [];
  var count = 1;

  for (var i = 0; i < colBValues.length; i++) {
    if (colBValues[i][0]) {
      ids.push([prefix + "-" + count]);
      count++;
    } else {
      ids.push([""]);
    }
  }
  sheet.getRange(2, 1, ids.length, 1).setValues(ids);
}

// Optional: Adds a menu for manual regeneration
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Expense Tools')
    .addItem('Regenerate IDs', 'menuRegenerateExpenseIDs')
    .addToUi();
}

function menuRegenerateExpenseIDs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet.getName().endsWith("-Expenses")) {
    SpreadsheetApp.getUi().alert("This script only works on sheets ending with '-Expenses'.");
    return;
  }
  regenerateExpenseIDs(sheet);
  SpreadsheetApp.getUi().alert("Expense IDs regenerated!");
}
