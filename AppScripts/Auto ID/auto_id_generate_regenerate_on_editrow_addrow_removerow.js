function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Only work on sheets ending with "-Expenses"
  if (!sheet.getName().endsWith("-Expenses")) return;

  var colB = 2; // Column B (Expense Date)
  var row = range.getRow();
  var col = range.getColumn();

  // Only trigger if editing column B (and not header)
  if (col === colB && row > 1) {
    generateExpenseIDs(sheet);
  }
}

// Function to generate IDs
function generateExpenseIDs(sheet) {
  var prefix = sheet.getName().substring(0, 2).toUpperCase();
  var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues(); // Column B, excluding header
  var ids = [];
  var count = 1;

  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      ids.push([`${prefix}-${count}`]);
      count++;
    } else {
      ids.push([""]);
    }
  }
  // Write all IDs at once for efficiency
  sheet.getRange(2, 1, ids.length, 1).setValues(ids);
}

// Optional: Add a menu to force re-numbering
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Expense Tools")
    .addItem("Regenerate Expense IDs", "regenerateIDsMenu")
    .addToUi();
  regenerateIDsMenu();
}

function regenerateIDsMenu() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Only work on sheets ending with "-Expenses"
  if (!sheet.getName().endsWith("-Expenses")) return;

  // Uncomment the following line if you want to enforce the user to switch to an '-Expenses' sheet
  // if (!sheet.getName().endsWith("-Expenses")) {
  //   SpreadsheetApp.getUi().alert("Please switch to an '-Expenses' sheet first.");
  //   return;
  // }
  generateExpenseIDs(sheet);
  SpreadsheetApp.getUi().alert("Expense IDs regenerated!");
}
