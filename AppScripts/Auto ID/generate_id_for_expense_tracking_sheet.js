function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Check if the sheet name ends with "Expenses"
  if (!sheet.getName().endsWith("-Expenses")) return;

  var colA = 1; // Column A for ID
  var colB = 2; // Column B to trigger ID generation
  var prefix = sheet.getName().substring(0, 2).toUpperCase(); // Prefix from sheet name

  // Check if the edit was made in Column B and not in the header row
  if (range.getColumn() === colB && range.getRow() > 1) {
    var row = range.getRow();
    var valueInB = range.getValue();

    // If there's a value in Column B, generate IDs in Column A
    if (valueInB) {
      var lastRow = sheet.getLastRow();
      var count = 1;

      for (var i = 2; i <= lastRow; i++) {
        var cellB = sheet.getRange(i, colB).getValue();

        if (cellB) {
          sheet.getRange(i, colA).setValue(`${prefix}-${count}`); // Set ID in Column A
          count++;
        } else {
          sheet.getRange(i, colA).clearContent(); // Clear ID if Column B is empty
        }
      }
    } else {
      sheet.getRange(row, colA).clearContent(); // Clear ID if Column B is cleared
    }
  }
}
