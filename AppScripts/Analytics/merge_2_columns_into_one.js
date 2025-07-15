/**
 * This script merges the "Job title" column into the "Title" column in the "CombinedData" sheet.
 * 
 * If "Title" is empty, it takes the value from "Job title".
 * If both columns have values, it highlights those rows in light red.
 * If there are no conflicts, it deletes the "Job title" column after merging.
 * 
 * @returns {void}
 */
function mergeTitleColumnsInCombinedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CombinedData");
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❗ CombinedData sheet not found!');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const titleColIndex = headers.indexOf("Title");
  const jobTitleColIndex = headers.indexOf("Job title");

  if (titleColIndex === -1 || jobTitleColIndex === -1) {
    SpreadsheetApp.getUi().alert('❗ Either "Title" or "Job title" column is missing in CombinedData.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('✅ No data rows to process.');
    return;
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const data = dataRange.getValues();

  const highlightRows = [];
  let mergedCount = 0;

  data.forEach((row, rowIndex) => {
    const title = row[titleColIndex];
    const jobTitle = row[jobTitleColIndex];

    if ((!title || title === "") && jobTitle) {
      // Merge: Move Job title to Title
      row[titleColIndex] = jobTitle;
      mergedCount++;
    } else if (title && jobTitle) {
      // Conflict: both have values
      highlightRows.push(rowIndex + 2); // +2 accounts for header + zero-based index
    }
  });

  // Update modified data back to the sheet
  dataRange.setValues(data);

  // Highlight conflict rows if any
  if (highlightRows.length > 0) {
    const ranges = highlightRows.map(r =>
      sheet.getRange(r, 1, 1, sheet.getLastColumn()).getA1Notation()
    );
    sheet.getRangeList(ranges).setBackground("#f4cccc");
  }

  // Conditionally delete "Job title" column if there are no conflicts
  let columnDeleted = false;
  if (highlightRows.length === 0) {
    sheet.deleteColumn(jobTitleColIndex + 1); // Always add 1 to zero-based index
    columnDeleted = true;
  }

  // Final message
  SpreadsheetApp.getUi().alert(
    `✅ Merge Summary:\n` +
    `• ${mergedCount} row(s) merged from "Job title" to "Title".\n` +
    (highlightRows.length > 0
      ? `• ${highlightRows.length} conflicted row(s) found and highlighted in light red.\n❗ "Job title" column NOT deleted for review.`
      : `• No conflicts found.\n✅ "Job title" column has been deleted.`)
  );
}
