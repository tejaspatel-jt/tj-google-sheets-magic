function onOpen() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Get the active sheet
    var lastRow = sheet.getLastRow(); // Get the last row with data in the sheet
    var lastActiveRow = sheet.getRange("C1:C" + lastRow).getValues() // Get values in column C
      .filter(String).length; // Filter out empty values to find the last active row in column C
    
    if (lastActiveRow > 0) {
      sheet.setActiveRange(sheet.getRange(lastActiveRow + 10, 3)); // Set the active range to the last active row in column C
    } else {
      Logger.log("Column C is empty."); // Log if column C is empty
    }
  
      // var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange('C:C');
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  
  function onEdit(e) {
    // Get the active sheet and the edited range
    var sheet = e.source.getActiveSheet();
    var range = e.range;
  
    // Set the wrap text option to true for the edited cell
    //range.setWrap(true);
  
   // Check if the edited cell is not empty
    if (range.getValue() !== "") {
      // Set the wrap text option to true for the edited cell
      range.setWrap(true);
    }
  
    // Check for Unique Bugs
    // Check if the active sheet is "My Bugs"
    if (true || sheet.getName() === "TRY") {
      const range = e.range;
  
      // Check if the edited cell is in column A
      if (range.getColumn() === 1) {
  
        const bugId = range.getValue();
        
        // If the cell is empty, do not check for duplicates
        if (bugId === "") {
          return;
        }
  
        const bugIds = sheet.getRange("A:A").getValues().flat();
  
        // Filter out empty values
        const uniqueBugIds = bugIds.filter(id => id !== "");
  
        // Check for duplicates
        const isDuplicate = uniqueBugIds.filter(id => id === bugId).length > 1;
  
        if (isDuplicate) {
          // If duplicate found, show an alert and clear the cell
          SpreadsheetApp.getUi().alert("😡 Bug ID already Exists.. ⚠️");
          range.setValue("");
          range.setBackground("yellow"); // Set the background color to red
          range.setValue("Duplicate");
        }
       
      }
    }
  
  
  
  }
  