function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tejas Notifications 🚀')
      .addItem('TOAST - Show Simple Toast', 'showSimpleToast')
      .addItem('TOAST - Show Toast with Title', 'showToastWithTitle')
      .addItem('TOAST - Show Toast with Custom Duration', 'showToastWithDuration')
      .addItem('TOAST - Show Toast Until Clicked', 'showToastUntilClicked')
      .addItem('ALERT - Show OK Alert', 'showOkAlert')
      .addItem('ALERT - Show OK/Cancel Alert', 'showOkCancelAlert')
      .addItem('ALERT - Show Yes/No Alert', 'showYesNoAlert')
      .addItem('ALERT - Show Yes/No/Cancel Alert', 'showYesNoCancelAlert')
      .addItem('ALERT - Show Alert with Title', 'showAlertWithTitle')
      .addToUi();
}

function showSimpleToast() {
  SpreadsheetApp.getActive().toast('This is a simple toast message.');
}

function showToastWithTitle() {
  SpreadsheetApp.getActive().toast('Operation completed!', 'Info');
}

function showToastWithDuration() {
  // Show a toast for 10 seconds
  SpreadsheetApp.getActive().toast('This message will disappear in 10 seconds.', 'Notice', 10);
}

function showToastUntilClicked() {
  // Show a toast that stays until clicked
  SpreadsheetApp.getActive().toast('I am sticky Toast message. Will be here until you Click me.', 'Attention', -1);
}

function showOkAlert() {
  SpreadsheetApp.getUi().alert("This is an OK alert.");
}

function showOkCancelAlert() {
  var result = SpreadsheetApp.getUi().alert("Do you want to proceed?", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (result === SpreadsheetApp.getUi().Button.OK) {
    SpreadsheetApp.getActive().toast("You chose to proceed.");
  } else {
    SpreadsheetApp.getActive().toast("You canceled the action.");
  }
}

function showYesNoAlert() {
  var result = SpreadsheetApp.getUi().alert("Do you approve the changes?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
  if (result === SpreadsheetApp.getUi().Button.YES) {
    SpreadsheetApp.getActive().toast("Changes approved.");
  } else {
    SpreadsheetApp.getActive().toast("Changes not approved.");
  }
}

function showYesNoCancelAlert() {
  var result = SpreadsheetApp.getUi().alert("Do you want to save your changes?", SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL);
  if (result === SpreadsheetApp.getUi().Button.YES) {
    SpreadsheetApp.getActive().toast("Changes saved.");
  } else if (result === SpreadsheetApp.getUi().Button.NO) {
    SpreadsheetApp.getActive().toast("Changes discarded.");
  } else {
    SpreadsheetApp.getActive().toast("Action canceled.");
  }
}

function showAlertWithTitle() {
  SpreadsheetApp.getUi().alert("Alert Title", "This is an alert with a title.", SpreadsheetApp.getUi().ButtonSet.OK);
}
