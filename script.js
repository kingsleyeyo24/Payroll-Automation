function pullActiveEmployees() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName('General Data');
  var target = ss.getSheetByName('PayrollData');

  // Get source data 
  var data = source.getDataRange().getValues();
  var headers = data[0];
  var isActiveCol = headers.indexOf("IsActive"); 
  if (isActiveCol === -1) throw new Error("IsActive column not found.");

  // Filter active employees
  var active = data.slice(1).filter(r => r[isActiveCol] === "Yes");

  // Clear old rows (only where new data goes)
  if (target.getLastRow() > 1) {
    target.getRange(2, 1, target.getLastRow() - 1, headers.length).clearContent();
  }

  // Write active employees to target sheet
  if (active.length) {
    target.getRange(2, 1, active.length, headers.length).setValues(active);
  }

  Logger.log("PayrollData updated with " + active.length + " active employees.");
}

