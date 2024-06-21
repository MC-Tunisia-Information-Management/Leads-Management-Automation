function onEdit(e) {
  var ss = e.source; // Get the active spreadsheet
  var sheet = ss.getSheetByName("Settings");

  // Check if the edit occurred in the "Settings" sheet and the edited cell is in column B (member names) or column F (emails)
  if (sheet && (e.range.getColumn() === 2 || e.range.getColumn() === 6)) {
    var memberName = sheet.getRange(e.range.getRow(), 2).getValue();
    var email = sheet.getRange(e.range.getRow(), 6).getValue();

    // Check if both the member name and email are not empty
    if (memberName && email) {
      // Check if a sheet with the same name doesn't already exist
      if (!ss.getSheetByName(memberName)) {
        // Duplicate the "Members Standardized Sheet" as a template
        var templateSheet = ss.getSheetByName("Draft leads for deps");
        var newSheet = templateSheet.copyTo(ss);

        // Rename the new sheet
        newSheet.setName(memberName);

        // Insert the member's name into the first row (A1:K1) of the new sheet
        var memberNameRange = newSheet.getRange("A1:J1");
        memberNameRange.setValue(memberName);

        // Protect the specified ranges in the new sheet
        protectSpecifiedRanges(newSheet);

        // Update the QUERY formula
        updateQueryFormula(ss);
      }
    }
  } else if (sheet && e.range.getA1Notation() === "B5") {
    updateQueryFormula(ss);
  }
}

function createQueryFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName("Settings");
  var sheetNamesRange = settingsSheet.getRange("B5:B");
  var sheetNames = sheetNamesRange.getValues().flat().filter(String);

  if (sheetNames.length > 0) {
    for (var i = 0; i < sheetNames.length; i++) {
      var formula =
        "=QUERY(Local All leads!A3:J, \"select Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, Col9, Col10 where Col11 = '" +
        sheetNames[i] +
        "'\", 0)";

      // Insert the formula into a cell in each sheet
      var memberSheet = ss.getSheetByName(sheetNames[i]);
      if (memberSheet) {
        var queryCell = memberSheet.getRange("A3"); // Change this to the cell where you want to place the formula
        queryCell.setValue(formula);
      }
    }
  }
}
