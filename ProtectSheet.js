/**
 * Protects all sheets in the active Google Spreadsheet that are not already protected.
 * Skips sheets that already have protection applied.
 * Applies strict protection (not warning-only).
 * Optionally, editors can be added to the protection.
 * Logs each protected sheet using CustomLogger.
 *
 * @function
 * @returns {void}
 */
function protectAllSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  sheets.forEach(function (sheet) {
    // Check if the sheet already has a protection
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (protections.length > 0) {
      return; // Skip this sheet
    }

    // Apply protection if none exists
    var protection = sheet.protect();
    protection.setWarningOnly(false); // Strict protection

    // Optional: Add editors
    // protection.addEditor("egalcantara@multisyscorp.com");

    CustomLogger.logInfo("Protected sheet: " + sheet.getName(), CONFIG.APP.NAME, 'protectAllSheets()');
  });
}

/**
 * Hides all sheets in the active spreadsheet starting from the 10th sheet onward.
 * Logs the action using CustomLogger.
 *
 * @function
 * @returns {void}
 */
function hideSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var visibleSheets = sheets.filter(function (sheet) {
    return sheet.isSheetHidden() === false;
  });

  // Ensure at least one sheet stays visible
  if (visibleSheets.length <= 1) {
    console.warn("Cannot hide any more sheets — at least one sheet must remain visible.");
    return;
  }

  // Hide sheets starting from index 10 (0-based, so this hides the 11th onward)
  for (var i = 10; i < sheets.length; i++) {
    // Only hide if it's currently visible
    if (!sheets[i].isSheetHidden()) {
      // Check again that there will be at least one visible sheet left
      var remainingVisible = sheets.filter(s => !s.isSheetHidden()).length;
      if (remainingVisible <= 1) {
        console.warn("Stopped hiding to prevent all sheets from being hidden.");
        break;
      }
      sheets[i].hideSheet();
    }
  }

  console.log("Total sheets: " + sheets.length);
  console.log("Visible sheets after hiding: " + sheets.filter(s => !s.isSheetHidden()).length);
}

/**
 * Hides columns C to H (columns 3 to 8) in the "Kiosk %" sheet of the active Google Spreadsheet.
 *
 * @function
 * @returns {void}
 */
function hideColumnsCtoH() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Kiosk %");
  sheet.hideColumns(3, 6); // Start hiding from column C (index 3), for 6 columns (C to H)

}

