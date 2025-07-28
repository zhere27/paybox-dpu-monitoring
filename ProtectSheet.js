
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

    CustomLogger.logInfo("Protected sheet: " + sheet.getName(), PROJECT_NAME, 'protectAllSheets()');
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

  for (var i = 9; i < sheets.length; i++) { // Index 3 corresponds to the fourth sheet
    sheets[i].hideSheet();
  }
  CustomLogger.logInfo("Hide sheets after Kiosk%",PROJECT_NAME, 'hideSheets()');
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

