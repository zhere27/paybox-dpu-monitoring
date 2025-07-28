// function protectAllSheets() {
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   var sheets = spreadsheet.getSheets();

//   sheets.forEach(function (sheet) {
//     var protection = sheet.protect();
//     protection.setWarningOnly(false); // Optional: Change to false to prevent edits without permission

//     // Optional: Add editors who can edit the protected sheets
//     // protection.addEditor("cvcabanilla@multisyscorp.com"); 
//     CustomLogger.logInfo("Protected sheet: " + sheet.getName(),PROJECT_NAME, 'protectAllSheets()');
//   });
// }
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

function hideSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  for (var i = 9; i < sheets.length; i++) { // Index 3 corresponds to the fourth sheet
    sheets[i].hideSheet();
  }
  CustomLogger.logInfo("Hide sheets after Kiosk%",PROJECT_NAME, 'hideSheets()');
}

function hideColumnsCtoH() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Kiosk %");
  sheet.hideColumns(3, 6); // Start hiding from column C (index 3), for 6 columns (C to H)

}

