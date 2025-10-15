function onOpen() {
  createMenu();
}

function onInstall(e) {
  createMenu();
}

function createMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Monitoring")
    .addItem("Refresh Stores", "refreshStores")
    // .addItem('Process Hourly', 'processMonitoringHourly')
    // .addItem('Update Remarks','lookupLastTwoEntries')
    // .addItem('Sort', 'sortLatestPercentage')
    .addToUi();
}
function getCollectedStores() {
  try {
    // Get or create the target sheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Collected Yesterday";
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      Logger.log(`Sheet '${sheetName}' created.`);
    }

    // Clear previous data
    sheet.getRange("A2:E300").clear();

    // SQL Query
    const QUERY = `
      SELECT DISTINCT created_at, machine_name, amount, count, collector_name
      FROM ms-paybox-prod-1.pldtsmart.collections
      WHERE DATE(created_at) >= DATE(CURRENT_DATE()) - 1
      ORDER BY created_at DESC
    `;

    // Execute query
    var queryResult = executeQueryAndWait(QUERY);
    if (!queryResult || !queryResult.rows) throw new Error("No data returned from the query.");

    // Process and populate data
    var rows = queryResult.rows;
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i].f;
      var formattedDate = formatTimestamp(row[0].v);
      setCellValue(sheet, i + 2, formattedDate, row[1].v, row[2].v, row[3].v, row[4].v);
    }

    CustomLogger.logInfo(`Populated collected stores.`, PROJECT_NAME, 'getCollectedStores()');

    // Ensure changes are visible in the sheet
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log(`Error in getCollectedStores: ${error.message}`);
  }
}

/**
 * Formats a timestamp (seconds since epoch) into a readable date string.
 * @param {number} timestamp - Timestamp in seconds since epoch.
 * @returns {string} - Formatted date string.
 */
function formatTimestamp(timestamp) {
  try {
    var dateValue = new Date(timestamp * 1000); // Convert seconds to milliseconds
    return Utilities.formatDate(dateValue, "UTC", "yyyy-MM-dd HH:mm:ss");
  } catch (error) {
    Logger.log(`Error formatting timestamp: ${error.message}`);
    return "";
  }
}

/**
 * Sets the value and formatting for a row in the target sheet.
 * @param {Sheet} sheet - The target sheet.
 * @param {number} rowIndex - The row index (1-based).
 * @param {string} date - Formatted date.
 * @param {string} machineName - Machine name.
 * @param {number} amount - Amount.
 * @param {number} count - Count.
 * @param {string} collectorName - Collector name.
 */
function setCellValue(sheet, rowIndex, date, machineName, amount, count, collectorName) {
  try {
    var row = [
      { value: date, alignment: "Left", format: "@STRING@" },
      { value: machineName, alignment: "Left", format: "@STRING@" },
      { value: amount, alignment: "Right", format: "###,###,##0.00" },
      { value: count, alignment: "Center", format: "###,###,##0" },
      { value: collectorName, alignment: "Left", format: "@STRING@" },
    ];

    row.forEach((cell, index) => {
      var range = sheet.getRange(rowIndex, index + 1);
      range.setValue(cell.value)
        .setFontFamily("Century Gothic")
        .setFontSize(9)
        .setHorizontalAlignment(cell.alignment)
        .setNumberFormat(cell.format);
    });
  } catch (error) {
    Logger.log(`Error setting cell value for row ${rowIndex}: ${error.message}`);
  }
}
const PROJECT_NAME = "Paybox - DPU Monitoring v2";

/**
 * Collections Helper Functions
 * This file contains utility functions for managing collection operations
 */

// === Constants === //

/**
 * Mapping of day indices to day abbreviations
 * @type {Object.<number, string>}
 */
const dayMapping = {
  0: 'Sun.',
  1: 'M.',
  2: 'T.',
  3: 'W.',
  4: 'Th.',
  5: 'F.',
  6: 'Sat.'
};

/**
 * Mapping of day abbreviations to day indices
 * @type {Object.<string, number>}
 */
const dayIndex = {
  'Sun.': 0,
  'M.': 1,
  'T.': 2,
  'W.': 3,
  'Th.': 4,
  'F.': 5,
  'Sat.': 6
};

/**
 * Amount thresholds for collection by day
 * @type {Object.<string, number>}
 */
const amountThresholds = {
  'M.': 300000,
  'T.': 310000,
  'W.': 310000,
  'Th.': 300000,
  'F.': 290000,
  'Sat.': 290000,
  'Sun.': 290000
};

// const amountThresholds = {
//   'M.': 250000,
//   'T.': 250000,
//   'W.': 250000,
//   'Th.': 250000,
//   'F.': 250000,
//   'Sat.': 250000,
//   'Sun.': 250000
// };

/**
 * Payday ranges for collection
 * @type {Array.<{start: number, end: number}>}
 */
const paydayRanges = [
  { start: 15, end: 16 },
  { start: 30, end: 31 }
];

/**
 * Due date cutoffs for collection
 * @type {Array.<{start: number, end: number}>}
 */
const dueDateCutoffs = [
  { start: 5, end: 6 },
  { start: 20, end: 21 }
];

/**
 * Amount threshold for payday collections
 * @type {number}
 */
const paydayAmount = 290000;
// const paydayAmount = 250000;

/**
 * Amount threshold for due date collections
 * @type {number}
 */
const dueDateCutoffsAmount = 290000;
// const dueDateCutoffsAmount = 250000;

/**
 * Email signature for all emails
 * @type {string}
 */
const emailSignature = `<div><div dir="ltr" class="gmail_signature"><div dir="ltr"><span><div dir="ltr" style="margin-left:0pt" align="left">Best Regards,</div><div dir="ltr" style="margin-left:0pt" align="left"><br><table style="border:none;border-collapse:collapse"><colgroup><col width="44"><col width="249"><col width="100"><col width="7"></colgroup><tbody><tr style="height:66.75pt"><td rowspan="3" style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;text-align:justify;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:47px;height:217px"><img src="https://lh3.googleusercontent.com/JNGaTKS3JTfQlHCnlwFXo14_knhu4v_WhlZCWOIFJfRPuUKjMMHWuj82yUQ0uUOxv9XNk1Nooae__kDJ1wS0st_Xe3SZvDdl3dkVSpX24SCtgfIt7ZfeTfIR8S93ndcLMdQSgm9Xyq1rykUOGv1sLo0" width="47" height="235.00000000000003" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td><td style="vertical-align:middle;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:12pt;font-family:Arial,sans-serif;background-color:transparent;font-weight:700;vertical-align:baseline">Support</span></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:12pt;font-family:Arial,sans-serif;background-color:transparent;font-weight:700;vertical-align:baseline">Paybox Operations</span></p><br></td><td colspan="2" rowspan="3" style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;margin-right:0.75pt;text-align:justify;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:118px;height:240px"><img src="https://lh6.googleusercontent.com/3QMqLONmIp2CUDF9DoGaWmEhai3RSB6fjgqFxwhxwcQoX58wlvnAVNMscDRgfOK-xv4S2bllMTzrKQSuvgqAi68syHzvqbNJziibdwTfx7A1pSWelqdkffPtJ9n6WC3JJEEcSqNYXrBthmb8cxIz5Dg" width="118" height="240" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td></tr><tr style="height:84pt"><td style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:14px;height:10px"><img src="https://lh4.googleusercontent.com/DM0kOfMz47NSJfROHj6mLS0ypJih5nHezY-SaBODOfd_oXMKxDagoXJG1WmGYaCgt0g_PLa4KQY1Btkuih7F2409F3-gjDxV8UGVeL_6bKF4l3Aze7QG33MalyKV0NmslPNz5aK3Fp8a8LX_8abfJoo" width="14" height="10" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><a href="mailto:support@paybox.ph" target="_blank"><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(17,85,204);vertical-align:baseline">support@paybox.ph</span></a></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:16px"><img src="https://lh6.googleusercontent.com/RFRWIimtGNr-ErnlmxYSpbrwRjPjBvZXbY5BdP3LD2ykCtMVGMYodZoIc-B7xHWXI3wYHcAr8FxK6d3L4hk12mbH-dTEZ1pU6pugBvzZeqvu2uLo_4BPb_zlAlT8ve3P2GD0CifMeZ_dX_qdWhPns5g" width="15" height="16" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline">+62 (2) 8835-9697</span></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:13px"><img src="https://lh4.googleusercontent.com/VuSTEwKlvSLg3DHUziWOwLBgmj1ctniwh9f7tsnECrzxxdMy1CRKnKgJaCi2m8xIRfrsdtsihCEvexpSHnDykgsRZ1WeMuxwHKQpbS-VUfBwoM2wC3oZU9i7B8vgAWf4JZPzq03GND-SOi9bD0RIIJM" width="15" height="13" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline">+62 (2) 8805-1066 (loc) 115&nbsp;</span></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:15px"><img src="https://lh3.googleusercontent.com/vLSCYEGN7oLtmhaSdBjzq9psnzQaLlY-fa7QgcbvWWjPvwrwCDr2qX7nHFglSmQHxwPPf0DmH17j6TgttmB54ke2L4x7BJYp3DiNISF5do3G2gsBdS1v9_KchSJAc-K_dh6FtGUxJtHmrOtZPDPZEnw" width="15" height="15" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><a href="https://www.multisyscorp.com/" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://www.multisyscorp.com/&amp;source=gmail&amp;ust=1731657202461000&amp;usg=AOvVaw0G3eJRh1oAS1IGmlpIrvQt"><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(17,85,204);vertical-align:baseline">www.multisyscorp.com</span></a></p></td></tr><tr style="height:30pt"><td style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;margin-top:0pt;margin-bottom:0pt"><span style="font-size:9pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:121px;height:27px"><img src="https://lh4.googleusercontent.com/m4vMFXp2L8_GL568U8By8vFWFacVg-4s_RAjIFlQMJJvTQfgik58VwVIJY8g1zW8g9n-__YYXwDYO9kZzDidrTIWrRieExlSnwvtRuEqikp5XgZ3G9xAyIXd8eFN-k42XJRXhK0APe5u1FD8si56y44" width="121" height="27" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td></tr></tbody></table></div><br><p dir="ltr" style="line-height:1.38;margin-top:8pt;margin-bottom:0pt"><span style="font-size:11pt;color:rgb(0,0,0);background-color:transparent;vertical-align:baseline;white-space:pre-wrap"><font face="times new roman, serif">DISCLAIMER:</font></span></p><p dir="ltr" style="line-height:1.38;margin-top:8pt;margin-bottom:0pt"><span style="font-size:11pt;color:rgb(0,0,0);background-color:transparent;vertical-align:baseline;white-space:pre-wrap"><font face="times new roman, serif">The content of this email is confidential and intended solely for the use of the individual or entity to whom it is addressed. If you received this email in error, please notify the sender or system manager. It is strictly forbidden to disclose, copy or distribute any part of this message.</font></span></p></span></div></div></div>`;

const storeSheetUrl = "https://docs.google.com/spreadsheets/d/1TJ10XqwS_cTQfkxKKWJaE5zhBdE2pDgIdZ_Zw9JQD_U";// === Date Functions === //

/**
 * Gets today's and tomorrow's dates in various formats
 * @return {Object} Object containing todayDate, tomorrowDate, todayDay, and tomorrowDateString
 * @property {Date} todayDate - Today's date as a Date object
 * @property {Date} tomorrowDate - Tomorrow's date as a Date object
 * @property {string} todayDateString - Today's day in "MMM d" format
 * @property {string} tomorrowDateString - Tomorrow's day in "MMM d" format
 */
function getTodayAndTomorrowDates() {
  const timeZone = Session.getScriptTimeZone();
  let todayDate = new Date();
  // todayDate = new Date(2025,9,3); //Month is 0-based in JS
  const tomorrowDate = new Date(todayDate);
  tomorrowDate.setDate(todayDate.getDate() + 1);

  return {
    todayDate,
    tomorrowDate,
    todayDateString: Utilities.formatDate(todayDate, timeZone, "MMM d"),
    tomorrowDateString: Utilities.formatDate(tomorrowDate, timeZone, "MMM d"),
    isTomorrowHoliday: isTomorrowHoliday(tomorrowDate),
  };
}

/**
 * Gets the next Monday from today
 * If today is Monday, returns today
 *
 * @returns {Date} Next Monday date
 */
function getNextMonday() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

  // Calculate days until next Monday
  // If today is Monday (1), daysUntilMonday = 0
  // If today is Tuesday (2), daysUntilMonday = 6
  // If today is Sunday (0), daysUntilMonday = 1
  const daysUntilMonday = dayOfWeek === 1 ? 0 : (8 - dayOfWeek) % 7;

  const nextMonday = new Date(today);
  nextMonday.setDate(today.getDate() + daysUntilMonday);

  return nextMonday;
}

/**
 * Formats a date according to the specified pattern
 * @param {Date} date - The date to format
 * @param {string} pattern - The format pattern
 * @return {string} Formatted date string
 */
function formatDate(date, pattern) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), pattern);
}

/**
 * Checks if execution should be skipped based on the day
 * @param {Date} todayDate - Today's date
 * @return {boolean} True if execution should be skipped, false otherwise
 */
function shouldSkipExecution(todayDate) {
  const day = todayDate.getDay();

  if (day === 0 || day === 6) {
    // 0-Sunday, 6-Saturday
    CustomLogger.logInfo("Skipping execution on weekends.", PROJECT_NAME, "shouldSkipExecution()");
    return true;
  }
  return false;
}

// === Collection Processing Functions === //

/**
 * Processes collections and sends emails
 * @param {Array} forCollections - Array of collections to process
 * @param {Date} tomorrowDate - Tomorrow's date
 * @param {string} emailTo - Email recipients (to)
 * @param {string} emailCc - Email recipients (cc)
 * @param {string} emailBcc - Email recipients (bcc)
 * @param {string} srvBank - Service bank
 */
function processCollectionsAndSendEmail(forCollections, tomorrowDate, emailTo, emailCc, emailBcc, srvBank) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo("No collections to process.", PROJECT_NAME, "processCollections()");
      return;
    }

    const collectionDate = new Date(tomorrowDate);

    // Helper: Sort by numeric value (desc), pushing NaN to the end
    const sortByAmountDescNaNLast = (a, b, index) => {
      const aAmount = Number(a[index]);
      const bAmount = Number(b[index]);

      if (Number.isNaN(aAmount) && Number.isNaN(bAmount)) return 0;
      if (Number.isNaN(aAmount)) return 1;
      if (Number.isNaN(bAmount)) return -1;

      return bAmount - aAmount; // Descending
    };

    // Helper: Sort by string (case-sensitive) at given index
    const sortByString = (a, b, index) => a[index].localeCompare(b[index]);

    // Main logic
    forCollections.sort(srvBank === "Brinks via BPI" ? (a, b) => sortByAmountDescNaNLast(a, b, 2) : (a, b) => sortByString(a, b, 0));

    // Saturday collection limit is only applicable to Brinks via BPI
    if (collectionDate.getDay() === 6 && srvBank === "Brinks via BPI") {
      // 6-Saturday
      if (forCollections.length > 4) {
        const collectionsForMonday = forCollections.slice(4);
        const collectionDateForMonday = getNextMonday();
        // collectionDateForMonday.setDate(collectionDateForMonday.getDate() + 3);
        sendEmailCollection(collectionsForMonday, collectionDateForMonday, emailTo, emailCc, emailBcc, srvBank);
      }

      sendEmailCollection(forCollections.slice(0, 4), collectionDate, emailTo, emailCc, emailBcc, srvBank);
    } else {
      sendEmailCollection(forCollections, collectionDate, emailTo, emailCc, emailBcc, srvBank);
    }
  } catch (error) {
    CustomLogger.logError(`Error in processCollections(): ${error.message}\nStack: ${error.stack}`, PROJECT_NAME, "processCollections()");
    throw error;
  }
}

// === Spreadsheet Functions === //

function addMachineToAdvanceNotice(machineName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Advance Notice`;

  // Check if the sheet already exists
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    sheet.appendRow([machineName]);
  }
}

/**
 * Creates a hidden worksheet and adds data
 * @param {Array} forCollections - Array of collections
 * @param {string} srvBank - Service bank
 */
function createHiddenWorksheetAndAddData(forCollections, srvBank) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo("No collection data to add to worksheet.", PROJECT_NAME, "createHiddenWorksheetAndAddData()");
      return;
    }

    // Open the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `For Collection -${srvBank}`;

    // Check if the sheet already exists
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      // If the sheet exists, clear its contents
      sheet.getRange(2, 1, sheet.getMaxRows(), 4).clear();
    } else {
      // If the sheet does not exist, create it
      sheet = spreadsheet.insertSheet(sheetName);
    }

    // Format the data removed the last column remarks
    forCollections = forCollections.map(function (item) {
      return item.slice(0, 4); // keep index 0 to 3
    });

    // Add data to the sheet
    const numRows = forCollections.length;
    const numColumns = forCollections[0].length; // Assume all rows have the same number of columns
    sheet.getRange(2, 1, numRows, numColumns).setValues(forCollections);

    // Hide the sheet
    sheet.hideSheet();

    // Adjust column A width for better readability
    sheet.autoResizeColumn(1);

    CustomLogger.logInfo(`Updated for collection worksheet "${sheetName}" with ${numRows} rows.`, PROJECT_NAME, "createHiddenWorksheetAndAddData()");
  } catch (error) {
    CustomLogger.logError(`Error in createHiddenWorksheetAndAddData(): ${error.message}\nStack: ${error.stack}`, PROJECT_NAME, "createHiddenWorksheetAndAddData()");
    throw error;
  }
}

function skipAlreadyCollected(machineName, forCollectionData) {
  try {
    if (!forCollectionData || forCollectionData.length === 0) {
      CustomLogger.logInfo("No collection data worksheet.", PROJECT_NAME, "skipAlreadyCollected()");
      return;
    }

    var result = null;
    for (var i = 0; i < forCollectionData.length; i++) {
      if (forCollectionData[i][0] === machineName) {
        result = forCollectionData[i][5]; // column 3
        break; // stop after first match
      }
    }

    if (result) {
      CustomLogger.logInfo(`Skipping collection for ${machineName}, it is already collected. `, PROJECT_NAME, "skipAlreadyCollected()");
      return result;
    }

    return;
  } catch (error) {
    CustomLogger.logError(`Error in skipAlreadyCollected(): ${error.message}\nStack: ${error.stack}`, PROJECT_NAME, "skipAlreadyCollected()");
    throw error;
  }
}

/**
 * Retrieves data from the "For Collection" worksheet
 * @param {string} srvBank - Service bank
 * @return {Array} Data from the "For Collection" worksheet
 * @throws {Error} If the "For Collection" worksheet is not found
 */
function getForCollections(srvBank) {
  // Open the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `For Collection -${srvBank}`;
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    CustomLogger.logInfo(`Sheet "${sheetName}" not found for reference of recently collected machines.`, PROJECT_NAME, "getForCollections()");
    throw new Error(`Sheet named "${sheetName}" not found.`);
  }

  //Iterate through the sheet and exclude machine name in column A and column F that has value TRUE from the forCollections
  const lastRow = sheet.getLastRow();

  // convert sheet range to array
  var forCollections = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  return forCollections;
}

// === Collection Logic Functions === //

/**
 * Checks if a machine should be included for collection
 * @param {string} machineName - Machine name
 * @param {number} amountValue - Amount value
 * @param {string} translatedBusinessDays - Translated business days
 * @param {Date} tomorrowDate - Tomorrow's date
 * @param {string} tomorrowDateString - Tomorrow's date string
 * @param {Date} todayDate - Today's date
 * @param {string} lastRemark - Last request
 * @param {string} srvBank - Service bank
 * @return {boolean} True if the machine should be included, false otherwise
 */
function shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank) {
  try {
    const collectionDay = dayMapping[tomorrowDate.getDay()];
    const dayOfWeek = tomorrowDate.getDay();

    // 1. Check if excluded store
    if (isExcludedStore(machineName)) {
      CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, part of the excluded stores.`, PROJECT_NAME, "shouldIncludeForCollection()");
      return false;
    }

    // 2. Check business days
    if (!translatedBusinessDays.includes(collectionDay)) {
      CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, not a collection day.`, PROJECT_NAME, "shouldIncludeForCollection()");
      return false;
    }

    // 3. Check weekend restrictions
    if (skipWeekendCollections(srvBank, tomorrowDate)) {
      CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, PROJECT_NAME, "shouldIncludeForCollection()");
      return false;
    }

    // Early returns for inclusion conditions (highest priority)

    // 1. Special collection conditions from last remarks
    if (hasSpecialCollectionConditions(lastRemark, tomorrowDateString)) {
      return true;
    }

    // 2. Schedule-based store collections
    if (CONFIG.SCHEDULE_BASED.stores.includes(machineName)) {
      return shouldCollectScheduleBasedStore(machineName, dayOfWeek);
    }

    // 3. Regular threshold-based collection
    if (meetsAmountThreshold(amountValue, collectionDay, srvBank)) {
      return true;
    }

    // 4. Payday collection
    if (isPaydayCollection(todayDate, amountValue)) {
      return true;
    }

    // 5. Due date collection
    if (isDueDateCollection(todayDate, amountValue)) {
      return true;
    }

    // Default: no collection needed
    return false;
  } catch (error) {
    CustomLogger.logError(`Error in shouldIncludeForCollection(): ${error.message}\nStack: ${error.stack}`, PROJECT_NAME, "shouldIncludeForCollection()");
    return false;
  }
}

/**
 * Checks if a machine should be excluded from collection
 * @param {string} lastRequest - Last request
 * @param {string} todayDay - Today's day
 * @return {boolean} True if the machine should be excluded, false otherwise
 */
function forExclusionBasedOnRemarks(lastRequest, todayDay, machineName = null) {
  try {
    // Early return for null/undefined/empty lastRequest
    if (!lastRequest?.trim()) {
      return false;
    }

    // Normalize inputs once
    const normalizedLastRequest = lastRequest.toLowerCase().trim();
    const normalizedTodayDay = todayDay?.toLowerCase().trim() || "";

    // Define exclusion criteria with better organization
    const exclusionReasons = [
      "waiting for collection items",
      "waiting for bank updates",
      "did not reset",
      "cassette removed",
      "manually collected",
      "for repair",
      "already collected",
      "store is closed",
      "store is not using the machine",
      "for relocation",
      normalizedTodayDay,
    ].filter((reason) => reason); // Remove empty strings

    // Check for exclusion with optimized search
    const isExcluded = exclusionReasons.some((reason) => normalizedLastRequest.includes(reason));

    // Log exclusion if machine name provided and excluded
    if (isExcluded && machineName) {
      CustomLogger.logInfo(`Skipping collection for machine ${machineName}, last remarks is "${lastRequest}".`, PROJECT_NAME, "forExclusionBasedOnRemarks()");
    }

    return isExcluded;
  } catch (error) {
    CustomLogger.logError(`Error in shouldExcludeFromCollection(): ${error.message}`, PROJECT_NAME, "forExclusionBasedOnRemarks()");
    return false; // Fail-safe: don't exclude on error
  }
}

/**
 * Checks if a remark indicates an exclusion is scheduled for a date *after* tomorrow.
 * * Returns:
 * - true: The remark contains "for collection on [Date]" AND that date is in the future 
 * (after the date provided by tomorrowDate).
 * - false: The remark does not contain the phrase, the date is invalid, or the date is
 * tomorrow or in the past.
 * * @param {string} machineName - Unused, retained for original function signature compatibility.
 * @param {Date} tomorrowDate - The reference date used for comparison (e.g., "tomorrow").
 * @param {string} lastRemark - The string to check for the exclusion phrase and date.
 */
function forExclusionNotYetScheduled(machineName, tomorrowDate, lastRemark) {
  // Configuration
  const EXCLUSION_PHRASE = 'for collection on';
  // Regex captures the month and day (e.g., "Oct 15")
  const DATE_PATTERN = /for collection on\s+([a-z]+\s+\d{1,2})/i;

  // Normalize input for basic phrase check
  const normalizedRemark = lastRemark.toLowerCase().trim();

  // 1. Check if the exclusion phrase exists
  if (!normalizedRemark.includes(EXCLUSION_PHRASE)) {
    return false;
  }

  // 2. Extract date from remark
  const match = DATE_PATTERN.exec(lastRemark);
  if (!match) {
    return false; // Phrase exists but date format is wrong
  }

  // 3. Parse the collection date
  const collectionDateStr = match[1]; // e.g., "Oct 15"
  // Use the year from the tomorrowDate parameter for accurate comparison
  const currentYear = tomorrowDate.getFullYear();
  // Attempt to create a Date object using the extracted month/day and the current year
  const collectionDate = new Date(`${collectionDateStr}, ${currentYear}`);

  // 4. Handle invalid date parsing
  if (isNaN(collectionDate.getTime())) {
    return false;
  }

  // 5. Normalize dates to midnight for comparison (remove time component)
  // This ensures comparison is purely based on the day
  const tomorrowMidnight = new Date(
    tomorrowDate.getFullYear(), 
    tomorrowDate.getMonth(), 
    tomorrowDate.getDate()
  );
  const collectionMidnight = new Date(
    collectionDate.getFullYear(), 
    collectionDate.getMonth(), 
    collectionDate.getDate()
  );

  // 6. Return true if collection date is strictly AFTER tomorrow's date (not yet scheduled)
  return collectionMidnight > tomorrowMidnight;
}


// function forExclusionNotYetScheduled(machineName, tomorrowDate, lastRemark) {
//   // Configuration
//   const EXCLUSION_PHRASE = 'for collection on';
//   const DATE_PATTERN = /for collection on\s+([a-z]+\s+\d{1,2})/i;

//   // Normalize input
//   const normalizedRemark = lastRemark.toLowerCase().trim();

//   // Check if exclusion phrase exists
//   if (!normalizedRemark.includes(EXCLUSION_PHRASE)) {
//     return false;
//   }

//   // Extract date from remark
//   const match = DATE_PATTERN.exec(lastRemark);
//   if (!match) {
//     return false;
//   }

//   // Parse the collection date
//   const collectionDateStr = match[1]; // e.g., "Oct 15"
//   const currentYear = tomorrowDate.getFullYear();
//   const collectionDate = new Date(`${collectionDateStr}, ${currentYear}`);

//   // Handle invalid date parsing
//   if (isNaN(collectionDate.getTime())) {
//     return false;
//   }

//   // Normalize dates to midnight for comparison (remove time component)
//   const tomorrowMidnight = new Date(tomorrowDate.getFullYear(), tomorrowDate.getMonth(), tomorrowDate.getDate());
//   const collectionMidnight = new Date(collectionDate.getFullYear(), collectionDate.getMonth(), collectionDate.getDate());

//   // Return true if collection date is AFTER tomorrow (not yet scheduled)
//   return collectionMidnight > tomorrowMidnight;
// }

function forExclusionRequestYesterday(previouslyRequestedMachines, machineName, srvBank) {
  // Use includes() for faster searching
  const isExcluded = previouslyRequestedMachines.includes(machineName);

  if (isExcluded) {
    CustomLogger.logInfo(`Skipping collection for machine: ${machineName}, requested for collection yesterday.`, PROJECT_NAME, "forExclusionRequestYesterday()");
  }

  return isExcluded;
}

function getPreviouslyRequestedMachineNamesByServiceBank(srvBank) {
  var machineNames = [];
  if (CONFIG.APP.ENVIRONMENT !== 'production') { return machineNames; }

  //validate srvBank if null or blank
  if (srvBank == null || srvBank == "") {
    CustomLogger.logInfo(`Service bank not found for reference of recently collected machines.`, PROJECT_NAME, "getPreviouslyRequestedMachineNamesByServiceBank()");
    return machineNames;
  }

  // Open the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `For Collection -${srvBank}`;
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    CustomLogger.logInfo(`Sheet "${sheetName}" not found for reference of recently collected machines.`, PROJECT_NAME, "getPreviouslyRequestedMachineNamesByServiceBank()");
    return machineNames;
  }

  // Get only column A data (machine names) for faster searching
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // No data rows

  machineNames = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(function (item) {
      return item != null && item !== "";
    });
  return machineNames;
}
/**
 * Checks if tomorrow is a holiday
 * @param {Date} tomorrow - Tomorrow's date
 * @return {boolean} True if tomorrow is a holiday, false otherwise
 */
function isTomorrowHoliday(tomorrow) {
  try {
    const sheetName = "StoreName Mapping";
    const rangeStart = "G3";
    const rangeEnd = "I";
    const validHolidayTypes = ["Regular Holiday", "Special Non-working Holiday"];

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet named "${sheetName}" not found.`);
    }

    const dataRange = sheet.getRange(`${rangeStart}:${rangeEnd}`);
    const data = dataRange.getValues();

    // Normalize tomorrow's date for comparison
    const normalizedTomorrow = new Date(tomorrow);
    normalizedTomorrow.setDate(normalizedTomorrow.getDate());
    normalizedTomorrow.setHours(0, 0, 0, 0);

    // Check if tomorrow is a holiday
    for (const row of data) {
      const holidayDate = row[0];
      const holiday = row[1];
      const holidayType = row[2];

      if (holidayDate instanceof Date && holidayType && validHolidayTypes.includes(holidayType)) {
        if (holidayDate.getTime() === normalizedTomorrow.getTime()) {
          // CustomLogger.logInfo(`Tomorrow (${formatDate(tomorrow, "MMM d, yyyy")}) is a holiday: ${holiday} - ${holidayType}`, PROJECT_NAME, "isTomorrowHoliday()");
          return true;
        }
      }
    }

    return false;
  } catch (error) {
    CustomLogger.logError(`Error in isTomorrowHoliday(): ${error.message}\nStack: ${error.stack}`, PROJECT_NAME, "isTomorrowHoliday()");
    return false;
  }
}

/**
 * Returns the custom abbreviation for the day of the week of the given date.
 * If no date is provided, uses today's date.
 * Custom abbreviations: ["Sun.", "M.", "T.", "W.", "Th.", "F.", "Sat."].
 *
 * @param {Date|string} [dateInput] - The input date as a Date object or a date string. If omitted, uses the current date.
 * @returns {string} The custom abbreviation for the day of the week.
 */
function getTomorrowDayCustomFormat(dateInput) {
  // Define the custom day abbreviations
  const customDays = ["Sun.", "M.", "T.", "W.", "Th.", "F.", "Sat."];

  // Parse the input date or use today's date if none is provided
  const inputDate = dateInput ? new Date(dateInput) : new Date();

  // Calculate tomorrow's date
  const tomorrow = new Date(inputDate);
  tomorrow.setDate(inputDate.getDate());

  // Get the day index (0 for Sunday, 1 for Monday, ..., 6 for Saturday)
  const dayIndex = tomorrow.getDay();

  // Get the custom abbreviation
  const tomorrowAbbreviation = customDays[dayIndex];

  // Log or return the result
  return tomorrowAbbreviation;
}

/**
 * Ensures that a string has only one space between words.
 * Removes leading and trailing spaces as well.
 *
 * @param {string} inputString - The string to process.
 * @return {string} - The cleaned-up string.
 */
function normalizeSpaces(inputString) {
  if (typeof inputString !== "string") {
    throw new Error("Input must be a string");
  }

  // Trim leading and trailing spaces and replace multiple spaces with a single space
  return inputString.trim().replace(/\s+/g, " ");
}

/**
 * Skips function execution based on a predefined UTC date range.
 * The function will not execute if the current UTC date falls within the specified range.
 */
function skipFunction() {
  const startDateUTC = Date.UTC(2025, 4, 30, 0, 0, 0); // May 30, 2025, 00:00:00 UTC (Months are 0-indexed)
  const endDateUTC = Date.UTC(2025, 5, 27, 23, 59, 59); // June 27, 2025, 23:59:59 UTC

  const nowUTC = Date.now(); // Get current timestamp in milliseconds since epoch UTC

  // Check if the current UTC timestamp is within the predefined range.
  if (nowUTC >= startDateUTC && nowUTC <= endDateUTC) {
    CustomLogger.logInfo("Function execution skipped due to date restriction.", PROJECT_NAME, "skipFunction()");
    return; // Skip execution if within the restricted time range.
  }
}

/**
 * Converts a 2D array into a structured JSON array.
 * Expected structure per row: [location, amount, channel, request, remarks]
 *
 * @param {Array} data - A 2D array of data.
 * @return {Array} - Array of JSON objects.
 */
function convertArrayToJson(data) {
  return data.map((row) => ({
    location: row[0] || "",
    amount: row[1] || 0,
    channel: row[2] || "",
    request: row[3] || "",
    remarks: row[4] || "",
  }));
}
/**
 * Helper Functions for Collection Logic
 */

/**
 * Checks if a machine should be excluded from collection
 * @param {string} machineName - The name of the machine
 * @returns {boolean} - True if machine should be excluded
 */
function isExcludedStore(machineName) {
  const excludedStores = ["SMART LIMKETKAI CDO 2"];
  return excludedStores.includes(machineName);
}

/**
 * Checks if collection should be skipped for BPI banks on weekends
 * @param {string} srvBank - The bank service name
 * @param {Date} date - The date to check
 * @returns {boolean} - True if should skip weekend collection
 */
function skipWeekendCollections(srvBank, date) {
  const dpuPartners = ["BPI", "BPI Internal"];
  const weekendDays = [dayIndex["Sat."], dayIndex["Sun."]];

  return dpuPartners.includes(srvBank) && weekendDays.includes(date.getDay());
}

/**
 * Checks if the last request contains special collection conditions
 * @param {string} lastRequest - The last request string
 * @param {string} tomorrowDateString - The formatted date string
 * @returns {boolean} - True if special conditions are met
 */
function hasSpecialCollectionConditions(lastRequest, tomorrowDateString) {
  if (!lastRequest) return false;
  const lowerDateString = tomorrowDateString.toLowerCase();

  const specialConditions = [
    `for replacement of cassette on ${lowerDateString}`,
    `for collection on ${lowerDateString}`,
    `resume collection on ${lowerDateString}`,
    `for revisit on ${lowerDateString}`,
  ];

  return specialConditions.some((condition) => lastRequest.toLowerCase().includes(condition));
}

/**
 * Configuration object for schedule-based stores
 */
const SCHEDULE_CONFIG = {
  stores: ["PLDT ROBINSONS DUMAGUETE", "SMART SM BACOLOD 1", "SMART SM BACOLOD 2", "SMART SM BACOLOD 3", "PLDT ILIGAN", "SMART GAISANO MALL OZAMIZ"],
  schedules: {
    "PLDT ROBINSONS DUMAGUETE": [dayIndex["M."], dayIndex["W."], dayIndex["Sat."]],
    "SMART SM BACOLOD 1": [dayIndex["T."], dayIndex["Sat."]],
    "SMART SM BACOLOD 2": [dayIndex["T."], dayIndex["Sat."]],
    "SMART SM BACOLOD 3": [dayIndex["T."], dayIndex["Sat."]],
    "PLDT ILIGAN": [dayIndex["M."], dayIndex["W."], dayIndex["F."]],
    "SMART GAISANO MALL OZAMIZ": [dayIndex["F."]],
  },
};

/**
 * Checks if a schedule-based store should be collected on the given day
 * @param {string} machineName - The name of the machine
 * @param {number} dayOfWeek - The day of week (0-6)
 * @returns {boolean} - True if collection should occur
 */
function shouldCollectScheduleBasedStore(machineName, dayOfWeek) {
  const schedule = CONFIG.SCHEDULE_BASED.schedules[machineName];
  return schedule ? schedule.includes(dayOfWeek) : false;
}

/**
 * Checks if the amount meets threshold requirements for regular collection
 * @param {number} amountValue - The amount to check
 * @param {string} collectionDay - The collection day
 * @returns {boolean} - True if threshold is met
 */
function meetsAmountThreshold(amountValue, collectionDay, srvBank) {
  // START SPECIAL CONDITIONS (requested by Sherwin Sept 25-26, 2025 to apply BPI dpu conditions of 150k threshold)
  // Alternative: Hardcoded specific dates (September 25-26, 2025)
  const currentDate = new Date();
  const targetDate1 = new Date("2025-09-25");
  const targetDate2 = new Date("2025-09-26");

  // Set all dates to start of day for accurate comparison
  currentDate.setHours(0, 0, 0, 0);
  targetDate1.setHours(0, 0, 0, 0);
  targetDate2.setHours(0, 0, 0, 0);

  const isValidDate = currentDate.getTime() === targetDate1.getTime() || currentDate.getTime() === targetDate2.getTime();

  // Only apply BPI condition on specified dates
  if (srvBank.includes("BPI") && amountValue >= 150000 && isValidDate) {
    return true;
  }
  // END SPECIAL CONDITIONS

  return amountValue >= amountThresholds[collectionDay];
}

/**
 * Checks if current date falls within payday ranges and meets amount requirement
 * @param {Date} todayDate - Today's date
 * @param {number} amountValue - The amount to check
 * @returns {boolean} - True if payday conditions are met
 */
function isPaydayCollection(todayDate, amountValue) {
  const currentDate = todayDate.getDate();
  return paydayRanges.some((range) => currentDate >= range.start && currentDate <= range.end && amountValue >= paydayAmount);
}

/**
 * Checks if current date falls within due date cutoffs and meets amount requirement
 * @param {Date} todayDate - Today's date
 * @param {number} amountValue - The amount to check
 * @returns {boolean} - True if due date conditions are met
 */
function isDueDateCollection(todayDate, amountValue) {
  const currentDate = todayDate.getDate();
  return dueDateCutoffs.some((range) => currentDate >= range.start && currentDate <= range.end && amountValue >= dueDateCutoffsAmount);
}

/**
 * Helper function to create standardized exclusion reasons
 * Useful if exclusion logic becomes more complex or needs to be shared
 * @param {string} todayDay - Current day name
 * @return {string[]} Array of exclusion reason strings
 */
function getExclusionReasons(todayDay) {
  const baseReasons = [
    "waiting for collection items",
    "waiting for bank updates",
    "did not reset",
    "cassette removed",
    "manually collected",
    "for repair",
    "already collected",
    "store is closed",
    "store is not using the machine",
    "for relocation",
  ];

  return todayDay ? [...baseReasons, todayDay.toLowerCase().trim()] : baseReasons;
}

function addMachineToAdvanceNotice(machineName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Advance Notice`;

  // Check if the sheet already exists
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    sheet.appendRow([machineName]);
  }
}

function forExclusionPartOfAdvanceNotice() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Advance Notice`;

  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return false;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null; // No data rows

  const machineNames = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(function (item) {
      return item != null && item !== "";
    });
  return machineNames;
}

function saveEligibleCollectionsToBQ(eligibleCollections) {

}const CONFIG = {
  APP: {
    NAME: 'Paybox - DPU Monitoring 2.1',
    VERSION: '1.0.0',

  },

  ETAP: {
    SERVICE_BANK: 'eTap',
    ENVIRONMENT: 'testing',
    EMAIL_RECIPIENTS: {
      production: {
        to: "christian@etapinc.com, jayr@etapinc.com, rodison.carapatan@etapinc.com, arnel@etapinc.com, ian.estrella@etapinc.com, roldan@etapinc.com, dante@etapinc.com, ramun.hamtig@etapinc.com, jellymae.osorio@etapinc.com, rojane@etapinc.com, A.jloro@etapinc.com, richard@etapinc.com, johnrandy.divina@etapinc.com",
        cc: "reinier@etapinc.com, miguel@etapinc.com, alvie@etapinc.com, rojane@etapinc.com, laila@etapinc.com, johnmarco@etapinc.com, ghie@etapinc.com, etap-recon@etapinc.com, Erwin Alcantara <egalcantara@multisyscorp.com>, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com",
        bcc: "mvolbara@pldt.com.ph, RBEspayos@smart.com.ph"
      },
      testing: {
        to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        bcc: ""
      }
    },
    SPECIAL_MACHINES: new Set(['PLDT BANTAY', 'SMART VIGAN'])
  },

  SCHEDULE_BASED: {
    stores: ["PLDT ROBINSONS DUMAGUETE", "SMART SM BACOLOD 1", "SMART SM BACOLOD 2", "SMART SM BACOLOD 3", "PLDT ILIGAN", "SMART GAISANO MALL OZAMIZ"],
    schedules: {
      "PLDT ROBINSONS DUMAGUETE": [dayIndex["M."], dayIndex["W."], dayIndex["Sat."]],
      "SMART SM BACOLOD 1": [dayIndex["T."], dayIndex["Sat."]],
      "SMART SM BACOLOD 2": [dayIndex["T."], dayIndex["Sat."]],
      "SMART SM BACOLOD 3": [dayIndex["T."], dayIndex["Sat."]],
      "PLDT ILIGAN": [dayIndex["M."], dayIndex["W."], dayIndex["F."]],
      "SMART GAISANO MALL OZAMIZ": [dayIndex["F."]],
    }
  },

}

Object.freeze(CONFIG);function createTimeDrivenTriggers() {
  // Deletes any existing triggers for the function to avoid duplicates
  deleteExistingTriggers("getCollectedStores");
  deleteExistingTriggers("processMonitoringHourly");
  deleteExistingTriggers("exportSheetAndSendEmail");
  deleteExistingTriggers("eTapCollectionsLogic");
  deleteExistingTriggers("bpiBrinkCollectionsLogic");
  deleteExistingTriggers("bpiInternalCollectionsLogic");
  deleteExistingTriggers("apeirosCollectionsLogic");
  deleteExistingTriggers("resetCollectionRequests");

  deleteExistingTriggers("sendAdvancedNotice");

  // Create new time-driven triggers for the specified times
  ScriptApp.newTrigger("getCollectedStores").timeBased().atHour(8).everyDays(1).create();

  ScriptApp.newTrigger("processMonitoringHourly").timeBased().everyDays(1).atHour(11).nearMinute(15).create();

  ScriptApp.newTrigger("processMonitoringHourly").timeBased().everyDays(1).atHour(15).nearMinute(15).create();

  ScriptApp.newTrigger("processMonitoringHourly").timeBased().everyDays(1).atHour(17).nearMinute(15).create();

  ScriptApp.newTrigger("exportSheetAndSendEmail").timeBased().everyDays(1).atHour(18).create();

  ScriptApp.newTrigger("bpiBrinkCollectionsLogic").timeBased().everyDays(1).atHour(15).nearMinute(30).create();

  ScriptApp.newTrigger("bpiInternalCollectionsLogic").timeBased().everyDays(1).atHour(15).nearMinute(30).create();

  ScriptApp.newTrigger("eTapCollectionsLogic").timeBased().everyDays(1).atHour(15).nearMinute(30).create();

  ScriptApp.newTrigger("sendAdvancedNotice").timeBased().everyDays(1).atHour(17).create();

  ScriptApp.newTrigger("apeirosCollectionsLogic").timeBased().everyDays(1).atHour(17).create();

  ScriptApp.newTrigger("resetCollectionRequests").timeBased().everyDays(1).atHour(17).nearMinute(45).create();
}

function deleteExistingTriggers(functionName) {
  // Gets all triggers for the current project
  var triggers = ScriptApp.getProjectTriggers();

  // Loops through all triggers and deletes those that call the specified function
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
function exportSheetAndSendEmail() {
  CustomLogger.logInfo("Running sending DPU monitoring...", PROJECT_NAME, 'exportSheetAndSendEmail()');

  var retries = 3;
  var sheetName = "For Sending";
  var subject = "[AUTO] Paybox DPU Monitoring";

  var emailTo = "RBEspayos@smart.com.ph, RACagbay@smart.com.ph, mvolbara@pldt.com.ph, avdeleon@pldt.com.ph, cvcabanilla@multisyscorp.com ";
  var emailCc = "AABenter@smart.com.ph, rtevangelista@pldt.com.ph, dmblanco@pldt.com.ph, nplimpiada@pldt.com.ph, kbmila@multisyscorp.com, npsandiego@pldt.com.ph, egalcantara@multisyscorp.com, support@paybox.ph";

  // var emailTo = "egalcantara@multisyscorp.com";
  // var emailCc = "egalcantara@multisyscorp.com";

  var day = new Date().getDay();
  if (day === 0 || day === 6) { // 0 = Sunday, 6 = Saturday
    CustomLogger.logInfo("Today is a weekend. Exiting script.", PROJECT_NAME, 'exportSheetAndSendEmail()');
    return; // Exit the function if today is a weekend
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    CustomLogger.logError(`Sheet "${sheetName}" not found.`, PROJECT_NAME, 'exportSheetAndSendEmail'); // Log error message to the log
    return;
  }

  sheet.showSheet();

  var token = ScriptApp.getOAuthToken();
  var fileName = "Paybox DPU Monitoring_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy") + ".pdf";
  var lastRow = sheet.getLastRow(); // Automatically get the last row

  // Prepare the export URL
  var url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?exportFormat=pdf&format=pdf` +
    `&size=A4&portrait=false&fitw=true&top_margin=0.1&bottom_margin=0.1` +
    `&left_margin=0.1&right_margin=0.1&sheetnames=false&printtitle=false` +
    `&pagenumbers=false&gridlines=false&fzr=false&range=A1:${sheet.getRange(lastRow, sheet.getLastColumn()).getA1Notation()}` +
    `&gid=${sheet.getSheetId()}`;

  while (retries > 0) {
    try {
      var response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      // Check if the response is successful
      if (response.getResponseCode() !== 200) {
        throw new Error(`Failed to fetch PDF. Response code: ${response.getResponseCode()}`);
      }

      var pdfBlob = response.getBlob().setName(fileName);
      var body = `Hi All,\n\nGood day! Please see the attached PDF file of Kiosk DPU Monitoring as of ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy")}.\n\nThank you,\n\n*** This is an auto-generated email, please do not reply to this email. ****`;

      // Send email with the PDF attachment
      GmailApp.sendEmail(emailTo, subject, body, {
        cc: emailCc,
        attachments: [pdfBlob],
      });

      sheet.hideSheet();

      CustomLogger.logInfo("Email sent successfully.", PROJECT_NAME, 'exportSheetAndSendEmail()');
      return; // Exit loop on success

    } catch (e) {
      CustomLogger.logError("Attempt " + (4 - retries) + " failed: " + e.message, 'exportSheetAndSendEmail', PROJECT_NAME, exportSheetAndSendEmail);
    }

    retries--;
    if (retries > 0) {
      CustomLogger.logInfo("Retrying in 5 seconds...", PROJECT_NAME, 'exportSheetAndSendEmail()');
      Utilities.sleep(5000); // Wait 5 seconds before retrying
    }
  }
  CustomLogger.logError("All attempts to send the email have failed.", PROJECT_NAME, 'exportSheetAndSendEmail()');
}
function exportSheetAndSendEmail() {
  CustomLogger.logInfo("Running sending DPU monitoring...", PROJECT_NAME, 'exportSheetAndSendEmail()');

  var retries = 3;
  var sheetName = "For Sending";
  var subject = "[AUTO] Paybox DPU Monitoring";

  var emailTo = "RBEspayos@smart.com.ph, RACagbay@smart.com.ph, mvolbara@pldt.com.ph, avdeleon@pldt.com.ph, cvcabanilla@multisyscorp.com ";
  var emailCc = "AABenter@smart.com.ph, rtevangelista@pldt.com.ph, nplimpiada@pldt.com.ph, egalcantara@multisyscorp.com, support@paybox.ph";

  // var emailTo = "egalcantara@multisyscorp.com";
  // var emailCc = "egalcantara@multisyscorp.com";

  var day = new Date().getDay();
  if (day === 0 || day === 6) { // 0 = Sunday, 6 = Saturday
    CustomLogger.logInfo("Today is a weekend. Exiting script.", PROJECT_NAME, 'exportSheetAndSendEmail()');
    return; // Exit the function if today is a weekend
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    CustomLogger.logError(`Sheet "${sheetName}" not found.`, PROJECT_NAME, 'exportSheetAndSendEmail'); // Log error message to the log
    return;
  }

  sheet.showSheet();

  var token = ScriptApp.getOAuthToken();
  var fileName = "Paybox DPU Monitoring_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy") + ".pdf";
  var lastRow = sheet.getLastRow(); // Automatically get the last row

  // Prepare the export URL
  var url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?exportFormat=pdf&format=pdf` +
    `&size=A4&portrait=false&fitw=true&top_margin=0.1&bottom_margin=0.1` +
    `&left_margin=0.1&right_margin=0.1&sheetnames=false&printtitle=false` +
    `&pagenumbers=false&gridlines=false&fzr=false&range=A1:${sheet.getRange(lastRow, sheet.getLastColumn()).getA1Notation()}` +
    `&gid=${sheet.getSheetId()}`;

  while (retries > 0) {
    try {
      var response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      // Check if the response is successful
      if (response.getResponseCode() !== 200) {
        throw new Error(`Failed to fetch PDF. Response code: ${response.getResponseCode()}`);
      }

      var pdfBlob = response.getBlob().setName(fileName);
      var body = `Hi All,\n\nGood day! Please see the attached PDF file of Kiosk DPU Monitoring as of ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy")}.\n\nThank you,\n\n*** This is an auto-generated email, please do not reply to this email. ****`;

      // Send email with the PDF attachment
      GmailApp.sendEmail(emailTo, subject, body, {
        cc: emailCc,
        attachments: [pdfBlob],
      });

      sheet.hideSheet();

      CustomLogger.logInfo("Email sent successfully.", PROJECT_NAME, 'exportSheetAndSendEmail()');
      return; // Exit loop on success

    } catch (e) {
      CustomLogger.logError("Attempt " + (4 - retries) + " failed: " + e.message, 'exportSheetAndSendEmail', PROJECT_NAME, exportSheetAndSendEmail);
    }

    retries--;
    if (retries > 0) {
      CustomLogger.logInfo("Retrying in 5 seconds...", PROJECT_NAME, 'exportSheetAndSendEmail()');
      Utilities.sleep(5000); // Wait 5 seconds before retrying
    }
  }
  CustomLogger.logError("All attempts to send the email have failed.", PROJECT_NAME, 'exportSheetAndSendEmail()');
}
function executeQueryAndWait(query) {
  const projectId = "ms-paybox-prod-1";
  const request = {
    query: query,
    useLegacySql: false,
  };

  try {
    const queryResults = BigQuery.Jobs.query(request, projectId);
    const jobId = queryResults.jobReference.jobId;

    // Wait for the query to complete with a timeout of 60 seconds
    let status;
    const maxRetries = 60;
    let retries = 0;
    do {
      Utilities.sleep(1000);
      status = BigQuery.Jobs.get(projectId, jobId);
      retries++;
    } while (status.status.state !== "DONE" && retries < maxRetries);

    if (status.status.state !== "DONE") {
      throw new Error("Query did not complete within the timeout period.");
    }

    const finalResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
    // Check for errors in the query results
    if (finalResults.errors && finalResults.errors.length > 0) {
      // Log the errors for debugging purposes
      Logger.log(JSON.stringify(finalResults.errors));
      throw new Error(
        "BigQuery query returned errors: " + JSON.stringify(finalResults.errors)
      );
    }

    return finalResults;
  } catch (e) {
    Logger.log("Error executing query: " + e);
    throw e; // Re-throw the error to be handled by the calling function
  }
}

/**
 * Retrieves the last monitoring hours for a given date and time.
 *
 * @param {*} trnDate
 * @param {*} trnTime
 * @return {*} An array of objects containing machine name, current amount, machine status, and bill validator health.
 * @throws {Error} If no data is found for the specified date and time.
 */
function getLastMonitoringHours(trnDate, trnTime) {
  try {
    const query = `
SELECT distinct machine_name, current_amount, machine_status, bill_validator_health 
FROM \`ms-paybox-prod-1.pldtsmart.monitoring_hourly\`
WHERE created_at >= '${trnDate} ${trnTime.substring(0, 2)}:00:00' 
AND created_at <= '${trnDate} ${trnTime.substring(0, 2)}:59:59';
`;

    const queryResults = executeQueryAndWait(query);
    const rows = queryResults.rows;

    if (!rows || rows.length === 0) {
      throw new Error("No data found");
    }

    return rows;
  } catch (error) {
    throw new Error(
      ("Failed to get last monitoring hours data: " + error.message) &
      ("\nStack Trace: " + error.stack)
    );
  }
}

/**
 * Retrieves the last hourly data from the monitoring_hourly table.
 *
 * @return {*} The last hourly data in a custom format.
 * @throws {Error} If no data is found or if there is an error during the query execution.
 */
function getLastHourly() {
  try {
    const query = `
    SELECT created_at 
    FROM \`ms-paybox-prod-1.pldtsmart.monitoring_hourly\` 
    ORDER BY created_at DESC
    LIMIT 1`;

    const queryResults = executeQueryAndWait(query);

    // Process the results
    const rows = queryResults.jobComplete ? queryResults.rows : [];

    if (!rows || rows.length === 0) {
      throw new Error("No data found");
    }

    const timestamp = rows[0].f[0].v; // Retrieve the timestamp value
    return convertToCustomFormat(timestamp);
  } catch (error) {
    throw new Error(
      ("Failed to get last hourly data: " + error.message) &
      ("\nStack Trace: " + error.stack)
    );
  }
}

function getStoreList() {
  try {
    const query = `
SELECT machine_name
FROM \`ms-paybox-prod-1.pldtsmart.machines\`
WHERE status = TRUE AND NOT machine_name IN ('Multisys Paybox Live','Test Kiosk - Paymaya VAPT','Kiosk Machine - Test','PLDT LAOAG 2')
  `;

    Logger.log(query);

    const queryResults = executeQueryAndWait(query);
    const rows = queryResults.rows;

    if (!rows || rows.length === 0) {
      throw new Error("No data found");
    }

    return rows;
  } catch (error) {
    throw new Error(
      ("Failed to retrieve store list:" + error.message) &
      ("\nStack Trace: " + error.stack)
    );
  }
}

function breakdownDateTime(datetimeStr) {
  if (datetimeStr.length !== 14) {
    throw new Error(
      "Invalid datetime string format. Expected format: YYYYMMDDHHMM"
    );
  }

  var year = datetimeStr.substring(0, 4);
  var month = datetimeStr.substring(4, 6);
  var day = datetimeStr.substring(6, 8);
  var hour = datetimeStr.substring(8, 10);
  var minute = datetimeStr.substring(10, 12);

  return {
    year: parseInt(year, 10),
    month: parseInt(month, 10),
    day: parseInt(day, 10),
    hour: parseInt(hour, 10),
    minute: parseInt(minute, 10),
  };
}

function extractNumbersFromText(text) {
  // Define a regular expression to match sequences of digits
  var regex = /\d+/g;

  // Find all matches in the text
  var matches = text.match(regex);

  // Convert matches to numbers
  var numbers = matches.map(function (match) {
    return parseInt(match, 10);
  });

  return numbers;
}

function saveLogsToFile(folderId, functionName) {
  // Get the execution log
  const log = Logger.getLog();

  // Check if there is anything in the log
  if (log.length === 0) {
    Logger.log("No logs to save.");
    return;
  }

  // Get the folder by ID
  const folder = DriveApp.getFolderById(folderId);

  // Check if the folder exists
  if (!folder) {
    Logger.log("Folder not found. Please check the folderId.");
    return;
  }

  // Create a timestamp to use in the filename
  const timestamp = new Date().toISOString().replace(/[-:.]/g, "");

  // Define the filename
  const filename = `${timestamp}_${functionName}_ExecutionLog.txt`;

  // Create the file in the specified folder
  const file = folder.createFile(filename, log);

  // Log the file creation
  Logger.log(`Log saved to file: ${file.getName()}`);
}

/**
 * Converts a Unix timestamp to a custom date format.
 *
 * @param {*} unixTimestamp
 * @return {*} A string representing the date in the format YYYYMMDDHHMMSS.
 * @throws {Error} If the input is not a valid Unix timestamp.
 */
function convertToCustomFormat(unixTimestamp) {
  // Convert the Unix timestamp (in seconds) to milliseconds
  const date = new Date(unixTimestamp * 1000);

  const year = date.getUTCFullYear().toString();
  const month = (date.getUTCMonth() + 1).toString().padStart(2, "0"); // Months are zero-based
  const day = date.getUTCDate().toString().padStart(2, "0");
  const hours = date.getUTCHours().toString().padStart(2, "0");
  const minutes = date.getUTCMinutes().toString().padStart(2, "0");
  const seconds = date.getUTCSeconds().toString().padStart(2, "0");

  return `${year}${month}${day}${hours}${minutes}${seconds}`;
}

function parseTime(timeStr) {
  // Create a Date object from the timeStr
  var date = new Date(timeStr);

  // Extract hours and minutes from the Date object
  var hours = date.getHours();
  var minutes = date.getMinutes();

  return { hours: hours, minutes: minutes };
}

/**
 * Formats a date object into a string in the format YYYYMMDDHHMMSS.
 *
 * @param {*} date
 * @return {*} A string representing the date in the format YYYYMMDDHHMMSS.
 * @throws {Error} If the input is not a valid Date object.
 */
function formatDateTime(date) {
  var year = date.getFullYear().toString();
  var month = (date.getMonth() + 1).toString().padStart(2, "0"); // Months are zero-based
  var day = date.getDate().toString().padStart(2, "0");
  var hours = date.getHours().toString().padStart(2, "0");
  var minutes = date.getMinutes().toString().padStart(2, "0");

  return `${year}${month}${day}${hours}${minutes}00`;
}

function getFirstFileInFolder(folderId) {
  // Get the folder by ID
  var folder = DriveApp.getFolderById(folderId);

  // Get the files iterator in the folder
  var files = folder.getFiles();

  // Check if the folder contains any files
  if (files.hasNext()) {
    // Get the first file
    var firstFile = files.next();
    return firstFile;
  } else {
    return null; // No files in the folder
  }
}

function extractDateFromFilename(filename) {
  // Use a regular expression to extract digits from the filename
  var match = filename.match(/\d{12}/);

  if (match) {
    var dateString = match[0]; // This should be '202408081708'

    // Break down the string into parts
    var year = dateString.substring(0, 4); // '2024'
    var month = dateString.substring(4, 6); // '08'
    var day = dateString.substring(6, 8); // '08'
    var hour = dateString.substring(8, 10); // '17'
    var minute = dateString.substring(10, 12); // '08'

    // Return the formatted date string
    return {
      year: parseInt(year, 10),
      month: parseInt(month, 10),
      day: parseInt(day, 10),
      hour: parseInt(hour, 10),
      minute: parseInt(minute, 10),
    };
  } else {
    return null; // Return null if the pattern is not found
  }
}

function getKioskPercentage() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Kiosk %");

  //get the column index of column with the string "Remarks"
  const remarksColumnIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat().findIndex(col => col === "Remarks") + 1;

  const lastRow = sheet.getLastRow();
  const machineNames = sheet.getRange(2, 1, lastRow - 1).getValues();
  const partnerAddress = sheet.getRange(2, 2, lastRow - 1).getValues(); // Column B - Collection Team
  const percentValues = sheet.getRange(2, 16, lastRow - 1).getValues(); // Column P - Current Percentage
  const amountValues = sheet.getRange(2, 17, lastRow - 1).getValues(); // Column Q - Current Cash Amount
  const lastRequests = sheet.getRange(2, remarksColumnIndex, lastRow - 1).getValues(); // Column R - Remarks
  const collectionSchedules = sheet.getRange(2, 23, lastRow - 1).getValues(); // Column W - Frequency
  const collectionPartners = sheet.getRange(2, 24, lastRow - 1).getValues(); // Column X - Servicing Bank
  const businessDays = sheet.getRange(2, 25, lastRow - 1).getValues(); // Column Y - Business Days

  return [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, partnerAddress, lastRequests, businessDays,];
}

function getKioskPercentage(minAmount = 250000) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Kiosk %");

  //get the column index of column with the string "Remarks"
  const remarksColumnIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat().findIndex(col => col === "Remarks") + 1;

  const lastRow = sheet.getLastRow();
  const machineNames = sheet.getRange(2, 1, lastRow - 1).getValues();
  const partnerAddress = sheet.getRange(2, 2, lastRow - 1).getValues(); // Column B - Collection Team
  const percentValues = sheet.getRange(2, 16, lastRow - 1).getValues(); // Column P - Current Percentage
  const amountValues = sheet.getRange(2, 17, lastRow - 1).getValues(); // Column Q - Current Cash Amount
  const lastRequests = sheet.getRange(2, remarksColumnIndex, lastRow - 1).getValues(); // Column R - Remarks
  const collectionSchedules = sheet.getRange(2, 23, lastRow - 1).getValues(); // Column W - Frequency
  const collectionPartners = sheet.getRange(2, 24, lastRow - 1).getValues(); // Column X - Servicing Bank
  const businessDays = sheet.getRange(2, 25, lastRow - 1).getValues(); // Column Y - Business Days

  // Filter arrays based on amount threshold
  const filteredData = [];
  for (let i = 0; i < amountValues.length; i++) {
    if (amountValues[i][0] >= minAmount) {
      filteredData.push({
        machineName: machineNames[i],
        percentValue: percentValues[i],
        amountValue: amountValues[i],
        collectionPartner: collectionPartners[i],
        collectionSchedule: collectionSchedules[i],
        partnerAddr: partnerAddress[i],
        lastRequest: lastRequests[i],
        businessDay: businessDays[i]
      });
    }
  }

  // Separate filtered data back into individual arrays
  const filtered = {
    machineNames: filteredData.map(item => item.machineName),
    percentValues: filteredData.map(item => item.percentValue),
    amountValues: filteredData.map(item => item.amountValue),
    collectionPartners: filteredData.map(item => item.collectionPartner),
    collectionSchedules: filteredData.map(item => item.collectionSchedule),
    partnerAddress: filteredData.map(item => item.partnerAddr),
    lastRequests: filteredData.map(item => item.lastRequest),
    businessDays: filteredData.map(item => item.businessDay)
  };

  return [
    filtered.machineNames,
    filtered.percentValues,
    filtered.amountValues,
    filtered.collectionPartners,
    filtered.collectionSchedules,
    filtered.partnerAddress,
    filtered.lastRequests,
    filtered.businessDays
  ];
}

function getMachineDataByPartner(srvBank = null) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Kiosk %");

  const lastRow = sheet.getLastRow();
  const machineNames = sheet.getRange(2, 1, lastRow - 1).getValues(); // Column A
  const partnerAddress = sheet.getRange(2, 2, lastRow - 1).getValues(); // Column B
  const percentValues = sheet.getRange(2, 16, lastRow - 1).getValues(); // Column P
  const amountValues = sheet.getRange(2, 17, lastRow - 1).getValues(); // Column Q
  const lastRemarks = sheet.getRange(2, 18, lastRow - 1).getValues(); // Column R
  const collectionSchedules = sheet.getRange(2, 23, lastRow - 1).getValues(); // Column W
  const collectionPartners = sheet.getRange(2, 24, lastRow - 1).getValues(); // Column X
  const businessDays = sheet.getRange(2, 25, lastRow - 1).getValues(); // Column Y

  // Apply filter where collectionPartners (Column X) == 'eTap'
  const filteredData = collectionPartners
    .map((partner, index) => ({
      index,
      partner: partner[0],
    }))
    .filter(item => srvBank == null || item.partner === srvBank)
    .map(item => item.index);

  // Map each field according to the filtered indices
  const filteredMachineNames = filteredData.map((i) => machineNames[i]);
  const filteredPercentValues = filteredData.map((i) => percentValues[i]);
  const filteredAmountValues = filteredData.map((i) => amountValues[i]);
  const filteredCollectionPartners = filteredData.map(
    (i) => collectionPartners[i]
  );
  const filteredCollectionSchedules = filteredData.map(
    (i) => collectionSchedules[i]
  );
  const filteredPartnerAddress = filteredData.map((i) => partnerAddress[i]);
  const filteredLastRemarks = filteredData.map((i) => lastRemarks[i]);
  const filteredBusinessDays = filteredData.map((i) => businessDays[i]);

  return [
    filteredMachineNames,
    filteredPercentValues,
    filteredAmountValues,
    filteredCollectionPartners,
    filteredCollectionSchedules,
    filteredPartnerAddress,
    filteredLastRemarks,
    filteredBusinessDays,
  ];
}

function translateDaysToAbbreviation(businessDays) {
  //const businessDays = "Monday - Friday";

  // Mapping full day names to their abbreviations
  const dayMap = {
    Monday: "M",
    Tuesday: "T",
    Wednesday: "W",
    Thursday: "Th",
    Friday: "F",
    Saturday: "Sat",
    Sunday: "Sun",
  };

  // Split the input by the separator (assuming " - ")
  const [startDay, endDay] = businessDays.split(" - ");

  // Get the keys (ordered days of the week) for proper traversal
  const daysOfWeek = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
  ];

  // Find the indices for startDay and endDay in the daysOfWeek array
  const startIndex = daysOfWeek.indexOf(startDay);
  const endIndex = daysOfWeek.indexOf(endDay);

  if (startIndex === -1 || endIndex === -1) {
    throw new Error("Invalid day name in input");
  }

  // Slice the days and map to abbreviations
  const abbreviations = daysOfWeek
    .slice(startIndex, endIndex + 1) // Get the range of days
    .map((day) => dayMap[day]) // Map each day to its abbreviation
    .join("."); // Join them with dots

  return abbreviations + "."; // Add trailing dot
}


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

  for (var i = 9; i < sheets.length; i++) { // Start hiding from the 10th sheet (index 9)
    sheets[i].hideSheet();
  }
  CustomLogger.logInfo("Hide sheets after Kiosk%", PROJECT_NAME, 'hideSheets()');
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

function lookupLastTwoEntries() {
  // Define the sheet names
  const formResponsesSheetName = "Form Responses";
  const kioskSheetName = "Kiosk %";

  // Define the column indices (adjust if necessary)
  const machineNameColumnInKiosk = 1; // Assuming Machine Name is in column A of Kiosk %
  const machineNameColumnInFormResponses = 3; // Assuming Machine Name is in column A of Form Responses
  const dateColumnInFormResponses = 1; // Assuming Date or Timestamp is in column B of Form Responses
  const outputColumnInKiosk = 18; // Define where you want to output the concatenated result in Kiosk % (e.g., column B)
  const dataColumnA = 1; // Column D in Form Responses (4th column)
  const dataColumnD = 4; // Column D in Form Responses (4th column)
  const dataColumnE = 5; // Column E in Form Responses (5th column)

  // Get the sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = ss.getSheetByName(formResponsesSheetName);
  const kioskSheet = ss.getSheetByName(kioskSheetName);

  // Get all machine names from Kiosk %
  const kioskData = kioskSheet.getRange(2, machineNameColumnInKiosk, kioskSheet.getLastRow() - 1, 1).getValues();

  // Get all data from Form Responses
  const formResponsesData = formResponsesSheet.getDataRange().getValues();

  // Loop through each machine in Kiosk % and find last two entries in Form Responses
  kioskData.forEach((row, index) => {
    const machineName = row[0];

    // Filter entries in Form Responses that match the current machine name
    const matchingEntries = formResponsesData
      .filter(entry => entry[machineNameColumnInFormResponses - 1] === machineName)
      .map(entry => {
        // Convert entry[dataColumnA - 1] to a Date object and format it as MM/dd
        const formattedDate = Utilities.formatDate(new Date(entry[dataColumnA - 1]), Session.getScriptTimeZone(), "MM/dd");

        // Concatenate formatted date, columns D and E
        return {
          date: new Date(entry[dateColumnInFormResponses - 1]), // For sorting purposes
          data: formattedDate + " - " + entry[dataColumnD - 1] + " " + entry[dataColumnE - 1]
        };
      });

    // Sort the entries by date in descending order (latest date first)
    matchingEntries.sort((a, b) => b.date - a.date);

    // Get the data for the last two entries after sorting
    const lastTwoEntries = matchingEntries.slice(0, 2).map(entry => entry.data).join(" | ");

    // Write the result in the output column of Kiosk %
    kioskSheet.getRange(index + 2, outputColumnInKiosk).setValue(lastTwoEntries);
  });
}

function lookupLastTwoEntries() {
  // Define the sheet names
  const formResponsesSheetName = "Form Responses";
  const kioskSheetName = "Kiosk %";

  // Define the column indices (adjust if necessary)
  const machineNameColumnInKiosk = 1; // Assuming Machine Name is in column A of Kiosk %
  const machineNameColumnInFormResponses = 3; // Assuming Machine Name is in column A of Form Responses
  const dateColumnInFormResponses = 1; // Assuming Date or Timestamp is in column B of Form Responses
  const outputColumnInKiosk = 18; // Define where you want to output the concatenated result in Kiosk % (e.g., column B)
  const dataColumnA = 1; // Column D in Form Responses (4th column)
  const dataColumnD = 4; // Column D in Form Responses (4th column)
  const dataColumnE = 5; // Column E in Form Responses (5th column)

  // Get the sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = ss.getSheetByName(formResponsesSheetName);
  const kioskSheet = ss.getSheetByName(kioskSheetName);

  // Get all machine names from Kiosk %
  const kioskData = kioskSheet.getRange(2, machineNameColumnInKiosk, kioskSheet.getLastRow() - 1, 1).getValues();

  // Get all data from Form Responses
  const formResponsesData = formResponsesSheet.getDataRange().getValues();

  // Loop through each machine in Kiosk % and find last two entries in Form Responses
  kioskData.forEach((row, index) => {
    const machineName = row[0];

    // Filter entries in Form Responses that match the current machine name
    const matchingEntries = formResponsesData
      .filter(entry => entry[machineNameColumnInFormResponses - 1] === machineName)
      .map(entry => {
        // Convert entry[dataColumnA - 1] to a Date object and format it as MM/dd
        const formattedDate = Utilities.formatDate(new Date(entry[dataColumnA - 1]), Session.getScriptTimeZone(), "MM/dd");

        // Concatenate formatted date, columns D and E
        return {
          date: new Date(entry[dateColumnInFormResponses - 1]), // For sorting purposes
          data: formattedDate + " - " + entry[dataColumnD - 1] + " " + entry[dataColumnE - 1]
        };
      });

    // Sort the entries by date in descending order (latest date first)
    matchingEntries.sort((a, b) => b.date - a.date);

    // Get the data for the last two entries after sorting
    const lastTwoEntries = matchingEntries.slice(0, 2).map(entry => entry.data).join(" | ");

    // Write the result in the output column of Kiosk %
    kioskSheet.getRange(index + 2, outputColumnInKiosk).setValue(lastTwoEntries);
  });
}
/**
 * Main function to process monitoring hourly data
 * Downloads logs, processes CSV files, and uploads data to BigQuery
 */
function processMonitoringHourly() {
  const environment = 'production'; //testing
  let processedCount = 0;
  let errorCount = 0;
  try {
    if (environment === 'production') {
      // Step 1: Download Paybox logs
      const downloadResult = runDownloadPayboxLogsMonitoringHourly();
      CustomLogger.logInfo(`Download completed: ${downloadResult.fileCount} files downloaded`, PROJECT_NAME, 'processMonitoringHourly()');

      // Step 2: Process CSV files
      const folderCollections = DriveApp.getFolderById('1knJiSYo3bO4B_V3BohCi5aBhFyhhXPSW');
      const files = folderCollections.getFilesByType(MimeType.CSV);

      // Process each CSV file
      while (files.hasNext()) {
        try {
          const file = files.next();
          runLoadCsvFromDrive(file.getId());
          processedCount++;
        } catch (error) {
          errorCount++;
          CustomLogger.logError(`Error processing file: ${error.message}`, PROJECT_NAME, 'processMonitoringHourly()');
        }
      }
    } else {
      //provide the fileid
      runLoadCsvFromDrive('1fVEIzPRsuCp15Lgzw3yphOtFEIBruwXG'); 
      processedCount++;
    }
    CustomLogger.logInfo(`Processing completed: ${processedCount} files processed, ${errorCount} errors`, PROJECT_NAME, 'processMonitoringHourly()');

    EmailSender.sendLogs('egalcantara@multisyscorp.com', 'Paybox - DPU Monitoring');
  } catch (error) {
    CustomLogger.logError(`Fatal error in processMonitoringHourly: ${error.message}`, PROJECT_NAME, 'processMonitoringHourly()');
    EmailSender.sendLogs('egalcantara@multisyscorp.com', 'FATAL ERROR IN processMonitoringHourly()');
  }
}

/**
 * Loads and processes a CSV file from Google Drive
 * @param {string} fileId - The ID of the file to process
 */
function runLoadCsvFromDrive(fileId) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Fetch the CSV file content from Google Drive
    const file = DriveApp.getFileById(fileId);
    const csvContent = file.getBlob().getDataAsString();
    const fileName = file.getName();
    CustomLogger.logInfo(`Processing file: ${fileName}`, PROJECT_NAME, 'runLoadCsvFromDrive()');

    // Extract and format sheet name from filename
    const sheetName = extractSheetNameFromFileName(fileName);
    if (!sheetName) {
      CustomLogger.logInfo(`Invalid filename format: ${fileName}`, PROJECT_NAME, 'runLoadCsvFromDrive()');
      return;
    }

    // Check if sheet already exists
    const sheet = spreadSheet.getSheetByName(sheetName);
    if (sheet) {
      CustomLogger.logInfo(`Sheet '${sheetName}' already exists. Skipping.`, PROJECT_NAME, 'runLoadCsvFromDrive()');
      return;
    }

    // Validate and process CSV content
    const csvString = validateLine(csvContent, fileName);
    if (!csvString) {
      CustomLogger.logInfo(`No valid data found in file: ${fileName}`, PROJECT_NAME, 'runLoadCsvFromDrive()');
      return;
    }

    // Parse CSV data
    const csvData = parseCsvData(csvString);

    // Upload to BigQuery
    uploadToBQ(csvData);

    // Refresh kiosk data
    refresh(file);

    CustomLogger.logInfo(`Successfully processed file: ${fileName}`, PROJECT_NAME, 'runLoadCsvFromDrive()');
  } catch (error) {
    CustomLogger.logError(`Error processing file ${fileId}: ${error.message}`, PROJECT_NAME, 'runLoadCsvFromDrive()');
    throw error; // Re-throw to be caught by the caller
  }
}

/**
 * Extracts and formats sheet name from filename
 * @param {string} fileName - The name of the file
 * @return {string|null} - Formatted sheet name or null if invalid
 */
function extractSheetNameFromFileName(fileName) {
  const match = fileName.match(/_(\d{12})/);
  if (!match || !match[1]) {
    return null;
  }

  const filenameRegex = match[1];
  return filenameRegex.replace(/20(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/g, '$2$3 $400$1');
}

/**
 * Parses CSV data into structured format
 * @param {string} csvString - The CSV content as a string
 * @return {Array} - Parsed CSV data
 */
function parseCsvData(csvString) {
  return Utilities.parseCsv(csvString)
    .map(row => [
      row[0],
      row[1],
      row[2],
      parseFloat(row[3]) || 0,
      row[4],
      parseFloat(row[5]) || 0
    ]);
}

/**
 * Validates and processes CSV content
 * @param {string} data - The CSV content
 * @param {string} fileName - The name of the file
 * @return {string} - Processed CSV content
 */
function validateLine(data, fileName) {
  const match = fileName.match(/_(\d{12})/);
  if (!match || !match[1]) {
    CustomLogger.logInfo(`Invalid filename format: ${fileName}`, PROJECT_NAME, 'validateLine()');
    return null;
  }

  const dateStr = match[1];
  const trnDate = parseDateFromString(dateStr);
  if (!trnDate) {
    CustomLogger.logInfo(`Invalid date format in filename: ${fileName}`, PROJECT_NAME, 'validateLine()');
    return null;
  }

  const formattedDate = trnDate.toISOString().split('T')[0];
  const validLines = [];

  // Filter out empty lines
  const lines = data.split("\n").filter(line => line.trim() !== "");

  // Process each line
  lines.forEach(line => {
    if (line.startsWith("Partner Name")) {
      return; // Skip header line
    }

    const processedLine = processLine(line, trnDate);
    if (processedLine) {
      validLines.push(processedLine);
    }
  });

  return validLines.length > 0 ? validLines.join("\n") : null;
}

/**
 * Parses date from string in format YYYYMMDDHHMM
 * @param {string} dateStr - Date string
 * @return {Date|null} - Parsed date or null if invalid
 */
function parseDateFromString(dateStr) {
  try {
    const year = dateStr.substring(0, 4);
    const month = dateStr.substring(4, 6);
    const day = dateStr.substring(6, 8);
    const hours = dateStr.substring(8, 10);
    const mins = dateStr.substring(10, 12);

    const trnDate = new Date(`${year}-${month}-${day} ${hours}:${mins}:00`);
    trnDate.setDate(trnDate.getDate());
    return trnDate;
  } catch (error) {
    CustomLogger.logError(`Error parsing date: ${error.message}`, PROJECT_NAME, 'parseDateFromString()');
    return null;
  }
}

/**
 * Processes a single line of CSV data
 * @param {string} line - The line to process
 * @param {Date} trnDate - The transaction date
 * @return {string|null} - Processed line or null if invalid
 */
function processLine(line, trnDate) {
  const regex = /^(?<partner_name>[^,]+),(?<machine_name>[^,]+),(?<current_cash_amount>[1-9]\d*)?,(?<machine_status>[^,]+),(?<bv_health>0|[1-9]\d*)(?<bv_healthDecimalPlace>\.\d*)?$/gm;

  const matches = line.matchAll(regex);
  for (const match of matches) {
    const bill_validator = match.groups.bv_healthDecimalPlace
      ? parseFloat(match.groups.bv_health + match.groups.bv_healthDecimalPlace)
      : parseFloat(match.groups.bv_health);

    const formattedDateTime = formatDateTime(trnDate);
    const current_cash_amount = isNaN(match[3]) ? 0 : match[3];

    return `${formattedDateTime},${match[1]},${match[2]},${current_cash_amount},${match[4]},${bill_validator}`;
  }

  return null;
}

/**
 * Formats a date object into a string in the format YYYY-MM-DD HH:MM:SS
 *
 * @param {*} date
 * @return {*} A string representing the date in the format YYYY-MM-DD HH:MM:SS.
 * @throws {Error} If the input is not a valid Date object.
 */
function formatDateTime(date) {
  if (isNaN(date.getMonth())) {
    return '1970-01-01 00:00:00';
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');

  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

/**
 * Downloads Paybox logs for monitoring hourly
 * @return {Object} - Result of the download operation
 */
function runDownloadPayboxLogsMonitoringHourly() {
  try {
    // Create folder structure
    const parentFolderId = GDriveFilesAPI.createFolderStructure('Paybox Temp');
    const monitoringFolderId = GDriveFilesAPI.getOrCreateFolder('+ Monitoring Hourly', parentFolderId);

    // Clear existing files
    clearMonitoringFolder(monitoringFolderId);

    // Set up email filter
    const now = new Date();
    const dateTo = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const emailFilter = createEmailFilter(dateTo);
    CustomLogger.logInfo(`Email filter: ${emailFilter}`, PROJECT_NAME, 'runDownloadPayboxLogsMonitoringHourly()');

    // Search for emails
    const threads = GmailApp.search(emailFilter);
    Utilities.sleep(1000);

    // Download attachments
    const fileCount = downloadAttachments(threads, monitoringFolderId, parentFolderId);

    // Manage CSV files
    if (fileCount > 0) {
      manageCsvFiles();
    } else {
      CustomLogger.logInfo('No attachments downloaded.', PROJECT_NAME, 'runDownloadPayboxLogsMonitoringHourly()');
    }

    return { success: true, fileCount };
  } catch (error) {
    CustomLogger.logError(`Error downloading Paybox logs: ${error.message}`, PROJECT_NAME, 'runDownloadPayboxLogsMonitoringHourly()');
    return { success: false, error: error.message, fileCount: 0 };
  }
}

/**
 * Clears all files in a folder
 * @param {string} folderId - The ID of the folder to clear
 */
function clearMonitoringFolder(folderId) {
  GDriveFilesAPI.deleteFilesInFolderById(folderId);

  CustomLogger.logInfo('Cleared files in the monitoring folder.', PROJECT_NAME, 'clearMonitoringFolder()');
}

/**
 * Creates an email filter for Gmail search
 * @param {string} dateTo - The date to filter to
 * @return {string} - The email filter
 */
function createEmailFilter(dateTo) {
  return `from:(no-reply@multisyscorp.io) subject:("PAYBOX Machine Unit Monitoring") has:attachment after:${dateTo}`;
}

/**
 * Downloads attachments from email threads
 * @param {Array} threads - The email threads
 * @param {string} monitoringFolderId - The ID of the monitoring folder
 * @param {string} parentFolderId - The ID of the parent folder
 * @return {number} - The number of files downloaded
 */
function downloadAttachments(threads, monitoringFolderId, parentFolderId) {
  let fileCount = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      message.getAttachments().forEach(attachment => {
        if (attachment.getName().startsWith('monitoring_')) {
          const attachmentBlob = attachment.copyBlob();
          const folderId = GDriveFilesAPI.getOrCreateFolder('+ Monitoring Hourly', parentFolderId);
          DriveApp.getFolderById(folderId).createFile(attachmentBlob);
          fileCount++;
          CustomLogger.logInfo(`Downloading ${attachment.getName()}`, PROJECT_NAME, 'downloadAttachments()');
        }
      });
    });
  });

  return fileCount;
}

/**
 * Manages CSV files in the collections folder
 */
function manageCsvFiles() {
  const folderCollections = DriveApp.getFolderById('1knJiSYo3bO4B_V3BohCi5aBhFyhhXPSW');
  const files = getCsvFiles(folderCollections);

  if (files.length === 0) {
    CustomLogger.logInfo('No CSV files found in the folder.', PROJECT_NAME, 'manageCsvFiles()');
    return;
  }

  deleteAllButLatestFile(files);
}

/**
 * Gets all CSV files from a folder
 * @param {Folder} folder - The folder to get files from
 * @return {Array} - The CSV files
 */
function getCsvFiles(folder) {
  const files = [];
  const filesIterator = folder.getFilesByType(MimeType.CSV);

  while (filesIterator.hasNext()) {
    files.push(filesIterator.next());
  }

  files.sort((a, b) => a.getName().localeCompare(b.getName()));
  return files;
}

/**
 * Deletes all but the latest file in an array of files
 * @param {Array} files - The files to process
 */
function deleteAllButLatestFile(files) {
  for (let i = 0; i < files.length - 1; i++) {
    files[i].setTrashed(true);
  }
  // CustomLogger.logInfo(`Deleted ${files.length - 1} older CSV files, kept the latest one.`, PROJECT_NAME, 'deleteAllButLatestFile');
  CustomLogger.logInfo(`Deleted ${files.length - 1} older CSV files, kept the latest one.`, PROJECT_NAME, 'deleteAllButLatestFile()');
}

function uploadToBQ(dataArray, maxRetries = 3, retryDelay = 2000) {
  // BigQuery configuration
  const config = {
    projectId: 'ms-paybox-prod-1',
    datasetId: 'pldtsmart',
    tableId: 'monitoring_hourly',
    schema: [
      { name: 'created_at', type: 'TIMESTAMP' },
      { name: 'partner_name', type: 'STRING' },
      { name: 'machine_name', type: 'STRING' },
      { name: 'current_amount', type: 'FLOAT' },
      { name: 'machine_status', type: 'STRING' },
      { name: 'bill_validator_health', type: 'FLOAT' }
    ]
  };

  let attempt = 0;
  
  while (attempt <= maxRetries) {
    try {
      CustomLogger.logInfo(`Upload attempt ${attempt + 1}/${maxRetries + 1}`, PROJECT_NAME, 'uploadToBQ');
      
      // Prepare the BigQuery job configuration
      const jobConfig = {
        configuration: {
          load: {
            destinationTable: {
              projectId: config.projectId,
              datasetId: config.datasetId,
              tableId: config.tableId
            },
            schema: { fields: config.schema },
            sourceFormat: 'CSV',
            writeDisposition: 'WRITE_APPEND',
            autodetect: false
          }
        }
      };

      // Insert data into BigQuery with timeout handling
      const job = BigQuery.Jobs.insert(jobConfig, config.projectId, Utilities.newBlob(dataArray.join('\n')));
      const jobId = job.jobReference.jobId;
      
      CustomLogger.logInfo(`BigQuery job started: ${jobId}`, PROJECT_NAME, 'uploadToBQ');

      // Wait for job completion with timeout handling
      const jobResult = waitForJobCompletionWithTimeout(config.projectId, jobId);

      if (jobResult.success) {
        CustomLogger.logInfo(`Job completed successfully: ${jobId}`, PROJECT_NAME, 'uploadToBQ');
        return; // Success, exit the function
      } else {
        throw new Error(`Job failed: ${jobResult.error}`);
      }
      
    } catch (error) {
      attempt++;
      
      // Check if it's a timeout error
      const isTimeoutError = isTimeoutRelatedError(error);
      
      if (isTimeoutError && attempt <= maxRetries) {
        CustomLogger.logWarning(
          `Timeout error on attempt ${attempt}/${maxRetries + 1}: ${error.message}. Retrying in ${retryDelay}ms...`, 
          PROJECT_NAME, 
          'uploadToBQ'
        );
        
        // Wait before retrying with exponential backoff
        Utilities.sleep(retryDelay * Math.pow(2, attempt - 1));
        continue;
        
      } else if (attempt > maxRetries) {
        // Max retries exceeded
        CustomLogger.logError(
          `Max retries (${maxRetries}) exceeded. Final error: ${error.message}`, 
          PROJECT_NAME, 
          'uploadToBQ'
        );
        throw new Error(`Upload failed after ${maxRetries} retries: ${error.message}`);
        
      } else {
        // Non-timeout error, don't retry
        CustomLogger.logError(`Non-retryable error: ${error.message}`, PROJECT_NAME, 'uploadToBQ');
        throw error;
      }
    }
  }
}

function waitForJobCompletionWithTimeout(projectId, jobId, timeoutMs = 300000) { // 5 minutes default
  const startTime = Date.now();
  const pollInterval = 5000; // 5 seconds
  
  try {
    while (Date.now() - startTime < timeoutMs) {
      const job = BigQuery.Jobs.get(projectId, jobId);
      
      if (job.status && job.status.state === 'DONE') {
        if (job.status.errors) {
          return {
            success: false,
            error: job.status.errors.map(e => e.message).join(', ')
          };
        }
        return { success: true };
      }
      
      // Wait before next poll
      Utilities.sleep(pollInterval);
    }
    
    // Timeout reached
    throw new Error(`Job ${jobId} timed out after ${timeoutMs}ms`);
    
  } catch (error) {
    if (error.message.includes('timed out') || error.message.includes('timeout')) {
      throw error; // Re-throw timeout errors
    }
    
    return {
      success: false,
      error: error.message
    };
  }
}

function isTimeoutRelatedError(error) {
  const timeoutKeywords = [
    'timeout',
    'timed out',
    'request timeout',
    'execution timeout',
    'deadline exceeded',
    'time limit exceeded',
    'connection timeout',
    'read timeout',
    'write timeout',
    'service timeout'
  ];
  
  const errorMessage = error.message ? error.message.toLowerCase() : '';
  const errorString = error.toString().toLowerCase();
  
  return timeoutKeywords.some(keyword => 
    errorMessage.includes(keyword) || errorString.includes(keyword)
  );
}

// Alternative helper function for checking specific Google Apps Script timeout errors
function isGASTimeoutError(error) {
  // Google Apps Script specific timeout patterns
  const gasTimeoutPatterns = [
    /exceeded maximum execution time/i,
    /script runtime exceeded/i,
    /execution timeout/i,
    /service invoked too many times/i,
    /quota exceeded/i
  ];
  
  const errorMessage = error.message || error.toString();
  return gasTimeoutPatterns.some(pattern => pattern.test(errorMessage));
}
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

  for (var i = 9; i < sheets.length; i++) { // Start hiding from the 10th sheet (index 9)
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

// ==================== OPTIMIZED VERSION ====================

function refresh(file) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");

  // Remove existing filter
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  if (!file) {
    CustomLogger.logError(`No file found in the + Monitoring Hourly folder.`, PROJECT_NAME, "refresh()");
    EmailSender.sendLogs("egalcantara@multisyscorp.com", "Paybox - DPU Monitoring");
    return;
  }

  const dateParts = extractDateFromFilename(file.getName());
  const now = new Date(dateParts.year, dateParts.month - 1, dateParts.day, dateParts.hour, dateParts.minute);
  const lastRow = sheet.getLastRow();
  const sheetName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMdd HH00yy");

  // Populate newly created sheet
  try {
    populateNewSheet(sheetName, now);
    const dpuSheet = spreadSheet.getSheetByName(sheetName);
    if (dpuSheet) dpuSheet.hideSheet();
  } catch (e) {
    CustomLogger.logError(e.message, PROJECT_NAME, "refresh()");
    EmailSender.sendLogs("egalcantara@multisyscorp.com", "Paybox - DPU Monitoring");
    return;
  }

  // Column management - single operation
  sheet.insertColumnsAfter(16, 1);
  sheet.deleteColumns(3, 1);

  // Setup formulas
  setupFormulas(sheet, lastRow, sheetName);

  // Update headers - batch operation
  updateHeaders(sheet, now);

  // Execute all remaining operations
  getCollectedStores();
  applyConditionalFormatting();
  hideColumnsCtoH();
  setRowsHeightAndAlignment(sheet, lastRow);
  sortLatestPercentage();

  SpreadsheetApp.flush(); // Single flush at the end

  // Cleanup operations
  deleteOldSheet();
  hideSheets();
  sheet.activate();
  GDriveFilesAPI.deleteFileByFileId(file.getId());
}

/**
 * Consolidated formula setup function
 * Reduces redundant code and batches operations
 */
function setupFormulas(sheet, lastRow, sheetName) {
  const storesSheet = `${storeSheetUrl}`;

  // Define all formulas
  const formulas = {
    sparkline: `=SPARKLINE(INDIRECT("D"&ROW()&":P"&ROW()))`,
    machineStatus: `=VLOOKUP(INDIRECT("A"&ROW()),INDIRECT(CONCATENATE("'",TEXT($P$1,"MM"),TEXT($P$1,"DD")," ",TEXT($P$1,"HH"),"00",TEXT($P$1,"YY"),"'!A:D")),3,false)`,
    noMovement: '=IF(INDIRECT("D"&ROW())<>"",IF(COUNTIF(INDIRECT("D"&ROW()&":P"&ROW()),"<>"& INDIRECT("D"&ROW()))=0, "No changes", ""),"")',
    collectedStores: "=IFNA(VLOOKUP(A2,'Collected Yesterday'!B:B,1,false))",
    frequency: `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),7,FALSE)`,
    servicingBank: `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),6,FALSE)`,
    businessDays: `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),3,FALSE)`,
    percentage: `=IFNA(TEXT(QUERY('${sheetName}'!$A:$D,"select D where A='" & TRIM(A2) & "'",0),"0.00%")/100,0)`,
    currentAmount: `=IFNA(QUERY('${sheetName}'!$A:$D,"select B where A='" & TRIM(A2) & "'",0),0)`
  };

  // Batch apply formulas with formatting
  const rangeConfigs = [
    { range: `P2:P${lastRow}`, formula: formulas.percentage, format: '0.00%', align: 'Center' },
    { range: `Q2:Q${lastRow}`, formula: formulas.currentAmount, format: '###,###,##0', clear: true },
    { range: `S2:S${lastRow}`, formula: formulas.machineStatus },
    { range: `T2:T${lastRow}`, formula: formulas.sparkline },
    { range: `U2:U${lastRow}`, formula: formulas.noMovement },
    { range: `V2:V${lastRow}`, formula: formulas.collectedStores },
    { range: `W2:W${lastRow}`, formula: formulas.frequency },
    { range: `X2:X${lastRow}`, formula: formulas.servicingBank },
    { range: `Y2:Y${lastRow}`, formula: formulas.businessDays }
  ];

  // Apply all formulas in batch
  rangeConfigs.forEach(config => {
    const range = sheet.getRange(config.range);
    if (config.clear) range.clear();

    range
      .setFontFamily("Century Gothic")
      .setFontSize(9)
      .setFormula(config.formula)
      .setVerticalAlignment("top");

    if (config.format) range.setNumberFormat(config.format);
    if (config.align) range.setHorizontalAlignment(config.align);
  });

  CustomLogger.logInfo(`Applied all formulas for columns P-Y.`, PROJECT_NAME, "setupFormulas()");
}

/**
 * Batch update all headers at once
 */
function updateHeaders(sheet, now) {
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:00:00");

  const headers = [
    { cell: 'P1', value: formattedDate, format: 'MM/dd HH:00AM/PM' },
    { cell: 'Q1', value: 'Current Cash Amount' },
    { cell: 'R1', value: 'Remarks' },
    { cell: 'S1', value: 'Machine Status as of last update' },
    { cell: 'T1', value: 'Movements' },
    { cell: 'U1', value: 'No Movement for 5 days' },
    { cell: 'V1', value: 'Collected Stores' },
    { cell: 'W1', value: 'Frequency' },
    { cell: 'X1', value: 'Servicing Bank' },
    { cell: 'Y1', value: 'Business Days' },
    { cell: 'Z1', value: 'Jira Tickets No. (Issue Type: Machine Issues)' }
  ];

  headers.forEach(h => {
    const range = sheet.getRange(h.cell).setValue(h.value).setBackground("#d9d9d9");
    if (h.format) range.setNumberFormat(h.format);
  });
}

function populateNewSheet(sheetName = null, now = null) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  // Delete existing sheet if present
  const existingSheet = spreadSheet.getSheetByName(sheetName);
  if (existingSheet) {
    spreadSheet.deleteSheet(existingSheet);
    CustomLogger.logError(`Deleted existing sheet: ${sheetName}`, PROJECT_NAME, "populateNewSheet()");
  }

  // Create new sheet
  let newSheet;
  try {
    newSheet = spreadSheet.insertSheet(sheetName);
    newSheet.showSheet();

    const kioskPercentSheetIndex = spreadSheet.getSheetByName("Kiosk %").getIndex();
    spreadSheet.moveActiveSheet(kioskPercentSheetIndex);
  } catch (e) {
    const errorMsg = `Failed to create new sheet: ${sheetName}. ${e.message}`;
    CustomLogger.logError(errorMsg, PROJECT_NAME, "populateNewSheet()");
    throw new Error(errorMsg);
  }

  if (!newSheet) {
    const errorMsg = `Failed to create new sheet: ${sheetName}`;
    CustomLogger.logError(errorMsg, PROJECT_NAME, "populateNewSheet()");
    throw new Error(errorMsg);
  }

  try {
    const trnDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const hours = String(now.getHours()).padStart(2, "0");
    const lastMonitoringHours = getLastMonitoringHours(trnDate, `${hours}:00:00`);

    if (!lastMonitoringHours?.length) {
      throw new Error("No data to populate the new sheet.");
    }

    // Validate data structure
    if (lastMonitoringHours[0]?.f?.length > 0) {
      // Single batch write operation
      const values = lastMonitoringHours.map(row => row.f.map(cell => cell.v));
      newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    } else {
      throw new Error("Rows array is not in the expected format.");
    }
  } catch (e) {
    const errorMsg = `Error populating sheet: ${e.message}`;
    CustomLogger.logError(errorMsg, PROJECT_NAME, "populateNewSheet()");
    throw new Error(errorMsg);
  }
}

function sortLatestPercentage() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kiosk %");
  sheet.getRange("A2:S").sort({ column: 16, ascending: false });
  CustomLogger.logInfo(`Sorted latest percentage.`, PROJECT_NAME, "sortLatestPercentage()");
}

function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kiosk %");

  if (!sheet) {
    CustomLogger.logError("Sheet 'Kiosk %' not found.", PROJECT_NAME, "applyConditionalFormatting()");
    return;
  }

  const lastRow = sheet.getLastRow();

  // Clear existing rules once
  sheet.setConditionalFormatRules([]);

  // Define all ranges
  const ranges = {
    main: sheet.getRange(`C2:P${lastRow}`),
    noMovement: sheet.getRange(`U2:U${lastRow}`),
    machineStatus: sheet.getRange(`S2:S${lastRow}`),
    collected: sheet.getRange(`A2:A${lastRow}`),
    rcbc: sheet.getRange(`Q2:Q${lastRow}`)
  };

  // Build all rules in array
  const allRules = [
    // Main range rules
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.9)
      .setBackground("#cc0000")
      .setFontColor("white")
      .setRanges([ranges.main])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.8)
      .setBackground("#e69138")
      .setFontColor("white")
      .setRanges([ranges.main])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.7)
      .setBackground("#ffe599")
      .setFontColor("black")
      .setRanges([ranges.main])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.6)
      .setBackground("#fff2cc")
      .setFontColor("black")
      .setRanges([ranges.main])
      .build(),

    // Machine status rule
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OFFLINE")
      .setBackground("white")
      .setFontColor('red')
      .setRanges([ranges.machineStatus])
      .build(),

    // No movement rule
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("No changes")
      .setBackground("#cc0000")
      .setFontColor("white")
      .setRanges([ranges.noMovement])
      .build(),

    // Collected stores rule
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=A2=V2")
      .setBackground("#ff9900")
      .setRanges([ranges.collected])
      .build(),

    // RCBC rule
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(X1="RCBC", Q1>180000)')
      .setBackground("#ff9900")
      .setRanges([ranges.rcbc])
      .build()
  ];

  // Apply all rules at once
  sheet.setConditionalFormatRules(allRules);
  CustomLogger.logInfo(`Applied conditional formatting.`, PROJECT_NAME, "applyConditionalFormatting()");
}

function deleteOldSheet() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadSheet.getSheets();

  //get last sheet index
  const lastSheetIndex = sheets.length - 1;

  const startIndex = lastSheetIndex;

  if (startIndex >= sheets.length) {
    CustomLogger.logError(`Start index ${startIndex} is out of range.`, PROJECT_NAME, "deleteOldSheet()");
    return;
  }

  // Delete from end to avoid index shifting
  for (let i = sheets.length - 1; i >= startIndex; i--) {
    const sheetName = sheets[i].getName();
    spreadSheet.deleteSheet(sheets[i]);
    CustomLogger.logInfo(`Deleted sheet: ${sheetName}`, PROJECT_NAME, "deleteOldSheet()");
  }
}

/**
 * Sets row height and vertical alignment efficiently
 */
function setRowsHeightAndAlignment(sheet, lastRow) {
  // Single operation for row heights
  sheet.setRowHeights(1, lastRow, 30);

  // Batch alignment operations
  sheet.getRange(`A2:Q${lastRow}`).setVerticalAlignment("middle");
  sheet.getRange(`R2:Z${lastRow}`).setVerticalAlignment("top");
}// // Configuration objects

// const LOCATION_PAIRS = {
//   PLDT_PACO: ['PLDT PACO 1', 'PLDT PACO 2'],
//   SMART_SM_NORTH: ['SMART SM NORTH EDSA 1', 'SMART SM NORTH EDSA 2']
// };

// const EXCLUSION_CONDITIONS = [
//   'replacement of cassette',
//   'for collection on',
//   'resume collection on'
// ];

// const SPECIAL_LOCATIONS = {
//   PLDT_LAOAG: 'PLDT LAOAG'
// };

// // Cache for memoized results
// const cache = {
//   kioskData: null,
//   lastFetchTime: null,
//   cacheDuration: 5 * 60 * 1000 // 5 minutes in milliseconds
// };

// // Main function
// function bdoCancellationLogic() {
//   skipFunction();
//   CustomLogger.logInfo("Running bdoCancellationLogic...",PROJECT_NAME, 'bdoCancellationLogic');
//   //If the current date is after May 31, 2025 then skip execution
//   if (new Date() > new Date(2025, 4, 31)) return;

//   const environment = "testing"; // or "testing"
//   const srvBank = 'BDO';
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();

//   const ENV_EMAIL_RECIPIENTS = {
//     production: {
//       to: "CMS-Corbank <cms-corbank@bdo.com.ph>,CHH Laoag Cash Hub <chh.laoag@bdo.com.ph>,BM SM City Bataan <bh.sm-city-bataan@bdo.com.ph>,BM SM City North EDSA C <bh.sm-city-north-edsa-c@bdo.com.ph>,CHBMH Tagaytay Cash Hub <chh.tagaytay@bdo.com.ph>",
//       cc: "CHH CCS Makati Cash Hub <chh.ccs-makati@bdo.com.ph>,CHH Las Piñas Cash Hub <chh.las-pinas@bdo.com.ph>,CHH Legazpi Cash Hub <chh.legazpi@bdo.com.ph>,BM SM Aura Premier <bh.sm-aura-premier@bdo.com.ph>,BM SM City Bacoor <bh.sm-city-bacoor@bdo.com.ph>,BM SM City Davao Annex <bh.sm-city-davao-annex@bdo.com.ph>,BM SM City Sta Rosa <bh.sm-city-sta-rosa@bdo.com.ph>,CHH Taft Cash Hub <chh.taft@bdo.com.ph>,CHH Tuguegarao Cash Hub <chh.tuguegarao@bdo.com.ph>, Erwin Alcantara <egalcantara@multisyscorp.com>,Ronald Florentino <raflorentino@multisyscorp.com>,cvcabanilla@multisyscorp.com",
//       bcc: ""
//     },
//     testing: {
//       to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       bcc: ""
//     }
//   };

//   // Usage
//   const emailRecipients = ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing;


//   //skip execution on weekends
//   if (shouldSkipExecution(todayDate)) return;

//   try {
//     // Get kiosk data with caching
//     const kioskData = getKioskDataWithCache();
//     const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, partnerAddresses, lastRequests, businessDays] = kioskData;

//     // Process cancellations
//     const forCancellation = processMachineCancellations(
//       machineNames,
//       amountValues,
//       collectionPartners,
//       collectionSchedules,
//       lastRequests,
//       businessDays,
//       partnerAddresses,
//       srvBank,
//       tomorrowDate,
//       todayDate,
//       tomorrowDateString
//     );

//     if (forCancellation.length > 0) {
//       processCancellation(forCancellation, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
//     }
//   } catch (error) {
//     handleError(error, 'bdoCancellationLogic');
//   }
// }

// // Helper function to get kiosk data with caching
// function getKioskDataWithCache() {
//   const now = new Date().getTime();

//   // Return cached data if it's still valid
//   if (cache.kioskData && cache.lastFetchTime && (now - cache.lastFetchTime < cache.cacheDuration)) {
//     return cache.kioskData;
//   }

//   // Fetch fresh data
//   const kioskData = getKioskPercentage();

//   // Update cache
//   cache.kioskData = kioskData;
//   cache.lastFetchTime = now;

//   return kioskData;
// }

// // Helper function to process machine cancellations
// function processMachineCancellations(machineNames, amountValues, collectionPartners, collectionSchedules, lastRequests, businessDays, partnerAddresses, srvBank, tomorrowDate, todayDate, tomorrowDateString) {
//   // Pre-filter machines by service bank for better performance
//   var bdoMachines = machineNames.reduce((acc, machineNameArr, i) => {
//     if (collectionPartners[i][0] === srvBank) {
//       if (collectionSchedules[i][0].includes(getTomorrowDayCustomFormat(tomorrowDate))) {
//         CustomLogger.logInfo(`Cancelling DPU for machine ${machineNameArr[0]} for tomorrow.`,PROJECT_NAME, 'bdoCancellationLogic');
//         Logger.log(`Cancelling DPU for machine ${machineNameArr[0]} (${collectionSchedules[i][0]}) - ${getTomorrowDayCustomFormat(tomorrowDate)}`);

//         acc.push({
//           index: i,
//           name: machineNameArr[0],
//           amountValue: amountValues[i][0],
//           collectionSchedule: collectionSchedules[i][0],
//           lastRequest: lastRequests[i][0],
//           partnerAddress: partnerAddresses[i][0]
//         });
//       }
//     }
//     return acc;
//   }, []);

//   // Process cancellations in a single pass
//   let forCancellation = bdoMachines
//     .filter(machine => shouldIncludeForCancellation(
//       machine.name,
//       machine.amountValue,
//       tomorrowDate,
//       machine.collectionSchedule,
//       todayDate,
//       tomorrowDateString,
//       machine.lastRequest
//     ))
//     .map(machine => ({
//       name: machine.name,
//       address: machine.partnerAddress,
//       currentCashAmount: machine.amountValue
//     }));

//   // Apply special validations
//   forCancellation = validatePairedLocations(forCancellation, Object.values(LOCATION_PAIRS));

//   return forCancellation;
// }

// // Helper function to check if a machine should be included for cancellation
// function shouldIncludeForCancellation(machineName, amountValue, tomorrowDate, collectionSchedule, todayDate, tomorrowDateString, lastRequest) {
//   const collectionDay = dayMapping[tomorrowDate.getDay()];

//   if (isTomorrowHoliday(tomorrowDate)) return true;

//   // Check for exclusion conditions in lastRequest
//   if (lastRequest) {
//     const hasExclusionCondition = EXCLUSION_CONDITIONS.some(condition => {
//       const searchText = condition === 'for collection on' || condition === 'resume collection on'
//         ? `${condition} ${tomorrowDateString.toLowerCase()}`
//         : condition;
//       return lastRequest.toLowerCase().includes(searchText);
//     });

//     if (hasExclusionCondition) {
//       return false;
//     }
//   }

//   // Check if machine meets cancellation criteria
//   return collectionSchedule.includes(collectionDay) && amountValue < amountThresholds[collectionDay];
// }

// // Helper function to validate paired locations
// function validatePairedLocations(forCancellation, locationPairs) {
//   // Create a map for faster lookups
//   const cancellationMap = new Map(
//     forCancellation.map(item => [item.name, item])
//   );

//   // Process each location pair
//   locationPairs.forEach(([location1, location2]) => {
//     const loc1 = cancellationMap.get(location1);
//     const loc2 = cancellationMap.get(location2);

//     if (loc1 && !loc2) {
//       cancellationMap.set(location2, {
//         name: location2,
//         address: loc1.address,
//         currentCashAmount: loc1.currentCashAmount
//       });
//     } else if (loc2 && !loc1) {
//       cancellationMap.set(location1, {
//         name: location1,
//         address: loc2.address,
//         currentCashAmount: loc2.currentCashAmount
//       });
//     }
//   });

//   // Convert map back to array
//   return Array.from(cancellationMap.values());
// }

// // Helper function to process cancellations and send emails
// function processCancellation(forCancellation, tomorrowDate, emailTo, emailCc, emailBcc, srvBank) {
//   if (forCancellation.length === 0) return;

//   const isSaturday = tomorrowDate.getDay() === 6;

//   // Sort once for better performance
//   forCancellation.sort((a, b) => a.name.localeCompare(b.name));

//   if (isSaturday) {
//     // Remove PLDT LAOAG for Saturday collections
//     forCancellation = forCancellation.filter(item => item.name !== SPECIAL_LOCATIONS.PLDT_LAOAG);

//     // Schedule Monday collection
//     const collectionDateForMonday = new Date(tomorrowDate);
//     collectionDateForMonday.setDate(collectionDateForMonday.getDate() + 2);
//     sendEmailCancellation(forCancellation, collectionDateForMonday, emailTo, emailCc, emailBcc, srvBank, isSaturday);
//   } else {
//     sendEmailCancellation(forCancellation, tomorrowDate, emailTo, emailCc, emailBcc, srvBank, isSaturday);
//   }
// }

// // Helper function to handle errors
// function handleError(error, functionName) {
//   const errorMessage = `Error in ${functionName}: ${error.message}`;
//   Logger.log(errorMessage);

//   // You could add additional error handling here, such as:
//   // - Sending error notifications
//   // - Logging to a monitoring service
//   // - Retrying the operation

//   // For now, we'll just log the error
//   console.error(errorMessage);
// }function processCancellationRequest() {
  // processCancellationLogic('BPI Internal');
  // processCancellationLogic('Brinks via BPI');
  // processCancellationLogic('eTap');
}

// function processCancellationLogic(srvBank) {
//   // Open the active spreadsheet
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   // Access the "For Collections" worksheet
//   const sheet = ss.getSheetByName("For Collection -" + srvBank);
//   if (!sheet) {
//     console.log("Sheet 'For Collections' not found.");
//     return;
//   }

//   // Get all the data in the sheet
//   const data = sheet.getDataRange().getValues();

//   // Store the results
//   const results = [];
//   var subject = '';

//   // Iterate through rows, starting from row 2 (skip the header)
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const machineName = row[0]; // Column A (index 0)
//     const emailSubject = row[3]; // Column D (index 3)
//     const isCollected = row[5]; // Column F (index 5)

//     subject = emailSubject;

//     const doNotCancel = [
//       'PLDT ROBINSONS DUMAGUETE',
//       'PLDT BANTAY',
//       'SMART VIGAN'
//     ];
//     if (doNotCancel.some(store => machineName.includes(store))) { continue; }

//     // Check if Column F (isCollected) is TRUE
//     if (machineName && isCollected === true) {
//       results.push({ machineName, emailSubject });
//     }
//   }

//   if (results.length > 0) {
//     var machineData = results.map(row => row.machineName);

//     var body = `
//     Hi All,<br><br>
//     Good day! Please <b>CANCEL</b> the collection for the following stores:<br><br>
//     ${machineData.map((machine) => machine).join('<br>')}<br>
//     *** Please acknowledge this email. ****<br><br>${emailSignature}
//   `;

//     replyToExistingThread(results[0].emailSubject, body);
//   }

// }// Configuration constants
const APEIROS_CONFIG = {
  SERVICE_BANK: 'Apeiros',
  ENVIRONMENT: 'testing' // Change to 'production' when ready
};

// Email recipients by environment
const APEIROS_EMAIL_RECIPIENTS = {
  production: {
    to: "mtcsantiago570@gmail.com, mtcsurigao@gmail.com, valdez.ezekiel23@gmail.com",
    cc: "sherwinamerica@yahoo.com, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com",
    bcc: "mvolbara@pldt.com.ph"
  },
  testing: {
    to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
    cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
    bcc: ""
  }
};

function apeirosCollectionsLogic() {
  CustomLogger.logInfo('Running Apeiros Collection Logic...', PROJECT_NAME, 'apeirosCollectionsLogic');

  try {
    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = dateInfo;

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipients(APEIROS_CONFIG.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubject(tomorrowDate, APEIROS_CONFIG.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollections(todayDate, tomorrowDate, todayDateString, tomorrowDateString, subject);

    // Process collections if any are eligible
    if (eligibleCollections.length > 0) {
      processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients);
    } else {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.', PROJECT_NAME, 'apeirosCollectionsLogic');
    }

  } catch (error) {
    CustomLogger.logError(`Error in apeirosCollectionsLogic: ${error}`, PROJECT_NAME, 'apeirosCollectionsLogic');
    throw error;
  }
}

/**
 * Get email recipients based on environment
 */
function getEmailRecipients(environment) {
  return APEIROS_EMAIL_RECIPIENTS[environment] || APEIROS_EMAIL_RECIPIENTS.testing;
}

/**
 * Generate email subject line
 */
function generateEmailSubject(tomorrowDate, serviceBank) {
  const formattedDate = Utilities.formatDate(
    tomorrowDate,
    Session.getScriptTimeZone(),
    "MMMM d, yyyy (EEEE)"
  );
  return `${serviceBank} DPU Request - ${formattedDate}`;
}

/**
 * Get eligible collections based on criteria
 */
function getEligibleCollections(todayDate, tomorrowDate, todayDateString, tomorrowDateString, subject) {
  const srvBank = APEIROS_CONFIG.SERVICE_BANK;

  // Fetch data
  const machineData = getMachineDataByPartner(srvBank);
  const forCollectionData = getForCollections(srvBank);
  const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(srvBank);

  // Destructure machine data
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRemarks, businessDays] = machineData;
  // Filter eligible collections
  const eligibleCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const lastRemark = normalizeSpaces(lastRemarks[i][0]);
    const businessDay = businessDays[i][0];

    // Apply exclusion filters
    if (shouldExcludeMachine(machineName, amountValue, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, srvBank)) {
      return; // Skip this machine
    }

    // Add to eligible collections
    eligibleCollections.push([machineName, amountValue, srvBank, subject]);
    addMachineToAdvanceNotice(machineName);
  });

  return eligibleCollections;
}

/**
 * Determine if machine should be excluded from collection
 */
function shouldExcludeMachine(machineName, amountValue, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, srvBank) {
  // Check 1: Already collected
  if (skipAlreadyCollected(machineName, forCollectionData)) {
    return true;
  }

  // Check 2: Excluded based on remarks
  if (forExclusionBasedOnRemarks(lastRemark, todayDateString, machineName)) {
    return true;
  }

  // Check 3: Requested yesterday
  if (forExclusionRequestYesterday(previouslyRequestedMachines, machineName, srvBank)) {
    return true;
  }

  // Check 4: If not yet scheduled for collection
  if(forExclusionNotYetScheduled(machineName,tomorrowDate,lastRemark)){
    return true;
  }

  // Check 4: Business day schedule
  const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

  if (!shouldIncludeForCollection(
    machineName,
    amountValue,
    translatedBusinessDays,
    tomorrowDate,
    tomorrowDateString,
    todayDate,
    lastRemark,
    srvBank
  )) {
    return true;
  }

  return false;
}

/**
 * Process eligible collections
 */
function processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients) {
  const srvBank = APEIROS_CONFIG.SERVICE_BANK;

  try {
    // Create hidden worksheet with collection data
    createHiddenWorksheetAndAddData(eligibleCollections, srvBank);

    // Process and send collections
    processCollectionsAndSendEmail(eligibleCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);

    CustomLogger.logInfo(
      `Processed ${eligibleCollections.length} eligible collections for ${srvBank}`,
      PROJECT_NAME,
      'processEligibleCollections'
    );

  } catch (error) {
    CustomLogger.logError(
      `Error processing eligible collections: ${error}`,
      PROJECT_NAME,
      'processEligibleCollections'
    );
    throw error;
  }
}
function bpiCollectionsLogic() {
  CustomLogger.logInfo("Running BPI Collections Logic...", PROJECT_NAME, 'bpiCollectionsLogic()');
  const environment = 'production';
  const srvBank = 'BPI';
  const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();

  // Skip weekends
  if (shouldSkipWeekendCollections(srvBank, tomorrowDate)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, PROJECT_NAME, "bpiCollectionsLogic");
    return false;
  }

  // Move collection schedule if tomorrow is holiday
  if (isTomorrowHoliday(tomorrowDate)) {
    tomorrowDate.setDate(tomorrowDate.getDate() + 1);
  }

  const machineData = getMachineDataByPartner(srvBank);
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRemarks, businessDays] = machineData;

  const ENV_EMAIL_RECIPIENTS = {
    production: {
      to: "mjdagasuhan@bpi.com.ph",
      cc: "malodriga@bpi.com.ph, julsales@bpi.com.ph, jcslingan@bpi.com.ph, eagmarayag@bpi.com.ph, jmdcantorna@bpi.com.ph, egcameros@bpi.com.ph, jdduque@bpi.com.ph, rdtayag@bpi.com.ph, rsmendoza1@bpi.com.ph, vjdtorcuator@bpi.com.ph, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com ",
      bcc: "mvolbara@pldt.com.ph,RBEspayos@smart.com.ph "
    },
    testing: {
      to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
      cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
      bcc: ""
    }
  };

  // Usage
  const emailRecipients = ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing;

  let forCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const lastRemark = normalizeSpaces(lastRemarks[i][0]);
    const businessDay = businessDays[i][0];

    // Skip if the last request should be excluded
    if (forExclusionBasedOnRemarks(lastRemark, todayDateString, machineName)) {
      return;
    }

    const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

    if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
      forCollections.push([machineName, amountValue, srvBank, subject]);
    }
  });

  if (environment === "production" && forCollections.length > 0) {
    createHiddenWorksheetAndAddData(forCollections, srvBank);
    processCollections(forCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
  } else {
    CustomLogger.logInfo('No eligible stores for collection tomorrow.', PROJECT_NAME, 'bpiCollectionsLogic()');
  }

}

function shouldIncludeForCollection_BPI(amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest) {
  const collectionDay = dayMapping[tomorrowDate.getDay()];

  if (['replacement of cassette', `for collection on ${tomorrowDateString.toLowerCase()}`, `resume collection on ${tomorrowDateString.toLowerCase()}`]
    .some(condition => lastRequest.toLowerCase().includes(condition))) {
    return true;
  }

  if (translatedBusinessDays.includes(collectionDay) && amountValue >= amountThresholds[collectionDay]) {
    return true;
  }

  return paydayRanges.some(range => todayDate.getDate() >= range.start && todayDate.getDate() <= range.end && amountValue >= paydayAmount) ||
    dueDateCutoffs.some(range => todayDate.getDate() >= range.start && todayDate.getDate() <= range.end && amountValue >= dueDateCutoffsAmount);
}
/**
 * Configuration object for the BPI Brinks collections script.
 * Centralizes settings for easy management and environment switching.
 */
const CONFIG_BPI = {
  ENVIRONMENT: "production", // Switch to "testing" for development
  SRV_BANK: "Brinks via BPI",
};

/**
 * Email recipients configuration, separated by environment.
 */
const EMAIL_RECIPIENTS_BPI = {
  production: {
    to: "mbocampo@bpi.com.ph",
    cc: "julsales@bpi.com.ph, mjdagasuhan@bpi.com.ph, eagmarayag@bpi.com.ph,rtorticio@bpi.com.ph, egcameros@bpi.com.ph, vjdtorcuator@bpi.com.ph, jdduque@bpi.com.ph, rsmendoza1@bpi.com.ph,jmdcantorna@bpi.com.ph, kdrepuyan@bpi.com.ph, dabayaua@bpi.com.ph, rdtayag@bpi.com.ph, vrvarellano@bpi.com.ph,mapcabela@bpi.com.ph, mvpenisa@bpi.com.ph, mbcernal@bpi.com.ph, cmmanalac@bpi.com.ph, mpdcastro@bpi.com.ph,rmdavid@bpi.com.ph, emflores@bpi.com.ph, apmlubaton@bpi.com.ph, smcarvajal@bpi.com.ph, avabarabar@bpi.com.ph,jcmontes@bpi.com.ph, jeobautista@bpi.com.ph, micaneda@bpi.com.ph, rrpacio@bpi.com.ph,mecdycueco@bpi.com.ph, tesruelo@bpi.com.ph, ssibon@bpi.com.ph, christine.sarong@brinks.com, icom2.ph@brinks.com,aillen.waje@brinks.com, rpsantiago@bpi.com.ph, jerome.apora@brinks.com, occ2supervisors.ph@brinks.com, mdtenido@bpi.com.ph, agmaiquez@bpi.com.ph, jsdamaolao@bpi.com.ph, cvcabanilla@multisyscorp.com, raflorentino@multisyscorp.com, egalcantara@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com",
    bcc: "mvolbara@pldt.com.ph, RBEspayos@smart.com.ph",
  },
  testing: {
    to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
    cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
    bcc: "",
  },
};

// =========================================================================================
// MAIN ORCHESTRATOR FUNCTION
// =========================================================================================

/**
 * Main logic for identifying and processing BPI Brinks collections.
 * This function orchestrates fetching data, filtering eligible machines,
 * and triggering the collection process in a structured and error-handled manner.
 */
function bpiBrinkCollectionsLogic() {
  try {
    CustomLogger.logInfo(`Running BPI Brinks Collections Logic in [${CONFIG_BPI.ENVIRONMENT}] mode...`, PROJECT_NAME, "bpiBrinkCollectionsLogic()");

    const { tomorrowDate, tomorrowDateString } = getTodayAndTomorrowDates();
    const emailRecipients = getEmailRecipientsBPI();
    const subject = generateEmailSubjectBPI(tomorrowDate);

    const eligibleCollections = getEligibleCollectionsBPI(subject, tomorrowDate, tomorrowDateString);

    processEligibleCollectionsBPI(eligibleCollections, tomorrowDate, emailRecipients);

  } catch (error) {
    const errorMessage = `An unexpected error occurred in BPI Brinks Logic: ${error.message} \nStack: ${error.stack}`;
    CustomLogger.logError(errorMessage, PROJECT_NAME, "bpiBrinkCollectionsLogic()");
    // Optionally, re-throw the error if you want it to be visible in the Apps Script dashboard
    // throw error;
  }
}

// =========================================================================================
// DATA FILTERING & PROCESSING FUNCTIONS
// =========================================================================================

/**
 * Filters and returns the list of machines eligible for collection.
 * @param {string} subject The email subject to be used for the collection request.
 * @param {Date} tomorrowDate The Date object for tomorrow.
 * @param {string} tomorrowDateString The formatted string for tomorrow's date.
 * @returns {Array<Array<string>>} A 2D array of eligible collection data.
 */
function getEligibleCollectionsBPI(subject, tomorrowDate, tomorrowDateString) {
  const { todayDate, todayDateString } = getTodayAndTomorrowDates();
  const machineData = getMachineDataByPartner(CONFIG_BPI.SRV_BANK);
  const [machineNames, , amountValues, , , , lastRemarks, businessDays] = machineData;

  const forCollectionData = getForCollections(CONFIG_BPI.SRV_BANK);
  const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(CONFIG_BPI.SRV_BANK);
  const eligibleCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machine = {
      name: machineNameArr[0],
      amount: amountValues[i][0],
      lastRemark: normalizeSpaces(lastRemarks[i][0]),
      businessDay: businessDays[i][0],
    };

    if (shouldExcludeMachineBPI(machine, forCollectionData, previouslyRequestedMachines, todayDateString)) {
      return; // Skip this machine
    }

    const translatedBusinessDays = translateDaysToAbbreviation(machine.businessDay.trim());
    if (shouldIncludeForCollection(machine.name, machine.amount, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, machine.lastRemark, CONFIG_BPI.SRV_BANK)) {
      const formattedName = formatMachineNameWithRemarkBPI(machine.name, machine.lastRemark, tomorrowDateString);
      eligibleCollections.push([formattedName, machine.amount, CONFIG_BPI.SRV_BANK, subject, machine.lastRemark]);
    }
  });

  return eligibleCollections;
}

/**
 * Processes the final list of eligible collections.
 * In production, it creates a worksheet and sends notifications.
 * In testing, it logs the data without sending emails.
 * @param {Array<Array<string>>} collections The list of eligible collections.
 * @param {Date} tomorrowDate The Date object for tomorrow, used in processing.
 * @param {object} recipients The email recipients object.
 */
function processEligibleCollectionsBPI(collections, tomorrowDate, recipients) {
  if (collections.length === 0) {
    CustomLogger.logInfo("No eligible stores for BPI Brinks collection tomorrow.", PROJECT_NAME, "processEligibleCollectionsBPI()");
    return;
  }

  if (CONFIG_BPI.ENVIRONMENT === "production") {
    CustomLogger.logInfo(`Processing ${collections.length} collections for production.`, PROJECT_NAME, "processEligibleCollectionsBPI()");
    createHiddenWorksheetAndAddData(collections, CONFIG_BPI.SRV_BANK);
    processCollectionsAndSendEmail(collections, tomorrowDate, recipients.to, recipients.cc, recipients.bcc, CONFIG_BPI.SRV_BANK);
  } else {
    CustomLogger.logInfo(`[TESTING MODE] Found ${collections.length} eligible collections. No emails will be sent.`, PROJECT_NAME, "processEligibleCollectionsBPI()");
    // Optional: Log the collections for verification during testing
    console.log(JSON.stringify(collections, null, 2));
    processCollectionsAndSendEmail(collections, tomorrowDate, recipients.to, recipients.cc, recipients.bcc, CONFIG_BPI.SRV_BANK);
  }
}


// =========================================================================================
// HELPER & UTILITY FUNCTIONS
// =========================================================================================

/**
 * Determines if a machine should be excluded based on various criteria.
 * @param {object} machine An object containing machine details.
 * @param {Array} forCollectionData Data on machines already slated for collection.
 * @param {Array} previouslyRequestedMachines Machines requested yesterday.
 * @param {string} todayDateString Formatted string for today's date.
 * @returns {boolean} True if the machine should be excluded, false otherwise.
 */
function shouldExcludeMachineBPI(machine, forCollectionData, previouslyRequestedMachines, todayDateString) {
  if (skipAlreadyCollected(machine.name, forCollectionData)) {
    return true;
  }
  if (forExclusionBasedOnRemarks(machine.lastRemark, todayDateString, machine.name)) {
    return true;
  }
  if (forExclusionRequestYesterday(previouslyRequestedMachines, machine.name, CONFIG_BPI.SRV_BANK)) {
    return true;
  }
  return false;
}

/**
 * Formats the machine name, adding a revisit remark if applicable.
 * @param {string} machineName The name of the machine.
 * @param {string} lastRemark The last remark associated with the machine.
 * @param {string} tomorrowDateString The formatted string for tomorrow's date.
 * @returns {string} The formatted machine name.
 */
function formatMachineNameWithRemarkBPI(machineName, lastRemark, tomorrowDateString) {
  const revisitRemark = `for revisit on ${tomorrowDateString.toLowerCase()}`;
  if (lastRemark.toLowerCase().includes(revisitRemark)) {
    return `${machineName} (<b>${lastRemark}</b>)`;
  }
  return machineName;
}

/**
 * Retrieves the email recipients based on the current environment setting.
 * @returns {object} An object with 'to', 'cc', and 'bcc' properties.
 */
function getEmailRecipientsBPI() {
  return EMAIL_RECIPIENTS_BPI[CONFIG_BPI.ENVIRONMENT] || EMAIL_RECIPIENTS_BPI.testing;
}

/**
 * Generates the email subject for the collection request.
 * @param {Date} tomorrowDate The Date object for tomorrow.
 * @returns {string} The formatted email subject.
 */
function generateEmailSubjectBPI(tomorrowDate) {
  const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
  return `${CONFIG_BPI.SRV_BANK} DPU Request - ${tomorrowDateFormatted}`;
}


// function bpiBrinkCollectionsLogic() {
//   CustomLogger.logInfo("Running bpiBrinkCollectionsLogic...", PROJECT_NAME, "bpiBrinkCollectionsLogic()");
//   const environment = "production";
//   const srvBank = "Brinks via BPI";
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = getTodayAndTomorrowDates();
//   const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
//   const subject = `${srvBank} DPU Request - ${tomorrowDateFormatted}`;

//   const machineData = getMachineDataByPartner(srvBank);
//   const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRemarks, businessDays] = machineData;

//   const ENV_EMAIL_RECIPIENTS = {
//     production: {
//       to: "mbocampo@bpi.com.ph",
//       cc: "julsales@bpi.com.ph, mjdagasuhan@bpi.com.ph, eagmarayag@bpi.com.ph,rtorticio@bpi.com.ph, egcameros@bpi.com.ph, vjdtorcuator@bpi.com.ph, jdduque@bpi.com.ph, rsmendoza1@bpi.com.ph,jmdcantorna@bpi.com.ph, kdrepuyan@bpi.com.ph, dabayaua@bpi.com.ph, rdtayag@bpi.com.ph, vrvarellano@bpi.com.ph,mapcabela@bpi.com.ph, mvpenisa@bpi.com.ph, mbcernal@bpi.com.ph, cmmanalac@bpi.com.ph, mpdcastro@bpi.com.ph,rmdavid@bpi.com.ph, emflores@bpi.com.ph, apmlubaton@bpi.com.ph, smcarvajal@bpi.com.ph, avabarabar@bpi.com.ph,jcmontes@bpi.com.ph, jeobautista@bpi.com.ph, micaneda@bpi.com.ph, rrpacio@bpi.com.ph,mecdycueco@bpi.com.ph, tesruelo@bpi.com.ph, ssibon@bpi.com.ph, christine.sarong@brinks.com, icom2.ph@brinks.com,aillen.waje@brinks.com, rpsantiago@bpi.com.ph, jerome.apora@brinks.com, occ2supervisors.ph@brinks.com, mdtenido@bpi.com.ph, agmaiquez@bpi.com.ph, jsdamaolao@bpi.com.ph, cvcabanilla@multisyscorp.com, raflorentino@multisyscorp.com, egalcantara@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com ",
//       bcc: "mvolbara@pldt.com.ph,RBEspayos@smart.com.ph ",
//     },
//     testing: {
//       to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       bcc: "",
//     },
//   };

//   // Usage
//   const emailRecipients = ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing;
//   const forCollectionData = getForCollections(srvBank);

//   const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(srvBank);
//   let forCollections = [];

//   machineNames.forEach((machineNameArr, i) => {
//     const machineName = machineNameArr[0];
//     const amountValue = amountValues[i][0];
//     const lastRemark = normalizeSpaces(lastRemarks[i][0]);
//     const businessDay = businessDays[i][0];
//     const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

// // if(machineName!=="PLDT TALISAY") {
// //   return;
// // }

//     // Skip if already collected
//     if(skipAlreadyCollected(machineName, forCollectionData)){
//       return;
//     }

//     // Skip if the last request should be excluded
//     if (forExclusionBasedOnRemarks(lastRemark, todayDateString, machineName)) {
//       return;
//     }

//     // Skip if already requested for collection yesterday
//     if(forExclusionRequestYesterday(previouslyRequestedMachines, machineName, srvBank)) {
//       return;
//     }

//     if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
//       if (lastRemark.toLowerCase().includes(`for revisit on ${tomorrowDateString.toLowerCase()}`)) {
//         forCollections.push([machineName + ` (<b>${lastRemark}</b>)`, amountValue, srvBank, subject, lastRemark]);
//       } else {
//         forCollections.push([machineName, amountValue, srvBank, subject, lastRemark]);
//       }
//     }
//   });

//   if (environment === "production" && forCollections.length > 0) {
//     createHiddenWorksheetAndAddData(forCollections, srvBank);
//     processCollections(forCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
//   } else {
//     CustomLogger.logInfo("No eligible stores for collection tomorrow.", PROJECT_NAME, "bpiBrinkCollectionsLogic()");
//   }
// }
// Configuration constants
const BPI_INTERNAL_CONFIG = {
  SERVICE_BANK: 'BPI Internal',
  ENVIRONMENT: 'production' // Change to 'testing' for testing
};

// Email recipients by environment
const BPI_INTERNAL_EMAIL_RECIPIENTS = {
  production: {
    to: "mjdagasuhan@bpi.com.ph",
    cc: "malodriga@bpi.com.ph, julsales@bpi.com.ph, jcslingan@bpi.com.ph, eagmarayag@bpi.com.ph, jmdcantorna@bpi.com.ph, egcameros@bpi.com.ph, jdduque@bpi.com.ph, rdtayag@bpi.com.ph, rsmendoza1@bpi.com.ph, vjdtorcuator@bpi.com.ph, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com",
    bcc: "mvolbara@pldt.com.ph, RBEspayos@smart.com.ph"
  },
  testing: {
    to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
    cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
    bcc: ""
  }
};

function bpiInternalCollectionsLogic() {
  CustomLogger.logInfo('Running BPI Internal Collections Logic...', PROJECT_NAME, 'bpiInternalCollectionsLogic');

  try {
    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    let { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = dateInfo;

    // Adjust collection date if tomorrow is a holiday
    if (isTomorrowHoliday) {
      tomorrowDate = adjustDateForHoliday(tomorrowDate);
      tomorrowDateString = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
      CustomLogger.logInfo(`Tomorrow is a holiday. Collection date adjusted to ${tomorrowDateString}`, PROJECT_NAME, 'bpiInternalCollectionsLogic');
    }

    // Check if adjusted date is weekend
    if (isWeekend(tomorrowDate)) {
      CustomLogger.logInfo(`Skipping collections - tomorrow date ${tomorrowDateString} falls on weekend.`, PROJECT_NAME, 'bpiInternalCollectionsLogic');
      return;
    }

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipientsInternal(BPI_INTERNAL_CONFIG.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubjectInternal(tomorrowDate, BPI_INTERNAL_CONFIG.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollectionsInternal(todayDate, tomorrowDate, todayDateString, tomorrowDateString, subject);

    // Process collections if any are eligible (production only)
    if (BPI_INTERNAL_CONFIG.ENVIRONMENT === 'production' && eligibleCollections.length > 0) {
      processEligibleCollectionsInternal(eligibleCollections, tomorrowDate, emailRecipients);
    } else if (eligibleCollections.length === 0) {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.', PROJECT_NAME, 'bpiInternalCollectionsLogic');
    } else {
      CustomLogger.logInfo(`Testing mode: ${eligibleCollections.length} collections identified but not sent.`, PROJECT_NAME, 'bpiInternalCollectionsLogic');
      console.log(JSON.stringify(collections, null, 2));
    }
  } catch (error) {
    CustomLogger.logError(`Error in bpiInternalCollectionsLogic: ${error}`, PROJECT_NAME, 'bpiInternalCollectionsLogic');
    throw error;
  }
}

/**
 * Adjust date by adding one day for holiday
 */
function adjustDateForHoliday(date) {
  const adjustedDate = new Date(date);
  adjustedDate.setDate(adjustedDate.getDate() + 1);
  return adjustedDate;
}

/**
 * Check if date falls on weekend
 */
function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6; // 0 = Sunday, 6 = Saturday
}

/**
 * Get email recipients based on environment
 */
function getEmailRecipientsInternal(environment) {
  return BPI_INTERNAL_EMAIL_RECIPIENTS[environment] || BPI_INTERNAL_EMAIL_RECIPIENTS.testing;
}

/**
 * Generate email subject line
 */
function generateEmailSubjectInternal(tomorrowDate, serviceBank) {
  const formattedDate = Utilities.formatDate(
    tomorrowDate,
    Session.getScriptTimeZone(),
    "MMMM d, yyyy (EEEE)"
  );
  return `${serviceBank} DPU Request - ${formattedDate}`;
}

/**
 * Get eligible collections based on criteria
 */
function getEligibleCollectionsInternal(todayDate, tomorrowDate, todayDateString, tomorrowDateString, subject) {
  const srvBank = BPI_INTERNAL_CONFIG.SERVICE_BANK;

  // Fetch data
  const machineData = getMachineDataByPartner(srvBank);
  const forCollectionData = getForCollections(srvBank);
  const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(srvBank);

  // Destructure machine data
  const [
    machineNames,
    percentValues,
    amountValues,
    collectionPartners,
    collectionSchedules,
    ,
    lastRemarks,
    businessDays
  ] = machineData;

  // Filter eligible collections
  const eligibleCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const lastRemark = normalizeSpaces(lastRemarks[i][0]);
    const businessDay = businessDays[i][0];

    // Apply exclusion filters
    if (shouldExcludeMachineInternal(machineName, amountValue, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, srvBank)) {
      return; // Skip this machine
    }

    // Add to eligible collections
    eligibleCollections.push([machineName, amountValue, srvBank, subject]);
  });

  return eligibleCollections;
}

/**
 * Determine if machine should be excluded from collection
 */
function shouldExcludeMachineInternal(machineName, amountValue, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, srvBank) {  // Check 1: Skip weekend collections
  if (skipWeekendCollections(srvBank, tomorrowDate)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, PROJECT_NAME, 'shouldExcludeMachineInternal');
    return true;
  }

  // Check 2: Already collected
  if (skipAlreadyCollected(machineName, forCollectionData)) {
    return true;
  }

  // Check 3: Excluded based on remarks
  if (forExclusionBasedOnRemarks(lastRemark, todayDateString, machineName)) {
    return true;
  }

  // Check 4: Requested yesterday
  if (forExclusionRequestYesterday(previouslyRequestedMachines, machineName, srvBank)) {
    return true;
  }

  // Check 5: Business day schedule
  const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

  if (!shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
    return true;
  }

  return false;
}

/**
 * Process eligible collections
 */
function processEligibleCollectionsInternal(eligibleCollections, tomorrowDate, emailRecipients) {
  const srvBank = BPI_INTERNAL_CONFIG.SERVICE_BANK;

  try {
    if (CONFIG_BPI.ENVIRONMENT === "production") {
      // Create hidden worksheet with collection data
      createHiddenWorksheetAndAddData(eligibleCollections, srvBank);

      // Save eligibleCollections to BigQuery
      saveEligibleCollectionsToBQ(eligibleCollections);

      // Process and send collections
      processCollectionsAndSendEmail(eligibleCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
      CustomLogger.logInfo(`Processed ${eligibleCollections.length} eligible collections for ${srvBank}`, PROJECT_NAME, 'processEligibleCollectionsInternal');
    } else {
      // Process and send collections
      processCollectionsAndSendEmail(eligibleCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
      CustomLogger.logInfo(`[TESTING MODE] Found ${eligibleCollections.length} eligible collections. No emails will be sent.`, PROJECT_NAME, "processEligibleCollectionsBPI()");
    }
  } catch (error) {
    CustomLogger.logError(`Error processing eligible collections: ${error}`, PROJECT_NAME, 'processEligibleCollectionsInternal');
    throw error;
  }
}

// function bpiInternalCollectionsLogic() {
//   CustomLogger.logInfo("Running BPI Internal Collections Logic...", PROJECT_NAME, "bpiInternalCollectionsLogic()");
//   const environment = "production";
//   const srvBank = "BPI Internal";
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = getTodayAndTomorrowDates();
//   const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
//   const subject = `${srvBank} DPU Request - ${tomorrowDateFormatted}`;

//   // Move collection schedule if tomorrow is holiday
//   if (isTomorrowHoliday) {
//     tomorrowDate.setDate(tomorrowDate.getDate() + 1);
//   }

//   const machineData = getMachineDataByPartner(srvBank);
//   const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRemarks, businessDays] = machineData;

//   const ENV_EMAIL_RECIPIENTS = {
//     production: {
//       to: "mjdagasuhan@bpi.com.ph",
//       cc: "malodriga@bpi.com.ph, julsales@bpi.com.ph, jcslingan@bpi.com.ph, eagmarayag@bpi.com.ph, jmdcantorna@bpi.com.ph, egcameros@bpi.com.ph, jdduque@bpi.com.ph, rdtayag@bpi.com.ph, rsmendoza1@bpi.com.ph, vjdtorcuator@bpi.com.ph, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com ",
//       bcc: "mvolbara@pldt.com.ph,RBEspayos@smart.com.ph ",
//     },
//     testing: {
//       to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       bcc: "",
//     },
//   };

//   // Usage
//   const emailRecipients = ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing;
//   const forCollectionData = getForCollections(srvBank);

//   const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(srvBank);
//   let forCollections = [];

//   machineNames.forEach((machineNameArr, i) => {
//     const machineName = machineNameArr[0];
//     const amountValue = amountValues[i][0];
//     const lastRemark = normalizeSpaces(lastRemarks[i][0]);
//     const businessDay = businessDays[i][0];
//     const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

//     // Skip weekends
//     if (skipWeekendCollections(srvBank, tomorrowDate)) {
//       CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, PROJECT_NAME, "bpiInternalCollectionsLogic()");
//       return false;
//     }

//     // Skip if already collected
//     if (skipAlreadyCollected(machineName, forCollectionData)) {
//       return;
//     }

//     // Skip if the last request should be excluded
//     if (forExclusionBasedOnRemarks(lastRemark, todayDateString, machineName)) {
//       return;
//     }

//     // Skip if already requested for collection yesterday
//     if (forExclusionRequestYesterday(previouslyRequestedMachines, machineName, srvBank)) {
//       return;
//     }


//     if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
//       forCollections.push([machineName, amountValue, srvBank, subject]);
//     }

//   });

//   if (environment === "production" && forCollections.length > 0) {
//     createHiddenWorksheetAndAddData(forCollections, srvBank);
//     processCollections(forCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
//   } else {
//     CustomLogger.logInfo("No eligible stores for collection tomorrow.", PROJECT_NAME, "bpiInternalCollectionsLogic()");
//   }
// }

function eTapCollectionsLogic() {
  CustomLogger.logInfo('Running eTap Collections Logic...', CONFIG.APP.NAME, 'eTapCollectionsLogic');

  try {
    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    const { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = dateInfo;

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipientsETap(CONFIG.ETAP.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubjectETap(tomorrowDate, CONFIG.ETAP.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollectionsETap(todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, subject);

    // Process collections if any are eligible (production only)
    if (CONFIG.ETAP.ENVIRONMENT === 'production' && eligibleCollections.length > 0) {
      processEligibleCollectionsETap(eligibleCollections, tomorrowDate, emailRecipients);
    } else if (eligibleCollections.length === 0) {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.', CONFIG.APP.NAME, 'eTapCollectionsLogic');
    } else {
      CustomLogger.logInfo(`Testing mode: ${eligibleCollections.length} collections identified but not sent.`, CONFIG.APP.NAME, 'eTapCollectionsLogic');
      console.log(JSON.stringify(eligibleCollections, null, 2));
    }

  } catch (error) {
    CustomLogger.logError(`Error in eTapCollectionsLogic: ${error.message}\nStack: ${error.stack}`, CONFIG.APP.NAME, 'eTapCollectionsLogic');
    throw error;
  }
}

/**
 * Get email recipients based on environment
 */
function getEmailRecipientsETap(environment) {
  return CONFIG.ETAP.EMAIL_RECIPIENTS[environment] || CONFIG.ETAP.EMAIL_RECIPIENTS.testing;
}

/**
 * Generate email subject line
 */
function generateEmailSubjectETap(tomorrowDate, serviceBank) {
  const formattedDate = Utilities.formatDate(
    tomorrowDate,
    Session.getScriptTimeZone(),
    "MMMM d, yyyy (EEEE)"
  );
  return `${serviceBank} DPU Request - ${formattedDate}`;
}

/**
 * Get eligible collections based on criteria
 */
function getEligibleCollectionsETap(todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, subject) {
  const srvBank = CONFIG.ETAP.SERVICE_BANK;

  // Fetch data
  const machineData = getMachineDataByPartner(srvBank);
  const forCollectionData = getForCollections(srvBank);
  const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(srvBank);

  // Destructure machine data
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, lastAddress, lastRemarks, businessDays] = machineData;

  // Filter eligible collections
  const eligibleCollections = [];
  const specialMachinesAdded = new Set();

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const collectionSchedule = collectionSchedules[i][0];
    const lastRemark = normalizeSpaces(lastRemarks[i][0]);
    const businessDay = businessDays[i][0];

    // Apply exclusion filters
    if (shouldExcludeMachineETap(machineName, amountValue, collectionSchedule, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, srvBank)) {
      return; // Skip this machine
    }

    // Handle special machines that must be collected together
    if (CONFIG.ETAP.SPECIAL_MACHINES.has(machineName)) {
      addSpecialMachines(eligibleCollections, specialMachinesAdded, amountValue, srvBank, subject);
    } else {
      // Format machine name with remark if it's a revisit
      const displayName = formatMachineNameWithRemarkETap(machineName, lastRemark, tomorrowDateString);
      eligibleCollections.push([displayName, amountValue, srvBank, subject]);
    }
  });

  return eligibleCollections;
}

/**
 * Determine if machine should be excluded from collection
 */
function shouldExcludeMachineETap(machineName, amountValue, collectionSchedule, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, srvBank) {
  // Check 1: Already collected
  if (skipAlreadyCollected(machineName, forCollectionData)) {
    return true;
  }

  // Check 2: No-Holiday schedule and tomorrow is holiday
  if (collectionSchedule.includes('No-Holiday') && isTomorrowHoliday) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store is closed during holidays.`, CONFIG.APP.NAME, 'shouldExcludeMachineETap');
    return true;
  }

  // Check 3: Excluded based on remarks
  if (forExclusionBasedOnRemarks(lastRemark, todayDateString, machineName)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store has remarks indicating no collection is needed.`, CONFIG.APP.NAME, 'shouldExcludeMachineETap');
    return true;
  }

  // Check 4: Requested yesterday
  if (forExclusionRequestYesterday(previouslyRequestedMachines, machineName, srvBank)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store was requested for collection yesterday.`, CONFIG.APP.NAME, 'shouldExcludeMachineETap');
    return true;
  }

    // Check 4: If not yet scheduled for collection
  if(forExclusionNotYetScheduled(machineName,tomorrowDate,lastRemark)){
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store is not yet scheduled for collection.`, CONFIG.APP.NAME, 'shouldExcludeMachineETap');
    return true;
  }

  // Check 5: Business day schedule
  const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

  if (!shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store is not scheduled for collection on ${tomorrowDateString}.`, CONFIG.APP.NAME, 'shouldExcludeMachineETap');
    return true;
  }

  return false;
}

/**
 * Add special machines that must be collected together
 */
function addSpecialMachines(eligibleCollections, specialMachinesAdded, amountValue, srvBank, subject) {
  // If one special machine qualifies, add both (but only once)
  CONFIG.ETAP.SPECIAL_MACHINES.forEach(specialMachine => {
    if (!specialMachinesAdded.has(specialMachine)) {
      eligibleCollections.push([specialMachine, amountValue, srvBank, subject]);
      specialMachinesAdded.add(specialMachine);
    }
  });
}

/**
 * Format machine name with remark if it's a revisit
 */
function formatMachineNameWithRemarkETap(machineName, lastRemark, tomorrowDateString) {
  const remarkLower = lastRemark.toLowerCase();
  const revisitPhrase = `for revisit on ${tomorrowDateString.toLowerCase()}`;

  if (remarkLower.includes(revisitPhrase)) {
    return `${machineName} (<b>${lastRemark}</b>)`;
  }

  return machineName;
}

/**
 * Process eligible collections
 */
function processEligibleCollectionsETap(eligibleCollections, tomorrowDate, emailRecipients) {
  const srvBank = CONFIG.ETAP.SERVICE_BANK;

  try {
    // Create hidden worksheet with collection data
    createHiddenWorksheetAndAddData(eligibleCollections, srvBank);

    // Process and send collections
    processCollectionsAndSendEmail(eligibleCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);

    CustomLogger.logInfo(`Processed ${eligibleCollections.length} eligible collections for ${srvBank}`, CONFIG.APP.NAME, 'processEligibleCollectionsETap');

  } catch (error) {
    CustomLogger.logError(`Error processing eligible collections: ${error}`, CONFIG.APP.NAME, 'processEligibleCollectionsETap');
    throw error;
  }
}/**
 * Optimized function to reset collection requests and revisits
 * Uses specific spreadsheet ID and batch operations for better performance
 * Handles both "for collection on" and "for revisit on" patterns
 */
function resetCollectionRequests() {
  const CONFIG = {
    spreadsheetId: '1YWYJl-0BOmfF-gf8FLpb0SPFG4xTnjhI5Ly77LA2UBc',
    sheetName: 'Kiosk %',
    daysBack: 2,
    patterns: {
      collection: 'for collection on ',
      revisit: 'for revisit on '
    }
  };
  
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet "${CONFIG.sheetName}" not found`);
  }
  
  const remarksIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()
    .flat()
    .findIndex(col => col === "Remarks") + 1;
  
  const { todayDate } = getTodayAndTomorrowDates();
  const kioskData = getMachineDataByPartner();
  const [machineNames, , , , , , lastRequests] = kioskData;

  const timeZone = Session.getScriptTimeZone();
  
  // Calculate target date (2 days prior to today)
  const targetDate = new Date(todayDate);
  targetDate.setDate(targetDate.getDate() - CONFIG.daysBack);
  const targetDateString = Utilities.formatDate(targetDate, timeZone, "MMM d");
  
  // Pre-compile search patterns for better performance
  const searchPatterns = [
    CONFIG.patterns.collection + targetDateString.toLowerCase(),
    CONFIG.patterns.revisit + targetDateString.toLowerCase()
  ];
  
  CustomLogger.logInfo(
    `Resetting entries matching: "${searchPatterns.join('" or "')}"`, 
    PROJECT_NAME, 
    'resetCollectionRequests'
  );

  // Collect rows to clear in batch
  const rowsToClear = [];
  const clearDetails = [];
  
  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const lastRequest = normalizeSpaces(lastRequests[i][0]);
    
    if (!lastRequest) return;
    
    const lastRequestLower = lastRequest.toLowerCase();
    const matchedPattern = searchPatterns.find(pattern => 
      lastRequestLower.startsWith(pattern)
    );
    
    if (matchedPattern) {
      rowsToClear.push(i + 2); // +2 because data starts at row 2
      clearDetails.push({
        machineName,
        pattern: matchedPattern,
        row: i + 2
      });
    }
  });

  // Batch clear all matching cells
  if (rowsToClear.length > 0) {
    rowsToClear.forEach((rowNum, idx) => {
      sheet.getRange(rowNum, remarksIndex).clearContent();
      CustomLogger.logInfo(        `Cleared "${clearDetails[idx].pattern}" for ${clearDetails[idx].machineName} (Row ${rowNum})`,         PROJECT_NAME,         'resetCollectionRequests'      );
    });
    
    CustomLogger.logInfo(
      `Total cleared: ${rowsToClear.length} machine(s)`, 
      PROJECT_NAME, 
      'resetCollectionRequests'
    );
  } else {
    CustomLogger.logInfo(
      `No machines found matching the target patterns for ${targetDateString}`, 
      PROJECT_NAME, 
      'resetCollectionRequests'
    );
  }
}



// /**
//  * Optimized function to reset collection requests
//  * Uses specific spreadsheet ID and batch operations for better performance
//  */
// function resetCollectionRequests() {

//   const CONFIG = {
//     spreadsheetId: '1YWYJl-0BOmfF-gf8FLpb0SPFG4xTnjhI5Ly77LA2UBc',
//     sheetName: 'Kiosk %'
//   };
  
//   const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
//   const sheet = ss.getSheetByName(CONFIG.sheetName);
//   const remarksIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat().findIndex(col => col === "Remarks") + 1;
  
//   if (!sheet) {
//     throw new Error(`Sheet "${CONFIG.sheetName}" not found`);
//   }
  
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
//   const kioskData = getMachineDataByPartner();
//   const [machineNames, percentValues, amountValues, collectionPartners, 
//          collectionSchedules, lastAddress, lastRequests, businessDays] = kioskData;

//   const timeZone = Session.getScriptTimeZone();
  
//   // Calculate yesterday's date (2 days prior to today)
//   const dayAfterYesterdayDate = new Date(todayDate);
//   dayAfterYesterdayDate.setDate(dayAfterYesterdayDate.getDate() - 2);
//   const yesterdayDateString = Utilities.formatDate(dayAfterYesterdayDate, timeZone, "MMM d");
  
//   const searchString = "for collection on " + yesterdayDateString.toLowerCase();
  
//   CustomLogger.logInfo(
//     `Resetting last request value [For collection on ${yesterdayDateString}]`, 
//     PROJECT_NAME, 
//     'resetCollectionRequests'
//   );

//   // Collect rows to clear in batch
//   const rowsToClear = [];
//   const machineNamesToClear = [];
  
//   machineNames.forEach((machineNameArr, i) => {
//     const machineName = machineNameArr[0];
//     const lastRequest = normalizeSpaces(lastRequests[i][0]);
    
//     if (lastRequest && lastRequest.toLowerCase() === searchString) {
//       rowsToClear.push(i + 2); // +2 because data starts at row 2
//       machineNamesToClear.push(machineName);
//     }
//   });

//   // Batch clear all matching cells
//   if (rowsToClear.length > 0) {
//     rowsToClear.forEach((rowNum, idx) => {
//       sheet.getRange(rowNum, remarksIndex).clearContent();
//       CustomLogger.logInfo(
//         `Cleared last request value of ${machineNamesToClear[idx]} (Row ${rowNum})`, 
//         PROJECT_NAME, 
//         'resetCollectionRequests'
//       );
//     });
    
//     CustomLogger.logInfo(
//       `Total cleared: ${rowsToClear.length} machine(s)`, 
//       PROJECT_NAME, 
//       'resetCollectionRequests'
//     );
//   } else {
//     CustomLogger.logInfo(
//       `No machines found matching "${searchString}"`, 
//       PROJECT_NAME, 
//       'resetCollectionRequests'
//     );
//   }
// }




// // function resetCollectionRequests() {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const sheet = ss.getSheetByName('Kiosk %');
// //   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
// //   const kioskData = getMachineDataByPartner();
// //   const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, lastAddress, lastRequests, businessDays] = kioskData;


// //   const timeZone = Session.getScriptTimeZone();

// //   const dayAafterYesterdayDate = new Date(todayDate); // 2 days prior
// //   dayAafterYesterdayDate.setDate(dayAafterYesterdayDate.getDate() - 2);
// //   const yesterdayDateString = Utilities.formatDate(dayAafterYesterdayDate, timeZone, "MMM d");

// //   CustomLogger.logInfo(`Resetting last request value [For collection on ${yesterdayDateString}]`, PROJECT_NAME, 'resetCollectionRequests');

// //   machineNames.forEach((machineNameArr, i) => {
// //     const machineName = machineNameArr[0];
// //     const lastRequest = normalizeSpaces(lastRequests[i][0]);
    
// //     if (lastRequest !== '') {
// //       if (lastRequest.toLocaleLowerCase() === "for collection on " + yesterdayDateString.toLowerCase()) {
// //         sheet.getRange(`R${i + 2}`).setValue('');
// //         CustomLogger.logInfo(`Cleared last request value [${sheet.getRange(`R${i + 2}`).getValue()}] of ${machineName}`, PROJECT_NAME, 'resetCollectionRequests');
// //       }
// //     }
// //   });

// // }
// // Configuration
// const CONFIG = {
//   spreadsheetId: "1TJ10XqwS_cTQfkxKKWJaE5zhBdE2pDgIdZ_Zw9JQD_U", // Your Google Sheet ID
//   sheetName: "Store Reps", // Name of your sheet tab
//   machineNameColumn: "A", // Column with machine names
//   emailColumn: "B", // Column with email addresses
//   startRow: 2, // Row where data starts (2 if you have headers)
//   environment: "testing",
// };

// function sendAdvancedNotice() {
//   CustomLogger.logInfo("Running sending DPU advanced notice...", PROJECT_NAME, "sendAdvancedNotice()");
//   var srvBank = "";
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
//   const collectionDay = dayMapping[tomorrowDate.getDay()];
//   const amountThreashold = amountThresholds[collectionDay];
//   const kioskData = getKioskPercentage(amountThreashold);
//   const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRequests, businessDays] = kioskData;
//   const formattedDate = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
//   const inAdvanceNotice = forExclusionPartOfAdvanceNotice();

//   var emailTo = "";
//   const pldtRecepient = "mvolbara@pldt.com.ph";
//   const smartRecepient = "RBEspayos@smart.com.ph, RACagbay@smart.com.ph";
//   let emailCc = "";
//   let emailBcc = "egalcantara@multisyscorp.com";

//   machineNames.forEach((machineNameArr, i) => {
//     const machineName = machineNameArr[0];
//     const amountValue = amountValues[i][0];
//     const percentValue = percentValues[i][0];
//     const collectionPartner = collectionPartners[i][0];
//     const lastRequest = lastRequests[i][0];
//     const businessDay = businessDays[i][0];
//     const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

//     if (inAdvanceNotice && inAdvanceNotice.includes(machineName)) {
//       CustomLogger.logInfo(`Skipping ${machineName} for advance notice, already in advance notice...`, PROJECT_NAME, "sendAdvancedNotice()");
//       return;
//     }

//     if (forExclusionBasedOnRemarks(lastRequest, todayDateString, machineName)) {
//       return;
//     }

//     if (CONFIG.environment === "production") {
//       emailTo = getEmailByMachineName(machineName);
//       if (machineName.includes("PLDT")) {
//         emailCc = pldtRecepient;
//       } else if (machineName.includes("SMART")) {
//         emailCc = smartRecepient;
//       }
//     } else {
//       emailTo = "egalcantara@multisyscorp.com";
//       emailCc = "zhere27@gmail.com";
//     }

//     if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank)) {
//       if (emailTo !== "") {
//         if (CONFIG.environment === "production") {
//           CustomLogger.logInfo(`Sending advance notice to ${machineName} to email addresses: ${emailTo}...`, PROJECT_NAME, "sendAdvancedNotice()");
//           addMachineToAdvanceNotice(machineName);
//           sendEmail(machineName, formattedDate, percentValue);
//         } else {
//           addMachineToAdvanceNotice(machineName);
//           CustomLogger.logInfo(`[TESTING] advance notice for ${machineName} to email addresses: ${emailTo}...`, PROJECT_NAME, "sendAdvancedNotice()");
//         }
//       }
//     }
//   });

//   function sendEmail(machineName, formattedDate, percentValue) {
//     const subject = `Paybox - DPU Collection advanced notice - ${formattedDate}`;

//     //percent value should be displayed as a percentage
//     percentValue = parseFloat(percentValue * 100.0).toFixed(2);

//     let body;

//     if (percentValue < 100) {
//       // Less than 100% capacity - Collection may take place
//       body = `Hi ${machineName},<br><br>
// The Paybox machine is at ${percentValue}% capacity. DPU collection has been requested and scheduled for pickup tomorrow ${formattedDate}.<br><br>
// Please have the required permit ready, if applicable. <br><br>
// Thank you.<br><br>

// *** This is an automated email. ****<br><br>${emailSignature}
// `;
//     } else {
//       // 100% or more capacity - Collection has been scheduled
//       body = `Hi ${machineName},<br><br>
// The Paybox machine is due for collection and has been requested and scheduled for pickup tomorrow ${formattedDate}.<br><br>
// Please have the required permit ready, if applicable.<br><br>
// Thank you.<br><br>

// *** This is an automated email. ****<br><br>${emailSignature}
// `;
//     }

//     GmailApp.sendEmail(emailTo, subject, "", {
//       cc: emailCc,
//       bcc: emailBcc,
//       htmlBody: body,
//       from: "support@paybox.ph",
//     });
//   }
// }

// /**
//  * Search for machine name and return corresponding email
//  */
// function getEmailByMachineName(machineName) {
//   const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
//   const sheet = ss.getSheetByName(CONFIG.sheetName);

//   if (!sheet) {
//     throw new Error('Sheet "' + CONFIG.sheetName + '" not found');
//   }

//   // Get all data from the sheet
//   const dataRange = sheet.getDataRange();
//   const data = dataRange.getValues();

//   // Search for the machine name (case-insensitive)
//   for (let i = CONFIG.startRow - 1; i < data.length; i++) {
//     const currentMachineName = data[i][0]; // Column A (index 0)

//     if (currentMachineName.toString().trim().toLowerCase() === machineName.toString().trim().toLowerCase()) {
//       const email = data[i][1]; // Column B (index 1)
//       return email;
//     }
//   }

//   return null; // Machine not found
// }

// /**
//  * Test function to verify the lookup works
//  */
// function testLookup() {
//   const testMachine = "SMART SM NAGA 2";
//   const email = getEmailByMachineName(testMachine);

//   if (email) {
//     Logger.log("Found: " + testMachine + " -> " + email);
//   } else {
//     Logger.log("Machine not found: " + testMachine);
//   }
// }
/**
 * Sends a collection email
 * @param {Array} machineData - Array of machine data
 * @param {Date} collectionDate - Collection date
 * @param {string} emailTo - Email recipients (to)
 * @param {string} emailCc - Email recipients (cc)
 * @param {string} emailBcc - Email recipients (bcc)
 * @param {string} srvBank - Service bank
 */
function sendEmailCollection(machineData, collectionDate, emailTo, emailCc, emailBcc, srvBank) {
  try {
    if (!machineData || machineData.length === 0) {
      CustomLogger.logInfo("No machine data to send in email.", PROJECT_NAME, "sendEmailCollection()");
      return;
    }

    // Helper: Sort by machine name (case-insensitive alphabetical)
    const sortByMachineName = (a, b, index) => {
      return a[index].toLowerCase().localeCompare(b[index].toLowerCase());
    };

    // Main logic
    machineData.sort((a, b) => sortByMachineName(a, b, 0));

    const formattedDate = formatDate(collectionDate, "MMMM d, yyyy (EEEE)");
    const subject = `${srvBank} DPU Request - ${formattedDate}`;

    let body;
    if (machineData.some((item) => item[0] === "PLDT BANTAY" || item[0] === "SMART VIGAN")) {
      body = `
        Hi All,<br><br>
        Good day! Please schedule <b>collection</b> on ${formattedDate} for the following stores:<br><br>
        ${machineData.map((row) => row[0]).join("<br>")}<br><br>
        For <b>PLDT BANTAY and SMART VIGAN</b>, kindly collect it on the same day.<br><br>
        *** Please acknowledge this email. ****<br><br>${emailSignature}
      `;
    } else {
      body = `
        Hi All,<br><br>
        Good day! Please schedule <b>collection</b> on ${formattedDate} for the following stores:<br><br>
        ${machineData.map((row) => row[0]).join("<br>")}<br><br>
        *** Please acknowledge this email. ****<br><br>${emailSignature}
      `;
    }

    GmailApp.sendEmail(emailTo, subject, "", {
      cc: emailCc,
      bcc: emailBcc,
      htmlBody: body,
      from: "support@paybox.ph",
    });
    CustomLogger.logInfo(`Collection email sent for ${machineData.length} machines.`, PROJECT_NAME, "sendEmailCollection()");
  } catch (error) {
    CustomLogger.logError(`Error in sendEmailCollection: ${error.message}\nStack: ${error.stack}`, PROJECT_NAME, "sendEmailCollection()");
    throw error;
  }
}

/**
 * Sends a cancellation email
 * @param {Array} forCancellation - Array of cancellations
 * @param {Date} collectionDate - Collection date
 * @param {string} emailTo - Email recipients (to)
 * @param {string} emailCc - Email recipients (cc)
 * @param {string} emailBcc - Email recipients (bcc)
 * @param {string} srvBank - Service bank
 * @param {boolean} isSaturday - Whether it's Saturday
 */
function sendEmailCancellation(forCancellation, collectionDate, emailTo, emailCc, emailBcc, srvBank, isSaturday) {
  try {
    // Cancel sending email if no records in the forCancellation
    if (!forCancellation || forCancellation.length === 0) {
      CustomLogger.logInfo("No cancellations to send in email.", PROJECT_NAME, "sendEmailCancellation()");
      return;
    }

    const formattedDate = formatDate(collectionDate, "EEEE, MMMM d, yyyy");
    const subject = `Cancellation Pickup for PLDT x Smart Locations | ${formattedDate} | MULTISYS TECHNOLOGIES CORPORATION`;

    const body = `
      Hi Team,<br><br>
      Good day! Please <b>cancel</b> the collection scheduled for ${formattedDate}, for the following stores:<br><br>
      <table>
      <tr>
        <th>Store</th>
        <th>&nbsp;</th>
        <th>Address</th>
      </tr>
      ${forCancellation
        .map(
          (store) => `
        <tr>
          <td>${store.name}</td>
          <td>&nbsp;</td>  
          <td>${store.address}</td>
        </tr>
      `
        )
        .join("")}
      </table>
      <br><br>
      *** Please acknowledge this email. ****<br><br>${emailSignature}
    `;

    GmailApp.sendEmail(emailTo, subject, "", {
      cc: emailCc,
      bcc: emailBcc,
      htmlBody: body,
      from: "support@paybox.ph",
    });
    CustomLogger.logInfo(`Cancellation email sent for ${forCancellation.length} machines.`, PROJECT_NAME, "sendEmailCancellation()");
  } catch (error) {
    CustomLogger.logError(`Error in sendEmailCancellation: ${error.message}\nStack: ${error.stack}`, PROJECT_NAME, "sendEmailCancellation()");
    throw error;
  }
}

/**
 * Replies to an existing email thread
 * @param {string} subject - Email subject
 * @param {string} messageBody - Message body
 */
function replyToExistingThread(subject, messageBody) {
  try {
    // Search for the thread by its subject
    const threads = GmailApp.search(`subject:"${subject}"`);

    if (threads.length === 0) {
      throw new Error(`No thread found with the subject: "${subject}"`);
    }

    // Get the first matching thread
    const thread = threads[0];

    // Reply to the thread
    thread.replyAll("", {
      htmlBody: messageBody,
      from: "support@paybox.ph",
    });
    CustomLogger.logInfo(`Reply sent to the thread with subject: "${subject}"`, PROJECT_NAME, "replyToExistingThread()");
  } catch (error) {
    CustomLogger.logError(`Error in replyToExistingThread: ${error.message}\nStack: ${error.stack}`, PROJECT_NAME, "replyToExistingThread()");
    throw error;
  }
}function test_hasSpecialCollectionConditions() {
 console.log(hasSpecialCollectionConditions('For collection on Aug 25','Aug 23')); 
}
function test_excludeAlreadyCollected(machineName='PLDT SURIGAO', srvBank='Apeiros') {
  console.log(skipAlreadyCollected(machineName, srvBank));
}

function test_isTomorrowHoliday() {
  var { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();

  console.log(isTomorrowHoliday(tomorrowDate));
}

function test_forExclusionRequestYesterday(machineName='SMART SM CEBU 2', srvBank='Brinks via BPI') {
  console.log(forExclusionRequestYesterday(machineName, srvBank))
}

function test_shouldExcludeFromCollection_stg() {
  const lastRequest = '';
  const todayDay = 'Aug 18';
  const machineName = 'PLDT TALISAY';
  const result = shouldExcludeFromCollection_stg(lastRequest, todayDay, machineName);
  console.log(`Result for "${lastRequest}" and "${todayDay}" on "${machineName}": ${result}`);
}

/**
 * Checks if a machine should be excluded from collection based on last request and current day
 * @param {string} lastRequest - Last request string to check against exclusion criteria
 * @param {string} todayDay - Today's day name (e.g., "monday", "tuesday")
 * @param {string} [machineName=null] - Optional machine name for logging purposes
 * @return {boolean} True if the machine should be excluded, false otherwise
 */
function shouldExcludeFromCollection_stg(lastRequest, todayDay, machineName = null) {
  try {
    // Early return for null/undefined/empty lastRequest
    if (!lastRequest?.trim()) {
      return false;
    }

    // Normalize inputs once
    const normalizedLastRequest = lastRequest.toLowerCase().trim();
    const normalizedTodayDay = todayDay?.toLowerCase().trim() || '';

    // Define exclusion criteria with better organization
    const exclusionReasons = [
      'waiting for collection items',
      'waiting for bank updates', 
      'did not reset',
      'cassette removed',
      'manually collected',
      'for repair',
      'already collected',
      'store is closed',
      'store is not using the machine',
      normalizedTodayDay
    ].filter(reason => reason); // Remove empty strings

    // Check for exclusion with optimized search
    const isExcluded = exclusionReasons.some(reason => 
      normalizedLastRequest.includes(reason)
    );

    // Log exclusion if machine name provided and excluded
    if (isExcluded && machineName) {
      CustomLogger.logInfo(`Machine "${machineName}" excluded - Last remarks: "${lastRequest}"`,PROJECT_NAME,'shouldExcludeFromCollection');
    }

    return isExcluded;

  } catch (error) {
    CustomLogger.logError(`Error in shouldExcludeFromCollection(): ${error.message}`,PROJECT_NAME,'shouldExcludeFromCollection');
    return false; // Fail-safe: don't exclude on error
  }
}

function test_meetsAmountThreshold(amountValue=151000, collectionDay, srvBank='Brinks via BPI') {
  console.log(meetsAmountThreshold(amountValue, collectionDay, srvBank));
}

function test_addMachineToAdvanceNotice(machineName='PLDT SURIGAO') {
  addMachineToAdvanceNotice(machineName);
}
function refreshStores() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("Kiosk %");

  // Clear the range from A2:B to the last row, as well as any formatting
  var lastRow = sheet.getLastRow();
  sheet.getRange("A2:A" + lastRow).clear().setFontFamily('Century Gothic').setFontSize(9);

  try {
    const rows = getStoreList();

    if (!rows || rows.length === 0) {
      Logger.log("No data to populate the new sheet.");
      return;
    }

    // Prepare the data rows for insertion
    const data = rows.map(row => row.f.map(cell => cell.v));

    // Define the range where new data will be appended
    var range = sheet.getRange(2, 1, data.length, data[0].length);

    // Set the new data values
    range.setValues(data);

    // Activate the sheet
    sheet.activate();

  } catch (e) {
    Logger.log("Error populating sheet: " + e.message);
  }

  sortLatestPercentage();
}function test_shouldIncludeForCollection() {
  const machineName="PLDT ROBINSONS DUMAGUETE";
  const amountValue=250000;
  const collectionSchedule=["M.W.Sat."];
  const tomorrowDate=new Date(2025,7,17);
  const tomorrowDateString="Aug 18";
  const todayDate="Aug 17";
  const lastRequest="";
  const srvBank="eTap";
  const translatedBusinessDays = translateDaysToAbbreviation("Monday - Sunday");

  if (shouldIncludeForCollection_1(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank, collectionSchedule)) {
    console.log("should include for collection");
  }
} 

/**
 * Main function to determine if a machine should be included for collection
 * @param {string} machineName - The name of the machine
 * @param {number} amountValue - The amount in the machine
 * @param {Array} translatedBusinessDays - Array of translated business days
 * @param {Date} tomorrowDate - Tomorrow's date object
 * @param {string} tomorrowDateString - Tomorrow's date as string
 * @param {Date} todayDate - Today's date object
 * @param {string} lastRequest - The last request string
 * @param {string} srvBank - The bank service name
 * @returns {boolean} - True if machine should be included for collection
 */
function shouldIncludeForCollection_1(  machineName,   amountValue,   translatedBusinessDays,   tomorrowDate,   tomorrowDateString,   todayDate,   lastRequest,   srvBank, collectionSchedule
) {
  try {
    const collectionDay = dayMapping[tomorrowDate.getDay()];
    const dayOfWeek = tomorrowDate.getDay();

    // Early returns for exclusion conditions
    
    // 1. Check if excluded store
    if (isExcludedStore(machineName)) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString}, part of the excluded stores.`, 
        PROJECT_NAME, 
        "shouldIncludeForCollection"
      );
      return false;
    }

    // 2. Check collection schedule
    if (!translatedBusinessDays.includes(collectionDay)) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString}, not a collection day.`, 
        PROJECT_NAME, 
        "shouldIncludeForCollection"
      );
      return false;
    }


    // 3. Check BPI weekend restrictions
    if (shouldSkipWeekendCollections(srvBank, tomorrowDate)) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, 
        PROJECT_NAME, 
        "shouldIncludeForCollection"
      );
      return false;
    }

    // Early returns for inclusion conditions (highest priority)
    
    // 1. Special collection conditions from last request
    if (hasSpecialCollectionConditions(lastRequest, tomorrowDateString)) {
      return true;
    }

    // 2. Schedule-based store collections
    if (SCHEDULE_CONFIG.stores.includes(machineName)) {
      return shouldCollectScheduleBasedStore(machineName, dayOfWeek);
    }

    // 3. Regular threshold-based collection
    if (meetsAmountThreshold(amountValue, collectionDay)) {
      return true;
    }

    // 4. Payday collection
    if (isPaydayCollection(todayDate, amountValue)) {
      return true;
    }

    // 5. Due date collection
    if (isDueDateCollection(todayDate, amountValue)) {
      return true;
    }

    // Default: no collection needed
    return false;

  } catch (error) {
    CustomLogger.logError(
      `Error in shouldIncludeForCollection(): ${error.message}\nStack: ${error.stack}`, 
      PROJECT_NAME, 
      "shouldIncludeForCollection()"
    );
    return false;
  }
}

function test_excludePastRequests() {
  var forCollections = [
    [
      "SMART SM BATAAN",
      963150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SAN FERNANDO LA UNION",
      393850.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT TALISAY",
      385390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SF LA UNION 1",
      496100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT OLONGAPO",
      378970.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT SM NORTH EDSA",
      428880.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART FESTIVAL MALL 2",
      454150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "SMART ROBINSONS GAPAN",
      370020.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT MANDAUE",
      343460.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 2",
      500400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SBMA",
      543400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT PASAY",
      350100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT MANAOAG",
      276200.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART MARKET MARKET",
      427810.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART AYALA CENTER CEBU",
      411120.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT JONES 1",
      354140.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 1",
      364840.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SM CEBU",
      376390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT DASMARINAS",
      311930.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART IBA ZAMBALES",
      356110.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM TARLAC",
      251710.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT GAPAN",
      100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
  ];

  forCollections = excludePastRequests(forCollections, "Jul 23");
  // forCollections = excludePreviouslyRequested(forCollections, 'Brinks via BPI');

  for (let i = 0; i < forCollections.length; i++) {
    console.log(forCollections[i][0]);
  }
}

function test_excludeRecentlyCollected(forCollections, srvBank) {
  var forCollections = [
    [
      "SMART SM BATAAN",
      963150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SAN FERNANDO LA UNION",
      393850.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT TALISAY",
      385390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SF LA UNION 1",
      496100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT OLONGAPO",
      378970.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT SM NORTH EDSA",
      428880.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART FESTIVAL MALL 2",
      454150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "SMART ROBINSONS GAPAN",
      370020.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT MANDAUE",
      343460.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 2",
      500400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SBMA",
      543400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT PASAY",
      350100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT MANAOAG",
      276200.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART MARKET MARKET",
      427810.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART AYALA CENTER CEBU",
      411120.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT JONES 1",
      354140.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 1",
      364840.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SM CEBU",
      376390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT DASMARINAS",
      311930.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART IBA ZAMBALES",
      356110.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM TARLAC",
      251710.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT GAPAN",
      100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
  ];

  forCollections = excludeRecentlyCollected(forCollections, "Brinks via BPI");

  for (let i = 0; i < forCollections.length; i++) {
    console.log(forCollections[i][0]);
  }
}

function test_createHiddenWorksheetAndAddData(forCollections, srvBank) {
  var forCollections = [
    [
      "SMART SM BATAAN",
      963150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SAN FERNANDO LA UNION",
      393850.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT TALISAY",
      385390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SF LA UNION 1",
      496100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT OLONGAPO",
      378970.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT SM NORTH EDSA",
      428880.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART FESTIVAL MALL 2",
      454150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "SMART ROBINSONS GAPAN",
      370020.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT MANDAUE",
      343460.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 2",
      500400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SBMA",
      543400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT PASAY",
      350100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT MANAOAG",
      276200.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART MARKET MARKET",
      427810.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART AYALA CENTER CEBU",
      411120.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT JONES 1",
      354140.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 1",
      364840.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SM CEBU",
      376390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT DASMARINAS",
      311930.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART IBA ZAMBALES",
      356110.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM TARLAC",
      251710.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT GAPAN",
      100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
  ];

  forCollections.sort(function (a, b) {
    const aAmount = parseFloat(a[1]);
    const bAmount = parseFloat(b[1]);
    if (isNaN(aAmount) && isNaN(bAmount)) return 0;
    if (isNaN(aAmount)) return 1;
    if (isNaN(bAmount)) return -1;
    return bAmount - aAmount;
  });

  createHiddenWorksheetAndAddData(forCollections, "Brinks via BPI");
}

function test_getTodayAndTomorrowDates() {
  const today = getTodayAndTomorrowDates();
  console.log(today[0]);
  console.log(today[1]);
}

function test_isTomorrowHoliday(tomorrow) {
  console.log(isTomorrowHoliday(tomorrow));
}

function test() {
  // console.log(isTomorrowHoliday(new Date('2025-04-17')));
  const environment = "testing"; // or "testing"
  console.log(
    ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing
  );
}

function test_skipFunction() {
  skipFunction();
}

function test_getLastHourly() {
  getLastHourly();
}

function test_setRowsHeightAndAlignment() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");

  const lastRow = sheet.getLastRow();

  setRowsHeightAndAlignment(sheet, lastRow);
}

function test_dryer_sorter() {
  var forCollections = [
    [
      "SMART SM BATAAN",
      963150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SAN FERNANDO LA UNION",
      393850.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT TALISAY",
      385390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SF LA UNION 1",
      496100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT OLONGAPO",
      378970.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT SM NORTH EDSA",
      428880.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART FESTIVAL MALL 2",
      454150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "SMART ROBINSONS GAPAN",
      370020.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT MANDAUE",
      343460.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 2",
      500400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SBMA",
      543400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT PASAY",
      350100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT MANAOAG",
      276200.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART MARKET MARKET",
      427810.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART AYALA CENTER CEBU",
      411120.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT JONES 1",
      354140.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 1",
      364840.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SM CEBU",
      376390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT DASMARINAS",
      311930.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART IBA ZAMBALES",
      356110.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM TARLAC",
      251710.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT GAPAN",
      100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
  ];
  srvBank="Brinks via BPI";

  // Helper: Sort by numeric value (desc), pushing NaN to the end
  const sortByAmountDescNaNLast = (a, b, index) => {
    const aAmount = Number(a[index]);
    const bAmount = Number(b[index]);

    if (Number.isNaN(aAmount) && Number.isNaN(bAmount)) return 0;
    if (Number.isNaN(aAmount)) return 1;
    if (Number.isNaN(bAmount)) return -1;

    return bAmount - aAmount; // Descending
  };

  // Helper: Sort by string (case-sensitive) at given index
  const sortByString = (a, b, index) => a[index].localeCompare(b[index]);

  // Main logic
  forCollections.sort(
    srvBank === "Brinks via BPI"
      ? (a, b) => sortByAmountDescNaNLast(a, b, 2)
      : (a, b) => sortByString(a, b, 0)
  );

    const jsonResult = convertArrayToJson(forCollections);
  Logger.log(JSON.stringify(jsonResult, null, 2)); // Pretty print
}


function test_shouldCollectScheduleBasedStore(machineName='PLDT ROBINSONS DUMAGUETE', dayOfWeek=1){
  var retVal = shouldCollectScheduleBasedStore(machineName, dayOfWeek);
  console.log(retVal);
}