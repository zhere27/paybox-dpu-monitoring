
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

/**
 * Filters machine data based on partner, amount threshold, and collection remarks
 * @param {Object} sheetData - Object containing all sheet column data
 * @param {string|null} srvBank - Optional partner to filter by
 * @param {Date} tomorrowDate - The target collection date
 * @returns {number[]} Array of indices that meet filter criteria
 */
function getFilteredIndices(sheetData, srvBank, collectionDay) {
  const amountThreshold = amountThresholds[collectionDay];

  const filteredIndices = [];

  sheetData.collectionPartners.forEach((partnerArray, index) => {
    const partner = partnerArray[0];
    const amount = sheetData.amountValues[index][0];
    const remark = sheetData.lastRemarks[index][0];

    // Check partner filter
    if (srvBank !== null && partner !== srvBank) return;

    // Check amount threshold or collection remark match
    if (!shouldIncludeMachine(amount, amountThreshold, remark, tomorrowDate)) return;

    filteredIndices.push(index);
  });

  return filteredIndices;
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

  for (var i = 9; i < sheets.length; i++) { // Start hiding from the 10th sheet (index 9)
    sheets[i].hideSheet();
  }
  CustomLogger.logInfo("Hide sheets after Kiosk%", CONFIG.APP.NAME, 'hideSheets()');
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




// ============================================================================
// CONFIGURATION MANAGEMENT
// ============================================================================

const SHEET_CONFIG = {
  SHEET_NAME: 'Kiosk %',
  START_ROW: 2,
  COLUMN_MAP: {
    machineNames: 1,      // A
    partnerAddress: 2,    // B
    percentValues: 16,    // P
    amountValues: 17,     // Q
    lastRemarks: 18,      // R
    collectionSchedules: 23, // W
    collectionPartners: 24,  // X
    businessDays: 25      // Y
  }
};

// Regex pattern for dynamic collection date matching
// Matches "For collection on Oct XX" where XX is any day
const COLLECTION_REMARK_PATTERN = /^For collection on\s+(\w+)\s+(\d{1,2})$/i;

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Extracts day number from a date remark string
 * @param {string} remark - The remark text (e.g., "For collection on Oct 17")
 * @param {Date} targetDate - The target date to compare against
 * @returns {boolean} True if the remark matches the target date
 */
function matchesCollectionRemark(remark, targetDate) {
  if (!remark || typeof remark !== 'string') return false;
  
  const match = remark.match(COLLECTION_REMARK_PATTERN);
  if (!match) return false;
  
  try {
    const remarkMonth = match[1];
    const remarkDay = parseInt(match[2]);
    
    const targetMonth = targetDate.toLocaleDateString('en-US', { month: 'short' });
    const targetDay = targetDate.getDate();
    
    return remarkMonth.toLowerCase() === targetMonth.toLowerCase() && 
           remarkDay === targetDay;
  } catch (error) {
    return false;
  }
}

/**
 * Determines if a machine should be included based on amount and threshold
 * @param {number} amount - The machine amount
 * @param {number} threshold - The amount threshold
 * @param {string} lastRemark - The last remark for the machine
 * @param {Date} collectionDate - The target collection date
 * @returns {boolean} True if machine meets criteria
 */
function shouldIncludeMachine(amount, threshold, lastRemark, collectionDate) {
  const amountNum = parseFloat(amount) || 0;
  
  // Include if amount meets threshold OR if remark matches collection date
  if (amountNum >= threshold) return true;
  if (matchesCollectionRemark(lastRemark, collectionDate)) return true;
  
  return false;
}

/**
 * Retrieves all data from the Kiosk sheet
 * @returns {Object} Object containing all column data
 */
function getSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const numRows = Math.max(0, lastRow - SHEET_CONFIG.START_ROW + 1);
  
  if (numRows === 0) {
    return null;
  }
  
  return {
    machineNames: sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.COLUMN_MAP.machineNames, numRows).getValues(),
    partnerAddress: sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.COLUMN_MAP.partnerAddress, numRows).getValues(),
    percentValues: sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.COLUMN_MAP.percentValues, numRows).getValues(),
    amountValues: sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.COLUMN_MAP.amountValues, numRows).getValues(),
    lastRemarks: sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.COLUMN_MAP.lastRemarks, numRows).getValues(),
    collectionSchedules: sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.COLUMN_MAP.collectionSchedules, numRows).getValues(),
    collectionPartners: sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.COLUMN_MAP.collectionPartners, numRows).getValues(),
    businessDays: sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.COLUMN_MAP.businessDays, numRows).getValues()
  };
}

/**
 * Filters machine data based on partner, amount threshold, and collection remarks
 * @param {Object} sheetData - Object containing all sheet column data
 * @param {string|null} srvBank - Optional partner to filter by
 * @param {number} amountThreshold - The amount threshold for the collection day
 * @param {Date} tomorrowDate - The target collection date
 * @returns {number[]} Array of indices that meet filter criteria
 */
function getFilteredIndices(sheetData, srvBank, amountThreshold, tomorrowDate) {
  const filteredIndices = [];
  
  sheetData.collectionPartners.forEach((partnerArray, index) => {
    const partner = partnerArray[0];
    const amount = sheetData.amountValues[index][0];
    const remark = sheetData.lastRemarks[index][0];
    
    // Check partner filter
    if (srvBank !== null && partner !== srvBank) return;
    
    // Check amount threshold or collection remark match
    if (!shouldIncludeMachine(amount, amountThreshold, remark, tomorrowDate)) return;
    
    filteredIndices.push(index);
  });
  
  return filteredIndices;
}

/**
 * Maps filtered indices across all data columns
 * @param {Object} sheetData - Object containing all sheet column data
 * @param {number[]} indices - Indices to extract
 * @returns {Object} Object containing filtered data for all columns
 */
function mapFilteredData(sheetData, indices) {
  return {
    machineNames: indices.map(i => sheetData.machineNames[i]),
    percentValues: indices.map(i => sheetData.percentValues[i]),
    amountValues: indices.map(i => sheetData.amountValues[i]),
    collectionPartners: indices.map(i => sheetData.collectionPartners[i]),
    collectionSchedules: indices.map(i => sheetData.collectionSchedules[i]),
    partnerAddress: indices.map(i => sheetData.partnerAddress[i]),
    lastRemarks: indices.map(i => sheetData.lastRemarks[i]),
    businessDays: indices.map(i => sheetData.businessDays[i])
  };
}

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Retrieves machine data filtered by partner and amount threshold
 * Includes machines with amounts below threshold if lastRemarks matches collection date
 * 
 * @param {string|null} srvBank - Optional partner name to filter by (null = all partners)
 * @param {Date} tomorrowDate - The target collection date for dynamic remark matching
 * @param {number} amountThreshold - The amount threshold for the collection day
 * @returns {Array} Array containing filtered data: [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, partnerAddress, lastRemarks, businessDays]
 * 
 * @example
 * const tomorrow = new Date();
 * tomorrow.setDate(tomorrow.getDate() + 1);
 * const [machineNames, percentValues, amountValues, ...] = getMachineDataByPartner('BDO', tomorrow, 300000);
 */
function getMachineDataByPartner(srvBank = null, tomorrowDate) {
  const collectionDay = dayMapping[tomorrowDate.getDay()];
  const amountThreshold = amountThresholds[collectionDay];
  try {
    // Validate inputs
    if (!tomorrowDate || !(tomorrowDate instanceof Date)) {
      throw new Error('tomorrowDate must be a valid Date object');
    }
    
    if (typeof amountThreshold !== 'number' || amountThreshold < 0) {
      throw new Error('amountThreshold must be a non-negative number');
    }
    
    // Retrieve all sheet data
    const sheetData = getSheetData();
    if (!sheetData) {
      return [[], [], [], [], [], [], [], []];
    }
    
    // Filter indices based on criteria
    const filteredIndices = getFilteredIndices(sheetData, srvBank, amountThreshold, tomorrowDate);
    
    // Map and return filtered data
    const filtered = mapFilteredData(sheetData, filteredIndices);
    
    return [
      filtered.machineNames,
      filtered.percentValues,
      filtered.amountValues,
      filtered.collectionPartners,
      filtered.collectionSchedules,
      filtered.partnerAddress,
      filtered.lastRemarks,
      filtered.businessDays
    ];
    
  } catch (error) {
    CustomLogger.logError(
      `Error in getMachineDataByPartner: ${error}`,
      'MachineDataModule',
      'getMachineDataByPartner'
    );
    throw error;
  }
}

// ============================================================================
// USAGE EXAMPLE
// ============================================================================

/*
const tomorrow = new Date();
tomorrow.setDate(tomorrow.getDate() + 1);
const amountThreshold = 300000;

// Get all partners - returns array
const allData = getMachineDataByPartner(null, tomorrow, amountThreshold);

// Get specific partner - returns array
const [
  machineNames,
  percentValues,
  amountValues,
  collectionPartners,
  collectionSchedules,
  partnerAddress,
  lastRemarks,
  businessDays
] = getMachineDataByPartner('BDO', tomorrow, amountThreshold);
*/