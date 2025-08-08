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

  const lastRow = sheet.getLastRow();
  const machineNames = sheet.getRange(2, 1, lastRow - 1).getValues();
  const partnerAddress = sheet.getRange(2, 2, lastRow - 1).getValues(); // Column B - Collection Team
  const percentValues = sheet.getRange(2, 16, lastRow - 1).getValues(); // Column P - Current Percentage
  const amountValues = sheet.getRange(2, 17, lastRow - 1).getValues(); // Column Q - Current Cash Amount
  const lastRequests = sheet.getRange(2, 19, lastRow - 1).getValues(); // Column R - Remarks
  const collectionSchedules = sheet.getRange(2, 23, lastRow - 1).getValues(); // Column W - Frequency
  const collectionPartners = sheet.getRange(2, 24, lastRow - 1).getValues(); // Column X - Servicing Bank
  const businessDays = sheet.getRange(2, 25, lastRow - 1).getValues(); // Column Y - Business Days

  return [
    machineNames,
    percentValues,
    amountValues,
    collectionPartners,
    collectionSchedules,
    partnerAddress,
    lastRequests,
    businessDays,
  ];
}

function getKioskData(srvBank = null) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Kiosk %");

  const lastRow = sheet.getLastRow();
  const machineNames = sheet.getRange(2, 1, lastRow - 1).getValues(); // Column A
  const partnerAddress = sheet.getRange(2, 2, lastRow - 1).getValues(); // Column B
  const percentValues = sheet.getRange(2, 16, lastRow - 1).getValues(); // Column P
  const amountValues = sheet.getRange(2, 17, lastRow - 1).getValues(); // Column Q
  const lastRequests = sheet.getRange(2, 18, lastRow - 1).getValues(); // Column R
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
  const filteredLastRequests = filteredData.map((i) => lastRequests[i]);
  const filteredBusinessDays = filteredData.map((i) => businessDays[i]);

  return [
    filteredMachineNames,
    filteredPercentValues,
    filteredAmountValues,
    filteredCollectionPartners,
    filteredCollectionSchedules,
    filteredPartnerAddress,
    filteredLastRequests,
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
