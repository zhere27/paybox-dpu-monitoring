// === Date Functions === //

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

}