// === Date Functions === //

/**
 * Gets today's and tomorrow's dates in various formats
 * @return {Object} Object containing todayDate, tomorrowDate, todayDay, and tomorrowDateString
 * @property {Date} todayDate - Today's date as a Date object
 * @property {Date} tomorrowDate - Tomorrow's date as a Date object
 * @property {string} todayDay - Today's day in "MMM d" format
 * @property {string} tomorrowDateString - Tomorrow's day in "MMM d" format
 */
function getTodayAndTomorrowDates() {
  const timeZone = Session.getScriptTimeZone();
  // const todayDate = new Date(2025,7,4); //August 4, 2025
  const todayDate = new Date();
  const tomorrowDate = new Date(todayDate);
  tomorrowDate.setDate(todayDate.getDate() + 1);

  return {
    todayDate,
    tomorrowDate,
    todayDay: Utilities.formatDate(todayDate, timeZone, "MMM d"),
    tomorrowDateString: Utilities.formatDate(tomorrowDate, timeZone, "MMM d"),
  };
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
    CustomLogger.logInfo(
      "Skipping execution on weekends.",
      PROJECT_NAME,
      "shouldSkipExecution()"
    );
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
function processCollections(
  forCollections,
  tomorrowDate,
  emailTo,
  emailCc,
  emailBcc,
  srvBank
) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo(
        "No collections to process.",
        PROJECT_NAME,
        "processCollections()"
      );
      return;
    }

    const collectionDate = new Date(tomorrowDate);

    if (srvBank === "Brinks via BPI") {
      // Sort by amount in descending order, handling NaN values by placing them at the end
      forCollections.sort(function (a, b) {
        const aAmount = parseFloat(a[2]);
        const bAmount = parseFloat(b[2]);
        if (isNaN(aAmount) && isNaN(bAmount)) return 0;
        if (isNaN(aAmount)) return 1;
        if (isNaN(bAmount)) return -1;
        return bAmount - aAmount;
      });
    } else {
      // sort forCollections by machineName
      forCollections.sort(function (a, b) {
        return a[0].localeCompare(b[0]);
      });

    }

    // Saturday collection limit is only applicable to Brinks via BPI
    if (collectionDate.getDay() === 6 && srvBank === "Brinks via BPI") {
      if (forCollections.length > 4) {
        const collectionsForMonday = forCollections.slice(4);
        const collectionDateForMonday = new Date();
        collectionDateForMonday.setDate(collectionDateForMonday.getDate() + 3);
        sendEmailCollection(
          collectionsForMonday,
          collectionDateForMonday,
          emailTo,
          emailCc,
          emailBcc,
          srvBank
        );
      }

      sendEmailCollection(
        forCollections.slice(0, 4),
        collectionDate,
        emailTo,
        emailCc,
        emailBcc,
        srvBank
      );
    } else {
      sendEmailCollection(
        forCollections,
        collectionDate,
        emailTo,
        emailCc,
        emailBcc,
        srvBank
      );
    }
  } catch (error) {
    CustomLogger.logError(
      `Error in processCollections(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "processCollections()"
    );
    throw error;
  }
}

// === Email Functions === //

/**
 * Sends a collection email
 * @param {Array} machineData - Array of machine data
 * @param {Date} collectionDate - Collection date
 * @param {string} emailTo - Email recipients (to)
 * @param {string} emailCc - Email recipients (cc)
 * @param {string} emailBcc - Email recipients (bcc)
 * @param {string} srvBank - Service bank
 */
function sendEmailCollection(
  machineData,
  collectionDate,
  emailTo,
  emailCc,
  emailBcc,
  srvBank
) {
  try {
    if (!machineData || machineData.length === 0) {
      CustomLogger.logInfo(
        "No machine data to send in email.",
        PROJECT_NAME,
        "sendEmailCollection()"
      );
      return;
    }

    const formattedDate = formatDate(collectionDate, "MMMM d, yyyy (EEEE)");
    const subject = `${srvBank} DPU Request - ${formattedDate}`;

    let body;
    if (
      machineData.some(
        (item) => item[0] === "PLDT BANTAY" || item[0] === "SMART VIGAN"
      )
    ) {
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
    CustomLogger.logInfo(
      `Collection email sent for ${machineData.length} machines.`,
      PROJECT_NAME,
      "sendEmailCollection()"
    );
  } catch (error) {
    CustomLogger.logError(
      `Error in sendEmailCollection: ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "sendEmailCollection()"
    );
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
function sendEmailCancellation(
  forCancellation,
  collectionDate,
  emailTo,
  emailCc,
  emailBcc,
  srvBank,
  isSaturday
) {
  try {
    // Cancel sending email if no records in the forCancellation
    if (!forCancellation || forCancellation.length === 0) {
      CustomLogger.logInfo(
        "No cancellations to send in email.",
        PROJECT_NAME,
        "sendEmailCancellation()"
      );
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
    CustomLogger.logInfo(
      `Cancellation email sent for ${forCancellation.length} machines.`,
      PROJECT_NAME,
      "sendEmailCancellation()"
    );
  } catch (error) {
    CustomLogger.logError(
      `Error in sendEmailCancellation: ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "sendEmailCancellation()"
    );
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
    CustomLogger.logInfo(
      `Reply sent to the thread with subject: "${subject}"`,
      PROJECT_NAME,
      "replyToExistingThread()"
    );
  } catch (error) {
    CustomLogger.logError(
      `Error in replyToExistingThread: ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "replyToExistingThread()"
    );
    throw error;
  }
}

// === Spreadsheet Functions === //

/**
 * Creates a hidden worksheet and adds data
 * @param {Array} forCollections - Array of collections
 * @param {string} srvBank - Service bank
 */
function createHiddenWorksheetAndAddData(forCollections, srvBank) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo(
        "No collection data to add to worksheet.",
        PROJECT_NAME,
        "createHiddenWorksheetAndAddData()"
      );
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

    // Adjust column width for better readability
    sheet.autoResizeColumns(1, 1);

    CustomLogger.logInfo(
      `Updated for collection worksheet "${sheetName}" with ${numRows} rows.`,
      PROJECT_NAME,
      "createHiddenWorksheetAndAddData()"
    );
    Logger.log(
      `Updated for collection worksheet "${sheetName}" with ${numRows} rows.`
    );
  } catch (error) {
    CustomLogger.logError(
      `Error in createHiddenWorksheetAndAddData(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "createHiddenWorksheetAndAddData()"
    );
    throw error;
  }
}

/**
 * Filters out past collection requests from the provided array, keeping only requests for the specified tomorrow date.
 * Also sorts the resulting array by machine name.
 *
 * @param {Array<Array>} forCollections - The array of collection request items, where each item is an array and the machine name is at index 0, and the collection info is at index 4.
 * @param {string} tomorrowDateString - The date string representing tomorrow, used to filter requests.
 * @returns {Array<Array>} The filtered and sorted array of collection requests.
 * @throws {Error} If an error occurs during filtering or sorting.
 */
function excludePastRequests(forCollections, tomorrowDateString) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo(
        "No collections to filter.",
        PROJECT_NAME,
        "excludePastRequests()"
      );
      return forCollections;
    }

    // using while forCollection and check forCollections[4] contains For Collection case incensitive
    forCollections = forCollections.filter((item) => {
      if (item[4].toLowerCase().includes("for collection")) {
        if (
          item[4].toLowerCase() !==
          "for collection on " + tomorrowDateString.toLowerCase()
        ) {
          CustomLogger.logInfo(
            `Excluded past request: ${item[0]} - ${item[4]}`,
            PROJECT_NAME,
            "excludePastRequests()"
          );
          return false;
        }
      }
      return true;
    });

    //Sort by machine name
    forCollections.sort((a, b) =>
      a[0].toLowerCase().localeCompare(b[0].toLowerCase())
    );

    CustomLogger.logInfo(
      `Done excluding past requested machines.`,
      PROJECT_NAME,
      "excludePastRequests()"
    );
    return forCollections;
  } catch (error) {
    CustomLogger.logError(
      `Error in excludePastRequests(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "excludePastRequests()"
    );
    throw error;
  }
}

/**
 * Filters out collections that have already been requested, based on machine names listed in a specific sheet.
 *
 * @param {Array<Array<any>>} forCollections - The list of collections to filter. Each collection is assumed to be an array, with the machine name in the first element.
 * @param {string} srvBank - The bank/service identifier used to determine the sheet name (e.g., "For Collection -{srvBank}").
 * @returns {Array<Array<any>>} The filtered list of collections, excluding those that have already been requested.
 * @throws {Error} If an error occurs during processing.
 */
function excludePreviouslyRequested(forCollections, srvBank) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo(
        "No collections to filter.",
        PROJECT_NAME,
        "excludePreviouslyRequested()"
      );
      return forCollections;
    }

    // Open the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `For Collection -${srvBank}`;
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      CustomLogger.logInfo(
        `Sheet "${sheetName}" not found. No previous collections to exclude.`,
        PROJECT_NAME,
        "excludePreviouslyRequested()"
      );
      return forCollections;
    }

    // Get previously requested machine names from Column A
    const previouslyRequested = getMachineNamesFromColumnA(sheet);
    if (!previouslyRequested || previouslyRequested.length === 0) {
      CustomLogger.logInfo(
        "No previously collected machines found.",
        PROJECT_NAME,
        "excludePreviouslyRequested()"
      );
      return forCollections;
    }

    // Use a Set for faster lookups
    const previouslyRequestedSet = new Set(previouslyRequested);

    // Filter out previously requested machines
    const filteredCollections = forCollections.filter((collection) => {
      const machineName = collection[0]; // Assuming machine names are in the first column
      var returnVal = !previouslyRequestedSet.has(machineName);
      return returnVal;
    });

    const removedCount = forCollections.length - filteredCollections.length;
    if (removedCount > 0) {
      CustomLogger.logInfo(
        `Removed ${removedCount} previously collected machines from collection list.`,
        PROJECT_NAME,
        "excludePreviouslyRequested()"
      );
    }

    return filteredCollections;
  } catch (error) {
    CustomLogger.logError(
      `Error in excludePreviouslyRequested(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "excludePreviouslyRequested()"
    );
    throw error;
  }
}

/**
 * Removes all occurrences of specified excluded items from the given collections array.
 *
 * @param {Array} forCollections - The array of collections to filter.
 * @param {Array} excludeExpiredForCollections - The array of items to exclude from forCollections.
 * @returns {Array} The filtered array with excluded items removed.
 */
function removeExcludedCollections(
  forCollections,
  excludeExpiredForCollections
) {
  excludeExpiredForCollections.forEach((excludeItem) => {
    let index;
    while ((index = forCollections.indexOf(excludeItem)) !== -1) {
      forCollections.splice(index, 1);
    }
  });
  return forCollections;
}

/**
 * Filters out machines from the forCollections array that have been recently collected,
 * based on data from the "For Collection -{srvBank}" sheet in the active spreadsheet.
 * Machines are excluded if their name appears in column A and column F has a TRUE value.
 *
 * @param {Array<Array<any>>} forCollections - Array of collections to filter. Each collection is expected to be an array where the first element is the machine name.
 * @param {string} srvBank - The service bank identifier used to select the appropriate sheet.
 * @returns {Array<Array<any>>} The filtered array of collections, excluding recently collected machines.
 * @throws {Error} If an error occurs during processing.
 */
function excludeRecentlyCollected(forCollections, srvBank) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo(
        "No collections to filter.",
        PROJECT_NAME,
        "excludeRecentlyCollected()"
      );
      return forCollections;
    }

    // Open the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `For Collection -${srvBank}`;
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      CustomLogger.logInfo(
        `Sheet "${sheetName}" not found for reference of recently collected machines.`,
        PROJECT_NAME,
        "excludeRecentlyCollected()"
      );
      return forCollections;
    }

    //Iterate through the sheet and exclude machine name in column A and column F that has value TRUE from the forCollections
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    data.forEach((row) => {
      const machineName = row[0];
      const isRecentlyCollected = row[5]; // index 5 for column 6

      if (isRecentlyCollected) {
        CustomLogger.logInfo(
          `Excluded recently collected machine: ${machineName}`,
          PROJECT_NAME,
          "excludeRecentlyCollected()"
        );
        forCollections = forCollections.filter(
          (collection) => collection[0] !== machineName
        );
      }
    });

    return forCollections;
  } catch (error) {
    CustomLogger.logError(
      `Error in excludeRecentlyCollected(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "excludeRecentlyCollected()"
    );
    throw error;
  }
}

/**
 * Fetches machine names from Column A of a given sheet.
 * @param {Sheet} sheet - The sheet to fetch machine names from
 * @return {Array<string>} List of machine names from Column A
 */
function getMachineNamesFromColumnA(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return []; // No data after the header row
    }

    // Get data from A2 to the last row
    const range = sheet.getRange(`A2:A${lastRow}`);
    const values = range.getValues();

    // Flatten the array and filter out empty cells
    return values.flat().filter((name) => name);
  } catch (error) {
    CustomLogger.logError(
      `Error in getMachineNamesFromColumnA(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "getMachineNamesFromColumnA()"
    );
    throw error;
  }
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
 * @param {string} lastRequest - Last request
 * @param {string} srvBank - Service bank
 * @return {boolean} True if the machine should be included, false otherwise
 */
function shouldIncludeForCollection(
  machineName,
  amountValue,
  translatedBusinessDays,
  tomorrowDate,
  tomorrowDateString,
  todayDate,
  lastRequest,
  srvBank
) {
  try {
    const collectionDay = dayMapping[tomorrowDate.getDay()];

    // Exclude no collection day
    if (!translatedBusinessDays.includes(collectionDay)) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString}, not a collection day.`,
        PROJECT_NAME,
        "shouldIncludeForCollection"
      );
      return false;
    }

    if (machineName === "PLDT OLONGAPO") {
      CustomLogger.logInfo(
        `Stop here.`,
        PROJECT_NAME,
        "shouldIncludeForCollection"
      );
    }

    const scheduleBased = [
      "PLDT ROBINSONS DUMAGUETE",
      "SMART SM BACOLOD 1",
      "SMART SM BACOLOD 2",
      "SMART SM BACOLOD 3",
      "PLDT ILIGAN",
      "SMART GAISANO MALL OZAMIZ",
    ];

    // Skip these stores
    const excludedStores = [
      "SMART LIMKETKAI CDO 2"
    ];

    if (excludedStores.includes(machineName)) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString}, part of the excluded stores.`,
        PROJECT_NAME,
        "shouldIncludeForCollection"
      );
      return false;
    }

    //For BPI Internal skip collection during weekends
    if (
      srvBank === "BPI Internal" &&
      (collectionDay === "Sat." || collectionDay === "Sun.")
    ) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString} as it is a weekend.`,
        PROJECT_NAME,
        "shouldIncludeForCollection"
      );
      return false;
    }

    // Check for special collection conditions in lastRequest
    if (lastRequest) {
      const specialConditions = [
        `for replacement of cassette on ${tomorrowDateString.toLowerCase()}`,
        `for collection on ${tomorrowDateString.toLowerCase()}`,
        `resume collection on ${tomorrowDateString.toLowerCase()}`,
        `for revisit on ${tomorrowDateString.toLowerCase()}`,
      ];

      if (
        specialConditions.some((condition) =>
          lastRequest.toLowerCase().includes(condition)
        )
      ) {
        return true;
      }
    }

    // Store requested for M.W.Sat. via email
    if (
      machineName === "PLDT ROBINSONS DUMAGUETE" &&
      (tomorrowDate.getDay() === dayIndex["M."] ||
        tomorrowDate.getDay() === dayIndex["W."] ||
        tomorrowDate.getDay() === dayIndex["Sat."])
    ) {
      return true;
    }

    // Store requested for M.W.Sat. via email
    if (
      (machineName === "SMART SM BACOLOD 1" ||
        machineName === "SMART SM BACOLOD 2" ||
        machineName === "SMART SM BACOLOD 3") &&
      (tomorrowDate.getDay() === dayIndex["T."] ||
        tomorrowDate.getDay() === dayIndex["Sat."])
    ) {
      return true;
    }

    if (
      machineName === "PLDT ILIGAN" &&
      (tomorrowDate.getDay() === dayIndex["M."] ||
        tomorrowDate.getDay() === dayIndex["W."] ||
        tomorrowDate.getDay() === dayIndex["F."])
    ) {
      return true;
    }

    if (
      machineName === "SMART GAISANO MALL OZAMIZ" &&
      tomorrowDate.getDay() === dayIndex["F."]
    ) {
      return true;
    }

    if (
      !scheduleBased.includes(machineName) &&
      amountValue >= amountThresholds[collectionDay]
    ) {
      return true;
    }

    // Check payday ranges
    const isPayday = paydayRanges.some(
      (range) =>
        todayDate.getDate() >= range.start &&
        todayDate.getDate() <= range.end &&
        amountValue >= paydayAmount
    );

    if (isPayday) {
      return true;
    }

    // Check due date cutoffs
    const isDueDate = dueDateCutoffs.some(
      (range) =>
        todayDate.getDate() >= range.start &&
        todayDate.getDate() <= range.end &&
        amountValue >= dueDateCutoffsAmount
    );

    if (isDueDate) {
      return true;
    }
  } catch (error) {
    CustomLogger.logError(
      `Error in shouldIncludeForCollection(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "shouldIncludeForCollection()"
    );
    return false;
  }
}

/**
 * Checks if a machine should be excluded from collection
 * @param {string} lastRequest - Last request
 * @param {string} todayDay - Today's day
 * @return {boolean} True if the machine should be excluded, false otherwise
 */
function shouldExcludeFromCollection(
  lastRequest,
  todayDay,
  machineName = null
) {
  try {
    if (!lastRequest) {
      return false;
    }

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
      todayDay,
    ];

    var excluded = exclusionReasons.some((reason) =>
      lastRequest.toLowerCase().includes(reason)
    );
    if (machineName !== null && excluded) {
      CustomLogger.logInfo(
        `Machine Name: ${machineName} was excluded due to last request: ${lastRequest} `,
        PROJECT_NAME,
        "shouldExcludeFromCollection()"
      );
    }

    return excluded;
  } catch (error) {
    CustomLogger.logError(
      `Error in shouldExcludeFromCollection(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "shouldExcludeFromCollection()"
    );
    return false;
  }
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
    const validHolidayTypes = [
      "Regular Holiday",
      "Special Non-working Holiday",
    ];

    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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
      const holidayType = row[2];

      if (
        holidayDate instanceof Date &&
        holidayType &&
        validHolidayTypes.includes(holidayType)
      ) {
        if (holidayDate.getTime() === normalizedTomorrow.getTime()) {
          CustomLogger.logInfo(
            `Tomorrow (${formatDate(
              tomorrow,
              "MMM d, yyyy"
            )}) is a holiday: ${holidayType}`,
            PROJECT_NAME,
            "isTomorrowHoliday()"
          );
          return true;
        }
      }
    }

    return false;
  } catch (error) {
    CustomLogger.logError(
      `Error in isTomorrowHoliday(): ${error.message}\nStack: ${error.stack}`,
      PROJECT_NAME,
      "isTomorrowHoliday()"
    );
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
    CustomLogger.logInfo(
      "Function execution skipped due to date restriction.",
      PROJECT_NAME,
      "skipFunction()"
    );
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
