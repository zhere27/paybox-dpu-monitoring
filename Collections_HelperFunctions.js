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

/**
 * Amount threshold for due date collections
 * @type {number}
 */
const dueDateCutoffsAmount = 290000;

/**
 * Email signature for all emails
 * @type {string}
 */
const emailSignature = `<div><div dir="ltr" class="gmail_signature"><div dir="ltr"><span><div dir="ltr" style="margin-left:0pt" align="left">Best Regards,</div><div dir="ltr" style="margin-left:0pt" align="left"><br><table style="border:none;border-collapse:collapse"><colgroup><col width="44"><col width="249"><col width="100"><col width="7"></colgroup><tbody><tr style="height:66.75pt"><td rowspan="3" style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;text-align:justify;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:47px;height:217px"><img src="https://lh3.googleusercontent.com/JNGaTKS3JTfQlHCnlwFXo14_knhu4v_WhlZCWOIFJfRPuUKjMMHWuj82yUQ0uUOxv9XNk1Nooae__kDJ1wS0st_Xe3SZvDdl3dkVSpX24SCtgfIt7ZfeTfIR8S93ndcLMdQSgm9Xyq1rykUOGv1sLo0" width="47" height="235.00000000000003" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td><td style="vertical-align:middle;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:12pt;font-family:Arial,sans-serif;background-color:transparent;font-weight:700;vertical-align:baseline">Support</span></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:12pt;font-family:Arial,sans-serif;background-color:transparent;font-weight:700;vertical-align:baseline">Paybox Operations</span></p><br></td><td colspan="2" rowspan="3" style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;margin-right:0.75pt;text-align:justify;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:118px;height:240px"><img src="https://lh6.googleusercontent.com/3QMqLONmIp2CUDF9DoGaWmEhai3RSB6fjgqFxwhxwcQoX58wlvnAVNMscDRgfOK-xv4S2bllMTzrKQSuvgqAi68syHzvqbNJziibdwTfx7A1pSWelqdkffPtJ9n6WC3JJEEcSqNYXrBthmb8cxIz5Dg" width="118" height="240" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td></tr><tr style="height:84pt"><td style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:14px;height:10px"><img src="https://lh4.googleusercontent.com/DM0kOfMz47NSJfROHj6mLS0ypJih5nHezY-SaBODOfd_oXMKxDagoXJG1WmGYaCgt0g_PLa4KQY1Btkuih7F2409F3-gjDxV8UGVeL_6bKF4l3Aze7QG33MalyKV0NmslPNz5aK3Fp8a8LX_8abfJoo" width="14" height="10" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><a href="mailto:support@paybox.ph" target="_blank"><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(17,85,204);vertical-align:baseline">support@paybox.ph</span></a></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:16px"><img src="https://lh6.googleusercontent.com/RFRWIimtGNr-ErnlmxYSpbrwRjPjBvZXbY5BdP3LD2ykCtMVGMYodZoIc-B7xHWXI3wYHcAr8FxK6d3L4hk12mbH-dTEZ1pU6pugBvzZeqvu2uLo_4BPb_zlAlT8ve3P2GD0CifMeZ_dX_qdWhPns5g" width="15" height="16" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline">+62 (2) 8835-9697</span></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:13px"><img src="https://lh4.googleusercontent.com/VuSTEwKlvSLg3DHUziWOwLBgmj1ctniwh9f7tsnECrzxxdMy1CRKnKgJaCi2m8xIRfrsdtsihCEvexpSHnDykgsRZ1WeMuxwHKQpbS-VUfBwoM2wC3oZU9i7B8vgAWf4JZPzq03GND-SOi9bD0RIIJM" width="15" height="13" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline">+62 (2) 8805-1066 (loc) 115&nbsp;</span></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:15px"><img src="https://lh3.googleusercontent.com/vLSCYEGN7oLtmhaSdBjzq9psnzQaLlY-fa7QgcbvWWjPvwrwCDr2qX7nHFglSmQHxwPPf0DmH17j6TgttmB54ke2L4x7BJYp3DiNISF5do3G2gsBdS1v9_KchSJAc-K_dh6FtGUxJtHmrOtZPDPZEnw" width="15" height="15" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><a href="https://www.multisyscorp.com/" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://www.multisyscorp.com/&amp;source=gmail&amp;ust=1731657202461000&amp;usg=AOvVaw0G3eJRh1oAS1IGmlpIrvQt"><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(17,85,204);vertical-align:baseline">www.multisyscorp.com</span></a></p></td></tr><tr style="height:30pt"><td style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;margin-top:0pt;margin-bottom:0pt"><span style="font-size:9pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:121px;height:27px"><img src="https://lh4.googleusercontent.com/m4vMFXp2L8_GL568U8By8vFWFacVg-4s_RAjIFlQMJJvTQfgik58VwVIJY8g1zW8g9n-__YYXwDYO9kZzDidrTIWrRieExlSnwvtRuEqikp5XgZ3G9xAyIXd8eFN-k42XJRXhK0APe5u1FD8si56y44" width="121" height="27" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td></tr></tbody></table></div><br><p dir="ltr" style="line-height:1.38;margin-top:8pt;margin-bottom:0pt"><span style="font-size:11pt;color:rgb(0,0,0);background-color:transparent;vertical-align:baseline;white-space:pre-wrap"><font face="times new roman, serif">DISCLAIMER:</font></span></p><p dir="ltr" style="line-height:1.38;margin-top:8pt;margin-bottom:0pt"><span style="font-size:11pt;color:rgb(0,0,0);background-color:transparent;vertical-align:baseline;white-space:pre-wrap"><font face="times new roman, serif">The content of this email is confidential and intended solely for the use of the individual or entity to whom it is addressed. If you received this email in error, please notify the sender or system manager. It is strictly forbidden to disclose, copy or distribute any part of this message.</font></span></p></span></div></div></div>`;

// === Date Functions === //

/**
 * Gets today's and tomorrow's dates in various formats
 * @return {Object} Object containing todayDate, tomorrowDate, todayDay, and tomorrowDateString
 */
function getTodayAndTomorrowDates() {
  try {
    const todayDate = new Date();
    const todayDay = Utilities.formatDate(todayDate, Session.getScriptTimeZone(), "MMM d");

    const tomorrowDate = new Date(todayDate);
    tomorrowDate.setDate(todayDate.getDate() + 1);
    const tomorrowDateString = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMM d");

    return { todayDate, tomorrowDate, todayDay, tomorrowDateString };
  } catch (error) {
    CustomLogger.logError(`Error in getTodayAndTomorrowDates(): ${error.message}`, PROJECT_NAME, 'getTodayAndTomorrowDates()');
    throw error;
  }
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

  if (day === 0 || day === 6) { // 0-Sunday, 6-Saturday
    CustomLogger.logInfo("Skipping execution on weekends.", PROJECT_NAME, 'shouldSkipExecution()');
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
function processCollections(forCollections, tomorrowDate, emailTo, emailCc, emailBcc, srvBank) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo("No collections to process.", PROJECT_NAME, 'processCollections()');
      return;
    }

    const collectionDate = new Date(tomorrowDate);
    forCollections.sort();

    // Saturday collection limit is only applicable to Brinks via BPI
    if (collectionDate.getDay() === 6 && srvBank === 'Brinks via BPI') {
      if (forCollections.length > 4) {
        const collectionsForMonday = forCollections.slice(4);
        const collectionDateForMonday = new Date();
        collectionDateForMonday.setDate(collectionDateForMonday.getDate() + 3);
        sendEmailCollection(collectionsForMonday, collectionDateForMonday, emailTo, emailCc, emailBcc, srvBank);
      }

      sendEmailCollection(forCollections.slice(0, 4), collectionDate, emailTo, emailCc, emailBcc, srvBank);
    } else {
      sendEmailCollection(forCollections, collectionDate, emailTo, emailCc, emailBcc, srvBank);
    }
  } catch (error) {
    CustomLogger.logError(`Error in processCollections(): ${error.message}`, PROJECT_NAME, 'processCollections()');
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
function sendEmailCollection(machineData, collectionDate, emailTo, emailCc, emailBcc, srvBank) {
  try {
    if (!machineData || machineData.length === 0) {
      CustomLogger.logInfo("No machine data to send in email.", PROJECT_NAME, 'sendEmailCollection()');
      return;
    }

    const formattedDate = formatDate(collectionDate, "MMMM d, yyyy (EEEE)");
    const subject = `${srvBank} DPU Request - ${formattedDate}`;

    let body;
    if (machineData.some(item => item[0] === 'PLDT BANTAY' || item[0] === 'SMART VIGAN')) {
      body = `
        Hi All,<br><br>
        Good day! Please schedule <b>collection</b> on ${formattedDate} for the following stores:<br><br>
        ${machineData.map(row => row[0]).join('<br>')}<br><br>
        For <b>PLDT BANTAY and SMART VIGAN</b>, kindly collect it on the same day.<br><br>
        *** Please acknowledge this email. ****<br><br>${emailSignature}
      `;
    } else {
      body = `
        Hi All,<br><br>
        Good day! Please schedule <b>collection</b> on ${formattedDate} for the following stores:<br><br>
        ${machineData.map(row => row[0]).join('<br>')}<br><br>
        *** Please acknowledge this email. ****<br><br>${emailSignature}
      `;
    }

    GmailApp.sendEmail(emailTo, subject, '', {
      cc: emailCc,
      bcc: emailBcc,
      htmlBody: body,
      from: "support@paybox.ph"
    });
    CustomLogger.logInfo(`Collection email sent for ${machineData.length} machines.`, PROJECT_NAME, 'sendEmailCollection()');
  } catch (error) {
    CustomLogger.logError(`Error in sendEmailCollection: ${error.message}`, PROJECT_NAME, 'sendEmailCollection()');
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
      CustomLogger.logInfo("No cancellations to send in email.", PROJECT_NAME, 'sendEmailCancellation()');
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
      ${forCancellation.map(store => `
        <tr>
          <td>${store.name}</td>
          <td>&nbsp;</td>  
          <td>${store.address}</td>
        </tr>
      `).join('')}
      </table>
      <br><br>
      *** Please acknowledge this email. ****<br><br>${emailSignature}
    `;

    GmailApp.sendEmail(emailTo, subject, '', {
      cc: emailCc,
      bcc: emailBcc,
      htmlBody: body,
      from: "support@paybox.ph"
    });
    CustomLogger.logInfo(`Cancellation email sent for ${forCancellation.length} machines.`, PROJECT_NAME, 'sendEmailCancellation()');
  } catch (error) {
    CustomLogger.logError(`Error in sendEmailCancellation: ${error.message}`, PROJECT_NAME, 'sendEmailCancellation()');
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
      from: "support@paybox.ph"
    });
    CustomLogger.logInfo(`Reply sent to the thread with subject: "${subject}"`, PROJECT_NAME, 'replyToExistingThread()');
  } catch (error) {
    CustomLogger.logError(`Error in replyToExistingThread: ${error.message}`, PROJECT_NAME, 'replyToExistingThread()');
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
      CustomLogger.logInfo("No collection data to add to worksheet.", PROJECT_NAME, 'createHiddenWorksheetAndAddData()');
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

    CustomLogger.logInfo(`Updated for collection worksheet "${sheetName}" with ${numRows} rows.`, PROJECT_NAME, 'createHiddenWorksheetAndAddData()');
    Logger.log(`Updated for collection worksheet "${sheetName}" with ${numRows} rows.`);
  } catch (error) {
    CustomLogger.logError(`Error in createHiddenWorksheetAndAddData(): ${error.message}`, PROJECT_NAME, 'createHiddenWorksheetAndAddData()');
    throw error;
  }
}

function excludePastRequests(forCollections, tomorrowDateString) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo("No collections to filter.", PROJECT_NAME, 'excludePastRequests()');
      return forCollections;
    }

    // using while forCollection and check forCollections[4] contains For Collection case incensitive
    forCollections = forCollections.filter(item => {
      if (item[4].toLowerCase().includes('for collection')) {
        if (item[4].toLowerCase() !== 'for collection on ' + tomorrowDateString.toLowerCase()) {
          CustomLogger.logInfo(`Excluded past request: ${item[0]} - ${item[4]}`, PROJECT_NAME, 'excludePastRequests()');
          return false;
        }
      }
      return true;
    });

    //Sort by machine name
    forCollections.sort((a, b) => a[0].toLowerCase().localeCompare(b[0].toLowerCase()));

    CustomLogger.logInfo(`Done excluding past requested machines.`, PROJECT_NAME, 'excludePastRequests()');
    return forCollections;

  } catch (error) {
    CustomLogger.logError(`Error in excludePastRequests(): ${error.message}`, PROJECT_NAME, 'excludePastRequests()');
    throw error;
  }
}

function excludePreviouslyRequested(forCollections, srvBank) {
  try {
    if (!forCollections || forCollections.length === 0) {
      CustomLogger.logInfo("No collections to filter.", PROJECT_NAME, 'excludePreviouslyRequested()');
      return forCollections;
    }

    // Open the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `For Collection -${srvBank}`;
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      CustomLogger.logInfo(`Sheet "${sheetName}" not found. No previous collections to exclude.`, PROJECT_NAME, 'excludePreviouslyRequested()');
      return forCollections;
    }

    // Get previously requested machine names from Column A
    const previouslyRequested = getMachineNamesFromColumnA(sheet);
    if (!previouslyRequested || previouslyRequested.length === 0) {
      CustomLogger.logInfo("No previously collected machines found.", PROJECT_NAME, 'excludePreviouslyRequested()');
      return forCollections;
    }

    // Use a Set for faster lookups
    const previouslyRequestedSet = new Set(previouslyRequested);

    // Filter out previously requested machines
    const filteredCollections = forCollections.filter(collection => {
      const machineName = collection[0]; // Assuming machine names are in the first column
      var returnVal = !previouslyRequestedSet.has(machineName)
      return returnVal;
    });

    const removedCount = forCollections.length - filteredCollections.length;
    if (removedCount > 0) {
      CustomLogger.logInfo(`Removed ${removedCount} previously collected machines from collection list.`, PROJECT_NAME, 'excludePreviouslyRequested()');
    }

    return filteredCollections;
  } catch (error) {
    CustomLogger.logError(`Error in excludePreviouslyRequested(): ${error.message}`, PROJECT_NAME, 'excludePreviouslyRequested()');
    throw error;
  }
}

function removeExcludedCollections(forCollections, excludeExpiredForCollections) {
  excludeExpiredForCollections.forEach(excludeItem => {
    let index;
    while ((index = forCollections.indexOf(excludeItem)) !== -1) {
      forCollections.splice(index, 1);
    }
  });
  return forCollections;
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
    return values.flat().filter(name => name);
  } catch (error) {
    CustomLogger.logError(`Error in getMachineNamesFromColumnA(): ${error.message}`, PROJECT_NAME, 'getMachineNamesFromColumnA()');
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
function shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank) {
  try {
    const collectionDay = dayMapping[tomorrowDate.getDay()];

    // Exclude no collection day
    if (!translatedBusinessDays.includes(collectionDay)) {
      CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, not a collection day.`, PROJECT_NAME, 'shouldIncludeForCollection');
      return false;
    }

    const scheduleBased = [
      'PLDT ROBINSONS DUMAGUETE',
      'SMART SM BACOLOD 1',
      'SMART SM BACOLOD 2',
      'SMART SM BACOLOD 3',
      'PLDT ILIGAN',
      'SMART GAISANO MALL OZAMIZ'
    ];

    // Skip these stores
    const excludedStores = [
      'SMART LIMKETKAI CDO 2',
      // 'SMART SM DAET',
      // 'SMART SOLANO',
      // 'SMART ROBINSONS BACOLOD 1'
    ];

    if (excludedStores.includes(machineName)) {
      CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, part of the excluded stores.`, PROJECT_NAME, 'shouldIncludeForCollection');
      return false;
    }

    //For BPI Internal skip collection during weekends
    if (srvBank === 'BPI Internal' && (collectionDay === 'Sat.' || collectionDay === 'Sun.')) {
      CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString} as it is a weekend.`, PROJECT_NAME, 'shouldIncludeForCollection');
      return false;
    }

    // Check for special collection conditions in lastRequest
    if (lastRequest) {
      const specialConditions = [
        `for replacement of cassette on ${tomorrowDateString.toLowerCase()}`,
        `for collection on ${tomorrowDateString.toLowerCase()}`,
        `resume collection on ${tomorrowDateString.toLowerCase()}`
      ];

      if (specialConditions.some(condition => lastRequest.toLowerCase().includes(condition))) {
        return true;
      }
    }

    // Store requested for M.W.Sat. via email
    if (machineName === 'PLDT ROBINSONS DUMAGUETE' &&
      (tomorrowDate.getDay() === dayIndex['M.'] ||
        tomorrowDate.getDay() === dayIndex['W.'] ||
        tomorrowDate.getDay() === dayIndex['Sat.'])) {
      return true;
    }

    // Store requested for M.W.Sat. via email
    if ((machineName === 'SMART SM BACOLOD 1' ||
      machineName === 'SMART SM BACOLOD 2' ||
      machineName === 'SMART SM BACOLOD 3') &&
      (tomorrowDate.getDay() === dayIndex['T.'] ||
        tomorrowDate.getDay() === dayIndex['Sat.'])) {
      return true;
    }

    if (machineName === 'PLDT ILIGAN' &&
      (tomorrowDate.getDay() === dayIndex['M.'] ||
        tomorrowDate.getDay() === dayIndex['W.'] ||
        tomorrowDate.getDay() === dayIndex['F.'])) {
      return true;
    }

    if (machineName === 'SMART GAISANO MALL OZAMIZ' &&
      tomorrowDate.getDay() === dayIndex['F.']) {
      return true;
    }

    if (!scheduleBased.includes(machineName) &&
      amountValue >= amountThresholds[collectionDay]) {
      return true;
    }

    // Check payday ranges
    const isPayday = paydayRanges.some(range =>
      todayDate.getDate() >= range.start &&
      todayDate.getDate() <= range.end &&
      amountValue >= paydayAmount
    );

    if (isPayday) {
      return true;
    }

    // Check due date cutoffs
    const isDueDate = dueDateCutoffs.some(range =>
      todayDate.getDate() >= range.start &&
      todayDate.getDate() <= range.end &&
      amountValue >= dueDateCutoffsAmount
    );

    if (isDueDate) {
      return true;
    }

  } catch (error) {
    CustomLogger.logError(`Error in shouldIncludeForCollection(): ${error.message}`, PROJECT_NAME, 'shouldIncludeForCollection()');
    return false;
  }
}

/**
 * Checks if a machine should be excluded from collection
 * @param {string} lastRequest - Last request
 * @param {string} todayDay - Today's day
 * @return {boolean} True if the machine should be excluded, false otherwise
 */
function shouldExcludeFromCollection(lastRequest, todayDay, machineName = null) {
  try {
    if (!lastRequest) {
      return false;
    }

    const exclusionReasons = [
      'waiting for collection items',
      'waiting for bank updates',
      'did not reset',
      'cassette removed',
      'manually collected',
      'for repair',
      'already collected',
      todayDay
    ];

    var excluded = exclusionReasons.some(reason => lastRequest.toLowerCase().includes(reason));
    if (machineName !== null && excluded) {
      CustomLogger.logInfo(`Machine Name: ${machineName} was excluded due to last request: ${lastRequest} `, PROJECT_NAME, 'shouldExcludeFromCollection()');
    }

    return excluded;
  } catch (error) {
    CustomLogger.logError(`Error in shouldExcludeFromCollection(): ${error.message}`, PROJECT_NAME, 'shouldExcludeFromCollection()');
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
      const holidayType = row[2];

      if (holidayDate instanceof Date && holidayType && validHolidayTypes.includes(holidayType)) {
        if (holidayDate.getTime() === normalizedTomorrow.getTime()) {
          CustomLogger.logInfo(`Tomorrow (${formatDate(tomorrow, "MMM d, yyyy")}) is a holiday: ${holidayType}`, PROJECT_NAME, 'isTomorrowHoliday()');
          return true;
        }
      }
    }

    return false;
  } catch (error) {
    CustomLogger.logError(`Error in isTomorrowHoliday(): ${error.message}`, PROJECT_NAME, 'isTomorrowHoliday()');
    return false;
  }
}

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
  if (typeof inputString !== 'string') {
    throw new Error('Input must be a string');
  }

  // Trim leading and trailing spaces and replace multiple spaces with a single space
  return inputString.trim().replace(/\s+/g, ' ');
}

function skipFunction() {
  // Define the UTC date range
  const startDate = new Date(2025, 4, 30, 0, 0, 0); // May 30, 2025, 00:00:00 UTC
  const endDate = new Date(2025, 5, 27, 23, 59, 0);  // June 27, 2025, 23:59:00 UTC

  // Get the current UTC date and time
  const now = new Date();
  const nowUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(),
    now.getUTCHours(), now.getUTCMinutes(), now.getUTCSeconds()));

  // Check if the current date is within the range
  if (nowUTC >= startDate && nowUTC <= endDate) {
    CustomLogger.logInfo('Function execution skipped due to date restriction.', PROJECT_NAME, 'skipFunction()');
    return; // Exit the function
  }
}
