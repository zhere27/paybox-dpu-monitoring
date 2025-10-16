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
      CustomLogger.logInfo("No machine data to send in email.", CONFIG.APP.NAME, "sendEmailCollection()");
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
    CustomLogger.logInfo(`Collection email sent for ${machineData.length} machines.`, CONFIG.APP.NAME, "sendEmailCollection()");
  } catch (error) {
    CustomLogger.logError(`Error in sendEmailCollection: ${error.message}\nStack: ${error.stack}`, CONFIG.APP.NAME, "sendEmailCollection()");
    throw error;
  }
}

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
      CustomLogger.logInfo("No collections to process.", CONFIG.APP.NAME, "processCollections()");
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
    CustomLogger.logError(`Error in processCollections(): ${error.message}\nStack: ${error.stack}`, CONFIG.APP.NAME, "processCollections()");
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
      CustomLogger.logInfo("No cancellations to send in email.", CONFIG.APP.NAME, "sendEmailCancellation()");
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
    CustomLogger.logInfo(`Cancellation email sent for ${forCancellation.length} machines.`, CONFIG.APP.NAME, "sendEmailCancellation()");
  } catch (error) {
    CustomLogger.logError(`Error in sendEmailCancellation: ${error.message}\nStack: ${error.stack}`, CONFIG.APP.NAME, "sendEmailCancellation()");
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
    CustomLogger.logInfo(`Reply sent to the thread with subject: "${subject}"`, CONFIG.APP.NAME, "replyToExistingThread()");
  } catch (error) {
    CustomLogger.logError(`Error in replyToExistingThread: ${error.message}\nStack: ${error.stack}`, CONFIG.APP.NAME, "replyToExistingThread()");
    throw error;
  }
}