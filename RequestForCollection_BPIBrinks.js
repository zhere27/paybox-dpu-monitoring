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
