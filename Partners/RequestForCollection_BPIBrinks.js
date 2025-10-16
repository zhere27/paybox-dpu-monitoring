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
    CustomLogger.logInfo(`Running BPI Brinks Collections Logic in [${CONFIG.BPI_BRINKS.ENVIRONMENT}] mode...`, CONFIG.APP.NAME, "bpiBrinkCollectionsLogic()");

    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    const { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = dateInfo;

    const emailRecipients = getEmailRecipientsBPI();
    const subject = generateEmailSubjectBPI(tomorrowDate);

    const eligibleCollections = getEligibleCollectionsBPI(todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, subject);

    // Process collections if any are eligible (production only)
    if (CONFIG.BPI_INTERNAL.ENVIRONMENT === 'production' && eligibleCollections.length > 0) {
      processEligibleCollectionsBPIBrinks(eligibleCollections, tomorrowDate, emailRecipients);
    } else if (eligibleCollections.length === 0) {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.', CONFIG.APP.NAME, 'bpiBrinkCollectionsLogic');
    } else {
      CustomLogger.logInfo(`Testing mode: ${eligibleCollections.length} collections identified but not sent.`, CONFIG.APP.NAME, 'bpiBrinkCollectionsLogic');
      console.log(JSON.stringify(collections, null, 2));
    }
    //Send Logs to Admin
    EmailSender.sendLogs(CONFIG.APP.ADMIN.email, CONFIG.APP.NAME);
  } catch (error) {
    const errorMessage = `An unexpected error occurred in BPI Brinks Logic: ${error.message} \nStack: ${error.stack}`;
    CustomLogger.logError(errorMessage, CONFIG.APP.NAME, "bpiBrinkCollectionsLogic()");
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
function getEligibleCollectionsBPI(todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, subject) {
  const machineData = getMachineDataByPartner(CONFIG.BPI_BRINKS.SERVICE_BANK);
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, lastAddress, lastRemarks, businessDays] = machineData;

  const forCollectionData = getForCollections(CONFIG.BPI_BRINKS.SERVICE_BANK);
  const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(CONFIG.BPI_BRINKS.SERVICE_BANK);
  const eligibleCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machine = {
      name: machineNameArr[0],
      amount: amountValues[i][0],
      lastRemark: normalizeSpaces(lastRemarks[i][0]),
      businessDay: businessDays[i][0],
      collectionSchedule: collectionSchedules[i][0]
    };

    if (excludeMachineBPI(machine, forCollectionData, previouslyRequestedMachines, todayDateString, tomorrowDate, tomorrowDateString, isTomorrowHoliday, CONFIG.BPI_BRINKS.SERVICE_BANK)) {
      return; // Skip this machine
    }

    const translatedBusinessDays = translateDaysToAbbreviation(machine.businessDay.trim());
    if (shouldIncludeForCollection(machine.name, machine.amount, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, machine.lastRemark, CONFIG.BPI_BRINKS.SERVICE_BANK)) {
      const formattedName = formatMachineNameWithRemarkBPI(machine.name, machine.lastRemark, tomorrowDateString);
      eligibleCollections.push([formattedName, machine.amount, CONFIG.BPI_BRINKS.SERVICE_BANK, subject, machine.lastRemark]);
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
function processEligibleCollectionsBPIBrinks(collections, tomorrowDate, recipients) {
  if (collections.length === 0) {
    CustomLogger.logInfo("No eligible stores for BPI Brinks collection tomorrow.", CONFIG.APP.NAME, "processEligibleCollectionsBPI()");
    return;
  }

  if (CONFIG.BPI_BRINKS.ENVIRONMENT === "production") {
    CustomLogger.logInfo(`Processing ${collections.length} collections for production.`, CONFIG.APP.NAME, "processEligibleCollectionsBPI()");
    createHiddenWorksheetAndAddData(collections, CONFIG.BPI_BRINKS.SERVICE_BANK);
    processCollectionsAndSendEmail(collections, tomorrowDate, recipients.to, recipients.cc, recipients.bcc, CONFIG.BPI_BRINKS.SERVICE_BANK);
  } else if (eligibleCollections.length === 0) {
    CustomLogger.logInfo('No eligible stores for collection tomorrow.', CONFIG.APP.NAME, 'bpiCollectionsLogic');
  } else {
    CustomLogger.logInfo(`[TESTING MODE] Found ${collections.length} eligible collections. No emails will be sent.`, CONFIG.APP.NAME, "processEligibleCollectionsBPI()");
    // Optional: Log the collections for verification during testing
    console.log(JSON.stringify(collections, null, 2));
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
function excludeMachineBPI(machine, forCollectionData, previouslyRequestedMachines, todayDateString, tomorrowDate, tomorrowDateString, isTomorrowHoliday, srvBank) {
  const collectionDay = dayMapping[tomorrowDate.getDay()];

  if (hasSpecialCollectionConditions(machine.lastRemark, tomorrowDateString)) {
    return false;
  }

  // Check: Exclude if amount did not meet
  if (!meetsAmountThreshold(machine.amount, collectionDay, srvBank)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the amount did not meet amount threshold.`, CONFIG.APP.NAME, 'excludeMachineBPI');
    return true;
  }

  if (excludeAlreadyCollected(machine.name, forCollectionData)) {
    return true;
  }

  // Check 2: No-Holiday schedule and tomorrow is holiday
  if (excludeHolidayCollection(machine.name, machine.collectionSchedule, isTomorrowHoliday)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, bacause store is closed during holidays.`, CONFIG.APP.NAME, 'excludeMachineBPI');
    return true;
  }

  if (excluceBasedOnRemarks(machine.lastRemark, todayDateString, machine.name)) {
    return true;
  }

  if (excludeRequestedYesterday(previouslyRequestedMachines, machine.name, CONFIG.BPI_BRINKS.SERVICE_BANK)) {
    return true;
  }

  // Check 4: If not yet scheduled for collection
  if (excludeNotYetScheduled(machine.name, tomorrowDateString, machine.lastRemark)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store is not yet scheduled for collection.`, CONFIG.APP.NAME, 'excludeMachineBPI');
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
  return CONFIG.BPI_BRINKS.EMAIL_RECIPIENTS[CONFIG.BPI_BRINKS.ENVIRONMENT] || CONFIG.BPI_BRINKS.EMAIL_RECIPIENTS.testing;
}

/**
 * Generates the email subject for the collection request.
 * @param {Date} tomorrowDate The Date object for tomorrow.
 * @returns {string} The formatted email subject.
 */
function generateEmailSubjectBPI(tomorrowDate) {
  const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
  return `${CONFIG.BPI_BRINKS.SERVICE_BANK} DPU Request - ${tomorrowDateFormatted}`;
}


// function bpiBrinkCollectionsLogic() {
//   CustomLogger.logInfo("Running bpiBrinkCollectionsLogic...", CONFIG.APP.NAME, "bpiBrinkCollectionsLogic()");
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
//     CustomLogger.logInfo("No eligible stores for collection tomorrow.", CONFIG.APP.NAME, "bpiBrinkCollectionsLogic()");
//   }
// }
