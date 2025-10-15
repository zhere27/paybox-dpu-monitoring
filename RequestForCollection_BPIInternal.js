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

