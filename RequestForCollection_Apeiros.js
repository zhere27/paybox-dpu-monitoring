// Configuration constants
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
