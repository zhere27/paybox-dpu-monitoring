function apeirosCollectionsLogic() {
  CustomLogger.logInfo('Running Apeiros Collection Logic...', CONFIG.APP.NAME, 'apeirosCollectionsLogic');

  try {
    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = dateInfo;

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipients(CONFIG.APEIROS.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubject(tomorrowDate, CONFIG.APEIROS.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollections(todayDate, tomorrowDate, todayDateString, tomorrowDateString, subject);

    // Process collections if any are eligible (production only)
    if (CONFIG.BPI_INTERNAL.ENVIRONMENT === 'production' && eligibleCollections.length > 0) {
      processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients);
    } else if (eligibleCollections.length === 0) {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.', CONFIG.APP.NAME, 'apeirosCollectionsLogic');
    } else {
      CustomLogger.logInfo(`Testing mode: ${eligibleCollections.length} collections identified but not sent.`, CONFIG.APP.NAME, 'apeirosCollectionsLogic');
      console.log(JSON.stringify(collections, null, 2));
    }
    //Send Logs to Admin
    EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);
  } catch (error) {
    CustomLogger.logError(`Error in apeirosCollectionsLogic: ${error}`, CONFIG.APP.NAME, 'apeirosCollectionsLogic');
    throw error;
  }
}

/**
 * Get email recipients based on environment
 */
function getEmailRecipients(environment) {
  return CONFIG.APEIROS.EMAIL_RECIPIENTS[environment] || CONFIG.APEIROS.EMAIL_RECIPIENTS.testing;
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
  const srvBank = CONFIG.APEIROS.SERVICE_BANK;

  // Fetch data
  const machineData = getMachineDataByPartner(srvBank, tomorrowDate);
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
    if (excludeMachineApeiros(machineName, amountValue, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, srvBank)) {
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
function excludeMachineApeiros(machineName, amountValue, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, srvBank) {
  // Check 1: Already collected
  if (excludeAlreadyCollected(machineName, forCollectionData)) {
    return true;
  }

  // Check 2: Excluded based on remarks
  if (excluceBasedOnRemarks(lastRemark, todayDateString, machineName)) {
    return true;
  }

  // Check 3: Requested yesterday
  if (excludeRequestedYesterday(previouslyRequestedMachines, machineName, srvBank)) {
    return true;
  }

  // Check 4: If not yet scheduled for collection
  if (excludeNotYetScheduled(machineName, tomorrowDate, lastRemark)) {
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
function processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients, srvBank) {
  try {
    // Create hidden worksheet with collection data
    createHiddenWorksheetAndAddData(eligibleCollections, srvBank);

    // Process and send collections
    processCollectionsAndSendEmail(eligibleCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);

    CustomLogger.logInfo(`Processed ${eligibleCollections.length} eligible collections for ${srvBank}`, CONFIG.APP.NAME, 'processEligibleCollections');

  } catch (error) {
    CustomLogger.logError(`Error processing eligible collections: ${error}`, CONFIG.APP.NAME, 'processEligibleCollections');
    throw error;
  }
}
