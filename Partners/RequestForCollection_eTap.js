function eTapCollectionsLogic() {
  CustomLogger.logInfo('Running eTap Collections Logic...', CONFIG.APP.NAME, 'eTapCollectionsLogic');

  try {
    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    const { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = dateInfo;

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipientsETap(CONFIG.ETAP.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubjectETap(tomorrowDate, CONFIG.ETAP.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollectionsETap(todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, subject);

    // Process collections if any are eligible (production only)
    if (CONFIG.ETAP.ENVIRONMENT === 'production' && eligibleCollections.length > 0) {
      processEligibleCollectionsETap(eligibleCollections, tomorrowDate, emailRecipients);
    } else if (eligibleCollections.length === 0) {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.', CONFIG.APP.NAME, 'eTapCollectionsLogic');
    } else {
      CustomLogger.logInfo(`Testing mode: ${eligibleCollections.length} collections identified but not sent.`, CONFIG.APP.NAME, 'eTapCollectionsLogic');
      console.log(JSON.stringify(eligibleCollections, null, 2));
    }
    //Send Logs to Admin
    EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);
  } catch (error) {
    CustomLogger.logError(`Error in eTapCollectionsLogic: ${error.message}\nStack: ${error.stack}`, CONFIG.APP.NAME, 'eTapCollectionsLogic');
    throw error;
  }
}

/**
 * Get email recipients based on environment
 */
function getEmailRecipientsETap(environment) {
  return CONFIG.ETAP.EMAIL_RECIPIENTS[environment] || CONFIG.ETAP.EMAIL_RECIPIENTS.testing;
}

/**
 * Generate email subject line
 */
function generateEmailSubjectETap(tomorrowDate, serviceBank) {
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
function getEligibleCollectionsETap(todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, subject) {
  // Fetch data
  const machineData = getMachineDataByPartner(CONFIG.ETAP.SERVICE_BANK, tomorrowDate);
  const forCollectionData = getForCollections(CONFIG.ETAP.SERVICE_BANK);
  const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(CONFIG.ETAP.SERVICE_BANK);

  // Destructure machine data
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, lastAddress, lastRemarks, businessDays] = machineData;

  // Filter eligible collections
  const eligibleCollections = [];
  const specialMachinesAdded = new Set();

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const collectionSchedule = collectionSchedules[i][0];
    const lastRemark = normalizeSpaces(lastRemarks[i][0]);
    const businessDay = businessDays[i][0];

    // Apply exclusion filters
    if (excludeMachineETap(machineName, amountValue, collectionSchedule, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, CONFIG.ETAP.SERVICE_BANK)) {
      return; // Skip this machine
    }

    // Handle special machines that must be collected together
    if (CONFIG.ETAP.SPECIAL_MACHINES.has(machineName)) {
      addSpecialMachines(eligibleCollections, specialMachinesAdded, amountValue, CONFIG.ETAP.SERVICE_BANK, subject);
    } else {
      // Format machine name with remark if it's a revisit
      const displayName = formatMachineNameWithRemarkETap(machineName, lastRemark, tomorrowDateString);
      eligibleCollections.push([displayName, amountValue, CONFIG.ETAP.SERVICE_BANK, subject]);
    }
  });

  return eligibleCollections;
}

/**
 * Determine if machine should be excluded from collection
 */
function excludeMachineETap(machineName, amountValue, collectionSchedule, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday, srvBank) {
  const collectionDay = dayMapping[tomorrowDate.getDay()];

  if (hasSpecialCollectionConditions(lastRemark, tomorrowDateString)) {
    return false;
  }

  // Check: Excluded stores
  if (isExcludedStore(machineName)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store is on the exclusion list.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  // Check: Exclude if amount did not meet
  if (!meetsAmountThreshold(amountValue, collectionDay, srvBank)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the amount did not meet amount threshold.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  // Check 1: Already collected
  if (excludeAlreadyCollected(machineName, forCollectionData)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store was already collected.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  // Check 2: No-Holiday schedule and tomorrow is holiday
  if (excludeHolidayCollection(machineName, collectionSchedule, isTomorrowHoliday)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, bacause store is closed during holidays.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  // Check 3: Excluded based on remarks
  if (excluceBasedOnRemarks(lastRemark, todayDateString, machineName)) {
    return true;
  }

  // Check 4: Requested yesterday
  if (excludeRequestedYesterday(previouslyRequestedMachines, machineName, srvBank)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store was requested for collection yesterday.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  // Check 4: If not yet scheduled for collection
  if (excludeNotYetScheduled(machineName, tomorrowDateString, lastRemark)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store is not yet scheduled for collection.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  // Check 5: Business day schedule
  const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

  if (!shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store is not scheduled for collection on ${tomorrowDateString}.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  return false;
}

/**
 * Add special machines that must be collected together
 */
function addSpecialMachines(eligibleCollections, specialMachinesAdded, amountValue, srvBank, subject) {
  // If one special machine qualifies, add both (but only once)
  CONFIG.ETAP.SPECIAL_MACHINES.forEach(specialMachine => {
    if (!specialMachinesAdded.has(specialMachine)) {
      eligibleCollections.push([specialMachine, amountValue, srvBank, subject]);
      specialMachinesAdded.add(specialMachine);
    }
  });
}

/**
 * Format machine name with remark if it's a revisit
 */
function formatMachineNameWithRemarkETap(machineName, lastRemark, tomorrowDateString) {
  const remarkLower = lastRemark.toLowerCase();
  const revisitPhrase = `for revisit on ${tomorrowDateString.toLowerCase()}`;

  if (remarkLower.includes(revisitPhrase)) {
    return `${machineName} (<b>${lastRemark}</b>)`;
  }

  return machineName;
}

/**
 * Process eligible collections
 */
function processEligibleCollectionsETap(eligibleCollections, tomorrowDate, emailRecipients) {
  const srvBank = CONFIG.ETAP.SERVICE_BANK;

  try {
    // Create hidden worksheet with collection data
    createHiddenWorksheetAndAddData(eligibleCollections, srvBank);

    // Process and send collections
    processCollectionsAndSendEmail(eligibleCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);

    CustomLogger.logInfo(`Processed ${eligibleCollections.length} eligible collections for ${srvBank}`, CONFIG.APP.NAME, 'processEligibleCollectionsETap');

  } catch (error) {
    CustomLogger.logError(`Error processing eligible collections: ${error}`, CONFIG.APP.NAME, 'processEligibleCollectionsETap');
    throw error;
  }
}