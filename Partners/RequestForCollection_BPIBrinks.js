/**
 * Main logic for identifying and processing BPI Brinks collections.
 * This function orchestrates fetching data, filtering eligible machines,
 * and triggering the collection process in a structured and error-handled manner.
 */
function bpiBrinkCollectionsLogic() {
  try {
    CustomLogger.logInfo(`Running BPI Brinks Collections Logic in [${CONFIG.BRINKS.ENVIRONMENT}] mode...`, CONFIG.APP.NAME, "bpiBrinkCollectionsLogic()");

    const { todayDate, todayDateString, tomorrowDate, tomorrowDateString } = getTodayAndTomorrowDates();
    const emailRecipients = getEmailRecipientsBRINKS();
    const subject = generateEmailSubjectBRINKS(tomorrowDate);

    const eligibleCollections = getEligibleCollectionsBRINKS(subject, todayDate, todayDateString,tomorrowDate, tomorrowDateString);

    processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients, CONFIG.BRINKS.ENVIRONMENT, CONFIG.BRINKS.SERVICE_BANK);
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
 * @param {Date} todayDate The Date object for today.
 * @param {string} todayDateString The formatted string for today's date.
 * @param {Date} tomorrowDate The Date object for tomorrow.
 * @param {string} tomorrowDateString The formatted string for tomorrow's date.
 * @returns {Array<Array<string>>} A 2D array of eligible collection data.
 */
function getEligibleCollectionsBRINKS(subject, todayDate, todayDateString, tomorrowDate, tomorrowDateString) {
  const machineData = getMachineDataByPartner(CONFIG.BRINKS.SERVICE_BANK, tomorrowDate);
  const [machineNames, , amountValues, , , , lastRemarks, businessDays] = machineData;
  const forCollectionData = getForCollections(CONFIG.BRINKS.SERVICE_BANK);
  const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(CONFIG.BRINKS.SERVICE_BANK, CONFIG.BRINKS.ENVIRONMENT);
  const eligibleCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machine = {
      name: machineNameArr[0],
      amount: amountValues[i][0],
      lastRemark: normalizeSpaces(lastRemarks[i][0]),
      businessDay: businessDays[i][0],
    };

    if (shouldExcludeMachines(machine, forCollectionData, previouslyRequestedMachines, todayDateString, CONFIG.BRINKS.SERVICE_BANK)) {
      return; // Skip this machine
    }

    const translatedBusinessDays = translateDaysToAbbreviation(machine.businessDay.trim());
    if (
      shouldIncludeForCollection(machine.name, machine.amount, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, machine.lastRemark, CONFIG.BRINKS.SERVICE_BANK)
    ) {
      const formattedName = formatMachineNameWithRemark(machine.name, machine.lastRemark, tomorrowDateString);
      eligibleCollections.push([formattedName, machine.amount, CONFIG.BRINKS.SERVICE_BANK, subject, machine.lastRemark]);
    }
  });

  return eligibleCollections;
}

/**
 * Retrieves the email recipients based on the current environment setting.
 * @returns {object} An object with 'to', 'cc', and 'bcc' properties.
 */
function getEmailRecipientsBRINKS() {
  return CONFIG.BRINKS.EMAIL_RECIPIENTS[CONFIG.BRINKS.ENVIRONMENT] || CONFIG.BRINKS.EMAIL_RECIPIENTS.testing;
}

/**
 * Generates the email subject for the collection request.
 * @param {Date} tomorrowDate The Date object for tomorrow.
 * @returns {string} The formatted email subject.
 */
function generateEmailSubjectBRINKS(tomorrowDate) {
  const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
  return `${CONFIG.BRINKS.SERVICE_BANK} DPU Request - ${tomorrowDateFormatted}`;
}
