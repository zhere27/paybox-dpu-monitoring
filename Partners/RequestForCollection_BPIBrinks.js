/**
 * Main logic for identifying and processing BPI Brinks collections.
 * This function orchestrates fetching data, filtering eligible machines,
 * and triggering the collection process in a structured and error-handled manner.
 */
function bpiBrinkCollectionsLogic() {
  try {
    CustomLogger.logInfo(`Running BPI Brinks Collections Logic in [${CONFIG.BRINKS.ENVIRONMENT}] mode...`, CONFIG.APP.NAME, "bpiBrinkCollectionsLogic()");

    const { todayDate, todayDateString, tomorrowDate, tomorrowDateString } = getTodayAndTomorrowDates();
    const emailRecipients = getEmailRecipients();
    const subject = generateEmailSubject(tomorrowDate);

    const eligibleCollections = getEligibleCollections(subject, todayDate, todayDateString, tomorrowDate, tomorrowDateString, CONFIG.BRINKS.SERVICE_BANK, CONFIG.BRINKS.ENVIRONMENT);

    processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients, CONFIG.BRINKS.ENVIRONMENT, CONFIG.BRINKS.SERVICE_BANK);

    EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);
  } catch (error) {
    const errorMessage = `An unexpected error occurred in BPI Brinks Logic: ${error.message} \nStack: ${error.stack}`;
    CustomLogger.logError(errorMessage, CONFIG.APP.NAME, "bpiBrinkCollectionsLogic()");
    // Optionally, re-throw the error if you want it to be visible in the Apps Script dashboard
    // throw error;
  }

  /**
   * Retrieves the email recipients based on the current environment setting.
   * @returns {object} An object with 'to', 'cc', and 'bcc' properties.
   */
  function getEmailRecipients() {
    return CONFIG.BRINKS.EMAIL_RECIPIENTS[CONFIG.BRINKS.ENVIRONMENT] || CONFIG.BRINKS.EMAIL_RECIPIENTS.testing;
  }

  /**
   * Generates the email subject for the collection request.
   * @param {Date} tomorrowDate The Date object for tomorrow.
   * @returns {string} The formatted email subject.
   */
  function generateEmailSubject(tomorrowDate) {
    const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
    return `${CONFIG.BRINKS.SERVICE_BANK} DPU Request - ${tomorrowDateFormatted}`;
  }

}