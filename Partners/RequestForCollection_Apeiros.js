function apeirosCollectionsLogic() {
  CustomLogger.logInfo('Running Apeiros Collection Logic...', CONFIG.APP.NAME, 'apeirosCollectionsLogic');

  try {
    // Get date information
    const { todayDate, todayDateString, tomorrowDate, tomorrowDateString } = getTodayAndTomorrowDates();

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipients(CONFIG.APEIROS.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubject(tomorrowDate, CONFIG.APEIROS.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollections(subject, todayDate, todayDateString, tomorrowDate, tomorrowDateString, CONFIG.APEIROS.SERVICE_BANK, CONFIG.APEIROS.ENVIRONMENT);

    processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients, CONFIG.APEIROS.ENVIRONMENT, CONFIG.APEIROS.SERVICE_BANK);

    EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);
  } catch (error) {
    CustomLogger.logError(`Error in apeirosCollectionsLogic: ${error}`, CONFIG.APP.NAME, 'apeirosCollectionsLogic');
    throw error;
  }

  /**
   * Retrieves the email recipients based on the current environment setting.
   * @returns {object} An object with 'to', 'cc', and 'bcc' properties.
   */
  function getEmailRecipients() {
    return CONFIG.APEIROS.EMAIL_RECIPIENTS[CONFIG.APEIROS.ENVIRONMENT] || CONFIG.APEIROS.EMAIL_RECIPIENTS.testing;
  }

  /**
   * Generates the email subject for the collection request.
   * @param {Date} tomorrowDate The Date object for tomorrow.
   * @returns {string} The formatted email subject.
   */
  function generateEmailSubject(tomorrowDate) {
    const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
    return `${CONFIG.APEIROS.SERVICE_BANK} DPU Request - ${tomorrowDateFormatted}`;
  }
}