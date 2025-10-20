function eTapCollectionsLogic() {
  CustomLogger.logInfo('Running eTap Collections Logic...', CONFIG.APP.NAME, 'eTapCollectionsLogic');

  try {
    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    const { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = dateInfo;

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipients(CONFIG.ETAP.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubject(tomorrowDate, CONFIG.ETAP.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollections(subject, todayDate, todayDateString, tomorrowDate, tomorrowDateString, CONFIG.ETAP.SERVICE_BANK, CONFIG.ETAP.ENVIRONMENT);

    processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients, CONFIG.ETAP.ENVIRONMENT, CONFIG.ETAP.SERVICE_BANK);

    //Send Logs to Admin
    EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);
  } catch (error) {
    CustomLogger.logError(`Error in eTapCollectionsLogic: ${error.message}\nStack: ${error.stack}`, CONFIG.APP.NAME, 'eTapCollectionsLogic');
    throw error;
  }

  /**
   * Retrieves the email recipients based on the current environment setting.
   * @returns {object} An object with 'to', 'cc', and 'bcc' properties.
   */
  function getEmailRecipients() {
    return CONFIG.ETAP.EMAIL_RECIPIENTS[CONFIG.ETAP.ENVIRONMENT] || CONFIG.ETAP.EMAIL_RECIPIENTS.testing;
  }

  /**
   * Generates the email subject for the collection request.
   * @param {Date} tomorrowDate The Date object for tomorrow.
   * @returns {string} The formatted email subject.
   */
  function generateEmailSubject(tomorrowDate) {
    const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
    return `${CONFIG.ETAP.SERVICE_BANK} DPU Request - ${tomorrowDateFormatted}`;
  }
}