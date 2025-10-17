function bpiInternalCollectionsLogic() {
  CustomLogger.logInfo('Running BPI Internal Collections Logic...', CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');

  try {
    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    let { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = dateInfo;

    // Adjust collection date to next valid working day (not weekend/holiday)
    if (isTomorrowHoliday) {
      tomorrowDate = adjustToNextWorkingDay(tomorrowDate);
      tomorrowDateString = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
      CustomLogger.logInfo(`Collection date adjusted to ${tomorrowDateString}`, CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');
    }

    // Check if adjusted date is weekend
    if (isWeekend(tomorrowDate)) {
      CustomLogger.logInfo(`Skipping collections - tomorrow date ${tomorrowDateString} falls on weekend.`, CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');
      //Send Logs to Admin
      EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);
      return;
    }

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipients(CONFIG.BPI_INTERNAL.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubject(tomorrowDate, CONFIG.BPI_INTERNAL.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollections(subject, todayDate, todayDateString, tomorrowDate, tomorrowDateString, CONFIG.BPI_INTERNAL.SERVICE_BANK, CONFIG.BPI_INTERNAL.ENVIRONMENT);

    processEligibleCollections(eligibleCollections, tomorrowDate, emailRecipients, CONFIG.BPI_INTERNAL.ENVIRONMENT, CONFIG.BPI_INTERNAL.SERVICE_BANK);

    //Send Logs to Admin
    EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);
  } catch (error) {
    CustomLogger.logError(`Error in bpiInternalCollectionsLogic: ${error}`, CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');
    throw error;
  }

  /**
   * Retrieves the email recipients based on the current environment setting.
   * @returns {object} An object with 'to', 'cc', and 'bcc' properties.
   */
  function getEmailRecipients() {
    return CONFIG.BPI_INTERNAL.EMAIL_RECIPIENTS[CONFIG.BRINKS.ENVIRONMENT] || CONFIG.BPI_INTERNAL.EMAIL_RECIPIENTS.testing;
  }

  /**
   * Generates the email subject for the collection request.
   * @param {Date} tomorrowDate The Date object for tomorrow.
   * @returns {string} The formatted email subject.
   */
  function generateEmailSubject(tomorrowDate) {
    const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
    return `${CONFIG.BPI_INTERNAL.SERVICE_BANK} DPU Request - ${tomorrowDateFormatted}`;
  }
}