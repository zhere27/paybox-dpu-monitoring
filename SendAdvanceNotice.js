function sendAdvancedNotice() {
  CustomLogger.logInfo("Running sending DPU advanced notice...", PROJECT_NAME, 'sendAdvancedNotice()');
  var srvBank = '';
  const { todayDate, tomorrowDate, todayDay, tomorrowDateString } = getTodayAndTomorrowDates();
  const kioskData = getKioskPercentage();
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRequests, businessDays] = kioskData;
  const formattedDate = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");

  var emailTo = "CJRonquillo@smart.com.ph, RDCorcega@smart.com.ph";
  var emailCc = "RBEspayos@smart.com.ph, RACagbay@smart.com.ph";
  var emailBcc = "Erwin Alcantara <egalcantara@multisyscorp.com>"

  // var emailTo = "Erwin Alcantara <egalcantara@multisyscorp.com>";
  // var emailCc = "Erwin Alcantara <egalcantara@multisyscorp.com>";
  // var emailBcc = "Erwin Alcantara <egalcantara@multisyscorp.com>";

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const collectionPartner = collectionPartners[i][0];
    const lastRequest = lastRequests[i][0];
    const businessDay = businessDays[i][0];
    const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

    if (shouldExcludeFromCollection(lastRequest, todayDay)) { return; }

    const eligibleMachines = new Set(['SMART SM SAN PABLO', 'SMART SM LAS PINAS', 'SMART SM EAST ORTIGAS']);

    if (eligibleMachines.has(machineName)) {
      const srvBank = 'eTap';
      if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank)) {
        sendEmail(machineName, tomorrowDate);
      }
    }

  });

  function sendEmail(machineName, collectionDate) {
    const formattedDate = Utilities.formatDate(collectionDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");

    const subject = `Paybox - DPU Collection advanced notice - ${formattedDate}`;
    const body = `Hi ${machineName},<br><br>
      Please be advised that Paybox collection pickup will occur tomorrow, ${formattedDate}.<br><br>
      Kindly prepare the nessary permit. <br><br>
      Thank you.<br><br>

      *** This is an automated email. ****<br><br>${emailSignature}
  `;

    GmailApp.sendEmail(emailTo, subject, '', {
      cc: emailCc,
      bcc: emailBcc,
      htmlBody: body,
      from: "support@paybox.ph"
    });
  }
}
