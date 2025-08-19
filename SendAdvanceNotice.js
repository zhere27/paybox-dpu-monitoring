function sendAdvancedNotice() {
  CustomLogger.logInfo("Running sending DPU advanced notice...", PROJECT_NAME, 'sendAdvancedNotice()');
  var srvBank = '';
  const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
  const kioskData = getKioskPercentage();
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRequests, businessDays] = kioskData;
  const formattedDate = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");

  var emailTo = "";
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

    if (forExclusionBasedOnRemarks(lastRequest, todayDateString, machineName)) { return; }

    const eligibleMachines = new Set(['SMART SM SAN PABLO', 'SMART SM LAS PINAS', 'SMART SM EAST ORTIGAS', 'SMART SM MUNTINLUPA']);

    if (machineName === 'SMART SM EAST ORTIGAS') {
      emailTo = "rbargano@smart.com.ph, kmabad@smart.com.ph, madelacruz@smart.com.ph, jrgragasin@smart.com.ph";
    } else if (machineName === 'SMART SM MUNTINLUPA') {
      emailTo = "CYArcinas@smart.com.ph, EBPadrigon@smart.com.ph";
    } else if (machineName === 'SMART SM SAN PABLO') {
      emailTo = "cjronquillo@smart.com.ph, rdcorcega@smart.com.ph";
    } else if (machineName === 'SMART SM LAS PINAS') {
      emailTo = "jajariel@smart.com.ph, srsantillan@smart.com.ph, crcarubio@smart.com.ph, andiongco+laspinas@smart.com.ph";
    }

    if (eligibleMachines.has(machineName)) {
      const srvBank = 'eTap';

      if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank)) {
        if (emailTo !== "") {
          CustomLogger.logInfo(`Sending advance notice email to ${emailTo} for ${machineName}...`, PROJECT_NAME, 'sendAdvancedNotice()');
        }

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
