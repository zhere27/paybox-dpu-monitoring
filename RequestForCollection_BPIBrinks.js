function bpiBrinkCollectionsLogic() {
  CustomLogger.logInfo('Running bpiBrinkCollectionsLogic...', PROJECT_NAME, 'bpiBrinkCollectionsLogic()');

  const environment = 'production';
  const srvBank = 'Brinks via BPI';
  var { todayDate, tomorrowDate, todayDay, tomorrowDateString } = getTodayAndTomorrowDates();

  if (shouldSkipExecution(todayDate)) return;

  if (isTomorrowHoliday(tomorrowDate)) {
    tomorrowDate.setDate(tomorrowDate.getDate() + 1);
  }

  const kioskData = getKioskData(srvBank);
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRequests, businessDays] = kioskData;
  const formattedDate = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
  const subject = `${srvBank} DPU Request - ${formattedDate}`

  const ENV_EMAIL_RECIPIENTS = {
    production: {
      to: "mbocampo@bpi.com.ph",
      cc: "julsales@bpi.com.ph, mjdagasuhan@bpi.com.ph, eagmarayag@bpi.com.ph,rtorticio@bpi.com.ph, egcameros@bpi.com.ph, vjdtorcuator@bpi.com.ph, jdduque@bpi.com.ph, rsmendoza1@bpi.com.ph,jmdcantorna@bpi.com.ph, kdrepuyan@bpi.com.ph, dabayaua@bpi.com.ph, rdtayag@bpi.com.ph, vrvarellano@bpi.com.ph,mapcabela@bpi.com.ph, mvpenisa@bpi.com.ph, mbcernal@bpi.com.ph, cmmanalac@bpi.com.ph, mpdcastro@bpi.com.ph,rmdavid@bpi.com.ph, emflores@bpi.com.ph, apmlubaton@bpi.com.ph, smcarvajal@bpi.com.ph, avabarabar@bpi.com.ph,jcmontes@bpi.com.ph, jeobautista@bpi.com.ph, micaneda@bpi.com.ph, jgramos@bpi.com.ph, rrpacio@bpi.com.ph,mecdycueco@bpi.com.ph, tesruelo@bpi.com.ph, ssibon@bpi.com.ph, christine.sarong@brinks.com, icom2.ph@brinks.com,aillen.waje@brinks.com, rpsantiago@bpi.com.ph, jerome.apora@brinks.com, occ2supervisors.ph@brinks.com, mdtenido@bpi.com.ph, agmaiquez@bpi.com.ph, cvcabanilla@multisyscorp.com, raflorentino@multisyscorp.com, egalcantara@multisyscorp.com",
      bcc: "mvolbara@pldt.com.ph"
    },
    testing: {
      to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
      cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
      bcc: ""
    }
  };

  // Usage
  const emailRecipients = ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing;

  let forCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const collectionPartner = collectionPartners[i][0];
    const lastRequest = normalizeSpaces(lastRequests[i][0]);
    const businessDay = businessDays[i][0];
    const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

    if (shouldExcludeFromCollection(lastRequest, todayDay, machineName)) return;

    if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank)) {
      if (lastRequest.toLowerCase().includes('for revisit')) {
        forCollections.push([machineName + ' (<b>FOR REVISIT</b>)', amountValue, srvBank, subject, lastRequest]);
      }
      else {
        forCollections.push([machineName, amountValue, srvBank, subject, lastRequest]);
      }
    }
  });

  forCollections = excludePastRequests(forCollections, tomorrowDateString);
  // Logger.log(forCollections);
  // forCollections = excludePreviouslyRequested(forCollections, srvBank);

  if (forCollections.length > 0) {
    createHiddenWorksheetAndAddData(forCollections, srvBank);
    processCollections(forCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
  } else {
    CustomLogger.logInfo('No eligible stores for collection tomorrow.', PROJECT_NAME, 'bpiBrinkCollectionsLogic()');
  }

}