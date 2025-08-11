function apeirosCollectionsLogic() {
  CustomLogger.logInfo("Running Apeiros Collection Logic...",PROJECT_NAME,'eTapCollectionsLogic()');
  const environment = 'testing';
  const srvBank = 'Apeiros';
  const { todayDate, tomorrowDate, todayDay, tomorrowDateString } = getTodayAndTomorrowDates();

  if (shouldSkipExecution(todayDate)) return;

  const kioskData = getKioskData(srvBank);
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRequests, businessDays] = kioskData;
  const formattedDate = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
  const subject = `${srvBank} DPU Request - ${formattedDate}`

  const ENV_EMAIL_RECIPIENTS = {
    production: {
      to: "mtcsantiago570@gmail.com, mtcsurigao@gmail.com, valdez.ezekiel23@gmail.com",
      cc: "sherwinamerica@yahoo.com, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com ",
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

    if (collectionPartner !== 'Apeiros' || shouldExcludeFromCollection(lastRequest, todayDay)) return;

    const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

    if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank)) {
      forCollections.push([machineName, amountValue, srvBank, subject]);
    }
  });

  if (forCollections.length > 0) {
    createHiddenWorksheetAndAddData(forCollections, srvBank);
    processCollections(forCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
  } else {
    CustomLogger.logInfo('No eligible stores for collection tomorrow.',PROJECT_NAME,'apeirosCollectionsLogic()');
  }
}
