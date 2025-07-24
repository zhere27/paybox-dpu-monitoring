function eTapCollectionsLogic() {
  CustomLogger.logInfo("Running eTap Collections Logic...",PROJECT_NAME,'eTapCollectionsLogic()');
  const environment = 'production'; //testing
  const srvBank = 'eTap';
  const { todayDate, tomorrowDate, todayDay, tomorrowDateString } = getTodayAndTomorrowDates();

  try {
    const kioskData = getKioskData(srvBank);
    const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, lastAddress, lastRequests, businessDays] = kioskData;
    const formattedDate = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
    const subject = `${srvBank} DPU Request - ${formattedDate}`

    const ENV_EMAIL_RECIPIENTS = {
      production: {
        to: "christian@etapinc.com,jayr@etapinc.com,rodison.carapatan@etapinc.com,arnel@etapinc.com,ian.estrella@etapinc.com,roldan@etapinc.com,dante@etapinc.com,ramun.hamtig@etapinc.com,jellymae.osorio@etapinc.com,rojane@etapinc.com, A.jloro@etapinc.com",
        cc: "reinier@etapinc.com,miguel@etapinc.com,alvie@etapinc.com,rojane@etapinc.com,laila@etapinc.com,johnmarco@etapinc.com,ghie@etapinc.com, etap-recon@etapinc.com, Erwin Alcantara <egalcantara@multisyscorp.com>, cvcabanilla@multisyscorp.com",
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

    // Special machines that need to be handled together
    const specialMachines = new Set(['PLDT BANTAY', 'SMART VIGAN']);
    let forCollections = [];

    machineNames.forEach((machineNameArr, i) => {
      const machineName = machineNameArr[0];
      const amountValue = amountValues[i][0];
      const collectionPartner = collectionPartners[i][0];
      const collectionSchedule = collectionSchedules[i][0];
      const lastRequest = normalizeSpaces(lastRequests[i][0]);
      const businessDay = businessDays[i][0];
      const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

      // Skip if it's a holiday and the machine has no-holiday schedule
      if (collectionSchedule.includes('No-Holiday') && isTomorrowHoliday(tomorrowDate)) {
        Customlogger.logInfo(`Exclude ${machineName} from DPU for tomorrow because it's a holiday and no store-in-charge.`,PROJECT_NAME, 'eTapCollectionsLogic');
        return;
      }

      // Skip if the last request should be excluded
      if (shouldExcludeFromCollection(lastRequest, todayDay, machineName)) {
        return;
      }

      // Check if we should include this machine in the collection
      if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank)) {
        const isForRevisit = lastRequest.toLowerCase().includes('for revisit');
        const displayName = `${machineName}${isForRevisit ? ' (<b>FOR REVISIT</b>)' : ''}`;

        if (specialMachines.has(machineName)) {
          // Add both special machines if either one qualifies
          specialMachines.forEach(specialMachine => {
            if (!forCollections.some(entry => entry[0] === specialMachine)) {
              forCollections.push([specialMachine, amountValue, srvBank, subject]);
            }
          });
        } else {
          forCollections.push([displayName, amountValue, srvBank, subject]);
        }
      }
    });

    if (forCollections.length > 0) {
      createHiddenWorksheetAndAddData(forCollections, srvBank);
      processCollections(forCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
    } else {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.',PROJECT_NAME, 'eTapCollectionsLogic');
    }
  } catch (error) {
    CustomLogger.logError(`Error in eTapCollectionsLogic: ${error.message}\nStack: ${error.stack}`,PROJECT_NAME, 'eTapCollectionsLogic');
  }

}
