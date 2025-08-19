function bpiInternalCollectionsLogic() {
  CustomLogger.logInfo("Running BPI Internal Collections Logic...", PROJECT_NAME, "bpiInternalCollectionsLogic()");
  const environment = "production";
  const srvBank = "BPI Internal";
  const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
  const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
  const subject = `${srvBank} DPU Request - ${tomorrowDateFormatted}`;

  // Skip weekends
  if (shouldSkipWeekendCollections(srvBank, tomorrowDate)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, PROJECT_NAME, "bpiInternalCollectionsLogic");
    return false;
  }

  // Move collection schedule if tomorrow is holiday
  if (isTomorrowHoliday(tomorrowDate)) {
    tomorrowDate.setDate(tomorrowDate.getDate() + 1);
  }

  const machineData = getMachineDataByPartner(srvBank);
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRemarks, businessDays] = machineData;

  const ENV_EMAIL_RECIPIENTS = {
    production: {
      to: "mjdagasuhan@bpi.com.ph",
      cc: "malodriga@bpi.com.ph, julsales@bpi.com.ph, jcslingan@bpi.com.ph, eagmarayag@bpi.com.ph, jmdcantorna@bpi.com.ph, egcameros@bpi.com.ph, jdduque@bpi.com.ph, rdtayag@bpi.com.ph, rsmendoza1@bpi.com.ph, vjdtorcuator@bpi.com.ph, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com ",
      bcc: "mvolbara@pldt.com.ph,RBEspayos@smart.com.ph ",
    },
    testing: {
      to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
      cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
      bcc: "",
    },
  };

  // Usage
  const emailRecipients = ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing;

  let forCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const lastRemark = normalizeSpaces(lastRemarks[i][0]);
    const businessDay = businessDays[i][0];
    const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

    // Skip if the last request should be excluded
    if (forExclusionBasedOnRemarks(lastRemark, todayDateString, machineName)) {
      return;
    }
    
    if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
      forCollections.push([machineName, amountValue, srvBank, subject]);
    }
    
  });

  if (environment === "production" && forCollections.length > 0) {
    createHiddenWorksheetAndAddData(forCollections, srvBank);
    processCollections(forCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
  } else {
    CustomLogger.logInfo("No eligible stores for collection tomorrow.", PROJECT_NAME, "bpiInternalCollectionsLogic()");
  }
}

function shouldIncludeForCollection_BPI(amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest) {
  const collectionDay = dayMapping[tomorrowDate.getDay()];

  if (
    ["replacement of cassette", `for collection on ${tomorrowDateString.toLowerCase()}`, `resume collection on ${tomorrowDateString.toLowerCase()}`].some((condition) =>
      lastRequest.toLowerCase().includes(condition)
    )
  ) {
    return true;
  }

  if (translatedBusinessDays.includes(collectionDay) && amountValue >= amountThresholds[collectionDay]) {
    return true;
  }

  return (
    paydayRanges.some((range) => todayDate.getDate() >= range.start && todayDate.getDate() <= range.end && amountValue >= paydayAmount) ||
    dueDateCutoffs.some((range) => todayDate.getDate() >= range.start && todayDate.getDate() <= range.end && amountValue >= dueDateCutoffsAmount)
  );
}
