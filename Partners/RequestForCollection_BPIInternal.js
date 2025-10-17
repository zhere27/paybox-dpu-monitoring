function bpiInternalCollectionsLogic() {
  CustomLogger.logInfo('Running BPI Internal Collections Logic...', CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');

  try {
    // Get date information
    const dateInfo = getTodayAndTomorrowDates();
    let { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = dateInfo;

    // Adjust collection date to next valid working day (not weekend/holiday)
    tomorrowDate = adjustToNextWorkingDay(tomorrowDate);
    tomorrowDateString = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy");

    CustomLogger.logInfo(
      `Collection date adjusted to ${tomorrowDateString}`,
      CONFIG.APP.NAME,
      'bpiInternalCollectionsLogic'
    );

    // // Adjust collection date if tomorrow is a holiday
    // if (isTomorrowHoliday) {
    //   tomorrowDate = adjustDateForHoliday(tomorrowDate);
    //   tomorrowDateString = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
    //   CustomLogger.logInfo(`Tomorrow is a holiday. Collection date adjusted to ${tomorrowDateString}`, CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');
    // }

    // Check if adjusted date is weekend
    if (isWeekend(tomorrowDate)) {
      CustomLogger.logInfo(`Skipping collections - tomorrow date ${tomorrowDateString} falls on weekend.`, CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');
      return;
    }

    // Get email recipients based on environment
    const emailRecipients = getEmailRecipientsInternal(CONFIG.BPI_INTERNAL.ENVIRONMENT);

    // Generate email subject
    const subject = generateEmailSubjectInternal(tomorrowDate, CONFIG.BPI_INTERNAL.SERVICE_BANK);

    // Get machine data and filter for eligible collections
    const eligibleCollections = getEligibleCollectionsInternal(todayDate, tomorrowDate, todayDateString, tomorrowDateString, subject);

    // Process collections if any are eligible (production only)
    if (CONFIG.BPI_INTERNAL.ENVIRONMENT === 'production' && eligibleCollections.length > 0) {
      processEligibleCollectionsInternal(eligibleCollections, tomorrowDate, emailRecipients);
    } else if (eligibleCollections.length === 0) {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.', CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');
    } else {
      CustomLogger.logInfo(`Testing mode: ${eligibleCollections.length} collections identified but not sent.`, CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');
      console.log(JSON.stringify(collections, null, 2));
    }
    //Send Logs to Admin
    EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);
  } catch (error) {
    CustomLogger.logError(`Error in bpiInternalCollectionsLogic: ${error}`, CONFIG.APP.NAME, 'bpiInternalCollectionsLogic');
    throw error;
  }
}

/**
 * Returns the next date that is not a weekend or holiday
 */
function adjustToNextWorkingDay(date) {
  let adjustedDate = new Date(date);

  while (isWeekend(adjustedDate) || isHoliday(adjustedDate)) {
    adjustedDate.setDate(adjustedDate.getDate() + 1);
  }

  return adjustedDate;
}

/**
 * Adjust date by adding one day for holiday
 */
function adjustDateForHoliday(date) {
  const adjustedDate = new Date(date);
  adjustedDate.setDate(adjustedDate.getDate() + 1);
  return adjustedDate;
}

/**
 * Check if date falls on weekend
 */
function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6; // 0 = Sunday, 6 = Saturday
}

/**
 * Returns true if the date matches one in the “Holidays” sheet.
 * Expects holidays in column A, one per row, in any date format.
 */
function isHoliday(date) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StoreName Mapping");
  if (!sheet) return false;

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return false;

  const holidayValues = sheet.getRange(3, 7, lastRow).getValues(); // Column A
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Convert sheet values to comparable string format
  return holidayValues.some(row => {
    const holiday = row[0];
    if (!holiday) return false;
    const holidayStr = Utilities.formatDate(new Date(holiday), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return holidayStr === formattedDate;
  });
}

/**
 * Get email recipients based on environment
 */
function getEmailRecipientsInternal(environment) {
  return CONFIG.BPI_INTERNAL.EMAIL_RECIPIENTS[environment] || CONFIG.BPI_INTERNAL.EMAIL_RECIPIENTS.testing;
}

/**
 * Generate email subject line
 */
function generateEmailSubjectInternal(tomorrowDate, serviceBank) {
  const formattedDate = Utilities.formatDate(
    tomorrowDate,
    Session.getScriptTimeZone(),
    "MMMM d, yyyy (EEEE)"
  );
  return `${serviceBank} DPU Request - ${formattedDate}`;
}

/**
 * Get eligible collections based on criteria
 */
function getEligibleCollectionsInternal(todayDate, tomorrowDate, todayDateString, tomorrowDateString, subject) {
  const srvBank = CONFIG.BPI_INTERNAL.SERVICE_BANK;

  // Fetch data
  const machineData = getMachineDataByPartner(srvBank, tomorrowDate);
  const forCollectionData = getForCollections(srvBank);
  const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(srvBank);

  // Destructure machine data
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRemarks, businessDays] = machineData;

  // Filter eligible collections
  const eligibleCollections = [];

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const amountValue = amountValues[i][0];
    const lastRemark = normalizeSpaces(lastRemarks[i][0]);
    const businessDay = businessDays[i][0];

    // Apply exclusion filters
    if (excludeMachineBpiInternal(machineName, amountValue, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, srvBank)) {
      return; // Skip this machine
    }

    // Add to eligible collections
    eligibleCollections.push([machineName, amountValue, srvBank, subject]);
  });

  return eligibleCollections;
}

/**
 * Determine if machine should be excluded from collection
 */
function excludeMachineBpiInternal(machineName, amountValue, lastRemark, businessDay, forCollectionData, previouslyRequestedMachines, todayDate, tomorrowDate, todayDateString, tomorrowDateString, srvBank) {  // Check 1: Skip weekend collections
  const collectionDay = dayMapping[tomorrowDate.getDay()];

  if (hasSpecialCollectionConditions(lastRemark, tomorrowDateString)) {
    return false;
  }

  // Check: Excluded stores
  if (isExcludedStore(machineName)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the store is on the exclusion list.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  // Check: Exclude if amount did not meet
  if (!meetsAmountThreshold(amountValue, collectionDay, srvBank)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName}, because the amount did not meet amount threshold.`, CONFIG.APP.NAME, 'excludeMachineETap');
    return true;
  }

  if (skipWeekendCollections(srvBank, tomorrowDate)) {
    CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, CONFIG.APP.NAME, 'shouldExcludeMachineInternal');
    return true;
  }

  // Check 2: Already collected
  if (excludeAlreadyCollected(machineName, forCollectionData)) {
    return true;
  }

  // Check 3: Excluded based on remarks
  if (excluceBasedOnRemarks(lastRemark, todayDateString, machineName)) {
    return true;
  }

  // Check 4: Requested yesterday
  if (excludeRequestedYesterday(previouslyRequestedMachines, machineName, srvBank)) {
    return true;
  }

  // Check 5: Business day schedule
  const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

  if (!shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
    return true;
  }

  return false;
}

/**
 * Process eligible collections
 */
function processEligibleCollectionsInternal(eligibleCollections, tomorrowDate, emailRecipients) {
  const srvBank = CONFIG.BPI_INTERNAL.SERVICE_BANK;

  try {
    if (CONFIG.BPI_INTERNAL.ENVIRONMENT === "production") {
      // Create hidden worksheet with collection data
      createHiddenWorksheetAndAddData(eligibleCollections, srvBank);

      // Save eligibleCollections to BigQuery
      saveEligibleCollectionsToBQ(eligibleCollections);

      // Process and send collections
      processCollectionsAndSendEmail(eligibleCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
      CustomLogger.logInfo(`Processed ${eligibleCollections.length} eligible collections for ${srvBank}`, CONFIG.APP.NAME, 'processEligibleCollectionsInternal');
    } else if (eligibleCollections.length === 0) {
      CustomLogger.logInfo('No eligible stores for collection tomorrow.', CONFIG.APP.NAME, 'processEligibleCollectionsInternal');
    } else {
      CustomLogger.logInfo(`[TESTING MODE] Found ${collections.length} eligible collections. No emails will be sent.`, CONFIG.APP.NAME, "processEligibleCollectionsInternal");
      // Optional: Log the collections for verification during testing
      console.log(JSON.stringify(collections, null, 2));
    }
  } catch (error) {
    CustomLogger.logError(`Error processing eligible collections: ${error}`, CONFIG.APP.NAME, 'processEligibleCollectionsInternal');
    throw error;
  }
}

// function bpiInternalCollectionsLogic() {
//   CustomLogger.logInfo("Running BPI Internal Collections Logic...", CONFIG.APP.NAME, "bpiInternalCollectionsLogic()");
//   const environment = "production";
//   const srvBank = "BPI Internal";
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString, isTomorrowHoliday } = getTodayAndTomorrowDates();
//   const tomorrowDateFormatted = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
//   const subject = `${srvBank} DPU Request - ${tomorrowDateFormatted}`;

//   // Move collection schedule if tomorrow is holiday
//   if (isTomorrowHoliday) {
//     tomorrowDate.setDate(tomorrowDate.getDate() + 1);
//   }

//   const machineData = getMachineDataByPartner(srvBank);
//   const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRemarks, businessDays] = machineData;

//   const ENV_EMAIL_RECIPIENTS = {
//     production: {
//       to: "mjdagasuhan@bpi.com.ph",
//       cc: "malodriga@bpi.com.ph, julsales@bpi.com.ph, jcslingan@bpi.com.ph, eagmarayag@bpi.com.ph, jmdcantorna@bpi.com.ph, egcameros@bpi.com.ph, jdduque@bpi.com.ph, rdtayag@bpi.com.ph, rsmendoza1@bpi.com.ph, vjdtorcuator@bpi.com.ph, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com ",
//       bcc: "mvolbara@pldt.com.ph,RBEspayos@smart.com.ph ",
//     },
//     testing: {
//       to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       bcc: "",
//     },
//   };

//   // Usage
//   const emailRecipients = ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing;
//   const forCollectionData = getForCollections(srvBank);

//   const previouslyRequestedMachines = getPreviouslyRequestedMachineNamesByServiceBank(srvBank);
//   let forCollections = [];

//   machineNames.forEach((machineNameArr, i) => {
//     const machineName = machineNameArr[0];
//     const amountValue = amountValues[i][0];
//     const lastRemark = normalizeSpaces(lastRemarks[i][0]);
//     const businessDay = businessDays[i][0];
//     const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

//     // Skip weekends
//     if (skipWeekendCollections(srvBank, tomorrowDate)) {
//       CustomLogger.logInfo(`Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, CONFIG.APP.NAME, "bpiInternalCollectionsLogic()");
//       return false;
//     }

//     // Skip if already collected
//     if (skipAlreadyCollected(machineName, forCollectionData)) {
//       return;
//     }

//     // Skip if the last request should be excluded
//     if (forExclusionBasedOnRemarks(lastRemark, todayDateString, machineName)) {
//       return;
//     }

//     // Skip if already requested for collection yesterday
//     if (forExclusionRequestYesterday(previouslyRequestedMachines, machineName, srvBank)) {
//       return;
//     }


//     if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRemark, srvBank)) {
//       forCollections.push([machineName, amountValue, srvBank, subject]);
//     }

//   });

//   if (environment === "production" && forCollections.length > 0) {
//     createHiddenWorksheetAndAddData(forCollections, srvBank);
//     processCollections(forCollections, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
//   } else {
//     CustomLogger.logInfo("No eligible stores for collection tomorrow.", CONFIG.APP.NAME, "bpiInternalCollectionsLogic()");
//   }
// }

