// // Configuration objects

// const LOCATION_PAIRS = {
//   PLDT_PACO: ['PLDT PACO 1', 'PLDT PACO 2'],
//   SMART_SM_NORTH: ['SMART SM NORTH EDSA 1', 'SMART SM NORTH EDSA 2']
// };

// const EXCLUSION_CONDITIONS = [
//   'replacement of cassette',
//   'for collection on',
//   'resume collection on'
// ];

// const SPECIAL_LOCATIONS = {
//   PLDT_LAOAG: 'PLDT LAOAG'
// };

// // Cache for memoized results
// const cache = {
//   kioskData: null,
//   lastFetchTime: null,
//   cacheDuration: 5 * 60 * 1000 // 5 minutes in milliseconds
// };

// // Main function
// function bdoCancellationLogic() {
//   skipFunction();
//   CustomLogger.logInfo("Running bdoCancellationLogic...",PROJECT_NAME, 'bdoCancellationLogic');
//   //If the current date is after May 31, 2025 then skip execution
//   if (new Date() > new Date(2025, 4, 31)) return;

//   const environment = "testing"; // or "testing"
//   const srvBank = 'BDO';
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();

//   const ENV_EMAIL_RECIPIENTS = {
//     production: {
//       to: "CMS-Corbank <cms-corbank@bdo.com.ph>,CHH Laoag Cash Hub <chh.laoag@bdo.com.ph>,BM SM City Bataan <bh.sm-city-bataan@bdo.com.ph>,BM SM City North EDSA C <bh.sm-city-north-edsa-c@bdo.com.ph>,CHBMH Tagaytay Cash Hub <chh.tagaytay@bdo.com.ph>",
//       cc: "CHH CCS Makati Cash Hub <chh.ccs-makati@bdo.com.ph>,CHH Las Piñas Cash Hub <chh.las-pinas@bdo.com.ph>,CHH Legazpi Cash Hub <chh.legazpi@bdo.com.ph>,BM SM Aura Premier <bh.sm-aura-premier@bdo.com.ph>,BM SM City Bacoor <bh.sm-city-bacoor@bdo.com.ph>,BM SM City Davao Annex <bh.sm-city-davao-annex@bdo.com.ph>,BM SM City Sta Rosa <bh.sm-city-sta-rosa@bdo.com.ph>,CHH Taft Cash Hub <chh.taft@bdo.com.ph>,CHH Tuguegarao Cash Hub <chh.tuguegarao@bdo.com.ph>, Erwin Alcantara <egalcantara@multisyscorp.com>,Ronald Florentino <raflorentino@multisyscorp.com>,cvcabanilla@multisyscorp.com",
//       bcc: ""
//     },
//     testing: {
//       to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
//       bcc: ""
//     }
//   };

//   // Usage
//   const emailRecipients = ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing;


//   //skip execution on weekends
//   if (shouldSkipExecution(todayDate)) return;

//   try {
//     // Get kiosk data with caching
//     const kioskData = getKioskDataWithCache();
//     const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, partnerAddresses, lastRequests, businessDays] = kioskData;

//     // Process cancellations
//     const forCancellation = processMachineCancellations(
//       machineNames,
//       amountValues,
//       collectionPartners,
//       collectionSchedules,
//       lastRequests,
//       businessDays,
//       partnerAddresses,
//       srvBank,
//       tomorrowDate,
//       todayDate,
//       tomorrowDateString
//     );

//     if (forCancellation.length > 0) {
//       processCancellation(forCancellation, tomorrowDate, emailRecipients.to, emailRecipients.cc, emailRecipients.bcc, srvBank);
//     }
//   } catch (error) {
//     handleError(error, 'bdoCancellationLogic');
//   }
// }

// // Helper function to get kiosk data with caching
// function getKioskDataWithCache() {
//   const now = new Date().getTime();

//   // Return cached data if it's still valid
//   if (cache.kioskData && cache.lastFetchTime && (now - cache.lastFetchTime < cache.cacheDuration)) {
//     return cache.kioskData;
//   }

//   // Fetch fresh data
//   const kioskData = getKioskPercentage();

//   // Update cache
//   cache.kioskData = kioskData;
//   cache.lastFetchTime = now;

//   return kioskData;
// }

// // Helper function to process machine cancellations
// function processMachineCancellations(machineNames, amountValues, collectionPartners, collectionSchedules, lastRequests, businessDays, partnerAddresses, srvBank, tomorrowDate, todayDate, tomorrowDateString) {
//   // Pre-filter machines by service bank for better performance
//   var bdoMachines = machineNames.reduce((acc, machineNameArr, i) => {
//     if (collectionPartners[i][0] === srvBank) {
//       if (collectionSchedules[i][0].includes(getTomorrowDayCustomFormat(tomorrowDate))) {
//         CustomLogger.logInfo(`Cancelling DPU for machine ${machineNameArr[0]} for tomorrow.`,PROJECT_NAME, 'bdoCancellationLogic');
//         Logger.log(`Cancelling DPU for machine ${machineNameArr[0]} (${collectionSchedules[i][0]}) - ${getTomorrowDayCustomFormat(tomorrowDate)}`);

//         acc.push({
//           index: i,
//           name: machineNameArr[0],
//           amountValue: amountValues[i][0],
//           collectionSchedule: collectionSchedules[i][0],
//           lastRequest: lastRequests[i][0],
//           partnerAddress: partnerAddresses[i][0]
//         });
//       }
//     }
//     return acc;
//   }, []);

//   // Process cancellations in a single pass
//   let forCancellation = bdoMachines
//     .filter(machine => shouldIncludeForCancellation(
//       machine.name,
//       machine.amountValue,
//       tomorrowDate,
//       machine.collectionSchedule,
//       todayDate,
//       tomorrowDateString,
//       machine.lastRequest
//     ))
//     .map(machine => ({
//       name: machine.name,
//       address: machine.partnerAddress,
//       currentCashAmount: machine.amountValue
//     }));

//   // Apply special validations
//   forCancellation = validatePairedLocations(forCancellation, Object.values(LOCATION_PAIRS));

//   return forCancellation;
// }

// // Helper function to check if a machine should be included for cancellation
// function shouldIncludeForCancellation(machineName, amountValue, tomorrowDate, collectionSchedule, todayDate, tomorrowDateString, lastRequest) {
//   const collectionDay = dayMapping[tomorrowDate.getDay()];

//   if (isTomorrowHoliday(tomorrowDate)) return true;

//   // Check for exclusion conditions in lastRequest
//   if (lastRequest) {
//     const hasExclusionCondition = EXCLUSION_CONDITIONS.some(condition => {
//       const searchText = condition === 'for collection on' || condition === 'resume collection on'
//         ? `${condition} ${tomorrowDateString.toLowerCase()}`
//         : condition;
//       return lastRequest.toLowerCase().includes(searchText);
//     });

//     if (hasExclusionCondition) {
//       return false;
//     }
//   }

//   // Check if machine meets cancellation criteria
//   return collectionSchedule.includes(collectionDay) && amountValue < amountThresholds[collectionDay];
// }

// // Helper function to validate paired locations
// function validatePairedLocations(forCancellation, locationPairs) {
//   // Create a map for faster lookups
//   const cancellationMap = new Map(
//     forCancellation.map(item => [item.name, item])
//   );

//   // Process each location pair
//   locationPairs.forEach(([location1, location2]) => {
//     const loc1 = cancellationMap.get(location1);
//     const loc2 = cancellationMap.get(location2);

//     if (loc1 && !loc2) {
//       cancellationMap.set(location2, {
//         name: location2,
//         address: loc1.address,
//         currentCashAmount: loc1.currentCashAmount
//       });
//     } else if (loc2 && !loc1) {
//       cancellationMap.set(location1, {
//         name: location1,
//         address: loc2.address,
//         currentCashAmount: loc2.currentCashAmount
//       });
//     }
//   });

//   // Convert map back to array
//   return Array.from(cancellationMap.values());
// }

// // Helper function to process cancellations and send emails
// function processCancellation(forCancellation, tomorrowDate, emailTo, emailCc, emailBcc, srvBank) {
//   if (forCancellation.length === 0) return;

//   const isSaturday = tomorrowDate.getDay() === 6;

//   // Sort once for better performance
//   forCancellation.sort((a, b) => a.name.localeCompare(b.name));

//   if (isSaturday) {
//     // Remove PLDT LAOAG for Saturday collections
//     forCancellation = forCancellation.filter(item => item.name !== SPECIAL_LOCATIONS.PLDT_LAOAG);

//     // Schedule Monday collection
//     const collectionDateForMonday = new Date(tomorrowDate);
//     collectionDateForMonday.setDate(collectionDateForMonday.getDate() + 2);
//     sendEmailCancellation(forCancellation, collectionDateForMonday, emailTo, emailCc, emailBcc, srvBank, isSaturday);
//   } else {
//     sendEmailCancellation(forCancellation, tomorrowDate, emailTo, emailCc, emailBcc, srvBank, isSaturday);
//   }
// }

// // Helper function to handle errors
// function handleError(error, functionName) {
//   const errorMessage = `Error in ${functionName}: ${error.message}`;
//   Logger.log(errorMessage);

//   // You could add additional error handling here, such as:
//   // - Sending error notifications
//   // - Logging to a monitoring service
//   // - Retrying the operation

//   // For now, we'll just log the error
//   console.error(errorMessage);
// }