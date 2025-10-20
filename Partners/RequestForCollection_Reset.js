/**
 * Optimized function to reset collection requests and revisits
 * Uses specific spreadsheet ID and batch operations for better performance
 * Handles both "for collection on" and "for revisit on" patterns
 */
function resetCollectionRequests() {
  const CONFIG = {
    spreadsheetId: '1YWYJl-0BOmfF-gf8FLpb0SPFG4xTnjhI5Ly77LA2UBc',
    sheetName: 'Kiosk %',
    daysBack: 2,
    patterns: {
      collection: 'for collection on ',
      revisit: 'for revisit on '
    }
  };
  
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet "${CONFIG.sheetName}" not found`);
  }
  
  const remarksIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()
    .flat()
    .findIndex(col => col === "Remarks") + 1;
  
  const { todayDate } = getTodayAndTomorrowDates();
  const kioskData = getMachineDataByPartner();
  const [machineNames, , , , , , lastRequests] = kioskData;

  const timeZone = Session.getScriptTimeZone();
  
  // Calculate target date (2 days prior to today)
  const targetDate = new Date(todayDate);
  targetDate.setDate(targetDate.getDate() - CONFIG.daysBack);
  const targetDateString = Utilities.formatDate(targetDate, timeZone, "MMM d");
  
  // Pre-compile search patterns for better performance
  const searchPatterns = [
    CONFIG.patterns.collection + targetDateString.toLowerCase(),
    CONFIG.patterns.revisit + targetDateString.toLowerCase()
  ];
  
  CustomLogger.logInfo(
    `Resetting entries matching: "${searchPatterns.join('" or "')}"`, 
    CONFIG.APP.NAME, 
    'resetCollectionRequests'
  );

  // Collect rows to clear in batch
  const rowsToClear = [];
  const clearDetails = [];
  
  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const lastRequest = normalizeSpaces(lastRequests[i][0]);
    
    if (!lastRequest) return;
    
    const lastRequestLower = lastRequest.toLowerCase();
    const matchedPattern = searchPatterns.find(pattern => 
      lastRequestLower.startsWith(pattern)
    );
    
    if (matchedPattern) {
      rowsToClear.push(i + 2); // +2 because data starts at row 2
      clearDetails.push({
        machineName,
        pattern: matchedPattern,
        row: i + 2
      });
    }
  });

  // Batch clear all matching cells
  if (rowsToClear.length > 0) {
    rowsToClear.forEach((rowNum, idx) => {
      sheet.getRange(rowNum, remarksIndex).clearContent();
      CustomLogger.logInfo(        `Cleared "${clearDetails[idx].pattern}" for ${clearDetails[idx].machineName} (Row ${rowNum})`,         CONFIG.APP.NAME,         'resetCollectionRequests'      );
    });
    
    CustomLogger.logInfo(      `Total cleared: ${rowsToClear.length} machine(s)`,       CONFIG.APP.NAME,       'resetCollectionRequests'    );
  } else {
    CustomLogger.logInfo(      `No machines found matching the target patterns for ${targetDateString}`,       CONFIG.APP.NAME,       'resetCollectionRequests'    );
  }
}



// /**
//  * Optimized function to reset collection requests
//  * Uses specific spreadsheet ID and batch operations for better performance
//  */
// function resetCollectionRequests() {

//   const CONFIG = {
//     spreadsheetId: '1YWYJl-0BOmfF-gf8FLpb0SPFG4xTnjhI5Ly77LA2UBc',
//     sheetName: 'Kiosk %'
//   };
  
//   const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
//   const sheet = ss.getSheetByName(CONFIG.sheetName);
//   const remarksIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat().findIndex(col => col === "Remarks") + 1;
  
//   if (!sheet) {
//     throw new Error(`Sheet "${CONFIG.sheetName}" not found`);
//   }
  
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
//   const kioskData = getMachineDataByPartner();
//   const [machineNames, percentValues, amountValues, collectionPartners, 
//          collectionSchedules, lastAddress, lastRequests, businessDays] = kioskData;

//   const timeZone = Session.getScriptTimeZone();
  
//   // Calculate yesterday's date (2 days prior to today)
//   const dayAfterYesterdayDate = new Date(todayDate);
//   dayAfterYesterdayDate.setDate(dayAfterYesterdayDate.getDate() - 2);
//   const yesterdayDateString = Utilities.formatDate(dayAfterYesterdayDate, timeZone, "MMM d");
  
//   const searchString = "for collection on " + yesterdayDateString.toLowerCase();
  
//   CustomLogger.logInfo(
//     `Resetting last request value [For collection on ${yesterdayDateString}]`, 
//     CONFIG.APP.NAME, 
//     'resetCollectionRequests'
//   );

//   // Collect rows to clear in batch
//   const rowsToClear = [];
//   const machineNamesToClear = [];
  
//   machineNames.forEach((machineNameArr, i) => {
//     const machineName = machineNameArr[0];
//     const lastRequest = normalizeSpaces(lastRequests[i][0]);
    
//     if (lastRequest && lastRequest.toLowerCase() === searchString) {
//       rowsToClear.push(i + 2); // +2 because data starts at row 2
//       machineNamesToClear.push(machineName);
//     }
//   });

//   // Batch clear all matching cells
//   if (rowsToClear.length > 0) {
//     rowsToClear.forEach((rowNum, idx) => {
//       sheet.getRange(rowNum, remarksIndex).clearContent();
//       CustomLogger.logInfo(
//         `Cleared last request value of ${machineNamesToClear[idx]} (Row ${rowNum})`, 
//         CONFIG.APP.NAME, 
//         'resetCollectionRequests'
//       );
//     });
    
//     CustomLogger.logInfo(
//       `Total cleared: ${rowsToClear.length} machine(s)`, 
//       CONFIG.APP.NAME, 
//       'resetCollectionRequests'
//     );
//   } else {
//     CustomLogger.logInfo(
//       `No machines found matching "${searchString}"`, 
//       CONFIG.APP.NAME, 
//       'resetCollectionRequests'
//     );
//   }
// }




// // function resetCollectionRequests() {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const sheet = ss.getSheetByName('Kiosk %');
// //   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
// //   const kioskData = getMachineDataByPartner();
// //   const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, lastAddress, lastRequests, businessDays] = kioskData;


// //   const timeZone = Session.getScriptTimeZone();

// //   const dayAafterYesterdayDate = new Date(todayDate); // 2 days prior
// //   dayAafterYesterdayDate.setDate(dayAafterYesterdayDate.getDate() - 2);
// //   const yesterdayDateString = Utilities.formatDate(dayAafterYesterdayDate, timeZone, "MMM d");

// //   CustomLogger.logInfo(`Resetting last request value [For collection on ${yesterdayDateString}]`, CONFIG.APP.NAME, 'resetCollectionRequests');

// //   machineNames.forEach((machineNameArr, i) => {
// //     const machineName = machineNameArr[0];
// //     const lastRequest = normalizeSpaces(lastRequests[i][0]);
    
// //     if (lastRequest !== '') {
// //       if (lastRequest.toLocaleLowerCase() === "for collection on " + yesterdayDateString.toLowerCase()) {
// //         sheet.getRange(`R${i + 2}`).setValue('');
// //         CustomLogger.logInfo(`Cleared last request value [${sheet.getRange(`R${i + 2}`).getValue()}] of ${machineName}`, CONFIG.APP.NAME, 'resetCollectionRequests');
// //       }
// //     }
// //   });

// // }
