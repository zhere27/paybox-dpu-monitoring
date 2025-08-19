function resetCollectionRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Kiosk %');
  const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
  const kioskData = getMachineDataByPartner();
  const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, lastAddress, lastRequests, businessDays] = kioskData;


  const timeZone = Session.getScriptTimeZone();

  const dayAafterYesterdayDate = new Date(todayDate); // 2 days prior
  dayAafterYesterdayDate.setDate(dayAafterYesterdayDate.getDate() - 2);
  const yesterdayDateString = Utilities.formatDate(dayAafterYesterdayDate, timeZone, "MMM d");

  CustomLogger.logInfo(`Resetting last request value [For collection on ${yesterdayDateString}]`, PROJECT_NAME, 'resetCollectionRequests');

  machineNames.forEach((machineNameArr, i) => {
    const machineName = machineNameArr[0];
    const lastRequest = normalizeSpaces(lastRequests[i][0]);
    
    if (lastRequest !== '') {
      if (lastRequest.toLocaleLowerCase() === "for collection on " + yesterdayDateString.toLowerCase()) {
        sheet.getRange(`R${i + 2}`).setValue('');
        CustomLogger.logInfo(`Cleared last request value [${sheet.getRange(`R${i + 2}`).getValue()}] of ${machineName}`, PROJECT_NAME, 'resetCollectionRequests');
      }
    }
  });

}
