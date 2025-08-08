function createTimeDrivenTriggers() {
  // Deletes any existing triggers for the function to avoid duplicates
  deleteExistingTriggers('getCollectedStores');
  deleteExistingTriggers('processMonitoringHourly');
  deleteExistingTriggers('exportSheetAndSendEmail');
  deleteExistingTriggers('eTapCollectionsLogic');
  deleteExistingTriggers('bpiBrinkCollectionsLogic');
  deleteExistingTriggers('bpiCollectionsLogic');
  deleteExistingTriggers('apeirosCollectionsLogic');
  deleteExistingTriggers('resetCollectionRequests');
  

  deleteExistingTriggers('sendAdvancedNotice');

  // Create new time-driven triggers for the specified times
  ScriptApp.newTrigger('getCollectedStores')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('processMonitoringHourly')
    .timeBased()
    .everyDays(1)
    .atHour(11)
    .nearMinute(15)
    .create();

  ScriptApp.newTrigger('processMonitoringHourly')
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .nearMinute(15)
    .create();

  ScriptApp.newTrigger('processMonitoringHourly')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(15)
    .create();

  ScriptApp.newTrigger('exportSheetAndSendEmail')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  ScriptApp.newTrigger('bpiBrinkCollectionsLogic')
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .nearMinute(20)
    .create();

  ScriptApp.newTrigger('bpiCollectionsLogic')
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .nearMinute(20)
    .create();

  ScriptApp.newTrigger('eTapCollectionsLogic')
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .nearMinute(30)
    .create();

  ScriptApp.newTrigger('sendAdvancedNotice')
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .nearMinute(30)
    .create();

  ScriptApp.newTrigger('apeirosCollectionsLogic')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .create();

  ScriptApp.newTrigger('resetCollectionRequests')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(45)
    .create();

}

function deleteExistingTriggers(functionName) {
  // Gets all triggers for the current project
  var triggers = ScriptApp.getProjectTriggers();

  // Loops through all triggers and deletes those that call the specified function
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}