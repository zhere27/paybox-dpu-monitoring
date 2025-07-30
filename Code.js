function onOpen() {
  createMenu();
}

function onInstall(e) {
  createMenu();
}

function createMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Monitoring")
    .addItem("Refresh Stores", "refreshStores")
    // .addItem('Process Hourly', 'processMonitoringHourly')
    // .addItem('Update Remarks','lookupLastTwoEntries')
    // .addItem('Sort', 'sortLatestPercentage')
    .addToUi();

  logger = console.log.bind(console); // Redirect console.log to Logger.log
}
