function refresh(file) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");
  // const file = GDriveFilesAPI.getFirstFileInFolder('1knJiSYo3bO4B_V3BohCi5aBhFyhhXPSW');

  var filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  if (!file) {
    CustomLogger.logError(`No file found in the + Monitoring Hourly folder.`, PROJECT_NAME, "refresh()");
    EmailSender.sendLogs("egalcantara@multisyscorp.com", "Paybox - DPU Monitoring");
    return;
  }

  const dateParts = extractDateFromFilename(file.getName());
  let now = new Date(dateParts.year, dateParts.month - 1, dateParts.day, dateParts.hour, dateParts.minute);
  const lastRow = sheet.getLastRow();
  const sheetName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMdd HH00yy");

  //Populate newly created sheet
  try {
    populateNewSheet(sheetName, now);
    var dpuSheet = spreadSheet.getSheetByName(sheetName);
    if (dpuSheet) {
      dpuSheet.hideSheet();
    }
  } catch (e) {
    // Logger.log(e.message);
    CustomLogger.logError(e.message, PROJECT_NAME, "refresh()");
    EmailSender.sendLogs("egalcantara@multisyscorp.com", "Paybox - DPU Monitoring");
    return;
  }

  // Define formulas
  sheet.insertColumnsAfter(16, 1);
  sheet.deleteColumns(3, 1);

  const storesSheet = `${storeSheetUrl}`;
  const formulaSparkline = `=SPARKLINE(INDIRECT("C"&ROW()&":P"&ROW()))`;
  const formulaMachineStatus =`=VLOOKUP(INDIRECT("A"&ROW()),INDIRECT(CONCATENATE("'",TEXT($P$1,"MM"),TEXT($P$1,"DD")," ",TEXT($P$1,"HH"),"00",TEXT($P$1,"YY"),"'!A:D")),3,false)`;
  const formulaNoMovement = '=IF(INDIRECT("C"&ROW())<>"",IF(COUNTIF(INDIRECT("C"&ROW()&":P"&ROW()),"<>"& INDIRECT("C"&ROW()))=0, "No changes", ""),"")';
  const formulaCollectedStores = "=IFNA(VLOOKUP(A2,'Collected Yesterday'!B:B,1,false))";
  const formulaFrequency = `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),7,FALSE)`;
  const formulaServicingBank = `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),6,FALSE)`;
  const formulaBusinessDays = `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),3,FALSE)`;

  sheet.getRange(`Q2:Q${lastRow}`).clear();
  SpreadsheetApp.flush();
  CustomLogger.logInfo(`Cleared column Q (Current Cash Amount)`, PROJECT_NAME, "refresh()");

  const formulaMachineStatusRange = sheet.getRange(`S2:S${lastRow}`).clear().setFontFamily("Century Gothic").setFontSize(9).setFormula(formulaMachineStatus);

  const formulaSparklineRange = sheet.getRange(`T2:T${lastRow}`).clear().setFontFamily("Century Gothic").setFontSize(9).setFormula(formulaSparkline);

  const formulaNoMovementRange = sheet.getRange(`U2:U${lastRow}`).clear().setFontFamily("Century Gothic").setFontSize(9).setFormula(formulaNoMovement);

  const formulaCollectedStoresRange = sheet.getRange(`V2:V${lastRow}`).clear().setFontFamily("Century Gothic").setFontSize(9).setFormula(formulaCollectedStores);

  const formulaFrequencyRange = sheet.getRange(`W2:W${lastRow}`).clear().setFontFamily("Century Gothic").setFontSize(9).setFormula(formulaFrequency);

  const formulaServicingBankRange = sheet.getRange(`X2:X${lastRow}`).clear().setFontFamily("Century Gothic").setFontSize(9).setFormula(formulaServicingBank);

  const formulaBusinessDaysRange = sheet.getRange(`Y2:Y${lastRow}`).clear().setFontFamily("Century Gothic").setFontSize(9).setFormula(formulaBusinessDays);

  // Update main sheet
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:00:00");
  sheet.getRange("P1").setValue(formattedDate).setNumberFormat("MM/dd HH:00AM/PM").setBackground("#d9d9d9");
  sheet.getRange("Q1").setValue("Current Cash Amount").setBackground("#d9d9d9");
  sheet.getRange("R1").setValue("Remarks").setBackground("#d9d9d9");
  sheet.getRange("S1").setValue("Machine Status as of last update").setBackground("#d9d9d9");
  sheet.getRange("T1").setValue("Movements").setBackground("#d9d9d9"); //sparkline
  sheet.getRange("U1").setValue("No Movement for 5 days").setBackground("#d9d9d9");
  sheet.getRange("V1").setValue("Collected Stores").setBackground("#d9d9d9");
  sheet.getRange("W1").setValue("Frequency").setBackground("#d9d9d9");
  sheet.getRange("X1").setValue("Servicing Bank").setBackground("#d9d9d9");
  sheet.getRange("Y1").setValue("Business Days").setBackground("#d9d9d9");
  sheet.getRange("Z1").setValue("Jira Tickets No. (Issue Type: Machine Issues)").setBackground("#d9d9d9");

  //Set Percentage and Current Amount Formula
  const formulaPercentage = `=IFNA(TEXT(QUERY('${sheetName}'!$A:$D,"select D where A='" & TRIM(A2) & "'",0),"0.00%")/100,0)`;
  const formulaCurrentAmount = `=IFNA(QUERY('${sheetName}'!$A:$D,"select B where A='" & TRIM(A2) & "'",0),0)`;
  
  // Set formulas with formatting
  const formulaRange = sheet
    .getRange(`P2:P${lastRow}`)
    .setFontFamily("Century Gothic")
    .setFontSize(9)
    .setHorizontalAlignment("Center")
    .setVerticalAlignment("top")
    .setFormula(formulaPercentage)
    .setNumberFormat("0.00%");

  const formulaCurrentAmountRange = sheet
    .getRange(`Q2:Q${lastRow}`)
    .clear()
    .setFontFamily("Century Gothic")
    .setFontSize(9)
    .setFormula(formulaCurrentAmount)
    .setVerticalAlignment("top")
    .setNumberFormat("###,###,##0");

  CustomLogger.logInfo(`Applied formula for Percentage and Current Amount.`, PROJECT_NAME, "refresh()");
  SpreadsheetApp.flush();

  //Retrieve the collected stores from yesterday
  getCollectedStores();

  // Apply conditional formatting
  applyConditionalFormatting();

  hideColumnsCtoH();

  //Set row height and vertical alignment
  setRowsHeightAndAlignment(sheet, lastRow);

  //Sort the range based on the latest percentage
  sortLatestPercentage();

  SpreadsheetApp.flush();

  //Delete old sheets
  deleteOldSheet();

  //Hide all other sheets
  hideSheets();

  //Activate sheet Kiosk %
  sheet.activate();

  //Delete monitoring hourly file
  GDriveFilesAPI.deleteFileByFileId(file.getId());
}

function populateNewSheet(sheetName = null, now = null) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  // Manage sheets
  const existingSheet = spreadSheet.getSheetByName(sheetName);
  if (existingSheet) {
    spreadSheet.deleteSheet(existingSheet);
    CustomLogger.logError(`Deleted existing sheet: ${sheetName}`, PROJECT_NAME, "populateNewSheet()");
  }

  // Insert the new sheet
  let newSheet;
  try {
    newSheet = spreadSheet.insertSheet(sheetName);
    newSheet.showSheet();
    spreadSheet.moveActiveSheet(10); // Move the new sheet to the 10th position
  } catch (e) {
    CustomLogger.logError(`Failed to create new sheet: ${sheetName}. ${e.message}`, PROJECT_NAME, "populateNewSheet()");
    throw new Error(`Failed to create new sheet: ${sheetName}. ${e.message}`);
  }

  if (!newSheet) {
    CustomLogger.logError(`Failed to create new sheet: ${sheetName}`, PROJECT_NAME, "populateNewSheet()");
    throw new Error(`Failed to create new sheet: ${sheetName}`);
  }

  try {
    const trnDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const hours = String(now.getHours()).padStart(2, "0");
    const lastMonitoringHours = getLastMonitoringHours(trnDate, `${hours}:00:00`);

    if (!lastMonitoringHours || lastMonitoringHours.length === 0) {
      CustomLogger.logError("No data to populate the new sheet.", PROJECT_NAME, "populateNewSheet()");
      throw new Error("No data to populate the new sheet.");
    }

    // Check the structure of the rows array
    if (lastMonitoringHours.length > 0 && lastMonitoringHours[0].f && lastMonitoringHours[0].f.length > 0) {
      // Populate the sheet with the mapped values
      newSheet.getRange(1, 1, lastMonitoringHours.length, lastMonitoringHours[0].f.length).setValues(lastMonitoringHours.map((row) => row.f.map((cell) => cell.v)));
    } else {
      CustomLogger.logError("Rows array is not in the expected format.", PROJECT_NAME, "populateNewSheet()");
      throw new Error("Rows array is not in the expected format.");
    }
  } catch (e) {
    CustomLogger.logError(`Error populating sheet: ${e.message}`, PROJECT_NAME, "populateNewSheet()");
    throw new Error(`Error populating sheet: ${e.message}`);
  }
}

function sortLatestPercentage() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");

  const range = sheet.getRange("A2:S");
  const columnToSortBy = 16; // Column to sort by (percentage)

  range.sort({ column: columnToSortBy, ascending: false });

  SpreadsheetApp.flush();

  CustomLogger.logInfo(`Sorted latest percentage.`, PROJECT_NAME, "sortLatestPercentage()");
}

function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kiosk %");
  const lastRow = sheet.getLastRow();

  if (!sheet) {
    CustomLogger.logError("Sheet 'Kiosk %' not found.", PROJECT_NAME, "applyConditionalFormatting()");
    Logger.log("Sheet 'Kiosk %' not found.");
    return;
  }

  // Clear existing conditional formatting rules
  sheet.setConditionalFormatRules([]);

  // Define ranges
  const range = sheet.getRange(`C2:P${lastRow}`); // Adjust the range as needed
  const rangeNoMovement = sheet.getRange(`U2:U${lastRow}`); // Adjust the range as needed
  const rangeMachineStatus = sheet.getRange(`S2:S${lastRow}`);
  const rangeCollected = sheet.getRange(`A2:A${lastRow}`);
  const rangeRCBC = sheet.getRange(`Q2:Q${lastRow}`);

  // Define conditional formatting rules for the main range
  const rules = [
    //Set color for 90%
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.9).setBackground("#cc0000").setFontColor("white").setRanges([range]).build(),

    //Set color for 80%
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.8).setBackground("#e69138").setFontColor("white").setRanges([range]).build(),

    //Set color for 70%
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.7).setBackground("#ffe599").setFontColor("black").setRanges([range]).build(),

    //Set color for 60%
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.6).setBackground("#fff2cc").setFontColor("black").setRanges([range]).build(),

    //Set color for 50%
    // SpreadsheetApp.newConditionalFormatRule()
    //   .whenNumberGreaterThan(0.5)
    //   .setBackground('#d9ead3')
    //   .setFontColor('black')
    //   .setRanges([range])
    //   .build()
  ];

  const rulesMachineStatus = [SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("OFFLINE").setBackground("white").setFontColor('red').setRanges([rangeMachineStatus]).build()];
 
  const rulesNoMovement = [SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("No changes").setBackground("#cc0000").setFontColor("white").setRanges([rangeNoMovement]).build()];

  const rulesCollected = [SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=A2=V2").setBackground("#ff9900").setRanges([rangeCollected]).build()];

  const rulesRCBC = [SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(X1="RCBC", Q1>180000)').setBackground("#ff9900").setRanges([rangeRCBC]).build()];

  // Combine all rules
  const allRules = rules.concat(rulesMachineStatus).concat(rulesNoMovement).concat(rulesCollected).concat(rulesRCBC);

  // Set combined conditional formatting rules
  sheet.setConditionalFormatRules(allRules);
  CustomLogger.logInfo(`Applied conditional formatting.`, PROJECT_NAME, "applyConditionalFormatting()");
}

function deleteOldSheet() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadSheet.getSheets();
  var startIndex = 23;

  if (startIndex >= sheets.length) {
    CustomLogger.logError(`Start index ${startIndex} is out of range.`, PROJECT_NAME, "deleteSheetsFromIndex23()");
    return;
  }

  // Delete sheets from the end to avoid index shifting
  for (var i = sheets.length - 1; i >= startIndex; i--) {
    var sheetName = sheets[i].getName();
    spreadSheet.deleteSheet(sheets[i]);
    CustomLogger.logInfo(`Deleted sheet: ${sheetName}`, PROJECT_NAME, "deleteSheetsFromIndex23()");
  }
}

/**
 * Sets the row height and vertical alignment for the specified sheet and last row.
 *
 * @param {*} sheet
 * @param {*} lastRow
 */
function setRowsHeightAndAlignment(sheet, lastRow) {
  // Set the row height to 30 for all rows at once
  {
    {
      sheet.setRowHeights(1, lastRow, 30);
    }
  }

  // Set the vertical alignment to Top for the specified range
  sheet.getRange(`A2:Q${lastRow}`).setVerticalAlignment("middle");

  // Set the vertical alignment to Top for the specified range
  sheet.getRange(`R2:Z${lastRow}`).setVerticalAlignment("top");
}
