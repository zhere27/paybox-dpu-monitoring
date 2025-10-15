// ==================== OPTIMIZED VERSION ====================

function refresh(file) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");

  // Remove existing filter
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  if (!file) {
    CustomLogger.logError(`No file found in the + Monitoring Hourly folder.`, PROJECT_NAME, "refresh()");
    EmailSender.sendLogs("egalcantara@multisyscorp.com", "Paybox - DPU Monitoring");
    return;
  }

  const dateParts = extractDateFromFilename(file.getName());
  const now = new Date(dateParts.year, dateParts.month - 1, dateParts.day, dateParts.hour, dateParts.minute);
  const lastRow = sheet.getLastRow();
  const sheetName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMdd HH00yy");

  // Populate newly created sheet
  try {
    populateNewSheet(sheetName, now);
    const dpuSheet = spreadSheet.getSheetByName(sheetName);
    if (dpuSheet) dpuSheet.hideSheet();
  } catch (e) {
    CustomLogger.logError(e.message, PROJECT_NAME, "refresh()");
    EmailSender.sendLogs("egalcantara@multisyscorp.com", "Paybox - DPU Monitoring");
    return;
  }

  // Column management - single operation
  sheet.insertColumnsAfter(16, 1);
  sheet.deleteColumns(3, 1);

  // Setup formulas
  setupFormulas(sheet, lastRow, sheetName);

  // Update headers - batch operation
  updateHeaders(sheet, now);

  // Execute all remaining operations
  getCollectedStores();
  applyConditionalFormatting();
  hideColumnsCtoH();
  setRowsHeightAndAlignment(sheet, lastRow);
  sortLatestPercentage();

  SpreadsheetApp.flush(); // Single flush at the end

  // Cleanup operations
  deleteOldSheet();
  hideSheets();
  sheet.activate();
  GDriveFilesAPI.deleteFileByFileId(file.getId());
}

/**
 * Consolidated formula setup function
 * Reduces redundant code and batches operations
 */
function setupFormulas(sheet, lastRow, sheetName) {
  const storesSheet = `${storeSheetUrl}`;

  // Define all formulas
  const formulas = {
    sparkline: `=SPARKLINE(INDIRECT("D"&ROW()&":P"&ROW()))`,
    machineStatus: `=VLOOKUP(INDIRECT("A"&ROW()),INDIRECT(CONCATENATE("'",TEXT($P$1,"MM"),TEXT($P$1,"DD")," ",TEXT($P$1,"HH"),"00",TEXT($P$1,"YY"),"'!A:D")),3,false)`,
    noMovement: '=IF(INDIRECT("D"&ROW())<>"",IF(COUNTIF(INDIRECT("D"&ROW()&":P"&ROW()),"<>"& INDIRECT("D"&ROW()))=0, "No changes", ""),"")',
    collectedStores: "=IFNA(VLOOKUP(A2,'Collected Yesterday'!B:B,1,false))",
    frequency: `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),7,FALSE)`,
    servicingBank: `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),6,FALSE)`,
    businessDays: `=VLOOKUP(A2,IMPORTRANGE("${storesSheet}","Stores!C:X"),3,FALSE)`,
    percentage: `=IFNA(TEXT(QUERY('${sheetName}'!$A:$D,"select D where A='" & TRIM(A2) & "'",0),"0.00%")/100,0)`,
    currentAmount: `=IFNA(QUERY('${sheetName}'!$A:$D,"select B where A='" & TRIM(A2) & "'",0),0)`
  };

  // Batch apply formulas with formatting
  const rangeConfigs = [
    { range: `P2:P${lastRow}`, formula: formulas.percentage, format: '0.00%', align: 'Center' },
    { range: `Q2:Q${lastRow}`, formula: formulas.currentAmount, format: '###,###,##0', clear: true },
    { range: `S2:S${lastRow}`, formula: formulas.machineStatus },
    { range: `T2:T${lastRow}`, formula: formulas.sparkline },
    { range: `U2:U${lastRow}`, formula: formulas.noMovement },
    { range: `V2:V${lastRow}`, formula: formulas.collectedStores },
    { range: `W2:W${lastRow}`, formula: formulas.frequency },
    { range: `X2:X${lastRow}`, formula: formulas.servicingBank },
    { range: `Y2:Y${lastRow}`, formula: formulas.businessDays }
  ];

  // Apply all formulas in batch
  rangeConfigs.forEach(config => {
    const range = sheet.getRange(config.range);
    if (config.clear) range.clear();

    range
      .setFontFamily("Century Gothic")
      .setFontSize(9)
      .setFormula(config.formula)
      .setVerticalAlignment("top");

    if (config.format) range.setNumberFormat(config.format);
    if (config.align) range.setHorizontalAlignment(config.align);
  });

  CustomLogger.logInfo(`Applied all formulas for columns P-Y.`, PROJECT_NAME, "setupFormulas()");
}

/**
 * Batch update all headers at once
 */
function updateHeaders(sheet, now) {
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:00:00");

  const headers = [
    { cell: 'P1', value: formattedDate, format: 'MM/dd HH:00AM/PM' },
    { cell: 'Q1', value: 'Current Cash Amount' },
    { cell: 'R1', value: 'Remarks' },
    { cell: 'S1', value: 'Machine Status as of last update' },
    { cell: 'T1', value: 'Movements' },
    { cell: 'U1', value: 'No Movement for 5 days' },
    { cell: 'V1', value: 'Collected Stores' },
    { cell: 'W1', value: 'Frequency' },
    { cell: 'X1', value: 'Servicing Bank' },
    { cell: 'Y1', value: 'Business Days' },
    { cell: 'Z1', value: 'Jira Tickets No. (Issue Type: Machine Issues)' }
  ];

  headers.forEach(h => {
    const range = sheet.getRange(h.cell).setValue(h.value).setBackground("#d9d9d9");
    if (h.format) range.setNumberFormat(h.format);
  });
}

function populateNewSheet(sheetName = null, now = null) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  // Delete existing sheet if present
  const existingSheet = spreadSheet.getSheetByName(sheetName);
  if (existingSheet) {
    spreadSheet.deleteSheet(existingSheet);
    CustomLogger.logError(`Deleted existing sheet: ${sheetName}`, PROJECT_NAME, "populateNewSheet()");
  }

  // Create new sheet
  let newSheet;
  try {
    newSheet = spreadSheet.insertSheet(sheetName);
    newSheet.showSheet();

    const kioskPercentSheetIndex = spreadSheet.getSheetByName("Kiosk %").getIndex();
    spreadSheet.moveActiveSheet(kioskPercentSheetIndex);
  } catch (e) {
    const errorMsg = `Failed to create new sheet: ${sheetName}. ${e.message}`;
    CustomLogger.logError(errorMsg, PROJECT_NAME, "populateNewSheet()");
    throw new Error(errorMsg);
  }

  if (!newSheet) {
    const errorMsg = `Failed to create new sheet: ${sheetName}`;
    CustomLogger.logError(errorMsg, PROJECT_NAME, "populateNewSheet()");
    throw new Error(errorMsg);
  }

  try {
    const trnDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const hours = String(now.getHours()).padStart(2, "0");
    const lastMonitoringHours = getLastMonitoringHours(trnDate, `${hours}:00:00`);

    if (!lastMonitoringHours?.length) {
      throw new Error("No data to populate the new sheet.");
    }

    // Validate data structure
    if (lastMonitoringHours[0]?.f?.length > 0) {
      // Single batch write operation
      const values = lastMonitoringHours.map(row => row.f.map(cell => cell.v));
      newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    } else {
      throw new Error("Rows array is not in the expected format.");
    }
  } catch (e) {
    const errorMsg = `Error populating sheet: ${e.message}`;
    CustomLogger.logError(errorMsg, PROJECT_NAME, "populateNewSheet()");
    throw new Error(errorMsg);
  }
}

function sortLatestPercentage() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kiosk %");
  sheet.getRange("A2:S").sort({ column: 16, ascending: false });
  CustomLogger.logInfo(`Sorted latest percentage.`, PROJECT_NAME, "sortLatestPercentage()");
}

function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kiosk %");

  if (!sheet) {
    CustomLogger.logError("Sheet 'Kiosk %' not found.", PROJECT_NAME, "applyConditionalFormatting()");
    return;
  }

  const lastRow = sheet.getLastRow();

  // Clear existing rules once
  sheet.setConditionalFormatRules([]);

  // Define all ranges
  const ranges = {
    main: sheet.getRange(`C2:P${lastRow}`),
    noMovement: sheet.getRange(`U2:U${lastRow}`),
    machineStatus: sheet.getRange(`S2:S${lastRow}`),
    collected: sheet.getRange(`A2:A${lastRow}`),
    rcbc: sheet.getRange(`Q2:Q${lastRow}`)
  };

  // Build all rules in array
  const allRules = [
    // Main range rules
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.9)
      .setBackground("#cc0000")
      .setFontColor("white")
      .setRanges([ranges.main])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.8)
      .setBackground("#e69138")
      .setFontColor("white")
      .setRanges([ranges.main])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.7)
      .setBackground("#ffe599")
      .setFontColor("black")
      .setRanges([ranges.main])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.6)
      .setBackground("#fff2cc")
      .setFontColor("black")
      .setRanges([ranges.main])
      .build(),

    // Machine status rule
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OFFLINE")
      .setBackground("white")
      .setFontColor('red')
      .setRanges([ranges.machineStatus])
      .build(),

    // No movement rule
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("No changes")
      .setBackground("#cc0000")
      .setFontColor("white")
      .setRanges([ranges.noMovement])
      .build(),

    // Collected stores rule
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=A2=V2")
      .setBackground("#ff9900")
      .setRanges([ranges.collected])
      .build(),

    // RCBC rule
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(X1="RCBC", Q1>180000)')
      .setBackground("#ff9900")
      .setRanges([ranges.rcbc])
      .build()
  ];

  // Apply all rules at once
  sheet.setConditionalFormatRules(allRules);
  CustomLogger.logInfo(`Applied conditional formatting.`, PROJECT_NAME, "applyConditionalFormatting()");
}

function deleteOldSheet() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadSheet.getSheets();

  //get last sheet index
  const lastSheetIndex = sheets.length - 1;

  const startIndex = lastSheetIndex;

  if (startIndex >= sheets.length) {
    CustomLogger.logError(`Start index ${startIndex} is out of range.`, PROJECT_NAME, "deleteOldSheet()");
    return;
  }

  // Delete from end to avoid index shifting
  for (let i = sheets.length - 1; i >= startIndex; i--) {
    const sheetName = sheets[i].getName();
    spreadSheet.deleteSheet(sheets[i]);
    CustomLogger.logInfo(`Deleted sheet: ${sheetName}`, PROJECT_NAME, "deleteOldSheet()");
  }
}

/**
 * Sets row height and vertical alignment efficiently
 */
function setRowsHeightAndAlignment(sheet, lastRow) {
  // Single operation for row heights
  sheet.setRowHeights(1, lastRow, 30);

  // Batch alignment operations
  sheet.getRange(`A2:Q${lastRow}`).setVerticalAlignment("middle");
  sheet.getRange(`R2:Z${lastRow}`).setVerticalAlignment("top");
}