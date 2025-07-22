function getCollectedStores() {
  try {
    // Get or create the target sheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Collected Yesterday";
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      Logger.log(`Sheet '${sheetName}' created.`);
    }

    // Clear previous data
    sheet.getRange("A2:E300").clear();

    // SQL Query
    const QUERY = `
      SELECT DISTINCT created_at, machine_name, amount, count, collector_name
      FROM ms-paybox-prod-1.pldtsmart.collections
      WHERE DATE(created_at) >= DATE(CURRENT_DATE()) - 1
      ORDER BY created_at DESC
    `;

    // Execute query
    var queryResult = executeQueryAndWait(QUERY);
    if (!queryResult || !queryResult.rows) throw new Error("No data returned from the query.");

    // Process and populate data
    var rows = queryResult.rows;
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i].f;
      var formattedDate = formatTimestamp(row[0].v);
      setCellValue(sheet, i + 2, formattedDate, row[1].v, row[2].v, row[3].v, row[4].v);
    }

    CustomLogger.logInfo(`Populated collected stores.`, PROJECT_NAME, 'getCollectedStores()');

    // Ensure changes are visible in the sheet
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log(`Error in getCollectedStores: ${error.message}`);
  }
}

/**
 * Formats a timestamp (seconds since epoch) into a readable date string.
 * @param {number} timestamp - Timestamp in seconds since epoch.
 * @returns {string} - Formatted date string.
 */
function formatTimestamp(timestamp) {
  try {
    var dateValue = new Date(timestamp * 1000); // Convert seconds to milliseconds
    return Utilities.formatDate(dateValue, "UTC", "yyyy-MM-dd HH:mm:ss");
  } catch (error) {
    Logger.log(`Error formatting timestamp: ${error.message}`);
    return "";
  }
}

/**
 * Sets the value and formatting for a row in the target sheet.
 * @param {Sheet} sheet - The target sheet.
 * @param {number} rowIndex - The row index (1-based).
 * @param {string} date - Formatted date.
 * @param {string} machineName - Machine name.
 * @param {number} amount - Amount.
 * @param {number} count - Count.
 * @param {string} collectorName - Collector name.
 */
function setCellValue(sheet, rowIndex, date, machineName, amount, count, collectorName) {
  try {
    var row = [
      { value: date, alignment: "Left", format: "@STRING@" },
      { value: machineName, alignment: "Left", format: "@STRING@" },
      { value: amount, alignment: "Right", format: "###,###,##0.00" },
      { value: count, alignment: "Center", format: "###,###,##0" },
      { value: collectorName, alignment: "Left", format: "@STRING@" },
    ];

    row.forEach((cell, index) => {
      var range = sheet.getRange(rowIndex, index + 1);
      range.setValue(cell.value)
        .setFontFamily("Century Gothic")
        .setFontSize(9)
        .setHorizontalAlignment(cell.alignment)
        .setNumberFormat(cell.format);
    });
  } catch (error) {
    Logger.log(`Error setting cell value for row ${rowIndex}: ${error.message}`);
  }
}
