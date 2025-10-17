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
