function lookupLastTwoEntries() {
  // Define the sheet names
  const formResponsesSheetName = "Form Responses";
  const kioskSheetName = "Kiosk %";

  // Define the column indices (adjust if necessary)
  const machineNameColumnInKiosk = 1; // Assuming Machine Name is in column A of Kiosk %
  const machineNameColumnInFormResponses = 3; // Assuming Machine Name is in column A of Form Responses
  const dateColumnInFormResponses = 1; // Assuming Date or Timestamp is in column B of Form Responses
  const outputColumnInKiosk = 18; // Define where you want to output the concatenated result in Kiosk % (e.g., column B)
  const dataColumnA = 1; // Column D in Form Responses (4th column)
  const dataColumnD = 4; // Column D in Form Responses (4th column)
  const dataColumnE = 5; // Column E in Form Responses (5th column)

  // Get the sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = ss.getSheetByName(formResponsesSheetName);
  const kioskSheet = ss.getSheetByName(kioskSheetName);

  // Get all machine names from Kiosk %
  const kioskData = kioskSheet.getRange(2, machineNameColumnInKiosk, kioskSheet.getLastRow() - 1, 1).getValues();

  // Get all data from Form Responses
  const formResponsesData = formResponsesSheet.getDataRange().getValues();

  // Loop through each machine in Kiosk % and find last two entries in Form Responses
  kioskData.forEach((row, index) => {
    const machineName = row[0];

    // Filter entries in Form Responses that match the current machine name
    const matchingEntries = formResponsesData
      .filter(entry => entry[machineNameColumnInFormResponses - 1] === machineName)
      .map(entry => {
        // Convert entry[dataColumnA - 1] to a Date object and format it as MM/dd
        const formattedDate = Utilities.formatDate(new Date(entry[dataColumnA - 1]), Session.getScriptTimeZone(), "MM/dd");

        // Concatenate formatted date, columns D and E
        return {
          date: new Date(entry[dateColumnInFormResponses - 1]), // For sorting purposes
          data: formattedDate + " - " + entry[dataColumnD - 1] + " " + entry[dataColumnE - 1]
        };
      });

    // Sort the entries by date in descending order (latest date first)
    matchingEntries.sort((a, b) => b.date - a.date);

    // Get the data for the last two entries after sorting
    const lastTwoEntries = matchingEntries.slice(0, 2).map(entry => entry.data).join(" | ");

    // Write the result in the output column of Kiosk %
    kioskSheet.getRange(index + 2, outputColumnInKiosk).setValue(lastTwoEntries);
  });
}
