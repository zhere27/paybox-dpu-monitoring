// function processCancellationLogic(srvBank) {
//   // Open the active spreadsheet
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   // Access the "For Collections" worksheet
//   const sheet = ss.getSheetByName("For Collection -" + srvBank);
//   if (!sheet) {
//     console.log("Sheet 'For Collections' not found.");
//     return;
//   }

//   // Get all the data in the sheet
//   const data = sheet.getDataRange().getValues();

//   // Store the results
//   const results = [];
//   var subject = '';

//   // Iterate through rows, starting from row 2 (skip the header)
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const machineName = row[0]; // Column A (index 0)
//     const emailSubject = row[3]; // Column D (index 3)
//     const isCollected = row[5]; // Column F (index 5)

//     subject = emailSubject;

//     const doNotCancel = [
//       'PLDT ROBINSONS DUMAGUETE',
//       'PLDT BANTAY',
//       'SMART VIGAN'
//     ];
//     if (doNotCancel.some(store => machineName.includes(store))) { continue; }

//     // Check if Column F (isCollected) is TRUE
//     if (machineName && isCollected === true) {
//       results.push({ machineName, emailSubject });
//     }
//   }

//   if (results.length > 0) {
//     var machineData = results.map(row => row.machineName);

//     var body = `
//     Hi All,<br><br>
//     Good day! Please <b>CANCEL</b> the collection for the following stores:<br><br>
//     ${machineData.map((machine) => machine).join('<br>')}<br>
//     *** Please acknowledge this email. ****<br><br>${emailSignature}
//   `;

//     replyToExistingThread(results[0].emailSubject, body);
//   }

// }