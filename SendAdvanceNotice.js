// // Configuration
// const CONFIG = {
//   spreadsheetId: "1TJ10XqwS_cTQfkxKKWJaE5zhBdE2pDgIdZ_Zw9JQD_U", // Your Google Sheet ID
//   sheetName: "Store Reps", // Name of your sheet tab
//   machineNameColumn: "A", // Column with machine names
//   emailColumn: "B", // Column with email addresses
//   startRow: 2, // Row where data starts (2 if you have headers)
//   environment: "testing",
// };

// function sendAdvancedNotice() {
//   CustomLogger.logInfo("Running sending DPU advanced notice...", PROJECT_NAME, "sendAdvancedNotice()");
//   var srvBank = "";
//   const { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();
//   const collectionDay = dayMapping[tomorrowDate.getDay()];
//   const amountThreashold = amountThresholds[collectionDay];
//   const kioskData = getKioskPercentage(amountThreashold);
//   const [machineNames, percentValues, amountValues, collectionPartners, collectionSchedules, , lastRequests, businessDays] = kioskData;
//   const formattedDate = Utilities.formatDate(tomorrowDate, Session.getScriptTimeZone(), "MMMM d, yyyy (EEEE)");
//   const inAdvanceNotice = forExclusionPartOfAdvanceNotice();

//   var emailTo = "";
//   const pldtRecepient = "mvolbara@pldt.com.ph";
//   const smartRecepient = "RBEspayos@smart.com.ph, RACagbay@smart.com.ph";
//   let emailCc = "";
//   let emailBcc = "egalcantara@multisyscorp.com";

//   machineNames.forEach((machineNameArr, i) => {
//     const machineName = machineNameArr[0];
//     const amountValue = amountValues[i][0];
//     const percentValue = percentValues[i][0];
//     const collectionPartner = collectionPartners[i][0];
//     const lastRequest = lastRequests[i][0];
//     const businessDay = businessDays[i][0];
//     const translatedBusinessDays = translateDaysToAbbreviation(businessDay.trim());

//     if (inAdvanceNotice && inAdvanceNotice.includes(machineName)) {
//       CustomLogger.logInfo(`Skipping ${machineName} for advance notice, already in advance notice...`, PROJECT_NAME, "sendAdvancedNotice()");
//       return;
//     }

//     if (forExclusionBasedOnRemarks(lastRequest, todayDateString, machineName)) {
//       return;
//     }

//     if (CONFIG.environment === "production") {
//       emailTo = getEmailByMachineName(machineName);
//       if (machineName.includes("PLDT")) {
//         emailCc = pldtRecepient;
//       } else if (machineName.includes("SMART")) {
//         emailCc = smartRecepient;
//       }
//     } else {
//       emailTo = "egalcantara@multisyscorp.com";
//       emailCc = "zhere27@gmail.com";
//     }

//     if (shouldIncludeForCollection(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank)) {
//       if (emailTo !== "") {
//         if (CONFIG.environment === "production") {
//           CustomLogger.logInfo(`Sending advance notice to ${machineName} to email addresses: ${emailTo}...`, PROJECT_NAME, "sendAdvancedNotice()");
//           addMachineToAdvanceNotice(machineName);
//           sendEmail(machineName, formattedDate, percentValue);
//         } else {
//           addMachineToAdvanceNotice(machineName);
//           CustomLogger.logInfo(`[TESTING] advance notice for ${machineName} to email addresses: ${emailTo}...`, PROJECT_NAME, "sendAdvancedNotice()");
//         }
//       }
//     }
//   });

//   function sendEmail(machineName, formattedDate, percentValue) {
//     const subject = `Paybox - DPU Collection advanced notice - ${formattedDate}`;

//     //percent value should be displayed as a percentage
//     percentValue = parseFloat(percentValue * 100.0).toFixed(2);

//     let body;

//     if (percentValue < 100) {
//       // Less than 100% capacity - Collection may take place
//       body = `Hi ${machineName},<br><br>
// The Paybox machine is at ${percentValue}% capacity. DPU collection has been requested and scheduled for pickup tomorrow ${formattedDate}.<br><br>
// Please have the required permit ready, if applicable. <br><br>
// Thank you.<br><br>

// *** This is an automated email. ****<br><br>${emailSignature}
// `;
//     } else {
//       // 100% or more capacity - Collection has been scheduled
//       body = `Hi ${machineName},<br><br>
// The Paybox machine is due for collection and has been requested and scheduled for pickup tomorrow ${formattedDate}.<br><br>
// Please have the required permit ready, if applicable.<br><br>
// Thank you.<br><br>

// *** This is an automated email. ****<br><br>${emailSignature}
// `;
//     }

//     GmailApp.sendEmail(emailTo, subject, "", {
//       cc: emailCc,
//       bcc: emailBcc,
//       htmlBody: body,
//       from: "support@paybox.ph",
//     });
//   }
// }

// /**
//  * Search for machine name and return corresponding email
//  */
// function getEmailByMachineName(machineName) {
//   const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
//   const sheet = ss.getSheetByName(CONFIG.sheetName);

//   if (!sheet) {
//     throw new Error('Sheet "' + CONFIG.sheetName + '" not found');
//   }

//   // Get all data from the sheet
//   const dataRange = sheet.getDataRange();
//   const data = dataRange.getValues();

//   // Search for the machine name (case-insensitive)
//   for (let i = CONFIG.startRow - 1; i < data.length; i++) {
//     const currentMachineName = data[i][0]; // Column A (index 0)

//     if (currentMachineName.toString().trim().toLowerCase() === machineName.toString().trim().toLowerCase()) {
//       const email = data[i][1]; // Column B (index 1)
//       return email;
//     }
//   }

//   return null; // Machine not found
// }

// /**
//  * Test function to verify the lookup works
//  */
// function testLookup() {
//   const testMachine = "SMART SM NAGA 2";
//   const email = getEmailByMachineName(testMachine);

//   if (email) {
//     Logger.log("Found: " + testMachine + " -> " + email);
//   } else {
//     Logger.log("Machine not found: " + testMachine);
//   }
// }
