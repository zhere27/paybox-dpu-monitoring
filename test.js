function test_excludePastRequests() {
  var forCollections = [
    ["SMART SM BATAAN", 963150.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["PLDT SAN FERNANDO LA UNION", 393850.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT TALISAY", 385390.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["SMART SF LA UNION 1", 496100.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT OLONGAPO", 378970.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 22"],
    ["PLDT SM NORTH EDSA", 428880.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART FESTIVAL MALL 2", 454150.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 22"],
    ["SMART ROBINSONS GAPAN", 370020.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 22"],
    ["PLDT MANDAUE", 343460.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART SM CLARK 2", 500400.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["PLDT SBMA", 543400.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT PASAY", 350100.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT MANAOAG", 276200.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["SMART MARKET MARKET", 427810.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART AYALA CENTER CEBU", 411120.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT JONES 1", 354140.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART SM CLARK 1", 364840.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["SMART SM CEBU", 376390.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT DASMARINAS", 311930.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART IBA ZAMBALES", 356110.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART SM TARLAC", 251710.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["PLDT GAPAN", 100.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"]
  ];

  forCollections = excludePastRequests(forCollections, 'Jul 23');
  // forCollections = excludePreviouslyRequested(forCollections, 'Brinks via BPI');

  for (let i = 0; i < forCollections.length; i++) {
    console.log(forCollections[i][0]);
  }
}

function test_excludeRecentlyCollected(forCollections, srvBank) {
  var forCollections = [
    ["SMART SM BATAAN", 963150.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["PLDT SAN FERNANDO LA UNION", 393850.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT TALISAY", 385390.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["SMART SF LA UNION 1", 496100.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT OLONGAPO", 378970.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 22"],
    ["PLDT SM NORTH EDSA", 428880.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART FESTIVAL MALL 2", 454150.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 22"],
    ["SMART ROBINSONS GAPAN", 370020.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 22"],
    ["PLDT MANDAUE", 343460.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART SM CLARK 2", 500400.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["PLDT SBMA", 543400.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT PASAY", 350100.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT MANAOAG", 276200.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["SMART MARKET MARKET", 427810.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART AYALA CENTER CEBU", 411120.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT JONES 1", 354140.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART SM CLARK 1", 364840.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["SMART SM CEBU", 376390.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["PLDT DASMARINAS", 311930.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART IBA ZAMBALES", 356110.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", ""],
    ["SMART SM TARLAC", 251710.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"],
    ["PLDT GAPAN", 100.0, "Brinks via BPI", "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)", "For collection on Jul 23"]
  ];

  forCollections = excludeRecentlyCollected(forCollections, 'Brinks via BPI');

  for (let i = 0; i < forCollections.length; i++) {
    console.log(forCollections[i][0]);
  }
}

// function isTomorrowHoliday(tomorrow) {
//   const sheetName = "StoreName Mapping";
//   const rangeStart = "G3";
//   const rangeEnd = "I";
//   const validHolidayTypes = ["Regular Holiday", "Special Non-working Holiday"];

//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
//   if (!sheet) {
//     throw new Error(`Sheet named "${sheetName}" not found.`);
//   }

//   const dataRange = sheet.getRange(`${rangeStart}:${rangeEnd}`);
//   const data = dataRange.getValues();
//   tomorrow.setDate(tomorrow.getDate());
//   tomorrow.setHours(0, 0, 0, 0); // Ensure only the date is compared

//   for (const row of data) {
//     const holidayDate = row[0];
//     const holidayType = row[2];

//     if (holidayDate instanceof Date && holidayType && validHolidayTypes.includes(holidayType)) {
//       if (holidayDate.getTime() === tomorrow.getTime()) {
//         return true;
//       }
//     }
//   }

//   return false;
// }



function test() {
  // console.log(isTomorrowHoliday(new Date('2025-04-17')));
  const environment = "testing"; // or "testing"
  console.log(ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing);
}




