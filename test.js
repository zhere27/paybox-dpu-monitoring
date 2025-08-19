function test_shouldIncludeForCollection() {
  const machineName="PLDT ROBINSONS DUMAGUETE";
  const amountValue=250000;
  const collectionSchedule=["M.W.Sat."];
  const tomorrowDate=new Date(2025,7,17);
  const tomorrowDateString="Aug 18";
  const todayDate="Aug 17";
  const lastRequest="";
  const srvBank="eTap";
  const translatedBusinessDays = translateDaysToAbbreviation("Monday - Sunday");

  if (shouldIncludeForCollection_1(machineName, amountValue, translatedBusinessDays, tomorrowDate, tomorrowDateString, todayDate, lastRequest, srvBank, collectionSchedule)) {
    console.log("should include for collection");
  }
} 

/**
 * Main function to determine if a machine should be included for collection
 * @param {string} machineName - The name of the machine
 * @param {number} amountValue - The amount in the machine
 * @param {Array} translatedBusinessDays - Array of translated business days
 * @param {Date} tomorrowDate - Tomorrow's date object
 * @param {string} tomorrowDateString - Tomorrow's date as string
 * @param {Date} todayDate - Today's date object
 * @param {string} lastRequest - The last request string
 * @param {string} srvBank - The bank service name
 * @returns {boolean} - True if machine should be included for collection
 */
function shouldIncludeForCollection_1(  machineName,   amountValue,   translatedBusinessDays,   tomorrowDate,   tomorrowDateString,   todayDate,   lastRequest,   srvBank, collectionSchedule
) {
  try {
    const collectionDay = dayMapping[tomorrowDate.getDay()];
    const dayOfWeek = tomorrowDate.getDay();

    // Early returns for exclusion conditions
    
    // 1. Check if excluded store
    if (isExcludedStore(machineName)) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString}, part of the excluded stores.`, 
        PROJECT_NAME, 
        "shouldIncludeForCollection"
      );
      return false;
    }

    // 2. Check collection schedule
    if (!translatedBusinessDays.includes(collectionDay)) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString}, not a collection day.`, 
        PROJECT_NAME, 
        "shouldIncludeForCollection"
      );
      return false;
    }


    // 3. Check BPI weekend restrictions
    if (shouldSkipWeekendCollections(srvBank, tomorrowDate)) {
      CustomLogger.logInfo(
        `Skipping collection for ${machineName} on ${tomorrowDateString}, during weekends.`, 
        PROJECT_NAME, 
        "shouldIncludeForCollection"
      );
      return false;
    }

    // Early returns for inclusion conditions (highest priority)
    
    // 1. Special collection conditions from last request
    if (hasSpecialCollectionConditions(lastRequest, tomorrowDateString)) {
      return true;
    }

    // 2. Schedule-based store collections
    if (SCHEDULE_CONFIG.stores.includes(machineName)) {
      return shouldCollectScheduleBasedStore(machineName, dayOfWeek);
    }

    // 3. Regular threshold-based collection
    if (meetsAmountThreshold(amountValue, collectionDay)) {
      return true;
    }

    // 4. Payday collection
    if (isPaydayCollection(todayDate, amountValue)) {
      return true;
    }

    // 5. Due date collection
    if (isDueDateCollection(todayDate, amountValue)) {
      return true;
    }

    // Default: no collection needed
    return false;

  } catch (error) {
    CustomLogger.logError(
      `Error in shouldIncludeForCollection(): ${error.message}\nStack: ${error.stack}`, 
      PROJECT_NAME, 
      "shouldIncludeForCollection()"
    );
    return false;
  }
}

function test_excludePastRequests() {
  var forCollections = [
    [
      "SMART SM BATAAN",
      963150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SAN FERNANDO LA UNION",
      393850.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT TALISAY",
      385390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SF LA UNION 1",
      496100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT OLONGAPO",
      378970.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT SM NORTH EDSA",
      428880.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART FESTIVAL MALL 2",
      454150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "SMART ROBINSONS GAPAN",
      370020.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT MANDAUE",
      343460.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 2",
      500400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SBMA",
      543400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT PASAY",
      350100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT MANAOAG",
      276200.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART MARKET MARKET",
      427810.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART AYALA CENTER CEBU",
      411120.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT JONES 1",
      354140.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 1",
      364840.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SM CEBU",
      376390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT DASMARINAS",
      311930.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART IBA ZAMBALES",
      356110.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM TARLAC",
      251710.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT GAPAN",
      100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
  ];

  forCollections = excludePastRequests(forCollections, "Jul 23");
  // forCollections = excludePreviouslyRequested(forCollections, 'Brinks via BPI');

  for (let i = 0; i < forCollections.length; i++) {
    console.log(forCollections[i][0]);
  }
}

function test_excludeRecentlyCollected(forCollections, srvBank) {
  var forCollections = [
    [
      "SMART SM BATAAN",
      963150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SAN FERNANDO LA UNION",
      393850.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT TALISAY",
      385390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SF LA UNION 1",
      496100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT OLONGAPO",
      378970.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT SM NORTH EDSA",
      428880.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART FESTIVAL MALL 2",
      454150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "SMART ROBINSONS GAPAN",
      370020.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT MANDAUE",
      343460.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 2",
      500400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SBMA",
      543400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT PASAY",
      350100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT MANAOAG",
      276200.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART MARKET MARKET",
      427810.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART AYALA CENTER CEBU",
      411120.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT JONES 1",
      354140.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 1",
      364840.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SM CEBU",
      376390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT DASMARINAS",
      311930.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART IBA ZAMBALES",
      356110.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM TARLAC",
      251710.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT GAPAN",
      100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
  ];

  forCollections = excludeRecentlyCollected(forCollections, "Brinks via BPI");

  for (let i = 0; i < forCollections.length; i++) {
    console.log(forCollections[i][0]);
  }
}

function test_createHiddenWorksheetAndAddData(forCollections, srvBank) {
  var forCollections = [
    [
      "SMART SM BATAAN",
      963150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SAN FERNANDO LA UNION",
      393850.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT TALISAY",
      385390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SF LA UNION 1",
      496100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT OLONGAPO",
      378970.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT SM NORTH EDSA",
      428880.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART FESTIVAL MALL 2",
      454150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "SMART ROBINSONS GAPAN",
      370020.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT MANDAUE",
      343460.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 2",
      500400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SBMA",
      543400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT PASAY",
      350100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT MANAOAG",
      276200.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART MARKET MARKET",
      427810.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART AYALA CENTER CEBU",
      411120.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT JONES 1",
      354140.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 1",
      364840.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SM CEBU",
      376390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT DASMARINAS",
      311930.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART IBA ZAMBALES",
      356110.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM TARLAC",
      251710.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT GAPAN",
      100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
  ];

  forCollections.sort(function (a, b) {
    const aAmount = parseFloat(a[1]);
    const bAmount = parseFloat(b[1]);
    if (isNaN(aAmount) && isNaN(bAmount)) return 0;
    if (isNaN(aAmount)) return 1;
    if (isNaN(bAmount)) return -1;
    return bAmount - aAmount;
  });

  createHiddenWorksheetAndAddData(forCollections, "Brinks via BPI");
}

function test_getTodayAndTomorrowDates() {
  const today = getTodayAndTomorrowDates();
  console.log(today[0]);
  console.log(today[1]);
}

function test_isTomorrowHoliday(tomorrow) {
  console.log(isTomorrowHoliday(tomorrow));
}

function test() {
  // console.log(isTomorrowHoliday(new Date('2025-04-17')));
  const environment = "testing"; // or "testing"
  console.log(
    ENV_EMAIL_RECIPIENTS[environment] || ENV_EMAIL_RECIPIENTS.testing
  );
}

function test_skipFunction() {
  skipFunction();
}

function test_getLastHourly() {
  getLastHourly();
}

function test_setRowsHeightAndAlignment() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");

  const lastRow = sheet.getLastRow();

  setRowsHeightAndAlignment(sheet, lastRow);
}

function test_dryer_sorter() {
  var forCollections = [
    [
      "SMART SM BATAAN",
      963150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SAN FERNANDO LA UNION",
      393850.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT TALISAY",
      385390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SF LA UNION 1",
      496100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT OLONGAPO",
      378970.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT SM NORTH EDSA",
      428880.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART FESTIVAL MALL 2",
      454150.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "SMART ROBINSONS GAPAN",
      370020.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 22",
    ],
    [
      "PLDT MANDAUE",
      343460.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 2",
      500400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT SBMA",
      543400.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT PASAY",
      350100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT MANAOAG",
      276200.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART MARKET MARKET",
      427810.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART AYALA CENTER CEBU",
      411120.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT JONES 1",
      354140.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM CLARK 1",
      364840.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "SMART SM CEBU",
      376390.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "PLDT DASMARINAS",
      311930.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART IBA ZAMBALES",
      356110.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "",
    ],
    [
      "SMART SM TARLAC",
      251710.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
    [
      "PLDT GAPAN",
      100.0,
      "Brinks via BPI",
      "Brinks via BPI DPU Request - July 23, 2025 (Wednesday)",
      "For collection on Jul 23",
    ],
  ];
  srvBank="Brinks via BPI";

  // Helper: Sort by numeric value (desc), pushing NaN to the end
  const sortByAmountDescNaNLast = (a, b, index) => {
    const aAmount = Number(a[index]);
    const bAmount = Number(b[index]);

    if (Number.isNaN(aAmount) && Number.isNaN(bAmount)) return 0;
    if (Number.isNaN(aAmount)) return 1;
    if (Number.isNaN(bAmount)) return -1;

    return bAmount - aAmount; // Descending
  };

  // Helper: Sort by string (case-sensitive) at given index
  const sortByString = (a, b, index) => a[index].localeCompare(b[index]);

  // Main logic
  forCollections.sort(
    srvBank === "Brinks via BPI"
      ? (a, b) => sortByAmountDescNaNLast(a, b, 2)
      : (a, b) => sortByString(a, b, 0)
  );

    const jsonResult = convertArrayToJson(forCollections);
  Logger.log(JSON.stringify(jsonResult, null, 2)); // Pretty print
}


function test_shouldCollectScheduleBasedStore(machineName='PLDT ROBINSONS DUMAGUETE', dayOfWeek=1){
  var retVal = shouldCollectScheduleBasedStore(machineName, dayOfWeek);
  console.log(retVal);
}