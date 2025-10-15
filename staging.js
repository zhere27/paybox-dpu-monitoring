function test_hasSpecialCollectionConditions() {
 console.log(hasSpecialCollectionConditions('For collection on Aug 25','Aug 23')); 
}
function test_excludeAlreadyCollected(machineName='PLDT SURIGAO', srvBank='Apeiros') {
  console.log(skipAlreadyCollected(machineName, srvBank));
}

function test_isTomorrowHoliday() {
  var { todayDate, tomorrowDate, todayDateString, tomorrowDateString } = getTodayAndTomorrowDates();

  console.log(isTomorrowHoliday(tomorrowDate));
}

function test_forExclusionRequestYesterday(machineName='SMART SM CEBU 2', srvBank='Brinks via BPI') {
  console.log(forExclusionRequestYesterday(machineName, srvBank))
}

function test_shouldExcludeFromCollection_stg() {
  const lastRequest = '';
  const todayDay = 'Aug 18';
  const machineName = 'PLDT TALISAY';
  const result = shouldExcludeFromCollection_stg(lastRequest, todayDay, machineName);
  console.log(`Result for "${lastRequest}" and "${todayDay}" on "${machineName}": ${result}`);
}

/**
 * Checks if a machine should be excluded from collection based on last request and current day
 * @param {string} lastRequest - Last request string to check against exclusion criteria
 * @param {string} todayDay - Today's day name (e.g., "monday", "tuesday")
 * @param {string} [machineName=null] - Optional machine name for logging purposes
 * @return {boolean} True if the machine should be excluded, false otherwise
 */
function shouldExcludeFromCollection_stg(lastRequest, todayDay, machineName = null) {
  try {
    // Early return for null/undefined/empty lastRequest
    if (!lastRequest?.trim()) {
      return false;
    }

    // Normalize inputs once
    const normalizedLastRequest = lastRequest.toLowerCase().trim();
    const normalizedTodayDay = todayDay?.toLowerCase().trim() || '';

    // Define exclusion criteria with better organization
    const exclusionReasons = [
      'waiting for collection items',
      'waiting for bank updates', 
      'did not reset',
      'cassette removed',
      'manually collected',
      'for repair',
      'already collected',
      'store is closed',
      'store is not using the machine',
      normalizedTodayDay
    ].filter(reason => reason); // Remove empty strings

    // Check for exclusion with optimized search
    const isExcluded = exclusionReasons.some(reason => 
      normalizedLastRequest.includes(reason)
    );

    // Log exclusion if machine name provided and excluded
    if (isExcluded && machineName) {
      CustomLogger.logInfo(`Machine "${machineName}" excluded - Last remarks: "${lastRequest}"`,PROJECT_NAME,'shouldExcludeFromCollection');
    }

    return isExcluded;

  } catch (error) {
    CustomLogger.logError(`Error in shouldExcludeFromCollection(): ${error.message}`,PROJECT_NAME,'shouldExcludeFromCollection');
    return false; // Fail-safe: don't exclude on error
  }
}

function test_meetsAmountThreshold(amountValue=151000, collectionDay, srvBank='Brinks via BPI') {
  console.log(meetsAmountThreshold(amountValue, collectionDay, srvBank));
}

function test_addMachineToAdvanceNotice(machineName='PLDT SURIGAO') {
  addMachineToAdvanceNotice(machineName);
}
