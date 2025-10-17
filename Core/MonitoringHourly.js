/**
 * Main function to process monitoring hourly data
 * Downloads logs, processes CSV files, and uploads data to BigQuery
 */
function processMonitoringHourly() {
  const environment = 'production'; //testing
  let processedCount = 0;
  let errorCount = 0;
  try {
    if (environment === 'production') {
      // Step 1: Download Paybox logs
      const downloadResult = runDownloadPayboxLogsMonitoringHourly();
      CustomLogger.logInfo(`Download completed: ${downloadResult.fileCount} files downloaded`, CONFIG.APP.NAME, 'processMonitoringHourly()');

      // Step 2: Process CSV files
      const folderCollections = DriveApp.getFolderById('1knJiSYo3bO4B_V3BohCi5aBhFyhhXPSW');
      const files = folderCollections.getFilesByType(MimeType.CSV);

      // Process each CSV file
      while (files.hasNext()) {
        try {
          const file = files.next();
          runLoadCsvFromDrive(file.getId());
          processedCount++;
        } catch (error) {
          errorCount++;
          CustomLogger.logError(`Error processing file: ${error.message}`, CONFIG.APP.NAME, 'processMonitoringHourly()');
        }
      }
    } else {
      //provide the fileid
      runLoadCsvFromDrive('1fVEIzPRsuCp15Lgzw3yphOtFEIBruwXG');
      processedCount++;
    }
    CustomLogger.logInfo(`Processing completed: ${processedCount} files processed, ${errorCount} errors`, CONFIG.APP.NAME, 'processMonitoringHourly()');
  } catch (error) {
    CustomLogger.logError(`Fatal error in processMonitoringHourly: ${error.message}`, CONFIG.APP.NAME, 'processMonitoringHourly()');
  }
  //Send Logs to Admin
  EmailSender.sendExecutionLogs(recipient = { to: CONFIG.APP.ADMIN.email }, CONFIG.APP.NAME);

}

/**
 * Loads and processes a CSV file from Google Drive
 * @param {string} fileId - The ID of the file to process
 */
function runLoadCsvFromDrive(fileId) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Fetch the CSV file content from Google Drive
    const file = DriveApp.getFileById(fileId);
    const csvContent = file.getBlob().getDataAsString();
    const fileName = file.getName();
    CustomLogger.logInfo(`Processing file: ${fileName}`, CONFIG.APP.NAME, 'runLoadCsvFromDrive()');

    // Extract and format sheet name from filename
    const sheetName = extractSheetNameFromFileName(fileName);
    if (!sheetName) {
      CustomLogger.logInfo(`Invalid filename format: ${fileName}`, CONFIG.APP.NAME, 'runLoadCsvFromDrive()');
      return;
    }

    // Check if sheet already exists
    const sheet = spreadSheet.getSheetByName(sheetName);
    if (sheet) {
      CustomLogger.logInfo(`Sheet '${sheetName}' already exists. Skipping.`, CONFIG.APP.NAME, 'runLoadCsvFromDrive()');
      return;
    }

    // Validate and process CSV content
    const csvString = validateLine(csvContent, fileName);
    if (!csvString) {
      CustomLogger.logInfo(`No valid data found in file: ${fileName}`, CONFIG.APP.NAME, 'runLoadCsvFromDrive()');
      return;
    }

    // Parse CSV data
    const csvData = parseCsvData(csvString);

    // Upload to BigQuery
    uploadToBQ(csvData);

    // Refresh kiosk data
    refresh(file);

    CustomLogger.logInfo(`Successfully processed file: ${fileName}`, CONFIG.APP.NAME, 'runLoadCsvFromDrive()');
  } catch (error) {
    CustomLogger.logError(`Error processing file ${fileId}: ${error.message}`, CONFIG.APP.NAME, 'runLoadCsvFromDrive()');
    throw error; // Re-throw to be caught by the caller
  }
}

/**
 * Extracts and formats sheet name from filename
 * @param {string} fileName - The name of the file
 * @return {string|null} - Formatted sheet name or null if invalid
 */
function extractSheetNameFromFileName(fileName) {
  const match = fileName.match(/_(\d{12})/);
  if (!match || !match[1]) {
    return null;
  }

  const filenameRegex = match[1];
  return filenameRegex.replace(/20(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/g, '$2$3 $400$1');
}

/**
 * Parses CSV data into structured format
 * @param {string} csvString - The CSV content as a string
 * @return {Array} - Parsed CSV data
 */
function parseCsvData(csvString) {
  return Utilities.parseCsv(csvString)
    .map(row => [
      row[0],
      row[1],
      row[2],
      parseFloat(row[3]) || 0,
      row[4],
      parseFloat(row[5]) || 0
    ]);
}

/**
 * Validates and processes CSV content
 * @param {string} data - The CSV content
 * @param {string} fileName - The name of the file
 * @return {string} - Processed CSV content
 */
function validateLine(data, fileName) {
  const match = fileName.match(/_(\d{12})/);
  if (!match || !match[1]) {
    CustomLogger.logInfo(`Invalid filename format: ${fileName}`, CONFIG.APP.NAME, 'validateLine()');
    return null;
  }

  const dateStr = match[1];
  const trnDate = parseDateFromString(dateStr);
  if (!trnDate) {
    CustomLogger.logInfo(`Invalid date format in filename: ${fileName}`, CONFIG.APP.NAME, 'validateLine()');
    return null;
  }

  const formattedDate = trnDate.toISOString().split('T')[0];
  const validLines = [];

  // Filter out empty lines
  const lines = data.split("\n").filter(line => line.trim() !== "");

  // Process each line
  lines.forEach(line => {
    if (line.startsWith("Partner Name")) {
      return; // Skip header line
    }

    const processedLine = processLine(line, trnDate);
    if (processedLine) {
      validLines.push(processedLine);
    }
  });

  return validLines.length > 0 ? validLines.join("\n") : null;
}

/**
 * Parses date from string in format YYYYMMDDHHMM
 * @param {string} dateStr - Date string
 * @return {Date|null} - Parsed date or null if invalid
 */
function parseDateFromString(dateStr) {
  try {
    const year = dateStr.substring(0, 4);
    const month = dateStr.substring(4, 6);
    const day = dateStr.substring(6, 8);
    const hours = dateStr.substring(8, 10);
    const mins = dateStr.substring(10, 12);

    const trnDate = new Date(`${year}-${month}-${day} ${hours}:${mins}:00`);
    trnDate.setDate(trnDate.getDate());
    return trnDate;
  } catch (error) {
    CustomLogger.logError(`Error parsing date: ${error.message}`, CONFIG.APP.NAME, 'parseDateFromString()');
    return null;
  }
}

/**
 * Processes a single line of CSV data
 * @param {string} line - The line to process
 * @param {Date} trnDate - The transaction date
 * @return {string|null} - Processed line or null if invalid
 */
function processLine(line, trnDate) {
  const regex = /^(?<partner_name>[^,]+),(?<machine_name>[^,]+),(?<current_cash_amount>[1-9]\d*)?,(?<machine_status>[^,]+),(?<bv_health>0|[1-9]\d*)(?<bv_healthDecimalPlace>\.\d*)?$/gm;

  const matches = line.matchAll(regex);
  for (const match of matches) {
    const bill_validator = match.groups.bv_healthDecimalPlace
      ? parseFloat(match.groups.bv_health + match.groups.bv_healthDecimalPlace)
      : parseFloat(match.groups.bv_health);

    const formattedDateTime = formatDateTime(trnDate);
    const current_cash_amount = isNaN(match[3]) ? 0 : match[3];

    return `${formattedDateTime},${match[1]},${match[2]},${current_cash_amount},${match[4]},${bill_validator}`;
  }

  return null;
}

/**
 * Formats a date object into a string in the format YYYY-MM-DD HH:MM:SS
 *
 * @param {*} date
 * @return {*} A string representing the date in the format YYYY-MM-DD HH:MM:SS.
 * @throws {Error} If the input is not a valid Date object.
 */
function formatDateTime(date) {
  if (isNaN(date.getMonth())) {
    return '1970-01-01 00:00:00';
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');

  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

/**
 * Downloads Paybox logs for monitoring hourly
 * @return {Object} - Result of the download operation
 */
function runDownloadPayboxLogsMonitoringHourly() {
  try {
    // Create folder structure
    const parentFolderId = GDriveFilesAPI.createFolderStructure('Paybox Temp');
    const monitoringFolderId = GDriveFilesAPI.getOrCreateFolder('+ Monitoring Hourly', parentFolderId);

    // Clear existing files
    clearMonitoringFolder(monitoringFolderId);

    // Set up email filter
    const now = new Date();
    const dateTo = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const emailFilter = createEmailFilter(dateTo);
    CustomLogger.logInfo(`Email filter: ${emailFilter}`, CONFIG.APP.NAME, 'runDownloadPayboxLogsMonitoringHourly()');

    // Search for emails
    const threads = GmailApp.search(emailFilter);
    Utilities.sleep(1000);

    // Download attachments
    const fileCount = downloadAttachments(threads, monitoringFolderId, parentFolderId);

    // Manage CSV files
    if (fileCount > 0) {
      manageCsvFiles();
    } else {
      CustomLogger.logInfo('No attachments downloaded.', CONFIG.APP.NAME, 'runDownloadPayboxLogsMonitoringHourly()');
    }

    return { success: true, fileCount };
  } catch (error) {
    CustomLogger.logError(`Error downloading Paybox logs: ${error.message}`, CONFIG.APP.NAME, 'runDownloadPayboxLogsMonitoringHourly()');
    return { success: false, error: error.message, fileCount: 0 };
  }
}

/**
 * Clears all files in a folder
 * @param {string} folderId - The ID of the folder to clear
 */
function clearMonitoringFolder(folderId) {
  GDriveFilesAPI.deleteFilesInFolderById(folderId);

  CustomLogger.logInfo('Cleared files in the monitoring folder.', CONFIG.APP.NAME, 'clearMonitoringFolder()');
}

/**
 * Creates an email filter for Gmail search
 * @param {string} dateTo - The date to filter to
 * @return {string} - The email filter
 */
function createEmailFilter(dateTo) {
  return `from:(no-reply@multisyscorp.io) subject:("PAYBOX Machine Unit Monitoring") has:attachment after:${dateTo}`;
}

/**
 * Downloads attachments from email threads
 * @param {Array} threads - The email threads
 * @param {string} monitoringFolderId - The ID of the monitoring folder
 * @param {string} parentFolderId - The ID of the parent folder
 * @return {number} - The number of files downloaded
 */
function downloadAttachments(threads, monitoringFolderId, parentFolderId) {
  let fileCount = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      message.getAttachments().forEach(attachment => {
        if (attachment.getName().startsWith('monitoring_')) {
          const attachmentBlob = attachment.copyBlob();
          const folderId = GDriveFilesAPI.getOrCreateFolder('+ Monitoring Hourly', parentFolderId);
          DriveApp.getFolderById(folderId).createFile(attachmentBlob);
          fileCount++;
          CustomLogger.logInfo(`Downloading ${attachment.getName()}`, CONFIG.APP.NAME, 'downloadAttachments()');
        }
      });
    });
  });

  return fileCount;
}

/**
 * Manages CSV files in the collections folder
 */
function manageCsvFiles() {
  const folderCollections = DriveApp.getFolderById('1knJiSYo3bO4B_V3BohCi5aBhFyhhXPSW');
  const files = getCsvFiles(folderCollections);

  if (files.length === 0) {
    CustomLogger.logInfo('No CSV files found in the folder.', CONFIG.APP.NAME, 'manageCsvFiles()');
    return;
  }

  deleteAllButLatestFile(files);
}

/**
 * Gets all CSV files from a folder
 * @param {Folder} folder - The folder to get files from
 * @return {Array} - The CSV files
 */
function getCsvFiles(folder) {
  const files = [];
  const filesIterator = folder.getFilesByType(MimeType.CSV);

  while (filesIterator.hasNext()) {
    files.push(filesIterator.next());
  }

  files.sort((a, b) => a.getName().localeCompare(b.getName()));
  return files;
}

/**
 * Deletes all but the latest file in an array of files
 * @param {Array} files - The files to process
 */
function deleteAllButLatestFile(files) {
  for (let i = 0; i < files.length - 1; i++) {
    files[i].setTrashed(true);
  }
  // CustomLogger.logInfo(`Deleted ${files.length - 1} older CSV files, kept the latest one.`, PROJECT_NAME, 'deleteAllButLatestFile');
  CustomLogger.logInfo(`Deleted ${files.length - 1} older CSV files, kept the latest one.`, CONFIG.APP.NAME, 'deleteAllButLatestFile()');
}

function uploadToBQ(dataArray, maxRetries = 3, retryDelay = 2000) {
  // BigQuery configuration
  const config = {
    projectId: 'ms-paybox-prod-1',
    datasetId: 'pldtsmart',
    tableId: 'monitoring_hourly',
    schema: [
      { name: 'created_at', type: 'TIMESTAMP' },
      { name: 'partner_name', type: 'STRING' },
      { name: 'machine_name', type: 'STRING' },
      { name: 'current_amount', type: 'FLOAT' },
      { name: 'machine_status', type: 'STRING' },
      { name: 'bill_validator_health', type: 'FLOAT' }
    ]
  };

  let attempt = 0;

  while (attempt <= maxRetries) {
    try {
      CustomLogger.logInfo(`Upload attempt ${attempt + 1}/${maxRetries + 1}`, CONFIG.APP.NAME, 'uploadToBQ');

      // Prepare the BigQuery job configuration
      const jobConfig = {
        configuration: {
          load: {
            destinationTable: {
              projectId: config.projectId,
              datasetId: config.datasetId,
              tableId: config.tableId
            },
            schema: { fields: config.schema },
            sourceFormat: 'CSV',
            writeDisposition: 'WRITE_APPEND',
            autodetect: false
          }
        }
      };

      // Insert data into BigQuery with timeout handling
      const job = BigQuery.Jobs.insert(jobConfig, config.projectId, Utilities.newBlob(dataArray.join('\n')));
      const jobId = job.jobReference.jobId;

      CustomLogger.logInfo(`BigQuery job started: ${jobId}`, CONFIG.APP.NAME, 'uploadToBQ');

      // Wait for job completion with timeout handling
      const jobResult = waitForJobCompletionWithTimeout(config.projectId, jobId);

      if (jobResult.success) {
        CustomLogger.logInfo(`Job completed successfully: ${jobId}`, CONFIG.APP.NAME, 'uploadToBQ');
        return; // Success, exit the function
      } else {
        throw new Error(`Job failed: ${jobResult.error}`);
      }

    } catch (error) {
      attempt++;

      // Check if it's a timeout error
      const isTimeoutError = isTimeoutRelatedError(error);

      if (isTimeoutError && attempt <= maxRetries) {
        CustomLogger.logWarning(
          `Timeout error on attempt ${attempt}/${maxRetries + 1}: ${error.message}. Retrying in ${retryDelay}ms...`,
          CONFIG.APP.NAME,
          'uploadToBQ'
        );

        // Wait before retrying with exponential backoff
        Utilities.sleep(retryDelay * Math.pow(2, attempt - 1));
        continue;

      } else if (attempt > maxRetries) {
        // Max retries exceeded
        CustomLogger.logError(
          `Max retries (${maxRetries}) exceeded. Final error: ${error.message}`,
          CONFIG.APP.NAME,
          'uploadToBQ'
        );
        throw new Error(`Upload failed after ${maxRetries} retries: ${error.message}`);

      } else {
        // Non-timeout error, don't retry
        CustomLogger.logError(`Non-retryable error: ${error.message}`, CONFIG.APP.NAME, 'uploadToBQ');
        throw error;
      }
    }
  }
}

function waitForJobCompletionWithTimeout(projectId, jobId, timeoutMs = 300000) { // 5 minutes default
  const startTime = Date.now();
  const pollInterval = 5000; // 5 seconds

  try {
    while (Date.now() - startTime < timeoutMs) {
      const job = BigQuery.Jobs.get(projectId, jobId);

      if (job.status && job.status.state === 'DONE') {
        if (job.status.errors) {
          return {
            success: false,
            error: job.status.errors.map(e => e.message).join(', ')
          };
        }
        return { success: true };
      }

      // Wait before next poll
      Utilities.sleep(pollInterval);
    }

    // Timeout reached
    throw new Error(`Job ${jobId} timed out after ${timeoutMs}ms`);

  } catch (error) {
    if (error.message.includes('timed out') || error.message.includes('timeout')) {
      throw error; // Re-throw timeout errors
    }

    return {
      success: false,
      error: error.message
    };
  }
}

function isTimeoutRelatedError(error) {
  const timeoutKeywords = [
    'timeout',
    'timed out',
    'request timeout',
    'execution timeout',
    'deadline exceeded',
    'time limit exceeded',
    'connection timeout',
    'read timeout',
    'write timeout',
    'service timeout'
  ];

  const errorMessage = error.message ? error.message.toLowerCase() : '';
  const errorString = error.toString().toLowerCase();

  return timeoutKeywords.some(keyword =>
    errorMessage.includes(keyword) || errorString.includes(keyword)
  );
}

// Alternative helper function for checking specific Google Apps Script timeout errors
function isGASTimeoutError(error) {
  // Google Apps Script specific timeout patterns
  const gasTimeoutPatterns = [
    /exceeded maximum execution time/i,
    /script runtime exceeded/i,
    /execution timeout/i,
    /service invoked too many times/i,
    /quota exceeded/i
  ];

  const errorMessage = error.message || error.toString();
  return gasTimeoutPatterns.some(pattern => pattern.test(errorMessage));
}