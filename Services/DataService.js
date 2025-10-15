/**
 * Upload Card data to BigQuery
 */
function uploadCardToBigQuery(dataArray, fileId) {
  if (!dataArray || dataArray.length === 0) {
    CustomLogger.logInfo('No data to upload to BigQuery', CONFIG.APP.NAME, 'uploadCardToBigQuery');
    return;
  }

  try {
    const jobConfig = {
      configuration: {
        load: {
          destinationTable: {
            projectId: CONFIG.BIGQUERY.PROJECT_ID,
            datasetId: CONFIG.BIGQUERY.DATASET_ID,
            tableId: CONFIG.BIGQUERY.TABLE_ID
          },
          schema: { fields: CONFIG.BIGQUERY.SCHEMA },
          sourceFormat: 'CSV',
          writeDisposition: 'WRITE_APPEND',
          autodetect: false
        }
      }
    };

    // Convert array to CSV blob
    const csvBlob = Utilities.newBlob(dataArray.map(row => row.join(',')).join('\n'));
    
    // Insert job
    let job = BigQuery.Jobs.insert(jobConfig, CONFIG.BIGQUERY.PROJECT_ID, csvBlob);
    const jobId = job.jobReference.jobId;

    // Poll for completion
    while (job.status.state !== "DONE") {
      Utilities.sleep(CONFIG.BIGQUERY.SLEEP_TIME_MS);
      job = BigQuery.Jobs.get(CONFIG.BIGQUERY.PROJECT_ID, jobId);
      CustomLogger.logInfo(`Job Status: ${job.status.state}`, CONFIG.APP.NAME, 'uploadCardToBigQuery');
    }

    // Check for errors
    if (job.status.errors && job.status.errors.length > 0) {
      const errorMsg = JSON.stringify(job.status.errors);
      CustomLogger.logError(`BigQuery job failed: ${errorMsg}`, CONFIG.APP.NAME, 'uploadCardToBigQuery');
      throw new Error(`BigQuery upload failed: ${errorMsg}`);
    }

    // Delete file on success
    GDriveFilesAPI.deleteFileByFileId(fileId);
    CustomLogger.logInfo('BigQuery job completed successfully.', CONFIG.APP.NAME, 'uploadCardToBigQuery');
    
  } catch (error) {
    CustomLogger.logError(`Error uploading to BigQuery: ${error}`, CONFIG.APP.NAME, 'uploadCardToBigQuery');
    throw error;
  }
}