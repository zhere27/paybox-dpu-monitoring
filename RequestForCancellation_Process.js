function processCancellationRequest() {
  processCancellationLogic('BPI Internal');
  processCancellationLogic('Brinks via BPI');
  processCancellationLogic('eTap');
}

