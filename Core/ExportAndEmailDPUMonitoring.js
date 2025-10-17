function exportSheetAndSendEmail() {
  CustomLogger.logInfo("Running sending DPU monitoring...", CONFIG.APP.NAME, 'exportSheetAndSendEmail()');

  var retries = 3;
  var sheetName = "For Sending";
  var subject = "[AUTO] Paybox DPU Monitoring";

  var emailTo = "RBEspayos@smart.com.ph, RACagbay@smart.com.ph, mvolbara@pldt.com.ph, avdeleon@pldt.com.ph, cvcabanilla@multisyscorp.com ";
  var emailCc = "AABenter@smart.com.ph, rtevangelista@pldt.com.ph, nplimpiada@pldt.com.ph, egalcantara@multisyscorp.com, support@paybox.ph";

  // var emailTo = "egalcantara@multisyscorp.com";
  // var emailCc = "egalcantara@multisyscorp.com";

  var day = new Date().getDay();
  if (day === 0 || day === 6) { // 0 = Sunday, 6 = Saturday
    CustomLogger.logInfo("Today is a weekend. Exiting script.", CONFIG.APP.NAME, 'exportSheetAndSendEmail()');
    return; // Exit the function if today is a weekend
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    CustomLogger.logError(`Sheet "${sheetName}" not found.`, CONFIG.APP.NAME, 'exportSheetAndSendEmail'); // Log error message to the log
    return;
  }

  sheet.showSheet();

  var token = ScriptApp.getOAuthToken();
  var fileName = "Paybox DPU Monitoring_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy") + ".pdf";
  var lastRow = sheet.getLastRow(); // Automatically get the last row

  // Prepare the export URL
  var url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?exportFormat=pdf&format=pdf` +
    `&size=A4&portrait=false&fitw=true&top_margin=0.1&bottom_margin=0.1` +
    `&left_margin=0.1&right_margin=0.1&sheetnames=false&printtitle=false` +
    `&pagenumbers=false&gridlines=false&fzr=false&range=A1:${sheet.getRange(lastRow, sheet.getLastColumn()).getA1Notation()}` +
    `&gid=${sheet.getSheetId()}`;

  while (retries > 0) {
    try {
      var response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      // Check if the response is successful
      if (response.getResponseCode() !== 200) {
        throw new Error(`Failed to fetch PDF. Response code: ${response.getResponseCode()}`);
      }

      var pdfBlob = response.getBlob().setName(fileName);
      const body = generateEmailBody();

      // Send email with the PDF attachment
      EmailSender.sendEmail(recipient = { to: emailTo, cc: emailCc }, subject, body, attachments = [pdfBlob], true);

      // GmailApp.sendEmail(emailTo, subject, body, {
      //   cc: emailCc,
      //   attachments: [pdfBlob],
      // });

      sheet.hideSheet();

      CustomLogger.logInfo("Email sent successfully.", CONFIG.APP.NAME, 'exportSheetAndSendEmail()');
      return; // Exit loop on success

    } catch (e) {
      CustomLogger.logError("Attempt " + (4 - retries) + " failed: " + e.message, 'exportSheetAndSendEmail', CONFIG.APP.NAME, exportSheetAndSendEmail);
    }

    retries--;
    if (retries > 0) {
      CustomLogger.logInfo("Retrying in 5 seconds...", CONFIG.APP.NAME, 'exportSheetAndSendEmail()');
      Utilities.sleep(5000); // Wait 5 seconds before retrying
    }
  }
  CustomLogger.logError("All attempts to send the email have failed.", CONFIG.APP.NAME, 'exportSheetAndSendEmail()');
}

function generateEmailBody() {
  const formattedDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "MMMM d, yyyy"
  );

  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <p>Hi All,</p>
        
        <p>Good day! Please see the attached PDF file of Kiosk DPU Monitoring as of <strong>${formattedDate}</strong>.</p>
        
        <p>Thank you,</p>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; color: #999;">
          <p>*** This is an auto-generated email, please do not reply to this email. ****</p>
        </div>
      </body>
    </html>
  `;

  return body.trim();
}
