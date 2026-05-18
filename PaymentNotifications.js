/**
 * Checks the "תשלומים" sheet for new payments and sends an aggregated email.
 * This should be triggered by a Time-Driven Trigger (e.g., every 15 minutes or 1 hour).
 */
function sendPendingPaymentEmails() {
  console.log('Entering sendPendingPaymentEmails');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('תשלומים');
  
  if (!sheet) {
    console.error("sendPendingPaymentEmails: Sheet 'תשלומים' not found.");
    Logger.log("Sheet 'תשלומים' not found.");
    console.log('Exiting sendPendingPaymentEmails (sheet not found)');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log('sendPendingPaymentEmails mid-step: No data found');
    console.log('Exiting sendPendingPaymentEmails (no data)');
    return; // No data
  }
  
  // Read all data from row 2 to lastRow
  // Assuming columns A to G: Name (1), Date (2), Amount (3), Handled (4), Comment (5), Empty (6), Email Sent (7)
  const range = sheet.getRange(2, 1, lastRow - 1, 7);
  const data = range.getValues();
  
  const pendingPayments = [];
  const rowsToUpdate = [];
  
  for (let i = 0; i < data.length; i++) {
    const name = data[i][0];
    const date = data[i][1];
    const amount = data[i][2];
    const emailSent = data[i][6];
    
    // Check if Name, Date, Amount are present and Email Sent is NOT true
    if (name && date && amount && emailSent !== true) {
      const formattedDate = date instanceof Date ? Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy") : date;
      const comment = data[i][4] ? data[i][4] : "אין";
      
      pendingPayments.push({
        name: name,
        date: formattedDate,
        amount: amount,
        comment: comment
      });
      
      rowsToUpdate.push(i + 2); // Save row number for later update
    }
  }
  
  if (pendingPayments.length === 0) {
    Logger.log("No new pending payments to send.");
    return;
  }
  
  // Create HTML body for the email
  let htmlBody = `
    <div dir="rtl" style="font-family: Arial, sans-serif;">
      <h2>התקבלו תשלומים חדשים</h2>
      <p>להלן פירוט התשלומים שהוזנו במערכת:</p>
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; max-width: 600px;">
        <tr style="background-color: #f2f2f2;">
          <th>תאריך</th>
          <th>שם שוכר</th>
          <th>סכום</th>
          <th>הערות</th>
        </tr>
  `;
  
  pendingPayments.forEach(payment => {
    htmlBody += `
        <tr>
          <td>${payment.date}</td>
          <td>${payment.name}</td>
          <td>${payment.amount}</td>
          <td>${payment.comment}</td>
        </tr>
    `;
  });
  
  htmlBody += `
      </table>
      <br>
      <p>בברכה,<br>מערכת הניהול</p>
    </div>
  `;
  
  // Send the email
  const recipient = "stardohot@gmail.com";
  const subject = "עדכון תשלומים חדשים מתוך מערכת הניהול";
  
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody
  });
  
  // Mark as sent in the sheet
  rowsToUpdate.forEach(rowNum => {
    sheet.getRange(rowNum, 7).setValue(true);
  });
  
  console.log(`sendPendingPaymentEmails mid-step: Marked ${rowsToUpdate.length} rows as sent`);
  Logger.log('Successfully sent email for ' + pendingPayments.length + ' payments.');
  console.log('Exiting sendPendingPaymentEmails');
}
