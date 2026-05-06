function doGet(e) {
  Logger.log("doGet triggered");
  return HtmlService.createHtmlOutputFromFile('WorkerPortal')
    .setTitle('מערכת טיפול בלקוחות')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getPendingTasks() {
  Logger.log("getPendingTasks triggered");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('לטיפול המשרד');
  if (!sheet) return { tasks: [], pendingPayments: {} };
  
  const tasks = [];
  const lastRow = sheet.getLastRow();
  
  if (lastRow >= 3) {
    const data = sheet.getRange(3, 1, lastRow - 2, 9).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const isDone = row[8]; // Col I (9) - Boolean
      const amountVal = row[6];
      const hasAmount = amountVal !== "" && amountVal != null;
      
      if (isDone !== true && hasAmount) {
        const dateVal = row[5]; // Col F (6) - Date
        let formattedDate = dateVal;
        if (dateVal instanceof Date) {
          formattedDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        
        tasks.push({
          rowIndex: i + 3,
          customerName: row[0] ? normalizeHebrew(row[0]) : '',
          carNumber: row[1] ? normalizeHebrew(row[1]) : '',
          carModel: row[2] ? normalizeHebrew(row[2]) : '',
          debtSource: row[3] ? normalizeHebrew(row[3]) : '',
          debtInfo: row[4] ? row[4].toString() : '',
          date: formattedDate,
          amount: Number(row[6]) || 0,
          notes: row[7] ? row[7].toString() : ''
        });
      }
    }
  }
  
  const paymentsSheet = ss.getSheetByName('תשלומים');
  const pendingPaymentsMap = {};
  if (paymentsSheet) {
    const pLastRow = paymentsSheet.getLastRow();
    if (pLastRow >= 2) {
      const pData = paymentsSheet.getRange(2, 1, pLastRow - 1, 4).getValues();
      for (let i = 0; i < pData.length; i++) {
        const row = pData[i];
        const name = row[0] ? normalizeHebrew(row[0]) : '';
        const amount = Number(row[2]) || 0;
        const handled = row[3]; // Col D (4)
        
        if (name && handled !== true && amount > 0) {
          if (!pendingPaymentsMap[name]) {
            pendingPaymentsMap[name] = 0;
          }
          pendingPaymentsMap[name] += amount;
        }
      }
    }
  }
  
  Logger.log(`Found ${tasks.length} pending tasks`);
  return { tasks: tasks, pendingPayments: pendingPaymentsMap };
}

function addPaymentToSheet(paymentData) {
  Logger.log("addPaymentToSheet triggered with data: " + JSON.stringify(paymentData));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('תשלומים');
  if (!sheet) throw new Error("Sheet 'תשלומים' not found.");
  
  // Create combined info string
  let combinedInfo = `אמצעי תשלום: ${paymentData.method}`;
  
  if (paymentData.method === 'אשראי') {
    combinedInfo += `\n4 ספרות אחרונות: ${paymentData.last4Digits || 'לא הוזן'}`;
  } else if (paymentData.method === 'העברה בנקאית') {
    combinedInfo += `\nפרטי חשבון: ${paymentData.bankInfo || 'לא הוזן'}`;
  }
  
  if (paymentData.comments) {
    combinedInfo += `\nהערות: ${paymentData.comments}`;
  }
  
  // Convert date if possible
  let dateObj = new Date(paymentData.date);
  let formattedDate = paymentData.date;
  if (!isNaN(dateObj.getTime())) {
    formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  
  const amount = Number(paymentData.amount) || 0;
  
  // Find the first row where columns A, B, C, and E are empty
  const lastRow = Math.max(sheet.getLastRow(), 1);
  let targetRow = lastRow + 1;
  
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const isAEmpty = (!row[0] || row[0].toString().trim() === "");
      const isBEmpty = (!row[1] || row[1].toString().trim() === "");
      const isCEmpty = (!row[2] || row[2].toString().trim() === "");
      const isEEmpty = (!row[4] || row[4].toString().trim() === "");
      
      if (isAEmpty && isBEmpty && isCEmpty && isEEmpty) {
        targetRow = i + 2; // +2 because data starts at row 2, and i is 0-indexed
        break;
      }
    }
  }
  
  Logger.log(`Target row selected: ${targetRow}`);
  
  // Write the data to the target row
  sheet.getRange(targetRow, 1, 1, 5).setValues([[
    paymentData.customerName,
    formattedDate,
    amount,
    "",
    combinedInfo
  ]]);
  
  return true;
}
