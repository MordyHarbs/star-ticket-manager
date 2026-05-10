// LEGACY — as of the Netlify migration the live portal calls
//   /.netlify/functions/getTasks  and  /.netlify/functions/submitPayment
// instead of this Apps Script Web App. The doPost handlers below are kept as
// a fallback only. The PIN now lives in Netlify env var WORKER_PIN; the
// constant below is left in place so the legacy endpoint still works during
// rollout, but should be removed (and rotated) once the new endpoints are
// confirmed stable.
const WORKER_PASSWORD = "x7#M9$kL2@pQ5&vN8*wZ";

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "Worker Portal API is running successfully." }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const pin = requestData.pin;
    
    if (action === 'getTasks') {
      const response = getPendingTasks(pin);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'addPayment') {
      const response = addPaymentToSheet(requestData.paymentData, pin);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ error: "Unknown action" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getPendingTasks(password) {
  Logger.log("getPendingTasks triggered");
  if (password !== WORKER_PASSWORD) {
    Logger.log("Unauthorized access attempt to getPendingTasks");
    return { error: 'Unauthorized' };
  }
  
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
          customerName: row[0] ? normalizeHebrew(row[0]).replace(/"/g, '״').replace(/'/g, '׳') : '',
          carNumber: row[1] ? normalizeHebrew(row[1]).replace(/"/g, '״').replace(/'/g, '׳') : '',
          carModel: row[2] ? normalizeHebrew(row[2]).replace(/"/g, '״').replace(/'/g, '׳') : '',
          debtSource: row[3] ? normalizeHebrew(row[3]).replace(/"/g, '״').replace(/'/g, '׳') : '',
          debtInfo: row[4] ? row[4].toString().replace(/"/g, '״').replace(/'/g, '׳') : '',
          date: formattedDate,
          amount: Number(row[6]) || 0,
          notes: row[7] ? row[7].toString().replace(/"/g, '״').replace(/'/g, '׳') : ''
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

function addPaymentToSheet(paymentData, password) {
  Logger.log("addPaymentToSheet triggered with data: " + JSON.stringify(paymentData));
  if (password !== WORKER_PASSWORD) {
    Logger.log("Unauthorized access attempt to addPaymentToSheet");
    return { error: 'Unauthorized' };
  }
  
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
