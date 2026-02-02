/**
 * Main function to trigger the check.
 */
function runTicketCheck() {
  const ABC = 14; // Days threshold
  showTicketWarnings(ABC);
}

function showTicketWarnings(abcDays) {
  const html = HtmlService.createTemplateFromFile('TicketAlert');
  html.abcDays = abcDays; 
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html.evaluate().setWidth(1400).setHeight(800), 'דוח סטטוס דוחות - התראות');
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'string') {
    const parts = val.trim().split('/');
    if (parts.length === 3) {
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }
  }
  return null;
}

function toMidnight(d) {
  if (!d) return null;
  const newD = new Date(d);
  newD.setHours(0, 0, 0, 0);
  return newD;
}

function updateLastChecked(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('דוחות');
  sheet.getRange(rowIndex, 23).setValue(new Date());
  return "Updated";
}

function updateTicketNote(rowIndex, newNote) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('דוחות');
  // Column Q is the 17th column
  sheet.getRange(rowIndex, 17).setValue(newNote); 
  return "Note Updated";
}

// --- NEW FUNCTION FOR TRANSFER APPROVAL ---
function updateTransferStatus(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('דוחות');
  
  // Column M is index 13
  const rangeM = sheet.getRange(rowIndex, 13);
  const valueM = rangeM.getValue();
  
  // Column N is index 14 (Status)
  const rangeN = sheet.getRange(rowIndex, 14);
  
  let newStatus = "";
  
  // Check if M is exactly 0 (strict check depending on data type, usually 0 or "0")
  if (valueM == 0) {
    newStatus = "סיום טיפול הוסב";
  } else {
    newStatus = "אושרה הסבה";
  }
  
  rangeN.setValue(newStatus);
  return newStatus;
}

function getBodyDetails(bodyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('רשימות');
  
  if (!sheet) return { error: "Sheet 'רשימות' not found" };
  
  const lastRow = sheet.getLastRow();
  
  // Headers are in Row 3. Data starts in Row 4.
  // We need at least 3 rows to have headers.
  if (lastRow < 3) return { error: "No data in lists sheet" };

  // Start at Row 3 (Headers), Column 7 (G)
  // Number of rows to fetch = lastRow - 2 (e.g. if lastRow is 10, we want rows 3 to 10, which is 8 rows)
  // Width is 5 columns (G, H, I, J, K)
  const range = sheet.getRange(3, 7, lastRow - 2, 5); 
  const data = range.getValues();
  
  const headers = data[0]; // This is Row 3
  const rows = data.slice(1); // This is Row 4+
  
  const searchStr = String(bodyName).trim();
  
  // Find the matching row (Index 0 is Column G - Body Name)
  const foundRow = rows.find(r => String(r[0]).trim() === searchStr);
  
  if (!foundRow) return null; 
  
  const result = {};
  headers.forEach((header, index) => {
    if (header && String(header).trim() !== "") {
      let val = foundRow[index];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      result[header] = val;
    }
  });
  
  return result;
}

function getTicketData(abcDays) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('דוחות');

  if (abcDays == null) abcDays = 0;
  if (!sheet) throw new Error("Sheet 'דוחות' not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { section1: [], section2: [] };

  const range = sheet.getRange(2, 1, lastRow - 1, 23);
  const data = range.getValues();

  const todayMidnight = toMidnight(new Date());

  const activeStatuses = [
    'הותחל טיפול',
    'נשלח פעם אחת ממתין לתגובה',
    'מוכן להסבה',
    'נשלחה בקשה להסבה',
    'ממתין לטיפול',
    'שונות',
    ''
  ];

  const excludedRequests = ['בקשה להפחתה', 'ערעור', 'בקשה להשפט'];

  const section1 = []; // Stale / Status Check
  const section2 = []; // Urgent / About to pass

  const fmt = (d) => {
    if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
    return d ? String(d) : "";
  };
  
  const fmtTime = (d) => {
    if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    return d ? String(d) : "";
  };

  const daysDiff = (d1, d2) => {
    if (!d1 || !d2) return 0;
    const diffTime = Math.abs(d2.getTime() - d1.getTime());
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  };

  data.forEach((row, index) => {
    const renterName = row[1];       
    const ticketOrigin = row[2];     
    const carModel = row[3];         
    const carNo = row[4];            
    const ticketNo = row[5];         
    const rawTicketDate = row[6];    
    const rawExpDate = row[8];       
    const ticketAmount = row[9];     
    let ticketStatus = row[13];      
    let requestNo = row[14];       
    const rawRequestDate = row[15];  
    const notes = row[16];           
    const rawLastChecked = row[22];  

    if (!ticketNo || String(ticketNo).trim() === "") return;
    if (typeof ticketStatus === 'string') ticketStatus = ticketStatus.trim();

    if (requestNo instanceof Date) {
      requestNo = fmt(requestNo);
    }

    const requestDate = toMidnight(parseDate(rawRequestDate));
    const lastChecked = toMidnight(parseDate(rawLastChecked));
    const expDate = toMidnight(parseDate(rawExpDate));

    const displayTicketDate = parseDate(rawTicketDate);
    const displayExpDate = parseDate(rawExpDate);
    
    const displayReqDate = fmt(parseDate(rawRequestDate) || rawRequestDate);
    
    let reqDetails = "";
    if (requestNo) {
      reqDetails = String(requestNo);
      if (displayReqDate && displayReqDate !== requestNo) {
         reqDetails += ` (${displayReqDate})`;
      }
    } else if (displayReqDate) {
      reqDetails = displayReqDate;
    }

    const realRowIndex = index + 2;

    if (!activeStatuses.includes(ticketStatus)) return;
    const reqStr = String(requestNo);
    if (excludedRequests.some(r => reqStr.includes(r))) return;

    const rowObj = {
      rowIndex: realRowIndex,
      renter: renterName,
      origin: ticketOrigin,
      car: `${carModel} (${carNo})`,
      ticketNo: ticketNo,
      offenseDate: fmtTime(displayTicketDate || rawTicketDate),
      expDate: fmt(displayExpDate || rawExpDate),
      amount: ticketAmount,
      status: ticketStatus,
      reqDetails: reqDetails, 
      notes: notes
    };

    let addedToSection = false;

    // --- SECTION 2 (Stale Checks) LOGIC ---
    let isRequestOld = false;
    if (requestDate && requestDate < todayMidnight && daysDiff(requestDate, todayMidnight) > abcDays) {
      isRequestOld = true;
    }

    let isRecentlyChecked = false;
    if (lastChecked) {
      if (lastChecked <= todayMidnight) {
        if (daysDiff(lastChecked, todayMidnight) <= abcDays) {
          isRecentlyChecked = true;
        }
      }
    }

    if (isRequestOld && !isRecentlyChecked) {
      section1.push(rowObj);
      addedToSection = true;
    }

    // --- SECTION 1 (Urgent / About to pass) LOGIC ---
    // Condition: Not already in Section 2 AND Request Date is empty
    if (!addedToSection && (!rawRequestDate || String(rawRequestDate).trim() === "")) {
      
      // CHANGE IS HERE:
      // We ONLY verify if expDate exists. If it is null/empty, we skip it.
      if (expDate) {
        const diff = (expDate.getTime() - todayMidnight.getTime()) / (1000 * 60 * 60 * 24);
        
        // Show if 15 days or less remain (or negative if passed)
        if (diff <= 15) {
          section2.push(rowObj);
        }
      }
    }
  });

  return { section1: section1, section2: section2 };
}