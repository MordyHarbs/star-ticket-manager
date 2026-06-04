
function runOnce_createOnOpenTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onOpenWithUi') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger('onOpenWithUi')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}

function removeDuplicateTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let found = false;
  let count = 0;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onOpenWithUi') {
      if (found) {
        ScriptApp.deleteTrigger(trigger);
        count++;
      }
      found = true;
    }
  }
  return `Removed ${count} duplicate triggers.`;
}

function onOpenWithUi(e) {
  console.log('Entering onOpenWithUi');
  buildCustomMenu();

  // --- DEBOUNCE PROTECTION: Prevent duplicate triggers from firing multiple popups ---
  const props = PropertiesService.getScriptProperties();
  const lastRun = parseInt(props.getProperty('LAST_ONOPEN_POPUP_TIME') || '0', 10);
  const now = new Date().getTime();
  if (now - lastRun < 2000) {
    console.log('onOpenWithUi mid-step: Debounced');
    console.log('Exiting onOpenWithUi (debounced)');
    return; // If script was triggered less than 2 seconds ago, ignore this run
  }
  props.setProperty('LAST_ONOPEN_POPUP_TIME', now.toString());

  try {
    const ticketCount = getTicketData(28, true); // Fast count only

    if (ticketCount > 0) {
      console.log('onOpenWithUi mid-step: Prompting user for ticketCount', { ticketCount });
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        "התראת דוחות 🚨",
        "ישנם " + ticketCount + " דוחות הממתינים לבדיקה או שמתקרבים להתיישנות.\nהאם להציג אותם כעת?",
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        runTicketCheck();
      }
    }
  } catch (err) {
    console.error("Error fetching fast ticket count on open: " + err);
  }
  console.log('Exiting onOpenWithUi');
}

function onOpen(e) {
  buildCustomMenu();
}

function buildCustomMenu() {
  const ui = SpreadsheetApp.getUi();

  // 2. תפריט דוחות
  ui.createMenu('תפריט דוחות')
    .addSubMenu(ui.createMenu('הצגת סטטוס דוחות')
      .addItem('הצג דוחות לבדיקת סטטוס', 'runTicketCheck')
      .addItem('הצג דוחות לבדיקת סטטוס (התאמה אישית)', 'runTicketCheckWithPrompt')
    )
    .addSubMenu(ui.createMenu('טבלת רשויות')
      .addItem('הצג טבלת רשויות לדוחות', 'showSourcesDialog')
    )
    .addToUi();

  // 3. תפריט לקוחות
  ui.createMenu('תפריט לקוחות')
    .addSubMenu(ui.createMenu('ניהול תזכורות')
      .addItem('הצג לקוחות ממתינים לטיפול תאריך עבר', 'checkFollowUpReminders')
      .addItem('הצג כל הלקוחות הממתינים לטיפול', 'ShowAllFollowUpReminders')
      .addItem('הוסף תאריך לטיפול', 'updateDateMenu')
    )
    .addItem('הוסף הערה למשרד', 'openAddCustomerNoteDialog')
    .addItem('סנכרון לקוחות', 'syncCustomerSheet')
    .addItem('חיפוש...', 'updateSearchInfo')
    .addToUi();

  // 1. תפריט ראשי
  ui.createMenu('⭐ תפריט ראשי')
    .addItem('סמן הכל כשולם', 'markAllAsPaid')
    .addItem('העבר חובות לטיפול המשרד', 'transferToOfficeCare')
    .addItem('סמן חובות משרד כשולם', 'markOfficeAsPaid')
    .addItem('סמן שורות ריקות כ"טופל נשלח בקבוצה"', 'markAllAsGroupSent')
    .addToUi();


}

function syncCustomerSheet() {
  CustomerSync.importCustomersData();
}

function checkFollowUpReminders() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('follow_up_date');
  if (!sheet) return;

  const rows = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues(); // A:C
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const sorted = rows
    .filter(([cust, d]) => cust && d instanceof Date)
    .sort((a, b) => a[1] - b[1]);

  const pending = sorted
    .filter(([cust, d]) => d instanceof Date && d.setHours(0, 0, 0, 0) <= today)
    .map(([cust, d, note]) => ({
      cust,
      date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      note
    }));

  if (pending.length === 0) return;

  const template = HtmlService.createTemplateFromFile("followup_dialog");
  template.pending = pending;

  const html = template.evaluate().setWidth(650).setHeight(450);
  SpreadsheetApp.getUi().showModelessDialog(html, "לקוחות ממתינים לטיפול");
}

function showDatePickerDialog(name) {
  const template = HtmlService.createTemplateFromFile("date_picker");
  template.name = name;

  const html = template.evaluate().setWidth(320).setHeight(185);
  SpreadsheetApp.getUi().showModalDialog(html, "בחר תאריך");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showSourcesDialog() {
  Logger.log('showSourcesDialog');
  const html = HtmlService.createHtmlOutputFromFile('SourcesDialog')
    .setWidth(900).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sources for Reports');
}

function updateDateMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('פירוט נסיעות לפי לקוח');

  // Get name from C6
  const name = sh.getRange('C6').getValue();

  // Call your existing dialog function with the name
  showDatePickerDialog(name);
}


// show all info from follow_up_date
function ShowAllFollowUpReminders() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('follow_up_date');
  if (!sheet) return;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues(); // A:C

  // Keep only rows with column A filled (customer name)
  const filtered = rows.filter(([cust, d, note]) =>
    cust && (d instanceof Date || (note && note.toString().trim() !== ''))
  );

  // Sort: rows with dates first, sorted earliest → latest; then rows without dates
  const sorted = filtered.sort((a, b) => {
    const d1 = a[1] instanceof Date ? a[1].getTime() : Infinity;
    const d2 = b[1] instanceof Date ? b[1].getTime() : Infinity;
    return d1 - d2;
  });

  // Prepare display
  const pending = sorted.map(([cust, d, note]) => ({
    cust,
    date: d instanceof Date
      ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy')
      : '(ללא תאריך)',
    note: note || ''
  }));

  // Build UI
  const html = HtmlService.createHtmlOutput(
    `<h3 style="padding-right:10px; direction: rtl; text-align: right; margin-top:0">כל הלקוחות ברשימת המעקב</h3>
     <ul style="padding-right:16px; direction: rtl; text-align: right; list-style-position: inside;">
       ${pending.map(p => `<li><strong>${p.cust}</strong> – ${p.date}${p.note ? ` – ${p.note}` : ''}</li>`).join('')}
     </ul>`
  )
    .setWidth(700)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, "רשימת מעקב מלאה");
}

function markAllAsPaid() {
  console.log('Entering markAllAsPaid');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('פירוט נסיעות לפי לקוח');
  const defaultName = sh.getRange(6, 3).getValue()

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "אישור פעולה",
    "האם הנך בטוח שברצונך לסמן את הכל כשולם עבור:\n\n" +
    "            ------ " + defaultName + " ------\n\n" +
    "ללקוח אחר לחץ 'לא'.",
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response == ui.Button.NO) {
    console.log('markAllAsPaid mid-step: User selected NO, opening namePicker');
    const html = HtmlService.createTemplateFromFile("namePicker");
    html.defaultName = defaultName;
    const dialog = html.evaluate()
      .setWidth(450)
      .setHeight(420)
      .setTitle("בחירת שם לסימון 'שולם'");

    ss.show(dialog);
  }
  else if (response == ui.Button.YES) {
    console.log('markAllAsPaid mid-step: User selected YES, running markAllForName');
    runMarkAllForName(defaultName)
  }
  else {
    console.log('markAllAsPaid mid-step: User cancelled');
    ui.alert("הפעולה בוטלה.");
  }
  console.log('Exiting markAllAsPaid');
}

function runMarkAllForName(selectedName) {
  console.log('Entering runMarkAllForName', { selectedName });
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("🟢 Running markAllAsPaid...");

  // Normalize input: always treat as array
  const names = Array.isArray(selectedName) ? selectedName : [selectedName];

  names.forEach(name => {
    ss.toast(`מסמן כעת כשולם את כל החובות של:  ⭐ ⭐ ${name} ⭐ ⭐ `, "סימון כשולם");
    markAllForName(name, ss);
    ss.toast(`כל החובות של הלקוח הבא: ⭐ ⭐ ${name} ⭐ ⭐ סומנו כשולם`, "סימון כשולם");
  });
  ss.toast(`כל הלקוחות שנבחרו סומנו כשולם`, "סימון כשולם");
  console.log('Exiting runMarkAllForName');
}

function getAllClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('לקוחות');
  const values = sh.getRange(1, 1, sh.getLastRow()).getValues();
  return values.map(r => r[0]).filter(x => x); // רק שמות לא ריקים
}

function getAllClientsWithDebts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clients = getAllClients();
  const debtMap = {};
  const normToExact = {};
  
  clients.forEach(c => {
    debtMap[c] = 0;
    normToExact[normalizeHebrew(c)] = c;
  });

  // -------- דוחות --------
  let sh = ss.getSheetByName('דוחות');
  let last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 13).getValues(); // A..M (13 columns)
    for (let i = 0; i < data.length; i++) {
      const normName = normalizeHebrew(data[i][1]); // Col B
      if (Number(data[i][12]) !== 0) { // Col M
        const exactName = normToExact[normName];
        if (exactName) {
          debtMap[exactName] += (Number(data[i][10]) || 0); // Col K
        }
      }
    }
  }

  // -------- כביש 6 / מנהרות --------
  sh = ss.getSheetByName('כביש 6/מנהרות');
  last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 12).getValues(); // A..L (12 columns)
    for (let i = 0; i < data.length; i++) {
      const normName = normalizeHebrew(data[i][4]); // Col E
      if (Number(data[i][10]) !== 0 && data[i][11] !== "טופל הועבר לטיפול המשרד") { // Col K, Col L
        const exactName = normToExact[normName];
        if (exactName) {
          debtMap[exactName] += (Number(data[i][10]) || 0);
        }
      }
    }
  }

  // -------- חוצה צפון / נתיב מהיר --------
  sh = ss.getSheetByName('חוצה צפון/נתיב מהיר');
  last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 10).getValues(); // A..J (10 columns)
    for (let i = 0; i < data.length; i++) {
      const normName = normalizeHebrew(data[i][4]); // Col E
      if (Number(data[i][8]) !== 0 && data[i][9] !== "טופל הועבר לטיפול המשרד") { // Col I, Col J
        const exactName = normToExact[normName];
        if (exactName) {
          debtMap[exactName] += (Number(data[i][8]) || 0);
        }
      }
    }
  }

  return { clients: clients, debtMap: debtMap };
}

function calculateDebtForName(name, ss) {
  const normName = normalizeHebrew(name);
  let totalDebt = 0;

  let sh = ss.getSheetByName('דוחות');
  let last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 13).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][1]) === normName && Number(data[i][12]) !== 0) {
        totalDebt += (Number(data[i][10]) || 0);
      }
    }
  }

  sh = ss.getSheetByName('כביש 6/מנהרות');
  last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 12).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][4]) === normName && Number(data[i][10]) !== 0 && data[i][11] !== "טופל הועבר לטיפול המשרד") {
        totalDebt += (Number(data[i][10]) || 0);
      }
    }
  }

  sh = ss.getSheetByName('חוצה צפון/נתיב מהיר');
  last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 10).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][4]) === normName && Number(data[i][8]) !== 0 && data[i][9] !== "טופל הועבר לטיפול המשרד") {
        totalDebt += (Number(data[i][8]) || 0);
      }
    }
  }

  return totalDebt;
}

function markAllForName(name, ss) {
  const normName = normalizeHebrew(name);

  // -------- דוחות --------
  let sh = ss.getSheetByName('דוחות');
  let last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 2, last - 1, 13).getValues();
    // columns B..N (13 columns)
    for (let i = 0; i < data.length; i++) {
      const rowName = normalizeHebrew(data[i][0]); // col B
      const total = Number(data[i][11]); // col M
      const colN = data[i][13 - 1]; // col N (index 12)
      if (rowName === normName) {
        const tr = i + 2;
        
        if (normalizeHebrew(colN) === 'אושרה הסבה') {
          sh.getRange(tr, 14).setValue('סיום טיפול הוסב'); // N
        }

        if (total != 0) {
          sh.getRange(tr, 12).setValue(true); // L = true
        }
      }
    }
  }

  // -------- כביש 6 / מנהרות --------
  sh = ss.getSheetByName('כביש 6/מנהרות');
  last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 5, last - 1, 7).getValues();
    // columns E..K (7 columns)
    for (let i = 0; i < data.length; i++) {
      const rowName = normalizeHebrew(data[i][0]); // E
      const total = Number(data[i][6]); // K (index 6 in this slice)
      if (rowName === normName && total != 0) {
        const tr = i + 2;
        sh.getRange(tr, 12).setValue('שולם'); // L
      }
    }
  }

  // -------- חוצה צפון / נתיב מהיר --------
  sh = ss.getSheetByName('חוצה צפון/נתיב מהיר');
  last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 5, last - 1, 5).getValues();
    // columns E..I (5 columns)
    for (let i = 0; i < data.length; i++) {
      const rowName = normalizeHebrew(data[i][0]); // E
      const total = Number(data[i][4]); // I (index 4)
      if (rowName === normName && total != 0) {
        const tr = i + 2;
        sh.getRange(tr, 10).setValue('שולם'); // J
      }
    }
  }
  // -------- סיכומי מחיר --------
  sh = ss.getSheetByName('סיכומי מחיר');
  if (sh) { // Safety check to ensure sheet exists
    last = sh.getLastRow();
    if (last >= 2) {
      // Get columns A through E
      // Column A (1) = Name, Column E (5) = Value
      const data = sh.getRange(2, 1, last - 1, 5).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowName = normalizeHebrew(data[i][0]); // Column A (index 0)
        const total = Number(data[i][4]);          // Column E (index 4)

        if (rowName === normName && total != 0) {
          const tr = i + 2; // Real row index in the sheet
          sh.getRange(tr, 4).setValue(true); // Column D (4) = true
        }
      }
    }
  }

}

function processPaymentRow(sh, r, ss) {
  const key = normalizeHebrew(sh.getRange(r, 2).getValue());
  let target, last, data, tr, curI, curD, curF;

  if (key === 'דוחות') {
    target = ss.getSheetByName('דוחות');
    last = target.getLastRow();
    if (last >= 2) {
      curI = normalizeHebrew(sh.getRange(r, 9).getValue());
      curF = normalizeHebrew(sh.getRange(r, 6).getValue());
      data = target.getRange(2, 6, last - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][1]) === curF &&
          normalizeHebrew(data[i][0]) === curI) {
          tr = i + 2;
          target.getRange(tr, 12).setValue(true);
          const colNVal = target.getRange(tr, 14).getValue();
          if (normalizeHebrew(colNVal) === 'אושרה הסבה') {
            target.getRange(tr, 14).setValue('סיום טיפול הוסב');
          }
          break;
        }
      }
    }
  } else if (key === 'כביש 6') {
    target = ss.getSheetByName('כביש 6/מנהרות');
    last = target.getLastRow();
    if (last >= 2) {
      curD = normalizeHebrew(sh.getRange(r, 4).getValue());
      curF = normalizeHebrew(sh.getRange(r, 6).getValue());
      data = target.getRange(2, 2, last - 1, 3).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][0]) === curD &&
          normalizeHebrew(data[i][2]) === curF) {
          tr = i + 2;
          target.getRange(tr, 12).setValue('שולם');
          break;
        }
      }
    }
  } else if (key === 'חוצה צפון') {
    target = ss.getSheetByName('חוצה צפון/נתיב מהיר');
    last = target.getLastRow();
    if (last >= 2) {
      curD = normalizeHebrew(sh.getRange(r, 4).getValue()); // D
      curF = normalizeHebrew(sh.getRange(r, 6).getValue()); // F
      const curG = normalizeHebrew(sh.getRange(r, 7).getValue()); // G
      // fetch cols B..F so index 4 = column F in target
      data = target.getRange(2, 2, last - 1, 5).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][0]) === curD &&      // target col B
          normalizeHebrew(data[i][2]) === curF &&      // target col D (existing)
          normalizeHebrew(data[i][4]) === curG) {      // target col F (new)
          tr = i + 2;
          target.getRange(tr, 10).setValue('שולם');
          break;
        }
      }
    }
  }
}

// =======================================================================
// העברת חובות לטיפול המשרד
// =======================================================================

function transferToOfficeCare() {
  console.log('Entering transferToOfficeCare');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('פירוט נסיעות לפי לקוח');
  const defaultName = sh.getRange(6, 3).getValue();
  const sum = calculateDebtForName(defaultName, ss);

  const html = HtmlService.createTemplateFromFile('transferConfirm');
  html.defaultName = defaultName;
  html.defaultSum = sum;
  const dialog = html.evaluate()
    .setWidth(400)
    .setHeight(460)
    .setTitle('העברה לטיפול המשרד');

  SpreadsheetApp.getUi().showModalDialog(dialog, 'העברה לטיפול המשרד');
  console.log('Exiting transferToOfficeCare');
}

function openNamePickerForTransfer(defaultName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const html = HtmlService.createTemplateFromFile('namePicker');
  html.defaultName = defaultName;
  html.action = 'transfer';
  const dialog = html.evaluate()
    .setWidth(450)
    .setHeight(420)
    .setTitle('בחירת שם להעברה למשרד');
  SpreadsheetApp.getUi().showModalDialog(dialog, 'בחירת שם להעברה למשרד');
}

function runTransferForName(selectedName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("🟢 Running transferToOfficeCare...");

  // Normalize input: always treat as array
  const names = Array.isArray(selectedName) ? selectedName : [selectedName];

  names.forEach(name => {
    ss.toast(`מעביר כעת לטיפול המשרד את כל החובות של:  ⭐ ⭐ ${name} ⭐ ⭐ `, "העברה לטיפול המשרד");
    processTransferToOffice(name, ss);
    ss.toast(`כל החובות של הלקוח הבא: ⭐ ⭐ ${name} ⭐ ⭐ הועברו לטיפול המשרד`, "העברה לטיפול המשרד");
  });
  ss.toast(`כל הלקוחות שנבחרו הועברו לטיפול המשרד`, "העברה לטיפול המשרד");
}

function processTransferToOffice(name, ss) {
  console.log('Entering processTransferToOffice', { name });
  const normName = normalizeHebrew(name);
  const targetSheet = ss.getSheetByName('לטיפול המשרד');
  if (!targetSheet) {
    console.error("processTransferToOffice: Sheet 'לטיפול המשרד' not found.");
    SpreadsheetApp.getUi().alert("שגיאה: הגיליון 'לטיפול המשרד' לא נמצא.");
    console.log('Exiting processTransferToOffice (sheet not found)');
    return;
  }

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const newRows = [];

  // -------- דוחות --------
  let sh = ss.getSheetByName('דוחות');
  let last = sh.getLastRow();
  if (last >= 2) {
    const dataRange = sh.getRange(2, 1, last - 1, 24); // A to X (24 columns)
    const data = dataRange.getValues();
    for (let i = 0; i < data.length; i++) {
      const rowName = normalizeHebrew(data[i][1]); // Col B
      const amount = Number(data[i][10]); // Col K
      const balanceToPay = Number(data[i][12]); // Col M
      
      if (rowName === normName) {
        const colN = data[i][13]; // Col N
        if (normalizeHebrew(colN) === 'אושרה הסבה') {
          sh.getRange(i + 2, 14).setValue('סיום טיפול הוסב'); // N
        }

        if (balanceToPay !== 0) {
          const sourceCity = data[i][2]; // Col C
          const plate = data[i][4]; // Col E
          const reportNumber = data[i][5]; // Col F
          const rawReportDate = data[i][6]; // Col G
          const reportDate = (rawReportDate instanceof Date) ? Utilities.formatDate(rawReportDate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : rawReportDate;
          const reportAmount = data[i][9]; // Col J
          const formattedReportAmount = !isNaN(Number(reportAmount)) && reportAmount !== "" ? Number(reportAmount).toFixed(2) : reportAmount;
          const comments = data[i][16]; // Col Q

          const details = `דוח מ${sourceCity}, מספר דוח: ${reportNumber}, תאריך ושעת דוח: ${reportDate}, סכום הדוח: ${formattedReportAmount}.`;

          newRows.push([
            data[i][1], // Customer name (original Col B)
            plate,
            "",
            "דוחות",
            details,
            todayStr,
            amount,
            comments,
            false, // Col I
            44 // Invoice code (Col J in office sheet) — fixed value for דוחות
          ]);

          sh.getRange(i + 2, 24).setValue(true);
        }
      }
    }
  }

  // -------- כביש 6 / מנהרות --------
  sh = ss.getSheetByName('כביש 6/מנהרות');
  last = sh.getLastRow();
  if (last >= 2) {
    const dataRange = sh.getRange(2, 1, last - 1, 21); // A to U (21 columns)
    const data = dataRange.getValues();
    for (let i = 0; i < data.length; i++) {
      const rowName = normalizeHebrew(data[i][4]); // Col E
      const amount = Number(data[i][10]); // Col K
      const isProcessed = data[i][11] === "טופל הועבר לטיפול המשרד"; // Col L
      if (rowName === normName && amount !== 0 && !isProcessed) {
        const plate = data[i][1]; // Col B
        const rawDateStr = data[i][3]; // Col D
        const dateStr = (rawDateStr instanceof Date) ? Utilities.formatDate(rawDateStr, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : rawDateStr;
        const sourceVal = String(data[i][5] || ""); // Col F
        const entrySegment = data[i][5]; // Col F
        const exitSegment = data[i][6]; // Col G
        const totalWithVat = Number(data[i][8]); // Col I
        const comments = data[i][12]; // Col M
        const invoiceCode = data[i][20]; // Col U — invoice code

        const source = sourceVal.includes("מנהרה") ? "מנהרות הכרמל" : "כביש 6";
        const commission = (amount - totalWithVat).toFixed(2);
        const formattedTotalWithVat = totalWithVat.toFixed(2);
        const details = `נסיעה בתאריך ושעה: ${dateStr}, מקטע כניסה: ${entrySegment}, מקטע יציאה: ${exitSegment}, סכום נסיעה כולל מע"מ (לפני עמלה): ${formattedTotalWithVat}, עמלה עבור נסיעה זו: ${commission}.`;

        newRows.push([
          data[i][4], // Customer name (original Col E)
          plate,
          "",
          source,
          details,
          todayStr,
          amount,
          comments,
          false, // Col I
          invoiceCode // Invoice code (Col J in office sheet) — from Col U
        ]);

        sh.getRange(i + 2, 12).setValue("טופל הועבר לטיפול המשרד");
      }
    }
  }

  // -------- חוצה צפון / נתיב מהיר --------
  sh = ss.getSheetByName('חוצה צפון/נתיב מהיר');
  last = sh.getLastRow();
  if (last >= 2) {
    const dataRange = sh.getRange(2, 1, last - 1, 18); // A to R (18 columns)
    const data = dataRange.getValues();
    for (let i = 0; i < data.length; i++) {
      const rowName = normalizeHebrew(data[i][4]); // Col E
      const amount = Number(data[i][8]); // Col I
      const isProcessed = data[i][9] === "טופל הועבר לטיפול המשרד"; // Col J
      if (rowName === normName && amount !== 0 && !isProcessed) {
        const plate = data[i][1]; // Col B
        const rawDateStr = data[i][3]; // Col D
        const dateStr = (rawDateStr instanceof Date) ? Utilities.formatDate(rawDateStr, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : rawDateStr;
        const sourceVal = String(data[i][5] || ""); // Col F
        const segment = data[i][5]; // Col F
        const totalWithVat = Number(data[i][6]); // Col G
        const comments = data[i][10]; // Col K
        const invoiceCode = data[i][17]; // Col R — invoice code

        const source = sourceVal.includes("נתיב המהיר") ? "נתיב המהיר" : "חוצה צפון";
        const commission = (amount - totalWithVat).toFixed(2);
        const formattedTotalWithVat = totalWithVat.toFixed(2);
        const details = `נסיעה בתאריך ושעה: ${dateStr}, מקטע נסיעה: ${segment}, סכום נסיעה כולל מע"מ (לפני עמלה): ${formattedTotalWithVat}, עמלה עבור נסיעה זו: ${commission}.`;

        newRows.push([
          data[i][4], // Customer name (original Col E)
          plate,
          "",
          source,
          details,
          todayStr,
          amount,
          comments,
          false, // Col I
          invoiceCode // Invoice code (Col J in office sheet) — from Col R
        ]);

        sh.getRange(i + 2, 10).setValue("טופל הועבר לטיפול המשרד");
      }
    }
  }

  // Append new rows to destination sheet
  if (newRows.length > 0) {
    let appendRow = targetSheet.getLastRow() + 1;
    const targetLast = targetSheet.getLastRow();
    if (targetLast > 0) {
      // Fetch columns A to J (10 columns)
      const targetData = targetSheet.getRange(1, 1, targetLast, 10).getValues();
      let found = false;
      for (let i = targetLast - 1; i >= 0; i--) {
        const row = targetData[i];
        const isEmpty = !String(row[0]).trim() && !String(row[1]).trim() &&
          !String(row[3]).trim() && !String(row[4]).trim() &&
          !String(row[5]).trim() && !String(row[6]).trim() &&
          !String(row[7]).trim();
        if (!isEmpty) {
          appendRow = i + 2;
          found = true;
          break;
        }
      }
      if (!found) {
        appendRow = 2; // Assuming row 1 is a header
      }
    } else {
      appendRow = 1;
    }
    targetSheet.getRange(appendRow, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
  console.log('Exiting processTransferToOffice', { newRowsCount: newRows.length });
}

// =======================================================================
// סימון חובות משרד כשולם (תשלומים, לטיפול המשרד)
// =======================================================================

function markOfficeAsPaid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('פירוט נסיעות לפי לקוח');
  const defaultName = sh.getRange(6, 3).getValue();

  const html = HtmlService.createTemplateFromFile('officeMarkConfirm');
  html.defaultName = defaultName;
  const dialog = html.evaluate()
    .setWidth(360)
    .setHeight(170)
    .setTitle('סמן חובות משרד כשולם');

  SpreadsheetApp.getUi().showModalDialog(dialog, 'סמן חובות משרד כשולם');
}

function openNamePickerForOfficeMark(defaultName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const html = HtmlService.createTemplateFromFile('namePicker');
  html.defaultName = defaultName;
  html.action = 'markOfficePaid';
  const dialog = html.evaluate()
    .setWidth(450)
    .setHeight(420)
    .setTitle("בחירת שם לסימון 'שולם' (משרד)");
  SpreadsheetApp.getUi().showModalDialog(dialog, "בחירת שם לסימון 'שולם' (משרד)");
}

function runMarkOfficeForName(selectedName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("🟢 Running markOfficeAsPaid...");

  // Normalize input: always treat as array
  const names = Array.isArray(selectedName) ? selectedName : [selectedName];

  names.forEach(name => {
    ss.toast(`מסמן כעת כשולם את חובות המשרד של:  ⭐ ⭐ ${name} ⭐ ⭐ `, "סימון כשולם (משרד)");
    markOfficeForName(name, ss);
    ss.toast(`חובות המשרד של הלקוח הבא: ⭐ ⭐ ${name} ⭐ ⭐ סומנו כשולם`, "סימון כשולם (משרד)");
  });
  ss.toast(`כל הלקוחות שנבחרו סומנו כשולם (משרד)`, "סימון כשולם (משרד)");
}

function markOfficeForName(name, ss) {
  const normName = normalizeHebrew(name);

  // -------- תשלומים --------
  let sh = ss.getSheetByName('תשלומים');
  if (sh) {
    let last = sh.getLastRow();
    if (last >= 2) {
      const data = sh.getRange(2, 1, last - 1, 4).getValues(); // A to D
      for (let i = 0; i < data.length; i++) {
        const rowName = normalizeHebrew(data[i][0]); // Col A
        const amount = Number(data[i][2]); // Col C
        const isHandled = data[i][3]; // Col D
        if (rowName === normName && amount !== 0 && isHandled !== true) {
          sh.getRange(i + 2, 4).setValue(true); // Col D
        }
      }
    }
  }

  // -------- לטיפול המשרד --------
  sh = ss.getSheetByName('לטיפול המשרד');
  if (sh) {
    let last = sh.getLastRow();
    if (last >= 2) {
      const data = sh.getRange(2, 1, last - 1, 9).getValues(); // A to I
      for (let i = 0; i < data.length; i++) {
        const rowName = normalizeHebrew(data[i][0]); // Col A
        const amount = Number(data[i][6]); // Col G
        const isHandled = data[i][8]; // Col I
        if (rowName === normName && amount !== 0 && isHandled !== true) {
          sh.getRange(i + 2, 9).setValue(true); // Col I
        }
      }
    }
  }

  // -------- הערות לקוחות לטיפול המשרד --------
  sh = ss.getSheetByName('הערות לקוחות לטיפול המשרד');
  if (sh) {
    let last = sh.getLastRow();
    if (last >= 2) {
      const data = sh.getRange(2, 1, last - 1, 1).getValues(); // Col A
      // Iterate backwards to safely delete rows
      for (let i = data.length - 1; i >= 0; i--) {
        const rowName = normalizeHebrew(data[i][0]);
        if (rowName === normName) {
          sh.deleteRow(i + 2);
        }
      }
    }
  }
}

// =======================================================================
// הערות לקוחות לטיפול המשרד
// =======================================================================

function openAddCustomerNoteDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('פירוט נסיעות לפי לקוח');
  let defaultName = "";
  if (sh) {
    defaultName = sh.getRange(6, 3).getValue();
  }

  const html = HtmlService.createTemplateFromFile('customerNoteDialog');
  html.defaultName = defaultName;
  const dialog = html.evaluate()
    .setWidth(450)
    .setHeight(400)
    .setTitle('הוספת הערה למשרד');

  SpreadsheetApp.getUi().showModalDialog(dialog, 'הוספת הערה למשרד');
}

function checkAndSaveCustomerNote(name, newNote, newReason, dealAccount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('הערות לקוחות לטיפול המשרד');
  if (!sh) return { status: 'error', message: 'הגיליון "הערות לקוחות לטיפול המשרד" לא נמצא' };

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const maxCols = Math.max(4, sh.getLastColumn());
    const data = sh.getRange(2, 1, lastRow - 1, maxCols).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][0]) === normalizeHebrew(name)) {
        const existingNote = data[i][1];
        const existingReason = data[i][2];
        const existingDealAccount = maxCols >= 4 ? data[i][3] : "";
        
        let hasNoteConflict = newNote && existingNote && String(existingNote).trim() !== "";
        let hasReasonConflict = newReason && existingReason && String(existingReason).trim() !== "";
        
        let newDealAccountToSave = dealAccount || "";
        if (dealAccount && existingDealAccount && String(existingDealAccount).trim() !== "") {
           newDealAccountToSave = String(existingDealAccount).trim() + ", " + dealAccount;
        } else if (existingDealAccount && String(existingDealAccount).trim() !== "") {
           newDealAccountToSave = existingDealAccount;
        }
        
        if (hasNoteConflict || hasReasonConflict) {
          return { 
            status: 'exists', 
            oldNote: existingNote || "",
            hasNoteConflict: !!hasNoteConflict,
            oldReason: existingReason || "",
            hasReasonConflict: !!hasReasonConflict,
            dealAccountToSave: newDealAccountToSave
          };
        } else {
          // No conflict.
          if (newNote) sh.getRange(i + 2, 2).setValue(newNote);
          if (newReason) sh.getRange(i + 2, 3).setValue(newReason);
          if (dealAccount) sh.getRange(i + 2, 4).setValue(newDealAccountToSave);
          return { status: 'success' };
        }
      }
    }
  }

  // Not found, append
  sh.appendRow([name, newNote || "", newReason || "", dealAccount || ""]);
  return { status: 'success' };
}

function overwriteCustomerNote(name, finalNote, finalReason, dealAccountToSave) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('הערות לקוחות לטיפול המשרד');
  if (!sh) return { status: 'error', message: 'הגיליון "הערות לקוחות לטיפול המשרד" לא נמצא' };

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const maxCols = Math.max(4, sh.getLastColumn());
    const data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][0]) === normalizeHebrew(name)) {
        if (finalNote !== undefined) sh.getRange(i + 2, 2).setValue(finalNote);
        if (finalReason !== undefined) sh.getRange(i + 2, 3).setValue(finalReason);
        if (dealAccountToSave !== undefined) sh.getRange(i + 2, 4).setValue(dealAccountToSave);
        return { status: 'success' };
      }
    }
  }

  // Should not happen if it existed, but just in case
  sh.appendRow([name, finalNote || "", finalReason || "", dealAccountToSave || ""]);
  return { status: 'success' };
}

function markAllAsGroupSent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('פירוט נסיעות לפי לקוח');
  const defaultName = sh.getRange(6, 3).getValue();

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "אישור פעולה",
    'האם לסמן את כל השורות ללא סטטוס כ"טופל נשלח בקבוצה" עבור:\n\n' +
    '            ------ ' + defaultName + ' ------',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    runGroupSentForName(defaultName);
  }
}

function runGroupSentForName(selectedName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = Array.isArray(selectedName) ? selectedName : [selectedName];

  names.forEach(name => {
    ss.toast(`מסמן כ"טופל נשלח בקבוצה" את כל החובות של:  ⭐ ⭐ ${name} ⭐ ⭐`, 'סימון נשלח בקבוצה');
    groupSentForName(name, ss);
    ss.toast(`כל החובות של ⭐ ⭐ ${name} ⭐ ⭐ סומנו כ"טופל נשלח בקבוצה"`, 'סימון נשלח בקבוצה');
  });
  ss.toast('כל הלקוחות שנבחרו סומנו כ"טופל נשלח בקבוצה"', 'סימון נשלח בקבוצה');
}

function groupSentForName(name, ss) {
  const normName = normalizeHebrew(name);
  const date = new Date();
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yy');

  // -------- דוחות --------
  let sh = ss.getSheetByName('דוחות');
  let last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 17).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][1]) !== normName) continue; // col B = name
      if (data[i][13] !== '') continue;                        // col N already has status
      const row = i + 2;
      sh.getRange(row, 14).setValue('טופל נשלח בקבוצה');
      sh.getRange(row, 15).setValue('נשלח בקבוצה');
      sh.getRange(row, 16).setValue(new Date());
      const oldComment = data[i][16]; // col Q
      sh.getRange(row, 17).setValue(oldComment ? `${oldComment}\n נשלח בקבוצה ${formattedDate}` : `נשלח בקבוצה ${formattedDate}`);
      sh.getRange(row, 12).setValue(true);
      sh.getRange(row, 11).setValue(0);
    }
  }

  // -------- כביש 6 / מנהרות --------
  sh = ss.getSheetByName('כביש 6/מנהרות');
  last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 13).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][4]) !== normName) continue; // col E = name
      if (data[i][11] !== '') continue;                        // col L already has status
      const row = i + 2;
      sh.getRange(row, 12).setValue('טופל נשלח בקבוצה');
      sh.getRange(row, 10).setValue('פטור');
      const oldComment = data[i][12]; // col M
      sh.getRange(row, 13).setValue(oldComment ? `${oldComment}\n נשלח בקבוצה ${formattedDate}` : `נשלח בקבוצה ${formattedDate}`);
    }
  }

  // -------- חוצה צפון / נתיב מהיר --------
  sh = ss.getSheetByName('חוצה צפון/נתיב מהיר');
  last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 11).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][4]) !== normName) continue; // col E = name
      if (data[i][9] !== '') continue;                         // col J already has status
      const row = i + 2;
      sh.getRange(row, 10).setValue('טופל נשלח בקבוצה');
      sh.getRange(row, 8).setValue('פטור');
      const oldComment = data[i][10]; // col K
      sh.getRange(row, 11).setValue(oldComment ? `${oldComment}\n נשלח בקבוצה ${formattedDate}` : `נשלח בקבוצה ${formattedDate}`);
    }
  }
}
