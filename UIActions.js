
function runOnce_createOnOpenTrigger() {
  ScriptApp.newTrigger('onOpenWithUi')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}

function onOpenWithUi(e) {
  buildCustomMenu();
  runTicketCheck();
  // checkFollowUpReminders();
}

function onOpen(e) {
  buildCustomMenu();
}

function buildCustomMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 כלים מותאמים')
    .addItem('הצג לקוחות ממתינים לטיפול תאריך עבר', 'checkFollowUpReminders')
    .addItem('הצג כל הלקוחות הממתינים לטיפול', 'ShowAllFollowUpReminders')
    .addItem('הוסף תאריך לטיפול', 'updateDateMenu')
    .addItem('הצג דוחות לבדיקת סטטוס', 'runTicketCheck')
    .addSeparator()
    .addItem('הצג טבלת רשויות לדוחות', 'showSourcesDialog')
    .addSeparator()
    .addItem('סמן הכול כשולם', 'markAllAsPaid')
    .addSeparator()
    .addItem('ביצוע חיפוש בפירוט לפי לקוח', 'updateSearchInfo')
    .addSeparator()
    .addItem('סנכרן לקוחות', 'syncCustomerSheet')
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
    const html = HtmlService.createTemplateFromFile("namePicker");
    html.defaultName = defaultName;
    const dialog = html.evaluate()
      .setWidth(450)
      .setHeight(420)
      .setTitle("בחירת שם לסימון 'שולם'");

    ss.show(dialog);
  }
  else if (response == ui.Button.YES) {
    runMarkAllForName(defaultName)
  }
  else {
    ui.alert("הפעולה בוטלה.");
  }

}

function runMarkAllForName(selectedName) {
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
}

function getAllClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('לקוחות');
  const values = sh.getRange(1, 1, sh.getLastRow()).getValues();
  return values.map(r => r[0]).filter(x => x); // רק שמות לא ריקים
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
      if (rowName === normName && total > 0) {
        const tr = i + 2;
        sh.getRange(tr, 12).setValue(true); // L = true

        if (normalizeHebrew(colN) === 'אושרה הסבה') {
          sh.getRange(tr, 14).setValue('סיום טיפול הוסב'); // N
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
      if (rowName === normName && total > 0) {
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
      if (rowName === normName && total > 0) {
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

        if (rowName === normName && total > 0) {
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
