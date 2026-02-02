// --- Search Interface Functions ---

function updateSearchInfo() {
  const html = HtmlService.createTemplateFromFile("SearchForm");
  const dialog = html.evaluate()
    .setWidth(600)
    .setHeight(800) // Taller to accommodate the form
    .setTitle("ביצוע חיפוש בפירוט לפי לקוח");
  
  SpreadsheetApp.getUi().showModalDialog(dialog, "חיפוש מתקדם");
}

function getSearchFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Helper to get unique non-empty values from a column
  const getValues = (sheetName, colLetter) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Fetch Column (Assuming row 1 is header, so start from 2)
    const values = sheet.getRange(colLetter + "2:" + colLetter + lastRow).getValues();
    const unique = [...new Set(values.flat().filter(v => v !== "" && v != null))];
    return unique.sort();
  };

  return {
    names: getValues('לקוחות', 'A'),
    debtTypes: ['דוחות', 'חוצה צפון', 'כביש 6', 'תשלומים', 'סיכומי מחיר'],
    carNumbers: getValues('רשימת רכבים', 'A'),
    carModels: getValues('רשימת רכבים', 'E'),
    entrySections: getValues('רשימות', 'Z'),
    exitSections: getValues('רשימות', 'AA'),
    reportNumbers: getValues('דוחות', 'F'),
    statuses: getValues('רשימות', 'Y'),
    // Note: Notes (Field 9) is free text, Paid (Field 10) is static
  };
}

function processSearchForm(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('פירוט נסיעות לפי לקוח');
  if (!sh) throw new Error("Sheet 'פירוט נסיעות לפי לקוח' not found");

  // *** Requirement 3: Activate the sheet ***
  sh.activate();

  // Helper to join array or return empty string
  const joinVal = (arr) => Array.isArray(arr) ? arr.join(',') : (arr || '');

  // 1. Client Name
  sh.getRange('C5').clearContent();
  sh.getRange('A1').clearContent();
  
  if (form.clientName && form.clientName.length > 0) {
    sh.getRange('C5').setValue(form.clientName[0]);
    if (form.clientName.length > 1) {
      const rest = form.clientName.slice(1);
      sh.getRange('A1').setValue(rest.join(','));
    }
  }

  // 2-10. Other fields
  sh.getRange('B5').setValue(joinVal(form.debtType));
  sh.getRange('D5').setValue(joinVal(form.carNumber));
  sh.getRange('E5').setValue(joinVal(form.carModel));
  sh.getRange('G5').setValue(joinVal(form.entrySection));
  sh.getRange('H5').setValue(joinVal(form.exitSection));
  sh.getRange('I5').setValue(joinVal(form.reportNumber));
  sh.getRange('P5').setValue(joinVal(form.status));
  sh.getRange('Q5').setValue(form.notes || '');
  sh.getRange('I1').setValue(form.paidStatus || '');
}