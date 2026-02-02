/**
 * Utility: Normalize text values by:
 *  - Converting to string
 *  - Replacing non-breaking spaces with normal spaces
 *  - Collapsing multiple whitespace into one
 *  - Trimming edges
 */
function normalizeHebrew(str) {
  return str
    .toString()
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * getSourcesTable: return rich values from named range “reprot_table”
 */
function getSourcesTable() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName('reprot_table');
  return range.getRichTextValues().map(row =>
    row.map(cell => {
      const url = cell.getLinkUrl();
      return url
        ? { text: cell.getText(), url }
        : cell.getText();
    })
  );
}

function updateFollowUpDate(key, val) {
  Logger.log(`saving date from menu, date selected is: ${val}, name is: ${key}`)
  const ss = SpreadsheetApp.getActive();
  const followSheet = ss.getSheetByName('follow_up_date');
  const data = followSheet.getRange('A:A').getValues();

  let found = false;

  // Update existing row
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      if (val === '') {
        // Clearing logic
        const cod_d = followSheet.getRange(i + 1, 4).getValue(); // Column D
        const col_c = followSheet.getRange(i + 1, 3).getValue(); // Column C
        const keepRow = (cod_d || col_c);

        if (keepRow) {
          followSheet.getRange(i + 1, 2).clearContent();
        } else {
          followSheet.deleteRow(i + 1);
        }
      } else {
        // Standard update
        followSheet.getRange(i + 1, 2).setValue(val);
      }

      found = true;
      break;
    }
  }

  // If not found and val not empty → append new
  if (!found && val !== '') {
    const lastRow = followSheet.getLastRow() + 1;
    followSheet.getRange(lastRow, 1).setValue(key);
    followSheet.getRange(lastRow, 2).setValue(val);
  }
}

// on edit auto trigger
function onEdit(e) {
  const ss  = e.source;
  const sh  = e.range.getSheet();
  const r   = e.range.getRow();
  const c   = e.range.getColumn();
  const val = e.range.getValue();

  // Log basic trigger info
  Logger.log(`onEdit triggered: sheet=${sh.getName()}, row=${r}, col=${c}, val=${val}`);

  //================================================================================
  // ===== update follow_up_date based on value added in פירוט נסיעות לפי לקוח =====
  //================================================================================
  // name added in C5
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && r === 5 && c === 3) {
    
    // 2. Define the ranges to check and clear
    const targetRanges = ['A1:B1', 'B5', 'E5:R5'];
    const rangeList = sh.getRangeList(targetRanges);
    
    // 3. Check if ANY of these cells have a value
    const ranges = rangeList.getRanges();
    let hasValue = ranges.some(range => {
      // flat() turns a 2D array into a 1D list; some() checks if any item is not empty
      return range.getValues().flat().some(cellValue => cellValue !== "");
    });

    // 4. If there is data, ask the user
    if (hasValue) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'ישנם סינונים נוספים מלבד השם שנוסף עכשיו,\nהאם לנקות נתוני סינון נוספים מלבד השם?', 
        ui.ButtonSet.YES_NO
      );

      // 5. If user clicks "YES"
      if (response == ui.Button.YES) {
        // Clear values only (keeps formatting and validation)
        rangeList.clearContent();
        
        // Update I1
        sh.getRange("I1").setValue("כולל שולם");
      }
    }
  }
  // date added in P3
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && r === 3 && c === 16) {
    Logger.log('P3 edited')
    const key = sh.getRange('C6').getValue();
    Logger.log(`Processing key: ${key}`);
    updateFollowUpDate(key, val)
    e.range.setValue('=LET(erorMessage, "אין תאריך לטיפול",lookupFuncion, VLOOKUP(C6,follow_up_date!A:B,2,FALSE),IFNA(IF(lookupFuncion="",erorMessage,lookupFuncion),erorMessage))')
  }

    // general comment for costomer added in R3
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && r === 3 && c === 14) {
    Logger.log('N3 edited')
    const key = sh.getRange('C6').getValue();
    Logger.log(`Processing key: ${key}`);

    const followSheet = ss.getSheetByName('follow_up_date');
    const data        = followSheet.getRange('A:A').getValues();

    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        Logger.log(`Key found in follow_up_date at row ${i + 1}. Updating column D to ${val}`);
        followSheet.getRange(i + 1, 4).setValue(val);
        found = true;
        break;
      }
    }

    if (!found) {
      const lastRow = followSheet.getLastRow() + 1;
      Logger.log(`Key not found. Appending new row ${lastRow} with (A: ${key}, B: ${val})`);
      followSheet.getRange(lastRow, 1).setValue(key);
      followSheet.getRange(lastRow, 4).setValue(val);
    }
    
    // If R3 is cleared → delete logic based on whether column C has value
    if (val === '') {
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === key) {
          // Get values from columns B and C (1 row, 2 columns)
          const col_b = followSheet.getRange(i + 1, 2).getValue(); // Column B
          const col_c = followSheet.getRange(i + 1, 3).getValue(); // Column C
          const hasValueInBOrC = (col_b || col_c); // true if either is non-empty

          if (hasValueInBOrC) {
            // Clear only column D
            followSheet.getRange(i + 1, 4).clearContent();
          } else {
            // Delete entire row
            followSheet.deleteRow(i + 1);
          }
          ;
        }
      }
      ;
    }

    e.range.setValue('=LET(erorMessage, "אין הערות כלליות ללקוח",lookupFuncion, VLOOKUP(C6,follow_up_date!A:D,4,FALSE),IFNA(IF(lookupFuncion="",erorMessage,lookupFuncion),erorMessage))')
  }

  // comment added in Q3
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && r === 3 && c === 17) { 
    Logger.log('Q3 edited');

    const key = sh.getRange('C6').getValue();
    const followSheet = ss.getSheetByName('follow_up_date');
    const data = followSheet.getRange('A:A').getValues();

    // Always restore VLOOKUP for Q3 after edit
    sh.getRange('Q3').setFormula(`=IFNA(IF(VLOOKUP($C$6,follow_up_date!A:C,3,false)="","אין הערות לטיפול להצגה",VLOOKUP($C$6,follow_up_date!A:C,3,false)),"אין הערות לטיפול להצגה")
    `);

    // If Q3 is cleared → delete logic based on whether column B has value
    if (val === '') {
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === key) {
          
          const col_b = followSheet.getRange(i + 1, 2).getValue(); // Column B
          const col_d = followSheet.getRange(i + 1, 4).getValue(); // Column D
          const hasValueInBOrD = (col_b || col_d); // true if either is non-empty

          if (hasValueInBOrD) {
            // Clear only column C
            followSheet.getRange(i + 1, 3).clearContent();
          } else {
            // Delete entire row
            followSheet.deleteRow(i + 1);
          }
          return;
        }
      }
      return;
    }

    // Otherwise, update or insert value in column C
    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        followSheet.getRange(i + 1, 3).setValue(val); // Column C
        found = true;
        Logger.log(`Updated existing row for key ${key} in column C with value ${val}`);
        break;
      }
    }

    if (!found) {
      const lastRow = followSheet.getLastRow() + 1;
      Logger.log(`Added new row for key ${key} with column C = ${val}`);
      followSheet.getRange(lastRow, 1).setValue(key); // Column A
      followSheet.getRange(lastRow, 3).setValue(val); // Column C
    }
  }

  // ===== A-column logic (col 1) =====
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && c === 1 && r >= 6) {
    if (val === true) {
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

      // reset flag back to false
      e.range.setValue(false);
    }
    return;
  }

  // ===== P-column logic (col 16) =====
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && c === 16 && r >= 6) {
    if (!val) return;
    const newVal = val;
    e.range.clearContent();
    Utilities.sleep(500);
    const key = normalizeHebrew(sh.getRange(r, 2).getValue());
    let target, writeCol;
    if (key === 'דוחות')       { target = ss.getSheetByName('דוחות');           writeCol = 14; }
    else if (key === 'כביש 6')  { target = ss.getSheetByName('כביש 6/מנהרות');   writeCol = 12; }
    else if (key === 'חוצה צפון'){ target = ss.getSheetByName('חוצה צפון/נתיב מהיר'); writeCol = 10; }
    else return;
    const last = target.getLastRow(); if (last < 2) return;
    if (key === 'דוחות') {
      const curI = normalizeHebrew(sh.getRange(r, 9).getValue());
      const curF = normalizeHebrew(sh.getRange(r, 6).getValue());
      const data = target.getRange(2, 6, last - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][1]) === curF &&
            normalizeHebrew(data[i][0]) === curI) {
          target.getRange(i + 2, writeCol).setValue(newVal);
          break;
        }
      }
    } else {
      const curD = normalizeHebrew(sh.getRange(r, 4).getValue());
      const curF = normalizeHebrew(sh.getRange(r, 6).getValue());
      const curG = normalizeHebrew(sh.getRange(r, 7).getValue());
      const data = target.getRange(2, 2, last - 1, 5).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][0]) === curD &&
            normalizeHebrew(data[i][2]) === curF &&
            normalizeHebrew(data[i][4]) === curG) {
          target.getRange(i + 2, writeCol).setValue(newVal);
          break;
        }
      }
    }
    return;
  }

  // ===== M-column logic (col 12) =====
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && c === 13 && r >= 6) {
    if (!val) return;
    const newVal = val === '.' ? '' : val;
    e.range.clearContent();
    Utilities.sleep(500);
    const key = normalizeHebrew(sh.getRange(r, 2).getValue());
    let target, writeCol;
    if (key === 'דוחות')       { target = ss.getSheetByName('דוחות');           writeCol = 14; }
    else if (key === 'כביש 6')  { target = ss.getSheetByName('כביש 6/מנהרות');   writeCol = 10; }
    else if (key === 'חוצה צפון'){ target = ss.getSheetByName('חוצה צפון/נתיב מהיר'); writeCol = 8; }
    else return;
    const last = target.getLastRow(); if (last < 2) return;
    if (key === 'דוחות') {
      const curI = normalizeHebrew(sh.getRange(r, 9).getValue());
      const curF = normalizeHebrew(sh.getRange(r, 6).getValue());
      const data = target.getRange(2, 6, last - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][1]) === curF &&
            normalizeHebrew(data[i][0]) === curI) {
          target.getRange(i + 2, writeCol).clearContent();
          break;
        }
      }
    } else {
      const curD = normalizeHebrew(sh.getRange(r, 4).getValue());
      const curF = normalizeHebrew(sh.getRange(r, 6).getValue());
      const curG = normalizeHebrew(sh.getRange(r, 7).getValue());
      const data = target.getRange(2, 2, last - 1, 5).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][0]) === curD &&
            normalizeHebrew(data[i][2]) === curF &&
            normalizeHebrew(data[i][4]) === curG) {
          target.getRange(i + 2, writeCol).setValue(newVal);
          break;
        }
      }
    }
    return;
  }

  // ===== O-column logic (col 15) =====
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && c === 15 && r >= 6) {
    const newVal = val;
    e.range.clearContent();
    Utilities.sleep(500);
    const key    = normalizeHebrew(sh.getRange(r, 2).getValue());
    const target = ss.getSheetByName('דוחות');
    if (!target) return;
    const last = target.getLastRow(); if (last < 2) return;
    const curI = normalizeHebrew(sh.getRange(r, 9).getValue());
    const curF = normalizeHebrew(sh.getRange(r, 6).getValue());
    const data = target.getRange(2, 6, last - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (normalizeHebrew(data[i][1]) === curF &&
          normalizeHebrew(data[i][0]) === curI) {
        const tr = i + 2;
        if (key === 'דוחות') {
          target.getRange(tr, 15).setValue(newVal);
          target.getRange(tr, 14).setValue('נשלחה בקשה להסבה');
        } else {
          target.getRange(tr, 15).clearContent();
        }
        break;
      }
    }
    return;
  }

  // ===== Q-column logic (col 17) =====
  if (sh.getName() === 'פירוט נסיעות לפי לקוח' && c === 17 && r >= 6) {
    if (!val) return;
    const qVal = val;
    e.range.clearContent();
    Utilities.sleep(500);
    const key  = normalizeHebrew(sh.getRange(r, 2).getValue());
    let target, writeCol;
    if (key === 'דוחות') {
      target   = ss.getSheetByName('דוחות');
      writeCol = 17;
    } else if (key === 'כביש 6') {
      target   = ss.getSheetByName('כביש 6/מנהרות');
      writeCol = 13;
    } else if (key === 'חוצה צפון') {
      target   = ss.getSheetByName('חוצה צפון/נתיב מהיר');
      writeCol = 11;
    } else {
      return;
    }
    const last = target.getLastRow(); if (last < 2) return;
    if (key === 'דוחות') {
      const curI = normalizeHebrew(sh.getRange(r, 9).getValue());
      const data = target.getRange(2, 6, last - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][1]) === normalizeHebrew(sh.getRange(r, 6).getValue()) &&
            normalizeHebrew(data[i][0]) === curI) {
          target.getRange(i + 2, writeCol).setValue(qVal);
          break;
        }
      }
    } else {
      const curG = normalizeHebrew(sh.getRange(r, 7).getValue());
      const data = target.getRange(2, 2, last - 1, 5).getValues();
      for (let i = 0; i < data.length; i++) {
        if (normalizeHebrew(data[i][0]) === normalizeHebrew(sh.getRange(r, 4).getValue()) &&
            normalizeHebrew(data[i][2]) === normalizeHebrew(sh.getRange(r, 6).getValue()) &&
            normalizeHebrew(data[i][4]) === curG) {
          target.getRange(i + 2, writeCol).setValue(qVal);
          break;
        }
      }
    }
    return;
  }

  // ===== Edits in 'כביש 6/מנהרות' =====
  if (sh.getName() === 'כביש 6/מנהרות') {
    let date = new Date();
    let formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");

    // for case of marking "טופל נשלח לשטריקר"
        if (c === 12 && typeof val === 'string' && val.includes('טופל נשלח לשטריקר')) {
      sh.getRange(r, 10).setValue("פטור");
      var oldComment = sh.getRange(r, 13).getValue()
      var newComment = oldComment ? `${oldComment}\n נשלח לאבי ${formattedDate}`: `נשלח לאבי ${formattedDate}`;
      sh.getRange(r, 13).setValue(newComment)
    }
  }

  // ===== Edits in 'חוצה צפון/נתיב מהיר' =====
  if (sh.getName() === 'חוצה צפון/נתיב מהיר') {
    Logger.log(`edit in "חוצה צפון/נתיב מהיר": row=${r}, column=${c}, value=${val}`)
    let date = new Date();
    let formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");

    // for case of marking "טופל נשלח לשטריקר"
        if (c === 10 && typeof val === 'string' && val.includes('טופל נשלח לשטריקר')) {
      sh.getRange(r, 8).setValue("פטור");
      var oldComment = sh.getRange(r, 11).getValue()
      var newComment = oldComment ? `${oldComment}\n נשלח לאבי ${formattedDate}`: `נשלח לאבי ${formattedDate}`;
      sh.getRange(r, 11).setValue(newComment)
    }
  }

  // ===== Edits in 'דוחות' =====
  if (sh.getName() === 'דוחות') {
    
    let date = new Date();
    let formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");
    
    // hyperlink links pasted in column A
    if (c === 1 && typeof val === 'string' && val.startsWith('http')) {
      sh.getRange(r, 1).setFormula(`=HYPERLINK("${val}","AA")`);
    }
    // for case of marking "טופל נשלח לשטריקר"
    if (c === 14 && typeof val === 'string' && val.includes('טופל נשלח לשטריקר')) {
      sh.getRange(r, 15).setValue("נשלח לשטריקר");
      sh.getRange(r, 16).setValue(new Date());
      var oldComment = sh.getRange(r, 17).getValue()
      var newComment = oldComment ? `${oldComment}\n נשלח לאבי ${formattedDate}`: `נשלח לאבי ${formattedDate}`;
      sh.getRange(r, 17).setValue(newComment)
      sh.getRange(r, 12).setValue(true)
      sh.getRange(r, 11).setValue(0)
    }
    // for case of adding מספר בקשה in column O
    if (c === 15) {
      sh.getRange(r, 14).setValue('נשלחה בקשה להסבה');
      sh.getRange(r, 16).setValue(new Date()); // Column P
    }
    // for case of marking column N (סטטוס) as סיום טיפול - mark as שולם
    if (c === 14 && typeof val === 'string' && val.includes('סיום טיפול')) {
      sh.getRange(r, 12).setValue(true);
    }
    // for case of marking column N (סטטוס) as אושרה הסבה - switch to סיום טיפול הוסב if שולם 
    if (c === 14 && normalizeHebrew(val) === 'אושרה הסבה') {
      if (sh.getRange(r, 13).getValue() === 0) {
        sh.getRange(r, 14).setValue('סיום טיפול הוסב');
      }
    }
    // for case of marking שולם in column L - if סטטוס is אושרה הסבה switch it to סיום טיפול הוסב
    if(c === 12 && val === true){
      col14Value = sh.getRange(r, 14).getValue()
      if (col14Value === 'אושרה הסבה'){
        sh.getRange(r, 14).setValue('סיום טיפול הוסב')
      }

    }
  }
}