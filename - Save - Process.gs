/////////////////////////////////////////////////////////////////////PROCESS SAVE/////////////////////////////////////////////////////////////////////

function processSaveGeneric(sheet_sr, SheetName, Save, Edit, exportFn) {
  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  const A1 = sheet_sr.getRange("A1").getValue();
  const A2 = sheet_sr.getRange("A2").getValue();
  const A5 = sheet_sr.getRange("A5").getValue();

  const Row1 = sheet_sr.getRange(1, 2, 1, 1).getValues()[0];
  const Row2 = sheet_sr.getRange(2, 2, 1, 1).getValues()[0];
  const Row5 = sheet_sr.getRange(5, 2, 1, 1).getValues()[0];

  // Handle SAVE = FALSE early
  if (Save !== "TRUE") {
    LogDebug(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`, 'MIN');
    return;
  }

  // Handle invalid A2 early
  if (ErrorValues.includes(A2)) {
    LogDebug(`ERROR SAVE: ${SheetName} - ErrorValues in A2 on processSave`, 'MIN');
    return;
  }

  const IsEqual = Row2.some((val, i) => val === Row1[i] || val === Row5[i]);

  if (A5 === "") {
    // Save only header
    const Data_Header = sheet_sr.getRange(2, 1, 1, LC).getValues();
    sheet_sr.getRange(5, 1, 1, LC).setValues(Data_Header);
    sheet_sr.getRange(1, 1, 1, LC).setValues(Data_Header);
    LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
    exportFn(SheetName);
    return;
  }

  if (A2 > A1 || A2 > A5) {
    // Save header and body
    const Data_Header = sheet_sr.getRange(2, 1, 1, LC).getValues();
    sheet_sr.getRange(5, 1, 1, LC).setValues(Data_Header);
    sheet_sr.getRange(1, 1, 1, LC).setValues(Data_Header);

    const Data_Body = sheet_sr.getRange(5, 1, LR - 4, LC).getValues();
    sheet_sr.getRange(6, 1, Data_Body.length, LC).setValues(Data_Body);

    LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
    exportFn(SheetName);
    return;
  }

  if (
    ((A2 === A5 || A2 === A1) && IsEqual) ||
    ErrorValues.includes(A1) || ErrorValues.includes(A5)
  ) {
    if (Edit === "TRUE") {
      doEditBasic(SheetName);
    } else {
      LogDebug(`ERROR SAVE: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
    }
    return;
  }

  LogDebug(`ERROR SAVE: ${SheetName} - Conditions aren't met on processSave`, 'MIN');
}

/////////////////////////////////////////////////////////////////////PROCESS BASIC AND EXTRA/////////////////////////////////////////////////////////////////////

function processSaveBasic(sheet_sr, SheetName, Save, Edit) {
  processSaveGeneric(sheet_sr, SheetName, Save, Edit, doExportBasic);
}

function processSaveExtra(sheet_sr, SheetName, Save, Edit) {
  processSaveGeneric(sheet_sr, SheetName, Save, Edit, doExportExtra);
}



/////////////////////////////////////////////////////////////////////PROCESS FINANCIAL/////////////////////////////////////////////////////////////////////

/**
 * Saves financial sheet data, backing up older columns and optionally triggering exports/edits.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|null} sheet_tr The “template” sheet (or null if write-back is on source).
 * @param {GoogleAppsScript.Spreadsheet.Sheet}      sheet_sr The source sheet where new data lives.
 * @param {number|string}                           New_tr    The new template date millis or blank.
 * @param {number|string}                           Old_tr    The old template date millis or blank.
 * @param {number|string}                           New_sr    The new source date millis or blank.
 * @param {number|string}                           Old_sr    The old source date millis or blank.
 * @param {string}                                  Save      “TRUE” if SAVE is enabled in config.
 * @param {string}                                  Edit      “TRUE” if EDIT is enabled in config.
 */
function processSaveFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Save, Edit) {
  var LR = sheet_tr ? sheet_tr.getLastRow() : sheet_sr.getLastRow();
  var LC = sheet_tr ? sheet_tr.getLastColumn() : sheet_sr.getLastColumn();
  var SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();

  // Main SAVE update variables:
  let save_range_sr, save_range_tr, mappingFunc;
  // Backup update variables:
  let backup_range_sr, backup_range_tr, backupMappingFunc;
  // EDIT update variables:
  let edit_range_sr, edit_range_tr, editMappingFunc;

  if ( Save == "TRUE" ) {

    //-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//
    if ([BLC, DRE, FLC, DVA].includes(SheetName)) {
      if (New_sr.valueOf() > Old_sr.valueOf()) {
        if (Old_sr.valueOf() == "") {
          save_range_sr = sheet_sr.getRange(1, 2, LR, 1);
          save_range_tr = sheet_tr.getRange(1, 2, LR, 1);
          mappingFunc = (source, target) => source;
        } else {
          backup_range_sr = sheet_tr.getRange(1, 2, LR, LC - 1);
          backup_range_tr = sheet_tr.getRange(1, 3, LR, LC - 1);
          backupMappingFunc = (source, target) => source;

          save_range_sr = sheet_sr.getRange(1, 2, LR, 1);
          save_range_tr = sheet_tr.getRange(1, 2, LR, 1);
          mappingFunc = (source, target) => source;
        }
      } else {
        LogDebug(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveFinancial`, 'MIN');
      }

      if ( Edit == "TRUE" ) {
        if (New_sr.valueOf() == New_tr.valueOf()) {
          edit_range_sr = sheet_sr.getRange("B1:B" + LR);
          edit_range_tr = sheet_tr.getRange("B1:B" + LR);
          editMappingFunc = (source, target) => source;
        }
      }
      if ( Edit != "TRUE" ) {
        LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
      }
    }
    //-------------------------------------------------------------------Balanco-------------------------------------------------------------------//
    else if (SheetName === Balanco) {
      if (New_sr.valueOf() > Old_sr.valueOf()) {
        if (Old_sr.valueOf() == "") {
          save_range_sr = sheet_sr.getRange(1, 2, LR, 1);
          save_range_tr = sheet_sr.getRange(1, 3, LR, 1);
          mappingFunc = (source, target) => source;
        } else {
          backup_range_sr = sheet_sr.getRange(1, 3, LR, LC - 2);
          backup_range_tr = sheet_sr.getRange(1, 4, LR, LC - 2);
          backupMappingFunc = (source, target) => source;

          save_range_sr = sheet_sr.getRange(1, 2, LR, 1);
          save_range_tr = sheet_sr.getRange(1, 3, LR, 1);
          mappingFunc = (source, target) => source;
        }
      } else {
        LogDebug(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveFinancial`, 'MIN');
      }

      if ( Edit == "TRUE" ) {
        if (New_sr.valueOf() == New_tr.valueOf()) {
          edit_range_sr = sheet_sr.getRange("B1:B" + LR);
          edit_range_tr = sheet_sr.getRange("C1:C" + LR);
          editMappingFunc = (source, target) => source;
        }
      }
      if ( Edit != "TRUE" ) {
        LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
      }
    }
    //-------------------------------------------------------------------Resultado / Valor / Fluxo-------------------------------------------------------------------//
    else if ([Resultado, Valor, Fluxo].includes(SheetName)) {
      if (New_sr.valueOf() > Old_sr.valueOf()) {
        if (Old_sr.valueOf() == "") {
          save_range_sr = sheet_sr.getRange(1, 3, LR, 1);
          save_range_tr = sheet_sr.getRange(1, 4, LR, 1);
          mappingFunc = (source, target) => source;
        } else {
          backup_range_sr = sheet_sr.getRange(1, 4, LR, LC - 3);
          backup_range_tr = sheet_sr.getRange(1, 5, LR, LC - 3);
          backupMappingFunc = (source, target) => source;

          save_range_sr = sheet_sr.getRange(1, 3, LR, 1);
          save_range_tr = sheet_sr.getRange(1, 4, LR, 1);
          mappingFunc = (source, target) => source;
        }
      } else {
        LogDebug(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveFinancial`, 'MIN');
      }

      if ( Edit == "TRUE" ) {
        if (New_sr.valueOf() == New_tr.valueOf()) {
          edit_range_sr = sheet_sr.getRange("C1:C" + LR);
          edit_range_tr = sheet_sr.getRange("D1:D" + LR);
          editMappingFunc = (source, target) => source;
        }
      }
      if ( Edit != "TRUE" ) {
        LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
      }
    }
//-------------------------------------------------------------------Foot-------------------------------------------------------------------//
  }
  if ( Save != "TRUE" ) {
    LogDebug(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`, 'MIN');
  }

  /////////////////////////////////////////////////////////////////////COMMON UPDATE BLOCK/////////////////////////////////////////////////////////////////////
  // Perform backup update first, if defined.
  if (backup_range_sr && backup_range_tr && backupMappingFunc) {
    const backupValues = backup_range_sr.getValues();
    backup_range_tr.setValues(backupMappingFunc(backupValues, backup_range_tr.getValues()));
  }

  // Then perform the main SAVE update.
  if (save_range_sr && save_range_tr && mappingFunc) {
    const values_sr = save_range_sr.getValues();
    const values_tr = save_range_tr.getValues();
    const updatedValues = mappingFunc(values_sr, values_tr);
    save_range_tr.setValues(updatedValues);
    LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
    doExportFinancial(SheetName);
  }

  // Finally, perform the common EDIT check.
  if (edit_range_sr && edit_range_tr && editMappingFunc) {
    const editValues_sr = edit_range_sr.getValues();
    const editValues_tr = edit_range_tr.getValues();
    const areEqual = editValues_tr.every((row, index) => row[0] === editValues_sr[index][0]);
    if (!areEqual) {
      doEditFinancial(SheetName);
    }
  }
}

/////////////////////////////////////////////////////////////////////SAVE PROCESS/////////////////////////////////////////////////////////////////////
