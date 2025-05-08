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
  const LR        = sheet_tr ? sheet_tr.getLastRow()    : sheet_sr.getLastRow();
  const LC        = sheet_tr ? sheet_tr.getLastColumn() : sheet_sr.getLastColumn();
  const SheetName = sheet_tr ? sheet_tr.getSheetName()  : sheet_sr.getSheetName();

  LogDebug(
    `DBG dates → New_tr=${New_tr} (${typeof New_tr}), Old_tr=${Old_tr} (${typeof Old_tr}), ` +
    `New_sr=${New_sr} (${typeof New_sr}), Old_sr=${Old_sr} (${typeof Old_sr})`,
    'MAX'
  );

  // bail out early if SAVE is disabled
  if (Save !== "TRUE") {
    LogDebug(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`, 'MIN');
    return;
  }

  // Configuration per‐sheet:
  // col_sr:        source column to save from
  // col_tr:        template column to save into (unless overridden by targetCol)
  // backupOffset:  how many cols back from last to leave intact
  // targetCol:     when present, write directly back into source at this column
  const financialMap = {
    [BLC]:       { col_sr: 2, col_tr: 2, backupOffset: 1 },
    [DRE]:       { col_sr: 2, col_tr: 2, backupOffset: 1 },
    [FLC]:       { col_sr: 2, col_tr: 2, backupOffset: 1 },
    [DVA]:       { col_sr: 2, col_tr: 2, backupOffset: 1 },

    [Balanco]:   { col_sr: 2, col_tr: 2, backupOffset: 2, targetCol: 3 },
    [Resultado]: { col_sr: 3, col_tr: 4, backupOffset: 3 },
    [Valor]:     { col_sr: 3, col_tr: 4, backupOffset: 3 },
    [Fluxo]:     { col_sr: 3, col_tr: 4, backupOffset: 3 }
  };

  const cfg = financialMap[SheetName];
  if (!cfg) {
    LogDebug(`ERROR SAVE: ${SheetName} not supported in processSaveFinancial`, 'MIN');
    return;
  }

  // pick write‐sheet: either the template (sheet_tr) or source (sheet_sr) when targetCol is set
  const sr       = sheet_sr;
  const tr       = cfg.targetCol != null ? sheet_sr : sheet_tr;
  const writeCol = cfg.targetCol != null ? cfg.targetCol : cfg.col_tr;

  let save_sr, save_tr, backup_sr, backup_tr, edit_sr, edit_tr;

  // only proceed when source date has advanced

  if (New_sr.valueOf() > New_tr.valueOf()) {
    if (New_tr.valueOf() === "") {
      // first‐time save entire column
      save_sr = sr.getRange(1, cfg.col_sr, LR, 1);
      save_tr = tr.getRange(1, writeCol,      LR, 1);
    }
    else if (New_sr.valueOf() > Old_sr.valueOf()) {
      // back up the existing “current” column(s)
      backup_sr = sr.getRange(1, cfg.col_sr + 1, LR, LC - cfg.backupOffset);
      backup_tr = tr.getRange(1, writeCol   + 1, LR, LC - cfg.backupOffset);

      // then overwrite with the new column
      save_sr = sr.getRange(1, cfg.col_sr,    LR, 1);
      save_tr = tr.getRange(1, writeCol,       LR, 1);
    }
  } else {
    LogDebug(`ERROR SAVE: ${SheetName} - Conditions aren’t met on processSaveFinancial`, 'MIN');
  }

  // if EDIT is enabled and new‐source equals new‐template, prepare an edit‐check
  if (Edit === "TRUE" && New_sr.valueOf() === New_tr.valueOf()) {
    edit_sr = sr.getRange(1, cfg.col_sr,     LR, 1);
    edit_tr = tr.getRange(1, writeCol + 1,   LR, 1);
  } else if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
  }

  // ——— perform backup ———
  if (backup_sr && backup_tr) {
    backup_tr.setValues(backup_sr.getValues());
  }

  // ——— perform main SAVE ———
  if (save_sr && save_tr) {
    save_tr.setValues(save_sr.getValues());
    LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
    doExportFinancial(SheetName);
  }

  // bail out early if EDIT is disabled
  if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
    return;
  }

  // ——— perform EDIT check ———
  if (edit_sr && edit_tr) {
    const src = edit_sr.getValues();
    const tgt = edit_tr.getValues();
    if (src.some((r,i) => r[0] !== tgt[i][0])) {
      doEditFinancial(SheetName);
    }
  }
}

/////////////////////////////////////////////////////////////////////SAVE PROCESS/////////////////////////////////////////////////////////////////////
