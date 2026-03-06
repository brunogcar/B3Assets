/////////////////////////////////////////////////////////////////////PROCESS SAVE/////////////////////////////////////////////////////////////////////

function processSaveGeneric(sheet_sr, SheetName, Save, Edit, exportFn) {

  if (Save !== "TRUE") {
    LogDebug(`❌ ERROR SAVE: ${SheetName} - SAVE is set to FALSE`, "MIN");
    return;
  }

  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  // Read rows 1–5, columns A–LC once
  const data = sheet_sr.getRange(1,1,5,LC).getValues();

  const A1 = data[0][0];
  const A2 = data[1][0];
  const A5 = data[4][0];

  // Full row comparisons starting from column B
  const Row1 = data[0].slice(1);
  const Row2 = data[1].slice(1);
  const Row5 = data[4].slice(1);

  const IsEqual = Row2.some((val, i) => val === Row1[i] || val === Row5[i]);

  const Header = [data[1]];

  if (ErrorValues.includes(A2)) {
    LogDebug(`❌ ERROR SAVE: ${SheetName} - ErrorValues in A2 ${A2}: processSaveGeneric`, "MIN");
    return;
  }

  if (A5 === "") {
    // Save only Header
    sheet_sr.getRange(5,1,1,LC).setValues(Header);
    sheet_sr.getRange(1,1,1,LC).setValues(Header);

    LogDebug(`✅ SUCCESS SAVE: ${SheetName}.`, "MIN");
    exportFn(SheetName);
    return;
  }

  if (A2 > A1 || A2 > A5) {
    // Save Header and Body
    const Body = sheet_sr.getRange(5,1,LR-4,LC).getValues();

    sheet_sr.getRange(6,1,Body.length,LC).setValues(Body);
    sheet_sr.getRange(5,1,1,LC).setValues(Header);
    sheet_sr.getRange(1,1,1,LC).setValues(Header);

    LogDebug(`✅ SUCCESS SAVE: ${SheetName}.`, "MIN");
    exportFn(SheetName);
    return;
  }

  if (
    ((A2 === A5 || A2 === A1) && IsEqual) ||
    ErrorValues.includes(A1) ||
    ErrorValues.includes(A5)
  ) {

    if (Edit === "TRUE") {
      doEditBasic(SheetName);
    } else {
      LogDebug(`❌ ERROR SAVE: ${SheetName} - EDIT is set to FALSE`, "MIN");
    }

    return;
  }

  LogDebug(`❌ ERROR SAVE: ${SheetName} - Conditions arent met: processSaveGeneric`, "MIN");
}

/////////////////////////////////////////////////////////////////////PROCESS BASIC AND EXTRA/////////////////////////////////////////////////////////////////////

function processSaveBasic(sheet_sr, SheetName, Save, Edit) {
  processSaveGeneric(sheet_sr, SheetName, Save, Edit, doExportBasic);
}

/**
 * Like processSaveBasic, but also trims the sheet for Swing.
 */
function processSaveSwing(sheet_sr, SheetName, Save, Edit) {
  processSaveGeneric(sheet_sr, SheetName, Save, Edit, doExportBasic);
  doTrimSheet(SheetName);
}

function processSaveExtra(sheet_sr, SheetName, Save, Edit) {
  processSaveGeneric(sheet_sr, SheetName, Save, Edit, doExportExtra);
}



/////////////////////////////////////////////////////////////////////PROCESS FINANCIAL/////////////////////////////////////////////////////////////////////

/**
 * Saves financial sheet data, backing up older columns and optionally triggering exports/edits.
 *
 * @param {Sheet}           sheet_tr  Target sheet (ticker)
 * @param {Sheet}           sheet_sr  Source sheet (template)
 * @param {Date|string}     New_tr  Parsed “new” date from target
 * @param {Date|string}     Old_tr  Parsed “old” date from target
 * @param {Date|string}     New_sr  Parsed “new” date from source
 * @param {Date|string}     Old_sr  Parsed “old” date from source
 * @param {boolean|string}  Save    “TRUE” if SAVE is enabled in config.
 * @param {boolean|string}  Edit    “TRUE” if EDIT is enabled in config.
 */
function processSaveFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr) {
  const SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();
  const cfg       = Object.values(financialMap)
                            .find(c => c.sh_tr === SheetName);
  if (!cfg) {
    LogDebug(`🚩 No financialMap entry: ${SheetName}`, 'MIN');
    return;
  }

  const LR = sheet_sr.getLastRow();
  const LC = cfg.recurse ? sheet_tr.getLastColumn() : sheet_sr.getLastColumn();

  let doSave = false;
  let doEdit = false;

  if (New_sr.valueOf() > Old_sr.valueOf()) {
    if (!cfg.recurse || New_sr.valueOf() > New_tr.valueOf()) {
      doSave = true;
    }
    else if (New_sr.valueOf() === New_tr.valueOf()) {
      doEdit = true;
    }
  }
  else if (New_sr.valueOf() === New_tr.valueOf()) {
    doEdit = true;
  }

  if (!doSave && !doEdit) {
    LogDebug(`⏩ SKIP ${SheetName}: no SAVE/EDIT condition met`, 'MID');
    return;
  }

  if (doSave) {
    const sheet_bk = cfg.recurse ? sheet_tr : sheet_sr;
    if (!isNaN(Old_sr.valueOf())) {
      const width    = LC - cfg.col_trg + 1;
      const backup_sr = sheet_bk.getRange(1, cfg.col_trg, LR, width);
      const backup_tr = sheet_bk.getRange(1, cfg.col_bak, LR, width);
      backup_tr.setValues(backup_sr.getValues());
      LogDebug(`✅ SUCCESS BACKUP: Range [${cfg.col_trg}→${cfg.col_trg+width-1}] → [${cfg.col_bak}→${cfg.col_bak+width-1}]: ${SheetName}`, 'MIN');
    }

    const save_sr = sheet_sr.getRange(1, cfg.col_src, LR, 1);
    const save_tr = sheet_tr.getRange(1, cfg.col_trg, LR, 1);
    save_tr.setValues(save_sr.getValues());
    LogDebug(`✅ SUCCESS SAVE: Column src=${cfg.col_src} → trg=${cfg.col_trg}: ${SheetName}`, 'MIN');

    if (cfg.recurse) {
      doExportFinancial(SheetName);
    }
  } else {
    LogDebug(`🏷️ Dates not advancing or aligned: ` + `Old_sr=${Old_sr}, New_sr=${New_sr}, New_tr=${New_tr}`, 'MIN');
  }

  // 3) EDIT branch
  if (doEdit) {
    // Compare the two columns in one shot:
    const diffs = getColumnDifferences(sheet_sr, sheet_tr, cfg.col_src, cfg.col_trg, LR);
    if (diffs.length) {
      LogDebug(`✏️ EDIT Detected: ${SheetName}`, 'MID');
      doEditFinancial(SheetName);
    } else {
      LogDebug(`🏷️ EDIT not needed: ${SheetName}`, 'MID');
    }
  }
}

/////////////////////////////////////////////////////////////////////SAVE PROCESS/////////////////////////////////////////////////////////////////////
