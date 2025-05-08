/////////////////////////////////////////////////////////////////////PROCESS EDIT/////////////////////////////////////////////////////////////////////

function processEditGeneric(sheet_sr, SheetName, Edit, exportFn) {
  const LC = sheet_sr.getLastColumn();

  const A1 = sheet_sr.getRange("A1").getValue();
  const A2 = sheet_sr.getRange("A2").getValue();
  const A5 = sheet_sr.getRange("A5").getValue();

  if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
    return;
  }

  if (ErrorValues.includes(A2)) {
    LogDebug(`ERROR EDIT: ${SheetName} - ErrorValues in A2 on processEdit`, 'MIN');
    return;
  }

  if (A5 === "" || A2 > A5 || A2 > A1) {
    doSaveBasic(SheetName);
    return;
  }

  if (
    A2 >= A5 || A2 >= A1 ||
    ErrorValues.includes(A1) || ErrorValues.includes(A5)
  ) {
    const condition = (SheetName === FUND);
    const columnCount = condition ? LC : LC - 4;

    const Data_Header = sheet_sr.getRange(2, 1, 1, columnCount).getValues();
    sheet_sr.getRange(5, 1, 1, columnCount).setValues(Data_Header);
    sheet_sr.getRange(1, 1, 1, columnCount).setValues(Data_Header);

    LogDebug(`SUCCESS EDIT. Sheet: ${SheetName}.`, 'MIN');
    exportFn(SheetName);
    return;
  }

  // Final fallback
  LogDebug(`ERROR EDIT: ${SheetName} - Conditions aren't met on processEdit`, 'MIN');
}


/////////////////////////////////////////////////////////////////////PROCESS BASIC AND EXTRA/////////////////////////////////////////////////////////////////////

function processEditBasic(sheet_sr, SheetName, Edit) {
  processEditGeneric(sheet_sr, SheetName, Edit, doExportBasic);
}

function processEditExtra(sheet_sr, SheetName, Edit) {
  processEditGeneric(sheet_sr, SheetName, Edit, doExportExtra);
}

/////////////////////////////////////////////////////////////////////PROCESS FINANCIAL/////////////////////////////////////////////////////////////////////

/**
 * Applies an “edit” sync to financial sheets when the source & template dates match.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|null} sheet_tr The template sheet (or null if source-only).
 * @param {GoogleAppsScript.Spreadsheet.Sheet}      sheet_sr The source sheet.
 * @param {number|string}                           New_tr    New template date millis or blank.
 * @param {number|string}                           Old_tr    Old template date millis or blank.
 * @param {number|string}                           New_sr    New source date millis or blank.
 * @param {number|string}                           Old_sr    Old source date millis or blank.
 * @param {string}                                  Edit      “TRUE” if EDIT is enabled.
 */
function processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit) {
  const SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();
  const LR        = sheet_tr ? sheet_tr.getLastRow()   : sheet_sr.getLastRow();

  if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
    return;
  }

  // pick which column‐pair to touch; same keys & fallback to col_tr=col_sr+1 when use_tr=false
  const editMap = {
    [BLC]:       { col_sr: 2, col_tr: 2, compareTo: 'New_tr' },
    [DRE]:       { col_sr: 2, col_tr: 2, compareTo: 'New_tr' },
    [FLC]:       { col_sr: 2, col_tr: 2, compareTo: 'New_tr' },
    [DVA]:       { col_sr: 2, col_tr: 2, compareTo: 'New_tr' },

    [Balanco]:   { col_sr: 2, col_tr: 3, compareTo: 'Old_sr' },
    [Resultado]: { col_sr: 3, col_tr: 4, compareTo: 'Old_sr' },
    [Valor]:     { col_sr: 3, col_tr: 4, compareTo: 'Old_sr' },
    [Fluxo]:     { col_sr: 3, col_tr: 4, compareTo: 'Old_sr' }
  };

  const cfg = editMap[SheetName];
  if (!cfg) {
    LogDebug(`ERROR EDIT: ${SheetName} - Unsupported sheet in processEditFinancial`, 'MIN');
    return;
  }

  // pick which date to compare against
  const refDate = cfg.compareTo === 'New_tr' ? New_tr : Old_sr;

  // if source has moved past reference, treat as a save
  if (New_sr.valueOf() > refDate.valueOf()) {
    doSaveFinancial(SheetName);
    return;
  }

  // if dates exactly match, then do a cell‐by‐cell patch
  if (New_sr.valueOf() === refDate.valueOf()) {
    // write‐sheet is template unless writing back into source for Balanco
    const tr = cfg.col_tr === 3 && sheet_tr == null ? sheet_sr : sheet_tr;
    const sr = sheet_sr;

    const range_sr = sr.getRange(1, cfg.col_sr, LR, 1);
    const range_tr = tr.getRange(1, cfg.col_tr, LR, 1);

    const src = range_sr.getValues();
    const tgt = range_tr.getValues();
    // only replace cells that differ
    const updated = src.map((r,i) => [ r[0] !== tgt[i][0] ? r[0] : tgt[i][0] ]);

    range_tr.setValues(updated);
    LogDebug(`SUCCESS EDIT. Sheet: ${SheetName}.`, 'MIN');
    doExportFinancial(SheetName);
  }
}

/////////////////////////////////////////////////////////////////////EDIT PROCESS/////////////////////////////////////////////////////////////////////
