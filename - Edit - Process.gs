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

function processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit) {
  const SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();
  const LR = sheet_tr ? sheet_tr.getLastRow() : sheet_sr.getLastRow();

  if (Edit !== "TRUE") {
    Logger.log(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`);
    return;
  }

  const financialMap = {
    [BLC]:       { col_sr: 2, col_tr: 2,     compareTo: 'New_tr',   mode: 'target' },
    [DRE]:       { col_sr: 2, col_tr: 2,     compareTo: 'New_tr',   mode: 'target' },
    [FLC]:       { col_sr: 2, col_tr: 2,     compareTo: 'New_tr',   mode: 'target' },
    [DVA]:       { col_sr: 2, col_tr: 2,     compareTo: 'New_tr',   mode: 'target' },

    [Balanco]:   { col_sr: 2, col_tr: 3,     compareTo: 'Old_sr',   mode: 'source' },
    [Resultado]: { col_sr: 3, col_tr: 4,     compareTo: 'Old_sr',   mode: 'source' },
    [Valor]:     { col_sr: 3, col_tr: 4,     compareTo: 'Old_sr',   mode: 'source' },
    [Fluxo]:     { col_sr: 3, col_tr: 4,     compareTo: 'Old_sr',   mode: 'source' }
  };

  const cfg = financialMap[SheetName];
  if (!cfg) {
    Logger.log(`ERROR EDIT: ${SheetName} - Conditions aren't met on processEditFinancial`);
    return;
  }

  const valRef = cfg.compareTo === 'New_tr' ? New_tr : Old_sr;

  if (New_sr.valueOf() > valRef.valueOf()) {
    doSaveFinancial(SheetName);
    return;
  }

  if (New_sr.valueOf() === valRef.valueOf()) {
    const range_sr = sheet_sr.getRange(1, cfg.col_sr, LR, 1);
    const range_tr = (sheet_tr || sheet_sr).getRange(1, cfg.col_tr, LR, 1);

    const values_sr = range_sr.getValues();
    const values_tr = range_tr.getValues();

    const updatedValues = cfg.mode === 'target'
      ? values_tr.map((row, i) => [row[0] !== values_sr[i][0] ? values_sr[i][0] : row[0]])
      : values_sr.map((row, i) => [row[0] !== values_tr[i][0] ? row[0] : values_tr[i][0]]);

    range_tr.setValues(updatedValues);
    Logger.log(`SUCCESS EDIT. Sheet: ${SheetName}.`);
    doExportFinancial(SheetName);
  }
}

/////////////////////////////////////////////////////////////////////EDIT PROCESS/////////////////////////////////////////////////////////////////////
