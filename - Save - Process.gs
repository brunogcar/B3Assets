/////////////////////////////////////////////////////////////////////PROCESS BASICS/////////////////////////////////////////////////////////////////////

function processSaveBasic(sheet_sr, SheetName, Save, Edit) {
  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  const A1 = sheet_sr.getRange("A1").getValue();
  const A2 = sheet_sr.getRange("A2").getValue();
  const A5 = sheet_sr.getRange("A5").getValue();

  const Row1 = sheet_sr.getRange(1, 2, 1, 1).getValues()[0];
  const Row2 = sheet_sr.getRange(2, 2, 1, 1).getValues()[0];
  const Row5 = sheet_sr.getRange(5, 2, 1, 1).getValues()[0];

  let IsEqual = false;
  for (let i = 0; i < Row2.length; i++) {
    if (Row2[i] == Row1[i] || Row2[i] == Row5[i]) {
      IsEqual = true;
      break;
    }
  }

  let Data_Header;
  let Data_Body;

  if (!ErrorValues.includes(A2)) {
    if (Save === "TRUE") {
      if (A5 === "") {
        Data_Header = sheet_sr.getRange(2, 1, 1, LC).getValues();
        sheet_sr.getRange(5, 1, Data_Header.length, Data_Header[0].length).setValues(Data_Header);
        sheet_sr.getRange(1, 1, Data_Header.length, Data_Header[0].length).setValues(Data_Header);

        LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
        doExportBasic(SheetName);

      } else if (A2.valueOf() > A1.valueOf() || A2.valueOf() > A5.valueOf()) {
        Data_Body = sheet_sr.getRange(5, 1, LR - 4, LC).getValues();
        Data_Header = sheet_sr.getRange(2, 1, 1, LC).getValues();

        sheet_sr.getRange(6, 1, Data_Body.length, Data_Body[0].length).setValues(Data_Body);
        sheet_sr.getRange(5, 1, Data_Header.length, Data_Header[0].length).setValues(Data_Header);
        sheet_sr.getRange(1, 1, Data_Header.length, Data_Header[0].length).setValues(Data_Header);

        LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
        doExportBasic(SheetName);

      } else if (
        ((A2.valueOf() == A5.valueOf() || A2.valueOf() == A1.valueOf()) && IsEqual) ||
        ErrorValues.includes(A1) || ErrorValues.includes(A5)
      ) {
        if (Edit === "TRUE") {
          doEditBasic(SheetName);
        } else {
          LogDebug(`ERROR SAVE: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
        }

      } else {
        LogDebug(`ERROR SAVE: ${SheetName} - Conditions aren't met on processSaveBasic`, 'MIN');
      }

    } else {
      LogDebug(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`, 'MIN');
    }

  } else {
    LogDebug(`ERROR SAVE: ${SheetName} - ErrorValues in A2 on processSaveBasic`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////PROCESS EXTRA/////////////////////////////////////////////////////////////////////

function processSaveExtra(sheet_sr, SheetName, Save, Edit)
{
  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  const A1 = sheet_sr.getRange("A1").getValue();
  const A2 = sheet_sr.getRange("A2").getValue();
  const A5 = sheet_sr.getRange("A5").getValue();

  const Row1 = sheet_sr.getRange(1, 2, 1, 1).getValues()[0];
  const Row2 = sheet_sr.getRange(2, 2, 1, 1).getValues()[0];
  const Row5 = sheet_sr.getRange(5, 2, 1, 1).getValues()[0];

  let IsEqual = false;
  for (let i = 0; i < Row2.length; i++)
  {
    if (Row2[i] == Row1[i] || Row2[i] == Row5[i])
    {
      IsEqual = true;
      break;
    }
  }

  let Data_Header;
  let Data_Body;

  if (!ErrorValues.includes(A2))
  {
    if (Save === "TRUE")
    {
      if (A5 === "")
      {
        Data_Header = sheet_sr.getRange(2, 1, 1, LC).getValues();
        sheet_sr.getRange(5, 1, 1, LC).setValues(Data_Header);
        sheet_sr.getRange(1, 1, 1, LC).setValues(Data_Header);

        LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
        doExportExtra(SheetName);
      }
      else if (A2 > A1 || A2 > A5)
      {
        Data_Body = sheet_sr.getRange(5, 1, LR - 4, LC).getValues();
        Data_Header = sheet_sr.getRange(2, 1, 1, LC).getValues();

        sheet_sr.getRange(6, 1, Data_Body.length, LC).setValues(Data_Body);
        sheet_sr.getRange(5, 1, 1, LC).setValues(Data_Header);
        sheet_sr.getRange(1, 1, 1, LC).setValues(Data_Header);

        LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
        doExportExtra(SheetName);
      }
      else if (
        ((A2 === A5 || A2 === A1) && IsEqual) ||
        ErrorValues.includes(A1) || ErrorValues.includes(A5)
      )
      {
        if (Edit === "TRUE")
        {
          doEditBasic(SheetName);
        }
        else
        {
          LogDebug(`ERROR SAVE: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
        }
      }
      else
      {
        LogDebug(`ERROR SAVE: ${SheetName} - Conditions aren't met on processSaveExtra`, 'MIN');
      }
    }
    else
    {
      LogDebug(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`, 'MIN');
    }
  }
  else
  {
    LogDebug(`ERROR SAVE: ${SheetName} - ErrorValues in A2 on processSaveExtra`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////PROCESS FINANCIAL/////////////////////////////////////////////////////////////////////

function processSaveFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Save, Edit) {
  const LR = sheet_tr ? sheet_tr.getLastRow() : sheet_sr.getLastRow();
  const LC = sheet_tr ? sheet_tr.getLastColumn() : sheet_sr.getLastColumn();
  const SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();

  if (Save !== "TRUE") {
    LogDebug(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`, 'MIN');
    return;
  }

  const financialMap = {
    [BLC]:       { baseCol: 2, backupOffset: 1, use_tr: true },
    [DRE]:       { baseCol: 2, backupOffset: 1, use_tr: true },
    [FLC]:       { baseCol: 2, backupOffset: 1, use_tr: true },
    [DVA]:       { baseCol: 2, backupOffset: 1, use_tr: true },

    [Balanco]:   { baseCol: 2, backupOffset: 2, use_tr: false },
    [Resultado]: { baseCol: 3, backupOffset: 3, use_tr: false },
    [Valor]:     { baseCol: 3, backupOffset: 3, use_tr: false },
    [Fluxo]:     { baseCol: 3, backupOffset: 3, use_tr: false }
  };

  const cfg = financialMap[SheetName];
  if (!cfg) {
    LogDebug(`ERROR: ${SheetName} not supported in processSaveFinancial`, 'MIN');
    return;
  }

  const sr = sheet_sr;
  const tr = cfg.use_tr ? sheet_tr : sheet_sr;

  let save_range_sr, save_range_tr, backup_range_sr, backup_range_tr;
  let edit_range_sr, edit_range_tr;

  if (New_sr.valueOf() > Old_sr.valueOf()) {
    if (Old_sr === "") {
      save_range_sr = sr.getRange(1, cfg.baseCol, LR, 1);
      save_range_tr = tr.getRange(1, cfg.baseCol, LR, 1);
    } else {
      backup_range_sr = sr.getRange(1, cfg.baseCol + 1, LR, LC - cfg.backupOffset);
      backup_range_tr = tr.getRange(1, cfg.baseCol + 2, LR, LC - cfg.backupOffset);

      save_range_sr = sr.getRange(1, cfg.baseCol, LR, 1);
      save_range_tr = tr.getRange(1, cfg.baseCol, LR, 1);
    }
  } else {
    LogDebug(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveFinancial`, 'MIN');
  }

  if (Edit === "TRUE") {
    if(New_sr.valueOf() === New_tr.valueOf()) {
      edit_range_sr = sr.getRange(1, cfg.baseCol, LR, 1);
      edit_range_tr = tr.getRange(1, cfg.baseCol + 1, LR, 1);
    }
  } else if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
  }

  // --- Perform backup ---
  if (backup_range_sr && backup_range_tr) {
    const backupValues = backup_range_sr.getValues();
    backup_range_tr.setValues(backupValues);
  }

  // --- Perform main SAVE ---
  if (save_range_sr && save_range_tr) {
    const values = save_range_sr.getValues();
    save_range_tr.setValues(values);
    LogDebug(`SUCCESS SAVE. Sheet: ${SheetName}.`, 'MIN');
    doExportFinancial(SheetName);
  }

  // --- Perform EDIT check ---
  if (edit_range_sr && edit_range_tr) {
    const values_sr = edit_range_sr.getValues();
    const values_tr = edit_range_tr.getValues();
    const areEqual = values_sr.every((row, i) => row[0] === values_tr[i][0]);
    if (!areEqual) {
      doEditFinancial(SheetName);
    }
  }
}

/////////////////////////////////////////////////////////////////////SAVE PROCESS/////////////////////////////////////////////////////////////////////
