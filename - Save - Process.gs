/////////////////////////////////////////////////////////////////////PROCESS SHEET/////////////////////////////////////////////////////////////////////

function processSaveSheet(sheet_sr, SheetName, Save, Edit)
{
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

    var A1 = sheet_sr.getRange('A1').getValue();
    var A2 = sheet_sr.getRange('A2').getValue();
    var A5 = sheet_sr.getRange('A5').getValue();

  // Get the values of rows 1, 2, and 5, column B, until last col add LC - 1 instead of 1 
    var Row1 = sheet_sr.getRange(1, 2, 1, 1).getValues()[0]; // Fetch as a 1D array
    var Row2 = sheet_sr.getRange(2, 2, 1, 1).getValues()[0]; // Fetch as a 1D array
    var Row5 = sheet_sr.getRange(5, 2, 1, 1).getValues()[0]; // Fetch as a 1D array

  // Compare each corresponding cell in rows 1, 2, and 5
  var IsEqual = false;
  for (var i = 0; i < Row2.length; i++) 
  {
    if (Row2[i] == Row1[i] || Row2[i] == Row5[i]) 
    {
      IsEqual = true;
      break;
    }
  }

  let Data;
  let Data_Backup;

  if (ErrorValues.includes(A2)) {
    Logger.log('ERROR SAVE:', SheetName, 'ErrorValues in A2 on processSaveSheet');
    return;
  }

  if (Save !== "TRUE") {
    Logger.log('ERROR SAVE:', SheetName, 'SAVE on config is set to FALSE');
    return;
  }

  if( A5 == "" )
  {
    Data = sheet_sr.getRange(2,1,1,LC).getValues();
    sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
    sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

    Logger.log(`SUCCESS SAVE. Sheet: ${SheetName}.`);

    doExportSheet(SheetName);
  }
  else if( A2.valueOf() > A1.valueOf() || A2.valueOf() > A5.valueOf() )
  {
    Data_Backup = sheet_sr.getRange(5,1,LR-4,LC).getValues();
    Data = sheet_sr.getRange(2,1,1,LC).getValues();

    sheet_sr.getRange(6, 1, Data_Backup.length, Data_Backup[0].length).setValues(Data_Backup);
    sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
    sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

    Logger.log(`SUCCESS SAVE. Sheet: ${SheetName}.`);

    doExportSheet(SheetName);
  }
  else if( ( ( A2.valueOf() == A5.valueOf() || A2.valueOf() == A1.valueOf() ) && 
             ( IsEqual ) ) ||
           ( ErrorValues.includes(A1) || ErrorValues.includes(A5) ) )
  {
    if( Edit == "TRUE" )
    {
      doEditSheet(SheetName);
    }
    if ( Edit != "TRUE" )
    {
      Logger.log('ERROR SAVE:', SheetName, 'EDIT on config is set to FALSE');
    }
  }
  else
  {
    Logger.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveSheet');
  }
}

/////////////////////////////////////////////////////////////////////PROCESS EXTRA/////////////////////////////////////////////////////////////////////

function processSaveExtra(sheet_sr, SheetName, Save, Edit)
{
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

    var A1 = sheet_sr.getRange('A1').getValue();
    var A2 = sheet_sr.getRange('A2').getValue();
    var A5 = sheet_sr.getRange('A5').getValue();

  // Get the values of rows 1, 2, and 5, column B, until last col add LC - 1 instead of 1 
    var Row1 = sheet_sr.getRange(1, 2, 1, 1).getValues()[0]; // Fetch as a 1D array
    var Row2 = sheet_sr.getRange(2, 2, 1, 1).getValues()[0]; // Fetch as a 1D array
    var Row5 = sheet_sr.getRange(5, 2, 1, 1).getValues()[0]; // Fetch as a 1D array

  // Compare each corresponding cell in rows 1, 2, and 5
  var IsEqual = false;
  for (var i = 0; i < Row2.length; i++) 
  {
    if (Row2[i] == Row1[i] || Row2[i] == Row5[i]) 
    {
      IsEqual = true;
      break;
    }
  }

  let Data;
  let Data_Backup;

  if (ErrorValues.includes(A2)) {
    Logger.log('ERROR SAVE:', SheetName, 'ErrorValues in A2 on processSaveSheet');
    return;
  }

  if (Save !== "TRUE") {
    Logger.log('ERROR SAVE:', SheetName, 'SAVE on config is set to FALSE');
    return;
  }

  if( A5 == "" )
  {
    Data = sheet_sr.getRange(2,1,1,LC).getValues();
    sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
    sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

    Logger.log(`SUCCESS SAVE. Sheet: ${SheetName}.`);

    doExportExtra(SheetName);
  }
  else if( A2.valueOf() > A1.valueOf() || A2.valueOf() > A5.valueOf() )
  {
    Data_Backup = sheet_sr.getRange(5,1,LR-4,LC).getValues();
    Data = sheet_sr.getRange(2,1,1,LC).getValues();

    sheet_sr.getRange(6, 1, Data_Backup.length, Data_Backup[0].length).setValues(Data_Backup);
    sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
    sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

    Logger.log(`SUCCESS SAVE. Sheet: ${SheetName}.`);

    doExportExtra(SheetName);
  }
  else if( ( ( A2.valueOf() == A5.valueOf() || A2.valueOf() == A1.valueOf() ) && 
             ( IsEqual ) ) ||
           ( ErrorValues.includes(A1) || ErrorValues.includes(A5) ) )
  {
    if( Edit == "TRUE" )
    {
      doEditSheet(SheetName);
    }
    if ( Edit != "TRUE" )
    {
      Logger.log('ERROR SAVE:', SheetName, 'EDIT on config is set to FALSE');
    }
  }
  else
  {
    Logger.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveExtra');
  }
}

/////////////////////////////////////////////////////////////////////PROCESS DATA/////////////////////////////////////////////////////////////////////

function processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit) {
  var LR = sheet_tr ? sheet_tr.getLastRow() : sheet_sr.getLastRow();
  var LC = sheet_tr ? sheet_tr.getLastColumn() : sheet_sr.getLastColumn();
  var SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();

  // Global SAVE check flag (applies to all cases)
  const save = (Save === "TRUE");
  if (!save) {
    Logger.log('ERROR SAVE:', SheetName, 'SAVE on config is set to FALSE');
  }

  // Global EDIT check flag (applies to all cases)
  const edit = (Edit === "TRUE");
  if (!edit) {
    Logger.log('ERROR EDIT:', SheetName, 'EDIT on config is set to FALSE');
  }

  // Main SAVE update variables:
  let save_range_sr, save_range_tr, mappingFunc;
  // Backup update variables:
  let backup_range_sr, backup_range_tr, backupMappingFunc;

  //-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//
  if (SheetName === BLC || SheetName === DRE || SheetName === FLC || SheetName === DVA) {
    if (save && New_S.valueOf() > Old_S.valueOf()) {
      if (Old_S.valueOf() == "") {
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
    } else if (!(New_S.valueOf() > Old_S.valueOf())) {
      Logger.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveData');
    }

    if (edit && New_S.valueOf() == New_T.valueOf()) {
      let edit_range_sr = sheet_sr.getRange("B1:B" + LR).getValues();
      let edit_range_tr = sheet_tr.getRange("B1:B" + LR).getValues();
      let areValuesEqual = edit_range_tr.map((row, index) => row[0] === edit_range_sr[index][0]);
      if (!areValuesEqual.every(Boolean)) {
        doEditData(SheetName);
      }
    }
  }
  //-------------------------------------------------------------------Balanco-------------------------------------------------------------------//
  else if (SheetName === Balanco) {
    if (save && New_S.valueOf() > Old_S.valueOf()) {
      if (Old_S.valueOf() == "") {
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
    } else if (!(New_S.valueOf() > Old_S.valueOf())) {
      Logger.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveData');
    }

    if (edit && New_S.valueOf() == New_T.valueOf()) {
      let edit_range_sr = sheet_sr.getRange("B1:B" + LR).getValues();
      let edit_range_tr = sheet_sr.getRange("C1:C" + LR).getValues();
      let areValuesEqual = edit_range_tr.map((row, index) => row[0] === edit_range_sr[index][0]);
      if (!areValuesEqual.every(Boolean)) {
        doEditData(SheetName);
      }
    }
  }
  //-------------------------------------------------------------------Resultado / Valor / Fluxo-------------------------------------------------------------------//
  else if (SheetName === Resultado || SheetName === Valor || SheetName === Fluxo) {
    if (save && New_S.valueOf() > Old_S.valueOf()) {
      if (Old_S.valueOf() == "") {
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
    } else if (!(New_S.valueOf() > Old_S.valueOf())) {
      Logger.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveData');
    }

    if (edit && New_S.valueOf() == New_T.valueOf()) {
      let edit_range_tr = sheet_sr.getRange("C1:C" + LR).getValues();
      let edit_range_sr = sheet_sr.getRange("D1:D" + LR).getValues();
      let areValuesEqual = edit_range_tr.map((row, index) => row[0] === edit_range_sr[index][0]);
      if (!areValuesEqual.every(Boolean)) {
        doEditData(SheetName);
      }
    }
  }

  /////////////////////////////////////////////////////////////////////COMMON UPDATE BLOCK/////////////////////////////////////////////////////////////////////
  if (backup_range_sr && backup_range_tr && backupMappingFunc) {
    const backupValues = backup_range_sr.getValues();
    backup_range_tr.setValues(backupMappingFunc(backupValues, backup_range_tr.getValues()));
  }

  if (save_range_sr && save_range_tr && mappingFunc) {
    const values_sr = save_range_sr.getValues();
    const values_tr = save_range_tr.getValues();
    const updatedValues = mappingFunc(values_sr, values_tr);
    save_range_tr.setValues(updatedValues);
    Logger.log(`SUCCESS SAVE. Sheet: ${SheetName}.`);
    doExportData(SheetName);
  }
}

/////////////////////////////////////////////////////////////////////SAVE PROCESS/////////////////////////////////////////////////////////////////////