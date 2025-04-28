/////////////////////////////////////////////////////////////////////PROCESS SHEET/////////////////////////////////////////////////////////////////////

function processSaveSheet(sheet_sr, SheetName, Save, Edit)
{
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

    var A1 = sheet_sr.getRange("A1").getValue();
    var A2 = sheet_sr.getRange("A2").getValue();
    var A5 = sheet_sr.getRange("A5").getValue();

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

  if( !ErrorValues.includes(A2) )
  {
    if ( Save == "TRUE" )
    {
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
          Logger.log(`ERROR SAVE: ${SheetName} - EDIT on config is set to FALSE`);
        }
      }
      else
      {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveSheet`);
      }
    }
    if ( Save != "TRUE" )
    {
      Logger.log(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`);
    }
  }
  else
  {
    Logger.log(`ERROR SAVE: ${SheetName} - ErrorValues in A2 on processSaveSheet`);
  }
}

/////////////////////////////////////////////////////////////////////PROCESS EXTRA/////////////////////////////////////////////////////////////////////

function processSaveExtra(sheet_sr, SheetName, Save, Edit)
{
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

    var A1 = sheet_sr.getRange("A1").getValue();
    var A2 = sheet_sr.getRange("A2").getValue();
    var A5 = sheet_sr.getRange("A5").getValue();

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

  if( !ErrorValues.includes(A2) )
  {
    if ( Save == "TRUE" )
    {
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
          Logger.log(`ERROR SAVE: ${SheetName} - EDIT on config is set to FALSE`);
        }
      }
      else
      {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveExtra`);
      }
    }
    if ( Save != "TRUE" )
    {
      Logger.log(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`);
    }
  }
  else
  {
    Logger.log(`ERROR SAVE: ${SheetName} - ErrorValues in A2 on processSaveExtra`);
  }
}

/////////////////////////////////////////////////////////////////////PROCESS DATA/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////PROCESS DATA/////////////////////////////////////////////////////////////////////

function processSaveData(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Save, Edit) {
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
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveData`);
      }

      if ( Edit == "TRUE" ) {
        if (New_sr.valueOf() == New_tr.valueOf()) {
          edit_range_sr = sheet_sr.getRange("B1:B" + LR);
          edit_range_tr = sheet_tr.getRange("B1:B" + LR);
          editMappingFunc = (source, target) => source; // Identity mapping for comparison
        }
      }
      if ( Edit != "TRUE" ) {
        Logger.log(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`);
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
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveData`);
      }

      if ( Edit == "TRUE" ) {
        if (New_sr.valueOf() == New_tr.valueOf()) {
          edit_range_sr = sheet_sr.getRange("B1:B" + LR);
          edit_range_tr = sheet_sr.getRange("C1:C" + LR);
          editMappingFunc = (source, target) => source;
        }
      }
      if ( Edit != "TRUE" ) {
        Logger.log(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`);
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
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on processSaveData`);
      }

      if ( Edit == "TRUE" ) {
        if (New_sr.valueOf() == New_tr.valueOf()) {
          edit_range_sr = sheet_sr.getRange("C1:C" + LR);
          edit_range_tr = sheet_sr.getRange("D1:D" + LR);
          editMappingFunc = (source, target) => source;
        }
      }
      if ( Edit != "TRUE" ) {
        Logger.log(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`);
      }
    }
//-------------------------------------------------------------------Foot-------------------------------------------------------------------//
  }
  if ( Save != "TRUE" ) {
    Logger.log(`ERROR SAVE: ${SheetName} - SAVE on config is set to FALSE`);
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
    Logger.log(`SUCCESS SAVE. Sheet: ${SheetName}.`);
    doExportData(SheetName);
  }

  // Finally, perform the common EDIT check.
  if (edit_range_sr && edit_range_tr && editMappingFunc) {
    const editValues_sr = edit_range_sr.getValues();
    const editValues_tr = edit_range_tr.getValues();
    const areEqual = editValues_tr.every((row, index) => row[0] === editValues_sr[index][0]);
    if (!areEqual) {
      doEditData(SheetName);
    }
  }
}

/////////////////////////////////////////////////////////////////////SAVE PROCESS/////////////////////////////////////////////////////////////////////