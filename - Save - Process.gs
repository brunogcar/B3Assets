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

  if (sheet_sr)
  {
    if( !ErrorValues.includes(A2) )
    {
      if ( Save == "TRUE" )
      {
        if( A5 == "" )
        {
          Data = sheet_sr.getRange(2,1,1,LC).getValues();
          sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
          sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);

          doExportSheet(SheetName);
        }
        else if( A2.valueOf() > A1.valueOf() || A2.valueOf() > A5.valueOf() )
        {
          Data_Backup = sheet_sr.getRange(5,1,LR-4,LC).getValues();
          Data = sheet_sr.getRange(2,1,1,LC).getValues();

          sheet_sr.getRange(6, 1, Data_Backup.length, Data_Backup[0].length).setValues(Data_Backup);
          sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
          sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);

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
            console.log('ERROR SAVE:', SheetName, 'EDIT on config is set to FALSE');
          }
        }
        else
        {
          console.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveSheet');
        }
      }
      if ( Save != "TRUE" )
      {
        console.log('ERROR SAVE:', SheetName, 'SAVE on config is set to FALSE');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, 'ErrorValues in A2 on processSaveSheet');
    }
  }
  else
  {
    console.log('ERROR SAVE:', SheetName, 'Sheet does not exist');
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

  if (sheet_sr)
  {
    if( !ErrorValues.includes(A2) )
    {
      if ( Save == "TRUE" )
      {
        if( A5 == "" )
        {
          Data = sheet_sr.getRange(2,1,1,LC).getValues();
          sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
          sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);

          doExportExtra(SheetName);
        }
        else if( A2.valueOf() > A1.valueOf() || A2.valueOf() > A5.valueOf() )
        {
          Data_Backup = sheet_sr.getRange(5,1,LR-4,LC).getValues();
          Data = sheet_sr.getRange(2,1,1,LC).getValues();

          sheet_sr.getRange(6, 1, Data_Backup.length, Data_Backup[0].length).setValues(Data_Backup);
          sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
          sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);

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
            console.log('ERROR SAVE:', SheetName, 'EDIT on config is set to FALSE');
          }
        }
        else
        {
          console.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveExtra');
        }
      }
      if ( Save != "TRUE" )
      {
        console.log('ERROR SAVE:', SheetName, 'SAVE on config is set to FALSE');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, 'ErrorValues in A2 on processSaveExtra');
    }
  }
  else
  {
    console.log('ERROR SAVE:', SheetName, 'Sheet does not exist');
  }
}

/////////////////////////////////////////////////////////////////////PROCESS DATA/////////////////////////////////////////////////////////////////////

function processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit)
{
    var LR = sheet_tr ? sheet_tr.getLastRow() : sheet_sr.getLastRow();
    var LC = sheet_tr ? sheet_tr.getLastColumn() : sheet_sr.getLastColumn();

    var SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();

  let Data;
  let Data_Backup;

//-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//

  if (SheetName === BLC || SheetName === DRE || SheetName === FLC || SheetName === DVA)
  {
    if ( Save == "TRUE" )
    {
      if (New_S.valueOf() > New_T.valueOf())
      {
        if (New_T.valueOf() == "")
        {
          Data = sheet_sr.getRange(1, 2, LR, 1).getValues();
          sheet_tr.getRange(1, 2, LR, 1).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);

          doExportData(SheetName);
        }
        else if (New_S.valueOf() > Old_S.valueOf())
        {
          Data_Backup = sheet_tr.getRange(1, 2, LR, LC - 1).getValues();
          sheet_tr.getRange(1, 3, LR, LC - 1).setValues(Data_Backup);

          Data = sheet_sr.getRange(1, 2, LR, 1).getValues();
          sheet_tr.getRange(1, 2, LR, 1).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);

        doExportData(SheetName);
        }
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveData');
      }
    }
    if ( Edit == "TRUE" )
    {
      if (New_S.valueOf() == New_T.valueOf())
      {
        var Range_S = sheet_sr.getRange("B1:B" + LR).getValues();
        var Range_T = sheet_tr.getRange("B1:B" + LR).getValues();

        var areValuesEqual = Range_T.map(function(row, index) {return row[0] === Range_S[index][0];});
        if (!areValuesEqual.every(Boolean))
        {
          doEditData(SheetName);
        }
      }
    }
    if ( Edit != "TRUE" )
    {
      console.log('ERROR EDIT:', SheetName, 'EDIT on config is set to FALSE');
    }
    if ( Save != "TRUE" )
    {
      console.log('ERROR SAVE:', SheetName, 'SAVE on config is set to FALSE');
    }
  }

//-------------------------------------------------------------------Balanco-------------------------------------------------------------------//

  else if (SheetName === Balanco)
  {
    if ( Save == "TRUE" )
    {
      if( New_S.valueOf() > Old_S.valueOf() )
      {
        if( Old_S.valueOf() == "" )
        {
          Data = sheet_sr.getRange(1,2,LR,1).getValues();
          sheet_sr.getRange(1,3,LR,1).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);
        }
        else
        {
          Data_Backup = sheet_sr.getRange(1,3,LR,LC-2).getValues();
          sheet_sr.getRange(1,4,LR,LC-2).setValues(Data_Backup);

          Data = sheet_sr.getRange(1,2,LR,1).getValues();
          sheet_sr.getRange(1,3,LR,1).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);
        }
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveData');
      }
    }
    if (Edit == "TRUE")
    {
      if ( New_S.valueOf() == Old_S.valueOf() )
      {
        var Range_S = sheet_sr.getRange("B1:B" + LR).getValues();
        var Range_T = sheet_sr.getRange("C1:C" + LR).getValues();

        var areValuesEqual = Range_T.map(function(row, index) {return row[0] === Range_S[index][0];});
        if (!areValuesEqual.every(Boolean))
        {
          doEditData(SheetName)
        }
      }
    }
    if ( Edit != "TRUE" )
    {
      console.log('ERROR EDIT:', SheetName, 'EDIT on config is set to FALSE');
    }
    if ( Save != "TRUE" )
    {
      console.log('ERROR SAVE:', SheetName, 'SAVE on config is set to FALSE');
    }

  }

//-------------------------------------------------------------------Resultado / Valor  / Fluxo-------------------------------------------------------------------//

  else if (SheetName === Resultado || SheetName === Valor || SheetName === Fluxo)
  {
    if ( Save == "TRUE" )
    {
      if( New_S.valueOf() > Old_S.valueOf())
      {
        if( Old_S.valueOf() == "" )
        {
          Data = sheet_sr.getRange(1,3,LR,1).getValues();
          sheet_sr.getRange(1,4,LR,1).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);
        }
        else
        {
          Data_Backup = sheet_sr.getRange(1,4,LR,LC-3).getValues();
          sheet_sr.getRange(1,5,LR,LC-3).setValues(Data_Backup);

          Data = sheet_sr.getRange(1,3,LR,1).getValues();
          sheet_sr.getRange(1,4,LR,1).setValues(Data);

          console.log(`SUCESS SAVE. Sheet: ${SheetName}.`);
        }
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on processSaveData');
      }
    }
    if (Edit == "TRUE")
    {
      if ( New_S.valueOf() == Old_S.valueOf() )
      {
        var Range_T = sheet_sr.getRange("C1:C" + LR).getValues();
        var Range_S = sheet_sr.getRange("D1:D" + LR).getValues();

        var areValuesEqual = Range_T.map(function(row, index) {return row[0] === Range_S[index][0];});

        if (!areValuesEqual.every(Boolean))
        {
          doEditData(SheetName)
        }
      }
    }
    if ( Edit != "TRUE" )
    {
      console.log('ERROR EDIT:', SheetName, 'EDIT on config is set to FALSE');
    }
    if ( Save != "TRUE" )
    {
      console.log('ERROR SAVE:', SheetName, 'SAVE on config is set to FALSE');
    }
  }
}

/////////////////////////////////////////////////////////////////////SAVE PROCESS/////////////////////////////////////////////////////////////////////