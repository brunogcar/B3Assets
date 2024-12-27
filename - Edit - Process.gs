/////////////////////////////////////////////////////////////////////PROCESS SHEET/////////////////////////////////////////////////////////////////////

function processEditSheet(sheet_sr, SheetName, Edit) 
{
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

    var A1 = sheet_sr.getRange('A1').getValue();
    var A2 = sheet_sr.getRange('A2').getValue();
    var A5 = sheet_sr.getRange('A5').getValue();

  if (sheet_sr)
  {
    if( !ErrorValues.includes(A2) )
    {
      if ( Edit == "TRUE" )
      { 
        if( A5 == "" || A2.valueOf() > A5.valueOf() || A2.valueOf() > A1.valueOf() )
        {
          doSaveSheet(SheetName);
        }
        else if( ( A2.valueOf() >= A5.valueOf() || A2.valueOf() >= A1.valueOf()) ||
                 ( ErrorValues.includes(A1) || ErrorValues.includes(A5) ) )
        {
          if (SheetName === FUND )
          {
            var Data = sheet_sr.getRange(2,1,1,LC).getValues();
            sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
            sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

            console.log(`SUCESS EDIT. Sheet: ${SheetName}.`);

            doExportSheet(SheetName);
          }
          else                                                                     // SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE 
          {
            var Data = sheet_sr.getRange(2,1,1,LC-4).getValues();                   // LC-4 to not overwrite Média data
            sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
            sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

            console.log(`SUCESS EDIT. Sheet: ${SheetName}.`);

            doExportSheet(SheetName);
          }
        }
        else
        {
          console.log('ERROR EDIT:', SheetName, 'Conditions arent met on processEditSheet');
        }
      }
       if ( Edit != "TRUE" )
      {
        console.log('ERROR EDIT:', SheetName, 'EDIT on config is set to FALSE');
      }
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'ErrorValues in A2 on processEditSheet');
    }
  }
  else
  {
    console.log('ERROR EDIT:', SheetName, 'Sheet does not exist');
  }
}

/////////////////////////////////////////////////////////////////////PROCESS EXTRA/////////////////////////////////////////////////////////////////////

function processEditExtra(sheet_sr, SheetName, Edit) 
{
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

    var A1 = sheet_sr.getRange('A1').getValue();
    var A2 = sheet_sr.getRange('A2').getValue();
    var A5 = sheet_sr.getRange('A5').getValue();

  if (sheet_sr)
  {
    if( !ErrorValues.includes(A2) )
    {
      if ( Edit == "TRUE" )
      { 
        if( A5 == "" || A2.valueOf() > A5.valueOf() || A2.valueOf() > A1.valueOf() )
        {
          doSaveSheet(SheetName);
        }
        else if( ( A2.valueOf() >= A5.valueOf() || A2.valueOf() >= A1.valueOf()) ||
                 ( ErrorValues.includes(A1) || ErrorValues.includes(A5) ) )
        {
          if (SheetName === FUND )
          {
            var Data = sheet_sr.getRange(2,1,1,LC).getValues();
            sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
            sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

            console.log(`SUCESS EDIT. Sheet: ${SheetName}.`);

            doExportExtra(SheetName);
          }
          else                                                                     // SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE 
          {
            var Data = sheet_sr.getRange(2,1,1,LC-4).getValues();                   // LC-4 to not overwrite Média data
            sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
            sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

            console.log(`SUCESS EDIT. Sheet: ${SheetName}.`);

            doExportExtra(SheetName);
          }
        }
        else
        {
          console.log('ERROR EDIT:', SheetName, 'Conditions arent met on processEditExtra');
        }
      }
       if ( Edit != "TRUE" )
      {
        console.log('ERROR EDIT:', SheetName, 'EDIT on config is set to FALSE');
      }
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'ErrorValues in A2 on processEditExtra');
    }
  }
  else
  {
    console.log('ERROR EDIT:', SheetName, 'Sheet does not exist');
  }
}

/////////////////////////////////////////////////////////////////////PROCESS DATA/////////////////////////////////////////////////////////////////////

function processEditData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Edit) 
{
    var LR = sheet_tr ? sheet_tr.getLastRow() : sheet_sr.getLastRow();
    var LC = sheet_tr ? sheet_tr.getLastColumn() : sheet_sr.getLastColumn();
    
    var SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();

  if ( Edit == "TRUE" )
  {

//-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//

    if (SheetName === BLC || SheetName === DRE || SheetName === FLC || SheetName === DVA) 
    {
      if( New_S.valueOf() > New_T.valueOf() )
      {
        doSaveData(SheetName) 
      }
      else if( New_S.valueOf() == New_T.valueOf() )
      {
        var B_S = sheet_sr.getRange("B1:B" + LR).getValues();
        var B = sheet_tr.getRange("B1:B" + LR).getValues();

        var Data = B.map(function(row, index) {return [row[0] != B_S[index][0] ? B_S[index][0] : row[0]];});
        sheet_tr.getRange(1,2,LR,1).setValues(Data);

        console.log(`SUCESS EDIT. Sheet: ${SheetName}.`);

        doExportData(SheetName);
      }
      else
      {
        console.log('ERROR EDIT:', SheetName, 'Conditions arent met on processEditSheet');
      }
    }

//-------------------------------------------------------------------Balanco-------------------------------------------------------------------//

    else if (SheetName === Balanco) 
    {
      if( New_S.valueOf() > Old_S.valueOf() )
      {
        doSaveData(SheetName) 
      }
      else if( New_S.valueOf() == Old_S.valueOf() )
      {
        var B = sheet_sr.getRange("B1:B" + LR).getValues();
        var C = sheet_sr.getRange("C1:C" + LR).getValues();

        var Data = B.map(function(row, index) {return [row[0] != C[index][0] ? row[0] : C[index][0]];});
        sheet_sr.getRange(1,3,LR,1).setValues(Data);

        console.log(`SUCESS EDIT. Sheet: ${SheetName}.`);
      }
      else
      {
        console.log('ERROR EDIT:', SheetName, 'Check conditions in processEditSheet');
      }
    }

//-------------------------------------------------------------------Resultado / Valor  / Fluxo-------------------------------------------------------------------//

    else if (SheetName === Resultado || SheetName === Valor || SheetName === Fluxo) 
    {
      if( New_S.valueOf() > Old_S.valueOf() )
      {
        doSaveData(SheetName) 
      }
      else if( New_S.valueOf() == Old_S.valueOf() )
      {
        var C = sheet_sr.getRange("C1:C" + LR).getValues();
        var D = sheet_sr.getRange("D1:D" + LR).getValues();

        var Data = C.map(function(row, index) {return [row[0] != D[index][0] ? row[0] : D[index][0]];});
        sheet_sr.getRange(1,4,LR,1).setValues(Data);

        console.log(`SUCESS EDIT. Sheet: ${SheetName}.`);
      }
      else
      {
        console.log('ERROR EDIT:', SheetName, 'Check conditions in processEditSheet');
      }
    }
  }

/////////////////////////////////////////////////////////////////////PROCESS -  END/////////////////////////////////////////////////////////////////////

  if ( Edit != "TRUE" )
  {
    console.log('ERROR EDIT:', SheetName, 'EDIT on config is set to FALSE');
  }
}

/////////////////////////////////////////////////////////////////////EDIT PROCESS/////////////////////////////////////////////////////////////////////