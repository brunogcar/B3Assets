/////////////////////////////////////////////////////////////////////PROCESS SHEET/////////////////////////////////////////////////////////////////////

function processEditSheet(sheet_sr, SheetName, Edit) 
{
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

    var A1 = sheet_sr.getRange("A1").getValue();
    var A2 = sheet_sr.getRange("A2").getValue();
    var A5 = sheet_sr.getRange("A5").getValue();



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

          Logger.log(`SUCCESS EDIT. Sheet: ${SheetName}.`);

          doExportSheet(SheetName);
        }
        else                                                                     // SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE 
        {
          var Data = sheet_sr.getRange(2,1,1,LC-4).getValues();                   // LC-4 to not overwrite Média data
          sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
          sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

          Logger.log(`SUCCESS EDIT. Sheet: ${SheetName}.`);

          doExportSheet(SheetName);
        }
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on processEditSheet`);
      }
    }
    if ( Edit != "TRUE" )
    {
      Logger.log(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`);
    }
  }
  else
  {
    Logger.log(`ERROR EDIT: ${SheetName} - ErrorValues in A2 on processEditSheet`);
  }
}

/////////////////////////////////////////////////////////////////////PROCESS EXTRA/////////////////////////////////////////////////////////////////////

function processEditExtra(sheet_sr, SheetName, Edit) 
{
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

    var A1 = sheet_sr.getRange("A1").getValue();
    var A2 = sheet_sr.getRange("A2").getValue();
    var A5 = sheet_sr.getRange("A5").getValue();

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

          Logger.log(`SUCCESS EDIT. Sheet: ${SheetName}.`);

          doExportExtra(SheetName);
        }
        else                                                                     // SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE 
        {
          var Data = sheet_sr.getRange(2,1,1,LC-4).getValues();                   // LC-4 to not overwrite Média data
          sheet_sr.getRange(5,1,Data.length,Data[0].length).setValues(Data);
          sheet_sr.getRange(1,1,Data.length,Data[0].length).setValues(Data);

          Logger.log(`SUCCESS EDIT. Sheet: ${SheetName}.`);

          doExportExtra(SheetName);
        }
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on processEditExtra`);
      }
    }
    if ( Edit != "TRUE" )
    {
      Logger.log(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`);
    }
  }
  else
  {
    Logger.log(`ERROR EDIT: ${SheetName} - ErrorValues in A2 on processEditExtra`);
  }
}

/////////////////////////////////////////////////////////////////////PROCESS DATA/////////////////////////////////////////////////////////////////////

function processEditData(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit) {

  const SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();
  const LR = sheet_tr ? sheet_tr.getLastRow() : sheet_sr.getLastRow();

  let range_sr, range_tr, mappingFunc;

  if ( Edit == "TRUE" ) {
    switch (SheetName) {
//-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//
      case BLC:
      case DRE:
      case FLC:
      case DVA:
        if (New_sr.valueOf() > New_tr.valueOf()) {
          doSaveData(SheetName);
          return;
        }
        if (New_sr.valueOf() == New_tr.valueOf()) {
          // For these sheets, the mapping is applied on the target values.
          range_sr = sheet_sr.getRange("B1:B" + LR);
          range_tr = sheet_tr.getRange("B1:B" + LR);
          mappingFunc = (source, target) =>
            target.map((row, index) => [row[0] !== source[index][0] ? source[index][0] : row[0]]);
        }
        break;
//-------------------------------------------------------------------Balanco-------------------------------------------------------------------//
      case Balanco:
        if (New_sr.valueOf() > Old_sr.valueOf()) {
          doSaveData(SheetName);
          return;
        }
        if (New_sr.valueOf() == Old_sr.valueOf()) {
          range_sr = sheet_sr.getRange("B1:B" + LR);
          range_tr = sheet_sr.getRange("C1:C" + LR);
          mappingFunc = (source, target) =>
            source.map((row, index) => [row[0] !== target[index][0] ? row[0] : target[index][0]]);
        }
        break;
//-------------------------------------------------------------------Resultado / Valor  / Fluxo-------------------------------------------------------------------//
      case Resultado:
      case Valor:
      case Fluxo:
        if (New_sr.valueOf() > Old_sr.valueOf()) {
          doSaveData(SheetName);
          return;
        }
        if (New_sr.valueOf() == Old_sr.valueOf()) {
          range_sr = sheet_sr.getRange("C1:C" + LR);
          range_tr = sheet_sr.getRange("D1:D" + LR);
          mappingFunc = (source, target) =>
            source.map((row, index) => [row[0] !== target[index][0] ? row[0] : target[index][0]]);
        }
        break;

      default:
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on processEditSheet`);
        return;
    }
//-------------------------------------------------------------------Foot-------------------------------------------------------------------//
  }
  if ( Edit != "TRUE" ) {
    Logger.log(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`);
  }

  /////////////////////////////////////////////////////////////////////PROCESS -  END/////////////////////////////////////////////////////////////////////
  // Common code block: update the values based on the mapping function
  if (range_sr && range_tr && mappingFunc) {
    const values_sr = range_sr.getValues();
    const values_tr = range_tr.getValues();
    const updatedValues = mappingFunc(values_sr, values_tr);
    range_tr.setValues(updatedValues);
    Logger.log(`SUCCESS EDIT. Sheet: ${SheetName}.`);
    doExportData(SheetName);
  }
}

/////////////////////////////////////////////////////////////////////EDIT PROCESS/////////////////////////////////////////////////////////////////////