/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doClearAll() {
  doClearProventos();
  doClearBasics();
  doClearFinancials();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////TRIM/////////////////////////////////////////////////////////////////////

function doTrimBasics() {
  const SheetNames = [SWING_4, SWING_12, SWING_52];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doTrimBasic(SheetName);
    } 
    catch (error) 
    {
      Logger.error(`Error saving sheet ${SheetName}: ${error}`);
    }
  });
}

function doTrimBasic(SheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = ss.getSheetByName(SheetName); // Target

  if (!sheet_sr) {Logger.error(`Sheet ${SheetName} not found.`); return;}

  var LR = sheet_sr.getLastRow();
  var LC = sheet_sr.getLastColumn();


  if (SheetName === SWING_4) 
  {
    if (LR > 81) 
    {
      sheet_sr.getRange(82, 1, LR - 81, LC).clearContent(); // Clear data below row 128
      Logger.log(`Cleared data below row 81 in ${SheetName}.`);
    }
  } 
  else if (SheetName === SWING_12) 
  {
    if (LR > 208) 
    {
      sheet_sr.getRange(209, 1, LR - 208, LC).clearContent(); // Clear data below row 128
      Logger.log(`Cleared data below row 208 in ${SheetName}.`);
    }
  } 
  else if (SheetName === SWING_52) 
  {

  } 
  else 
  {
    // Default logic for other sheets
    Logger.log(`No specific logic defined for ${SheetName}.`);
  }
}

/////////////////////////////////////////////////////////////////////COPY/////////////////////////////////////////////////////////////////////

function doCopyBasic(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);

  var LR = ss.getLastRow();
  var LC = ss.getLastColumn();

  ss.getRange(5, 1, LR - 4, LC).activate();
}

function doCopyFinancial(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);

  if (!ss) 
  {
    Logger.log(`ERROR COPY: ${SheetName} Does not exist`);
    return;
  }

  var LR = ss.getLastRow();
  var LC = ss.getLastColumn();

  if (SheetName === BLC||SheetName === DRE || SheetName === FLC || SheetName === DVA) 
  {
    ss.getRange(1, 2, LR, LC - 1).activate();
  } 
  else if (SheetName === Balanco) 
  {
    ss.getRange(1, 3, LR, LC - 2).activate();
  } 
  else if (SheetName === Resultado || SheetName === Valor || SheetName === Fluxo) 
  {
    ss.getRange(1, 4, LR, LC - 3).activate();
  }
  else 
  {
    Logger.error(`Unsupported sheet name: ${SheetName}`);
  }
}

/////////////////////////////////////////////////////////////////////CLEAR/////////////////////////////////////////////////////////////////////

function doClearBasics() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND, FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
    doClearBasic(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      Logger.error(`Error Clearing  ${SheetName}: ${error}`);
    }
  });
}

function doClearBasic(SheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);

  var LR = ss.getLastRow();
  var LC = ss.getLastColumn();

  Logger.log(`Clear: ${SheetName}`);

  ss.getRange(5, 1, LR, LC).clear({ contentsOnly: true, skipFilteredRows: false });
  ss.getRange(1, 1, 1, LC).clear({ contentsOnly: true, skipFilteredRows: false });

  Logger.log(`Data Cleared successfully. Sheet: ${SheetName}.`);

}

function doClearFinancials() {
  const SheetNames = [BLC, DRE, FLC, DVA];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
    doClearFinancial(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      Logger.error(`Error Clearing  ${SheetName}: ${error}`);
    }
  });
}

function doClearFinancial(SheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);

  Logger.log(`Clear: ${SheetName}`);

  var LR = ss.getLastRow();
  var LC = ss.getLastColumn();

  if (SheetName === BLC || SheetName === DRE || SheetName === FLC || SheetName === DVA) 
  {
    ss.getRange(1, 2, LR, LC - 1).clear({ contentsOnly: true, skipFilteredRows: false });

    Logger.log(`Data Cleared successfully. Sheet: ${SheetName}.`);
  } 
  else if 
  (SheetName === Balanco) 
  {
    ss.getRange(1, 3, LR, LC - 2).clear({ contentsOnly: true, skipFilteredRows: false });

    Logger.log(`Data Cleared successfully. Sheet: ${SheetName}.`);
  } 
  else if 
  (SheetName === Resultado || SheetName === Valor || SheetName === Fluxo) 
  {
    ss.getRange(1, 4, LR, LC - 3).clear({ contentsOnly: true, skipFilteredRows: false });

    Logger.log(`Data Cleared successfully. Sheet: ${SheetName}.`);
  } 
  else 
  {
    Logger.error(`Unsupported sheet name: ${SheetName}`);
  }

  if (SheetName === BLC) {doClearFinancial(Balanco);} 
  else if (SheetName === DRE) {doClearFinancial(Resultado );} 
  else if (SheetName === FLC) {doClearFinancial(Fluxo);} 
  else if (SheetName === DVA) {doClearFinancial(Valor);}
}


function doClearProventos()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Prov');

  var LR = ss.getLastRow();
  var LC = ss.getLastColumn();

  ss.getRange(PRV).clear({contentsOnly: true, skipFilteredRows: false});                             // PRV = Provento Range
};

/////////////////////////////////////////////////////////////////////ALTERNATIVE CLEAR/////////////////////////////////////////////////////////////////////

function doRecycleTrade()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = ss.getSheetByName(TRADE);

    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();

  const sheet_co = ss.getSheetByName('Config');

    var AX = sheet_co.getRange(PDT).getDisplayValue();        // PDT = Periodo de Trade
    var AX_ = sheet_sr.getRange("A" + AX ).getValue();

//  Logger.log(AX_);

  if( AX_ !== "" )
  {
    sheet_sr.getRange(AX,1,LR,LC).clear({contentsOnly: true, skipFilteredRows: false});
  }
};

/////////////////////////////////////////////////////////////////////CLEAN/////////////////////////////////////////////////////////////////////

function doCleanBasics() 
{
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND, FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doCleanBasic(SheetName);
    } 
    catch (error) 
    {
      Logger.error(`Error cleaning sheet ${SheetName}: ${error}`);
    }
  });
}


function doCleanBasic(SheetName) 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);

  Logger.log(`CLEAN: ${SheetName}`);

  if (!sheet) 
  {
    Logger.log(`ERROR CLEAN: ${SheetName} Does not exist`);
    return;
  }

  var LR = sheet.getLastRow();
  var LC = sheet.getLastColumn();

  sheet.getRange(5, 1, LR, LC).setValue('');
  sheet.createTextFinder("-").matchEntireCell(true).replaceAllWith("");
  sheet.createTextFinder("0").matchEntireCell(true).replaceAllWith("");
  sheet.createTextFinder("0,00").matchEntireCell(true).replaceAllWith("");
  sheet.createTextFinder("0,0000").matchEntireCell(true).replaceAllWith("");

  Logger.log(`SUCESS CLEAN. Sheet: ${SheetName}.`);
}

/////////////////////////////////////////////////////////////////////SPLIT/////////////////////////////////////////////////////////////////////

function fixSplit()
{
  fixSWING_4Split();
  fixSWING_12Split();
  fixSWING_52Split();
  fixOptionsSplit()
  fixBTCSplit()
  fixTermoSplit()
  fixFundSplit()
  fixFUTPlusSplits();
  fixEXTRASplits();
};

//-------------------------------------------------------------------Swing-------------------------------------------------------------------//

function fixSWING_4Split() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SWING_4);

  var Multiplier = sheet.getRange("AB4").getValue();
  var SR = sheet.getRange("AA4").getValue();                                  //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":Y" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);

}

function fixSWING_12Split() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SWING_12);

  var Multiplier = sheet.getRange("AB4").getValue();
  var SR = sheet.getRange("AA4").getValue();                                  //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":Y" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);

}

function fixSWING_52Split() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SWING_52);

  var Multiplier = sheet.getRange("AB4").getValue();
  var SR = sheet.getRange("AA4").getValue();                                  //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":Y" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

//-------------------------------------------------------------------Opçoes-------------------------------------------------------------------//

function fixOptionsSplit() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPCOES);

  var Multiplier = sheet.getRange("Z4").getValue();
  var SR = sheet.getRange("Y4").getValue();                                   //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":B" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }

  Range.setValues(Values);

  var Range = sheet.getRange("D" + SR + ":D" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }

  Range.setValues(Values);

  var Range = sheet.getRange("F" + SR + ":F" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
//......................................................................................................................................//

  Range.setValues(Values);

  var Range = sheet.getRange("K" + SR + ":N" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("T" + SR + ":W" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

//-------------------------------------------------------------------BTC-------------------------------------------------------------------//

function fixBTCSplit() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BTC);

  var Multiplier = sheet.getRange("Z4").getValue();
  var SR = sheet.getRange("Y4").getValue();                                   //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":C" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("P" + SR + ":S" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("D" + SR + ":D" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] / Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

//-------------------------------------------------------------------Termo-------------------------------------------------------------------//

function fixTermoSplit() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TERMO);

  var Multiplier = sheet.getRange("Z4").getValue();
  var SR = sheet.getRange("Y4").getValue();                                   //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":C" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("P" + SR + ":S" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("D" + SR + ":D" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] / Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("I" + SR + ":I" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] / Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

//-------------------------------------------------------------------Future-------------------------------------------------------------------//

function fixFutureSplit() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FUTURE);

  var Multiplier = sheet.getRange("Z4").getValue();
  var SR = sheet.getRange("Y4").getValue();                                   //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":C" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("P" + SR + ":S" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("E" + SR + ":E" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("G" + SR + ":G" + LR);
  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

//-------------------------------------------------------------------Fut-------------------------------------------------------------------//

function fixFUTPlusSplits()
{
  const SheetNames = [FUTURE_1, FUTURE_2, FUTURE_3];

  SheetNames.forEach(SheetName =>
  {
    try
    {
      fixFUTPlusSplit(SheetName);
    }
    catch (error)
    {
      Logger.error(`Error saving sheet ${SheetName}:`, error);
    }
  });
}

function fixFUTPlusSplit(SheetName) 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);

  var Multiplier = sheet.getRange("Z4").getValue();
  var SR = sheet.getRange("Y4").getValue();                                   //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":C" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("P" + SR + ":S" + LR);

  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("H" + SR + ":H" + LR);

  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] / Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

//-------------------------------------------------------------------Extra-------------------------------------------------------------------//

function fixEXTRASplits()
{
  const SheetNames = [RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];

  SheetNames.forEach(SheetName =>
  {
    try
    {
      fixEXTRASplit(SheetName);
    }
    catch (error)
    {
      Logger.error(`Error saving sheet ${SheetName}:`, error);
    }
  });
}

function fixEXTRASplit(SheetName) 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);

  var Multiplier = sheet.getRange("Z4").getValue();
  var SR = sheet.getRange("Y4").getValue();                                   //startRow
  var LR = sheet.getLastRow();                                                //lastRow

  var Range = sheet.getRange("B" + SR + ":C" + LR);
  var Values = Range.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("P" + SR + ":S" + LR);

  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("E" + SR + ":F" + LR);

  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("J" + SR + ":K" + LR);

  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Range.setValues(Values);

//......................................................................................................................................//

  var Range = sheet.getRange("D" + SR + ":D" + LR);

  var Values = Range.getValues();

  for (var i = 0; i < Values.length; i++) {
    for (var j = 0; j < Values[i].length; j++) {
      if (Values[i][j] != "" && Values[i][j] != 0) { // Skip blank or zero values
      Values[i][j] = Values[i][j] / Multiplier;
      }
    }
  }
  Range.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

//-------------------------------------------------------------------Fund-------------------------------------------------------------------//

function fixFundSplit() {
  multiplyFundSplit();
  divideFundSplit();
}

function multiplyFundSplit() {
  multiplyColumn("B");
  multiplyColumn("E");
  multiplyColumn("G");
  multiplyColumn("BE");

}

function divideFundSplit() {
  divideColumn("AO");
  divideColumn("BK");
  divideColumn("BL");
}

//......................................................................................................................................//

function multiplyColumn(columnLetter) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FUND);

  var Multiplier = sheet.getRange("BT4").getValue();

  var SR = sheet.getRange("BS4").getValue();
  var LR = sheet.getLastRow();

  var Column = sheet.getRange(columnLetter + SR +":" +columnLetter + LR);
  var Values = Column.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) 
      { // Skip blank or zero values
        Values[i][j] = Values[i][j] * Multiplier;
      }
    }
  }
  Column.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

//......................................................................................................................................//

function divideColumn(columnLetter) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FUND);

  var Multiplier = sheet.getRange("BT4").getValue();

  var SR = sheet.getRange("BS4").getValue();
  var LR = sheet.getLastRow();

  var Column = sheet.getRange(columnLetter + SR +":" +columnLetter + LR);
  var Values = Column.getValues();

  Logger.log(`FIX:  ${sheet.getName()}.`);

  for (var i = 0; i < Values.length; i++) 
  {
    for (var j = 0; j < Values[i].length; j++) 
    {
      if (Values[i][j] != "" && Values[i][j] != 0) 
      { // Skip blank or zero values
        Values[i][j] = Values[i][j] / Multiplier;
      }
    }
  }
  Column.setValues(Values);
  Logger.log(`SUCESS FIX. Sheet: ${sheet.getName()}.`);
}

/////////////////////////////////////////////////////////////////////COPY / CLEAR / CLEAN / FIX /////////////////////////////////////////////////////////////////////