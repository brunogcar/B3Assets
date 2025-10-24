/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doClearAll() {
  doClearProventos();
  doClearBasics();
  doClearFinancials();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////COPY/////////////////////////////////////////////////////////////////////

function doCopyBasic(SheetName) {
  const sheet = getSheet(SheetName);
  if (!sheet) return;

  var LR = sheet.getLastRow();
  var LC = sheet.getLastColumn();

  sheet.getRange(5, 1, LR - 4, LC).activate();
}

function doCopyFinancial(SheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);

  if (!sheet) {
    LogDebug(`❌ ERROR COPY: ${SheetName} Does not exist`, 'MIN');
    return;
  }

  var LR = sheet.getLastRow();
  var LC = sheet.getLastColumn();

  if (SheetName === BLC || SheetName === DRE || SheetName === FLC || SheetName === DVA) {
    sheet.getRange(1, 2, LR, LC - 1).activate();
  } else if (SheetName === Balanco) {
    sheet.getRange(1, 3, LR, LC - 2).activate();
  } else if (SheetName === Resultado || SheetName === Valor || SheetName === Fluxo) {
    sheet.getRange(1, 4, LR, LC - 3).activate();
  } else {
    LogDebug(`Unsupported sheet name: ${SheetName}`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////CLEAR/////////////////////////////////////////////////////////////////////


function doClearBasics() {
  const SheetNames = [...SheetsBasic,...SheetsExtra];

  _doGroup(SheetNames, doClearBasic, "Clearing", "cleared", "basic");
}

function doClearBasic(SheetName) {
  const sheet = getSheet(SheetName);
  if (!sheet) return;

  var LR = sheet.getLastRow();
  var LC = sheet.getLastColumn();

  LogDebug(`Clear: ${SheetName}`, 'MIN');

  sheet.getRange(5, 1, LR, LC).clear({ contentsOnly: true, skipFilteredRows: false });
  sheet.getRange(1, 1, 1, LC).clear({ contentsOnly: true, skipFilteredRows: false });

  LogDebug(`Data Cleared successfully. Sheet: ${SheetName}.`, 'MIN');
}

function doClearFinancials() {
  const SheetNames = SheetsFinancial;

  _doGroup(SheetNames, doClearFinancial, "Clearing", "cleared", "financial");
}

function doClearFinancial(SheetName) {
  const sheet = getSheet(SheetName);
  if (!sheet) return;

  LogDebug(`Clear: ${SheetName}`, 'MIN');

  var LR = sheet.getLastRow();
  var LC = sheet.getLastColumn();

  if (SheetName === BLC || SheetName === DRE || SheetName === FLC || SheetName === DVA)
  {
    sheet.getRange(1, 2, LR, LC - 1).clear({ contentsOnly: true, skipFilteredRows: false });

    LogDebug(`Data Cleared successfully. Sheet: ${SheetName}.`, 'MIN');
  }
  else if
  (SheetName === Balanco)
  {
    sheet.getRange(1, 3, LR, LC - 2).clear({ contentsOnly: true, skipFilteredRows: false });

    LogDebug(`Data Cleared successfully. Sheet: ${SheetName}.`, 'MIN');
  }
  else if
  (SheetName === Resultado || SheetName === Valor || SheetName === Fluxo)
  {
    sheet.getRange(1, 4, LR, LC - 3).clear({ contentsOnly: true, skipFilteredRows: false });

    LogDebug(`Data Cleared successfully. Sheet: ${SheetName}.`, 'MIN');
  }
  else
  {
    LogDebug(`Unsupported sheet name: ${SheetName}`, 'MIN');
  }

  if      (SheetName === BLC) {doClearFinancial(Balanco);}
  else if (SheetName === DRE) {doClearFinancial(Resultado );}
  else if (SheetName === FLC) {doClearFinancial(Fluxo);}
  else if (SheetName === DVA) {doClearFinancial(Valor);}
}


function doClearProventos() {
  const sheet = getSheet(PROV);
  if (!sheet) return;

  var LR = sheet.getLastRow();
  var LC = sheet.getLastColumn();

  sheet.getRange(PRV).clear({contentsOnly: true, skipFilteredRows: false});                             // PRV = Provento Range
};

/////////////////////////////////////////////////////////////////////ALTERNATIVE CLEAR/////////////////////////////////////////////////////////////////////

function doRecycleTrade() {
  const sheet = getSheet(TRADE);
  if (!sheet) return;

    var LR = sheet.getLastRow();
    var LC = sheet.getLastColumn();

  const AX = getConfigValue(PDT, 'Config');                                                               // PDT = Periodo de Trade
  const AX_ = sheet.getRange("A" + AX ).getValue();

//  Logger.log(AX_);

  if( AX_ !== "" )
  {
    sheet.getRange(AX,1,LR,LC).clear({contentsOnly: true, skipFilteredRows: false});
  }
};

/////////////////////////////////////////////////////////////////////CLEAN/////////////////////////////////////////////////////////////////////

function doCleanBasics() {
  const SheetNames = [...SheetsBasic,...SheetsExtra];

  _doGroup(SheetNames, doCleanBasic, "Cleaning", "cleaned", "basic");
}

function doCleanBasic(SheetName) {
  const sheet = getSheet(SheetName);
  if (!sheet) return;

  LogDebug(`CLEAN: ${SheetName}`, 'MIN');

  var LR = sheet.getLastRow();
  var LC = sheet.getLastColumn();

  sheet.getRange(5, 1, LR, LC).setValue('');
  sheet.createTextFinder("-").matchEntireCell(true).replaceAllWith("");
  sheet.createTextFinder("0").matchEntireCell(true).replaceAllWith("");
  sheet.createTextFinder("0,00").matchEntireCell(true).replaceAllWith("");
  sheet.createTextFinder("0,0000").matchEntireCell(true).replaceAllWith("");

  LogDebug(`SUCESS CLEAN. Sheet: ${SheetName}.`, 'MIN');
}

/////////////////////////////////////////////////////////////////////SPLIT/////////////////////////////////////////////////////////////////////

function fixSplit() {
  fixSWING_4Split();
  fixSWING_12Split();
  fixSWING_52Split();
  fixOptionsSplit();
  fixBTCSplit();
  fixTermoSplit();
  fixAfterSplit();
  fixFundSplit();
  fixFUTPlusSplits();
  fixEXTRASplits();
};

/**
 * Applies a multiply/divide “split” to one or more column blocks in a sheet.
 *
 * @param {string} SheetName        The sheet to process.
 * @param {string} multiplierA1     A1-notation cell containing the multiplier.
 * @param {string} startRowA1       A1-notation cell containing the start row.
 * @param {Array<{from: string, to?: string, op?: string}>} blocks
 *   - from: column letter for the first column in the block (e.g. "B")
 *   - to:   column letter for the last column (e.g. "Y").
 *            If omitted, only “from” is processed.
 *   - op:   `"mul"` (default) or `"div"`
 */
function processSplitBlocks(SheetName, multiplierA1, startRowA1, blocks) {
  const sheet = getSheet(SheetName);
  if (!sheet) return;

  const M  = sheet.getRange(multiplierA1).getValue();
  const SR = sheet.getRange(startRowA1).getValue();
  const LR = sheet.getLastRow();
  LogDebug(`FIX: ${SheetName} starting at row ${SR} with multiplier ${M}`, 'MIN');

  blocks.forEach(({from, to = from, op = 'mul'}) => {
    const RangeA1 = `${from}${SR}:${to}${LR}`;
    const Values = sheet.getRange(RangeA1).getValues();

    for (let i = 0; i < Values.length; i++) {
      for (let j = 0; j < Values[i].length; j++) {
        const v = Values[i][j];
        if (v !== '' && v !== 0) {
          Values[i][j] = (op === 'div') ? v / M : v * M;
        }
      }
    }
    sheet.getRange(RangeA1).setValues(Values);
  });

  LogDebug(`✅ SUCCESS FIX: ${SheetName}`, 'MIN');
}

// Generic split-processor has been defined separately as processSplitBlocks
//-------------------------------------------------------------------Swing-------------------------------------------------------------------//
function fixSWING_4Split() {
  processSplitBlocks(SWING_4, 'AB4', 'AA4', [
    { from: 'B', to: 'Y', op: 'mul' }
  ]);
}

function fixSWING_12Split() {
  processSplitBlocks(SWING_12, 'AB4', 'AA4', [
    { from: 'B', to: 'Y', op: 'mul' }
  ]);
}

function fixSWING_52Split() {
  processSplitBlocks(SWING_52, 'AB4', 'AA4', [
    { from: 'B', to: 'Y', op: 'mul' }
  ]);
}
//-------------------------------------------------------------------Opçoes-------------------------------------------------------------------//
function fixOptionsSplit() {
  processSplitBlocks(OPCOES, 'Z4', 'Y4', [
    { from: 'B',           op: 'mul' },
    { from: 'D',           op: 'mul' },
    { from: 'F',           op: 'mul' },
    { from: 'K', to: 'N',  op: 'mul' },
    { from: 'T', to: 'W',  op: 'mul' }
  ]);
}
//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
function fixBTCSplit() {
  processSplitBlocks(BTC, 'Z4', 'Y4', [
    { from: 'B', to: 'C',  op: 'mul' },
    { from: 'P', to: 'S',  op: 'mul' },
    { from: 'D',           op: 'div' }
  ]);
}
//-------------------------------------------------------------------Termo-------------------------------------------------------------------//
function fixTermoSplit() {
  processSplitBlocks(TERMO, 'Z4', 'Y4', [
    { from: 'B', to: 'C',  op: 'mul' },
    { from: 'P', to: 'S',  op: 'mul' },
    { from: 'D',           op: 'div' },
    { from: 'I',           op: 'div' }
  ]);
}
//-------------------------------------------------------------------After-------------------------------------------------------------------//
function fixAfterSplit() {
  processSplitBlocks(AFTER, 'Z4', 'Y4', [
    { from: 'B', to: 'C',  op: 'mul' },
    { from: 'P', to: 'S',  op: 'mul' },
    { from: 'D',           op: 'div' }
  ]);
}
//-------------------------------------------------------------------Future-------------------------------------------------------------------//

function fixFutureSplit() {
  processSplitBlocks(FUTURE, 'Z4', 'Y4', [
    { from: 'B', to: 'C',  op: 'mul' },
    { from: 'P', to: 'S',  op: 'mul' },
    { from: 'E',           op: 'mul' },
    { from: 'G',           op: 'mul' }
  ]);
}
//......................................................................................................................................//
function fixFUTPlusSplits() {
  const SheetNames = [FUTURE_1, FUTURE_2, FUTURE_3];
  SheetNames.forEach(name => {
    processSplitBlocks(name, 'Z4', 'Y4', [
      { from: 'B', to: 'C', op: 'mul' },
      { from: 'P', to: 'S', op: 'mul' },
      { from: 'H',        op: 'div' }
    ]);
  });
}
//-------------------------------------------------------------------Extra-------------------------------------------------------------------//
function fixEXTRASplits() {
  const SheetNames = SheetsExtra;
  SheetNames.forEach(name => {
    processSplitBlocks(name, 'Z4', 'Y4', [
      { from: 'B', to: 'C', op: 'mul' },
      { from: 'P', to: 'S', op: 'mul' },
      { from: 'E', to: 'F', op: 'mul' },
      { from: 'J', to: 'K', op: 'mul' },
      { from: 'D',        op: 'div' }
    ]);
  });
}
//-------------------------------------------------------------------Fund-------------------------------------------------------------------//
function fixFundSplit() {
  // Multiply columns
  processSplitBlocks(FUND, 'BT4', 'BS4', [
    { from: 'B',          op: 'mul' },
    { from: 'E',          op: 'mul' },
    { from: 'G',          op: 'mul' },
    { from: 'BE',         op: 'mul' }
  ]);
  // Divide columns
  processSplitBlocks(FUND, 'BT4', 'BS4', [
    { from: 'AO',         op: 'div' },
    { from: 'BK',         op: 'div' },
    { from: 'BL',         op: 'div' }
  ]);
}

/////////////////////////////////////////////////////////////////////COPY / CLEAR / CLEAN / FIX /////////////////////////////////////////////////////////////////////
