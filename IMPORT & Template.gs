//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////

function Import(){
  const sheet_co = fetchSheetByName('Config');                                   // Config sheet
  if (!sheet_co) {Logger.log("ERROR: 'Config' sheet not found."); return;}

  // Check if L2 has the expected colors
  if (!checkAutorizeScript()) {
    Logger.log("Import aborted: L2 does not have the correct background and font colors.");
    return;
  }

  const Source_Id = getConfigValue(SIR, 'Config');                               // SIR = Source ID
  if (!Source_Id) { Logger.log("ERROR IMPORT: Source ID is empty."); return; }

  const Option = getConfigValue(OPR, 'Config');                                  // OPR = Option

  if (Option === "AUTO")
  {
    // Check for specific sheets
    const hasSwing4 = fetchSheetByName(SWING_4) !== null;
    const hasSwing12 = fetchSheetByName(SWING_12) !== null;
    const hasSwing52 = fetchSheetByName(SWING_52) !== null;
    const hasTrade = fetchSheetByName('Trade') !== null;                         //only present in versions < 15

    if (hasSwing4 && hasSwing12 && hasSwing52)
    {
      import_Current();
    }
    else if (hasSwing12 && hasSwing52)
    {
      import_15x_to_161();                                                       // not in use anymore - to be deleted
    }
    else if (hasTrade)
    {
      import_14x_to_161();                                                       // not in use anymore - to be deleted
    }
    else
    {
      Logger.log(`No matching sheets found for AUTO mode.`);
    }
  }
  else
  {
    // Manual Option Handling
    if (Option == 1)
    {
      import_Current();
    }
    else if (Option == 2)
    {
      import_15x_to_16x();
    }
    else
    {
      Logger.log(`Invalid Option: ${Option}`);
    }
  }
}

/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function import_Current(){
  console.log('Import: import_Current');

  import_config();

  doImportProventos();
  doImportShares();

  doImportBasics();
  doImportFinancials();

  doCheckTriggers();
  update_form();

// doCleanZeros();

  console.log('Import: Finished');
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

function doImportGroup(SheetNames, importFunction, label) {
  _doGroup(SheetNames, importFunction, "Importing", "imported", label);
}

//-------------------------------------------------------------------BASICS-------------------------------------------------------------------//

function doImportBasics() {
  const SheetNames = [
    SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND,
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ];
  doImportGroup(SheetNames, doImportBasic, 'basic');
}

//-------------------------------------------------------------------FINANCIALS-------------------------------------------------------------------//

function doImportFinancials() {
  const SheetNames = [BLC, Balanco, DRE, Resultado, FLC, Fluxo, DVA, Valor];
  doImportGroup(SheetNames, doImportFinancial, 'financial');
}

//-------------------------------------------------------------------PROVENTOS-------------------------------------------------------------------//

function doImportProventos() {
  const ProvNames = ['Proventos'];
  doImportGroup(ProvNames, doImportProv, 'proventos');
}

/////////////////////////////////////////////////////////////////////Update Form/////////////////////////////////////////////////////////////////////

function update_form() {
  const sheet_co = fetchSheetByName('Config');                                        // Config sheet
  const Update_Form = getConfigValue(UFR, 'Config');                                  // UFR = Update Form

  switch (Update_Form)
  {
    case 'EDIT':
      doEditAll();
      break;
    case 'SAVE':
      doSaveAll();
      break;

    default:
      Logger.log(`Invalid update form value: ${Update_Form}`);
      break;
  }
}

/////////////////////////////////////////////////////////////////////Functions/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////Config/////////////////////////////////////////////////////////////////////

function import_config() {
  const sheet_co = fetchSheetByName('Config');
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  if (!Source_Id) { Logger.log("ERROR IMPORT: Source ID is empty."); return; }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('Config');       // Source Sheet
  {
    var Data = sheet_sr.getRange(COR).getValues();                                    // Does not use getConfigValue because it gets data from another spreadsheet
    sheet_co.getRange(COR).setValues(Data);
  }
};

/////////////////////////////////////////////////////////////////////SHARES and FF/////////////////////////////////////////////////////////////////////

function doImportShares() {
  const sheet_co = fetchSheetByName('Config');
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  if (!Source_Id) { Logger.log("ERROR IMPORT: Source ID is empty."); return; }
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('DATA');         // Source Sheet
    var L1 = sheet_sr.getRange("L1").getValue();
    var L2 = sheet_sr.getRange("L2").getValue();
  const sheet_tr = fetchSheetByName('DATA');                                          // Target Sheet
    var SheetName = sheet_tr.getName()

  Logger.log(`IMPORT: Shares and FF`);

  if (!ErrorValues.includes(L1) && !ErrorValues.includes(L2))
  {
    var Data = sheet_sr.getRange("L1:L2").getValues();
    sheet_tr.getRange("L1:L2").setValues(Data);
  }
  else
  {
    Logger.log(`ERROR IMPORT: ${SheetName} - ErrorValues on L1 or L2 on doImportShares`);
  }
Logger.log(`SUCCESS IMPORT: Shares and FF`);
}

/////////////////////////////////////////////////////////////////////Proventos/////////////////////////////////////////////////////////////////////

function doImportProv(ProvName){
  const sheet_co = fetchSheetByName('Config');
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  if (!Source_Id) { Logger.log("ERROR IMPORT: Source ID is empty."); return; }
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('Prov');         // Source Sheet
  const sheet_tr = fetchSheetByName('Prov');

  Logger.log(`IMPORT: ${ProvName}`);

  if (ProvName == 'Proventos')
  {
    var Check = sheet_sr.getRange("B3").getDisplayValue();

    if( Check == "Proventos" )  // check if error
    {
      var Data = sheet_sr.getRange(PRV).getValues();                              // PRV = Provento Range
      sheet_tr.getRange(PRV).setValues(Data);
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${ProvName} - B3 != Proventos on doImportProv`);
    }
  }
}

/////////////////////////////////////////////////////////////////////BASIC/////////////////////////////////////////////////////////////////////

const basicImportMap = {
  [SWING_4]:   { flag: ITR },
  [SWING_12]:  { flag: ITR },
  [SWING_52]:  { flag: ITR },

  [OPCOES]:    { flag: IOP },

  [BTC]:       { flag: IBT },

  [TERMO]:     { flag: ITE },

  [FUTURE]:    { flag: IFT },
  [FUTURE_1]:  { flag: IFT },
  [FUTURE_2]:  { flag: IFT },
  [FUTURE_3]:  { flag: IFT },

  [FUND]:      { flag: IFU },

  [RIGHT_1]:   { flag: IRT },
  [RIGHT_2]:   { flag: IRT },

  [RECEIPT_9]:  { flag: IRC },
  [RECEIPT_10]: { flag: IRC },

  [WARRANT_11]: { flag: IWT },
  [WARRANT_12]: { flag: IWT },
  [WARRANT_13]: { flag: IWT },

  [BLOCK]:     { flag: IBK },
};

function doImportBasic(SheetName) {
  Logger.log(`IMPORT: ${SheetName}`);

  const sheet_co  = fetchSheetByName('Config');
  const sheet_se  = fetchSheetByName('Settings');
  const Source_Id = getConfigValue(SIR, 'Config');
  if (!sheet_co || !sheet_se || !Source_Id) { Logger.log('ERROR IMPORT: Missing Config/Settings/Source ID. Aborting.'); return; }

  const cfg = basicImportMap[SheetName];
  if (!cfg) { Logger.log(`ERROR IMPORT: No import schema defined for "${SheetName}".`); return; }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);
  if (!sheet_sr) { Logger.log(`ERROR IMPORT: Source sheet ${SheetName} not found in ${Source_Id}.`); return; }

  const sheet_tr = fetchSheetByName(SheetName);
  if (!sheet_tr) { Logger.log(`ERROR IMPORT: Target sheet ${SheetName} not found. Skipping.`); return; }

  const Import = getConfigValue(cfg.flag, 'Config');
  if (Import !== "TRUE") { Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE.`); return; }

  const Check = sheet_sr.getRange("A5").getValue();
  if (Check === "") { Logger.log(`ERROR IMPORT: ${SheetName} - A5 is blank on doImportBasic.`); return; }

  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  // Copy body rows 5→LR, cols 1→LC
  const DataBody = sheet_sr.getRange(5, 1, LR - 4, LC).getValues();
  sheet_tr.getRange(5, 1, LR - 4, LC).setValues(DataBody);

  // Copy header row 1, cols 1→LC
  const DataHeader = sheet_sr.getRange(1, 1, 1, LC).getValues();
  sheet_tr.getRange(1, 1, 1, LC).setValues(DataHeader);

  Logger.log(`SUCCESS IMPORT for sheet ${SheetName}.`);
}

/////////////////////////////////////////////////////////////////////FINANCIAL////////////////////////////////////////////////////////////////////

const financialImportMap = {
  [BLC]:      { flag: IBL, checkCell: "B1", dataOffset: { colStart: 2,   colTrim: 1   } },
  [Balanco]:  { flag: IBL, checkCell: "C1", dataOffset: { colStart: 3,   colTrim: 2   } },
  [DRE]:      { flag: IDE, checkCell: "B1", dataOffset: { colStart: 2,   colTrim: 1   } },
  [Resultado]:{ flag: IDE, checkCell: "D1", dataOffset: { colStart: 4,   colTrim: 3   } },
  [FLC]:      { flag: IFL, checkCell: "B1", dataOffset: { colStart: 2,   colTrim: 1   } },
  [Fluxo]:    { flag: IFL, checkCell: "D1", dataOffset: { colStart: 4,   colTrim: 3   } },
  [DVA]:      { flag: IDV, checkCell: "B1", dataOffset: { colStart: 2,   colTrim: 1   } },
  [Valor]:    { flag: IDV, checkCell: "D1", dataOffset: { colStart: 4,   colTrim: 3   } },
};

function doImportFinancial(SheetName) {                                                      // TODO improve more functions like this one
  Logger.log(`IMPORT: ${SheetName}`);

  const sheet_co   = fetchSheetByName('Config');
  const sheet_se   = fetchSheetByName('Settings');
  const Source_Id  = getConfigValue(SIR, 'Config');
  if (!sheet_co || !sheet_se || !Source_Id) { Logger.log("ERROR IMPORT: Missing Config/Settings/Source ID. Aborting."); return; }

  const cfg = financialImportMap[SheetName];
  if (!cfg) { Logger.log(`ERROR IMPORT: No import schema defined for "${SheetName}".`); return; }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);
  if (!sheet_sr) { Logger.log(`ERROR IMPORT: "${SheetName}" not found in source ${Source_Id}.`); return; }

  const sheet_tr = fetchSheetByName(SheetName);
  if (!sheet_tr) { Logger.log(`ERROR IMPORT: Target sheet "${SheetName}" does not exist. Skipping.`); return; }

  const Import = getConfigValue(cfg.flag, 'Config');
  if (Import !== "TRUE") { Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE.`); return; }

  const Check = sheet_sr.getRange(cfg.checkCell).getValue();
  if (Check === "") { Logger.log(`ERROR IMPORT: ${SheetName} - ${cfg.checkCell} is blank on doImportFinancial.`); return; }

  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();
  const width = LC - cfg.dataOffset.colTrim;
  const data = sheet_sr
    .getRange(1, cfg.dataOffset.colStart, LR, width)
    .getValues();

  sheet_tr.getRange(1, cfg.dataOffset.colStart, LR, width)
          .setValues(data);

  Logger.log(`SUCCESS IMPORT for sheet ${SheetName}.`);
}

/////////////////////////////////////////////////////////////////////IMPORT TEMPLATE/////////////////////////////////////////////////////////////////////
