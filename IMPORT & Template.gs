//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////

function Import(){
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
    import_Current();
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

/////////////////////////////////////////////////////////////////////Proventos/////////////////////////////////////////////////////////////////////

function doImportProv(ProvName){
  Logger.log(`IMPORT: ${ProvName}`);

  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  if (!Source_Id) { Logger.log("ERROR IMPORT: Source ID is empty."); return; }
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('Prov');         // Source Sheet
  const sheet_tr = fetchSheetByName('Prov');
  if (!sheet_tr) return;

  if (ProvName == 'Proventos')
  {
    var Check = sheet_sr.getRange("B3").getDisplayValue();

    if( Check == "Proventos" )  // check if error
    {
      var Data = sheet_sr.getRange(PRV).getValues();                                  // PRV = Provento Range
      sheet_tr.getRange(PRV).setValues(Data);

      Logger.log(`SUCCESS IMPORT: ${ProvName}.`);
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${ProvName} - B3 != Proventos on doImportProv`);
    }
  }
}

/////////////////////////////////////////////////////////////////////Update Form/////////////////////////////////////////////////////////////////////

function update_form() {
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
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  if (!Source_Id) { Logger.log("ERROR IMPORT: Source ID is empty."); return; }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('Config');       // Source Sheet
  {
    const sheet_co = fetchSheetByName('Config');                                      // cant be deleted because of sheet_co.getRange(COR)
    if (!sheet_co) return;

    var Data = sheet_sr.getRange(COR).getValues();                                    // Does not use getConfigValue because it gets data from another spreadsheet
    sheet_co.getRange(COR).setValues(Data);
  }
};

/////////////////////////////////////////////////////////////////////SHARES and FF/////////////////////////////////////////////////////////////////////

function doImportShares() {
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  if (!Source_Id) { Logger.log("ERROR IMPORT: Source ID is empty."); return; }
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('DATA');         // Source Sheet
    var L1 = sheet_sr.getRange("L1").getValue();
    var L2 = sheet_sr.getRange("L2").getValue();
  const sheet_tr = fetchSheetByName('DATA');                                          // Target Sheet
  if (!sheet_tr) return;

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

/////////////////////////////////////////////////////////////////////BASIC/////////////////////////////////////////////////////////////////////

const basicImportMap = {
  [SWING_4]:    { flag: ITR },              // ITR = Import to SWING
  [SWING_12]:   { flag: ITR },              // ITR = Import to SWING
  [SWING_52]:   { flag: ITR },              // ITR = Import to SWING

  [OPCOES]:     { flag: IOP },              // IOP = Import to OPCOES (Options)

  [BTC]:        { flag: IBT },              // IBT = Import to BTC

  [TERMO]:      { flag: ITE },              // ITE = Import to TERMO

  [FUTURE]:     { flag: IFT },              // IFT = Import to FUTURE
  [FUTURE_1]:   { flag: IFT },              // IFT = Import to FUTURE
  [FUTURE_2]:   { flag: IFT },              // IFT = Import to FUTURE
  [FUTURE_3]:   { flag: IFT },              // IFT = Import to FUTURE

  [FUND]:       { flag: IFU },              // IFU = Import to FUND

  [RIGHT_1]:    { flag: IRT },              // IRT = Import to RIGHT
  [RIGHT_2]:    { flag: IRT },              // IRT = Import to RIGHT

  [RECEIPT_9]:  { flag: IRC },              // IRC = Import to RECEIPT
  [RECEIPT_10]: { flag: IRC },              // IRC = Import to RECEIPT

  [WARRANT_11]: { flag: IWT },              // IWT = Import to WARRANT
  [WARRANT_12]: { flag: IWT },              // IWT = Import to WARRANT
  [WARRANT_13]: { flag: IWT },              // IWT = Import to WARRANT

  [BLOCK]:      { flag: IBK },              // IBK = Import to BLOCK
};

function doImportBasic(SheetName) {
  Logger.log(`IMPORT: ${SheetName}`);

  const Source_Id = getConfigValue(SIR, 'Config');
  if (!Source_Id) { Logger.log('ERROR IMPORT: Source ID. Aborting.'); return; }

  const cfg = basicImportMap[SheetName];
  if (!cfg) { Logger.log(`ERROR IMPORT: No import schema defined for "${SheetName}".`); return; }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);
  if (!sheet_sr) { Logger.log(`ERROR IMPORT: Source sheet ${SheetName} not found in ${Source_Id}.`); return; }

  const sheet_tr = fetchSheetByName(SheetName);
  if (!sheet_tr) return;

  const Import = getConfigValue(cfg.flag, 'Config');
  if (Import !== "TRUE") { Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE.`); return; }

  const Check = sheet_sr.getRange("A5").getValue();
  if (Check === "") { Logger.log(`ERROR IMPORT: ${SheetName} - A5 is blank on doImportBasic.`); return; }

  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  // Copy body rows 5→LR, cols 1→LC
  const Data_Body = sheet_sr.getRange(5, 1, LR - 4, LC).getValues();
  sheet_tr.getRange(5, 1, LR - 4, LC).setValues(Data_Body);

  // Copy header row 1, cols 1→LC
  const Data_Header = sheet_sr.getRange(1, 1, 1, LC).getValues();
  sheet_tr.getRange(1, 1, 1, LC).setValues(Data_Header);

  Logger.log(`SUCCESS IMPORT for sheet ${SheetName}.`);
}

/////////////////////////////////////////////////////////////////////FINANCIAL////////////////////////////////////////////////////////////////////

const financialImportMap = {
  [BLC]:       { flag: IBL, checkCell: "B1", dataOffset: { colStart: 2, colTrim: 1 } },               // IBL = Import to BALANCO
  [Balanco]:   { flag: IBL, checkCell: "C1", dataOffset: { colStart: 3, colTrim: 2 } },               // IBL = Import to BALANCO

  [DRE]:       { flag: IDE, checkCell: "B1", dataOffset: { colStart: 2, colTrim: 1 } },               // IDE = Import to DRE
  [Resultado]: { flag: IDE, checkCell: "D1", dataOffset: { colStart: 4, colTrim: 3 } },               // IDE = Import to DRE

  [FLC]:       { flag: IFL, checkCell: "B1", dataOffset: { colStart: 2, colTrim: 1 } },               // IFL = Import to FLC (Cash Flow)
  [Fluxo]:     { flag: IFL, checkCell: "D1", dataOffset: { colStart: 4, colTrim: 3 } },               // IFL = Import to FLC (Cash Flow)

  [DVA]:       { flag: IDV, checkCell: "B1", dataOffset: { colStart: 2, colTrim: 1 } },               // IDV = Import to DVA
  [Valor]:     { flag: IDV, checkCell: "D1", dataOffset: { colStart: 4, colTrim: 3 } },               // IDV = Import to DVA
};


function doImportFinancial(SheetName) {                                                      // TODO improve more functions like this one
  Logger.log(`IMPORT: ${SheetName}`);

  const Source_Id  = getConfigValue(SIR, 'Config');
  if (!Source_Id) { Logger.log("ERROR IMPORT: Missing Config/Settings/Source ID. Aborting."); return; }

  const cfg = financialImportMap[SheetName];
  if (!cfg) { Logger.log(`ERROR IMPORT: No import schema defined for "${SheetName}".`); return; }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);
  if (!sheet_sr) { Logger.log(`ERROR IMPORT: "${SheetName}" not found in source ${Source_Id}.`); return; }

  const sheet_tr = fetchSheetByName(SheetName);
  if (!sheet_tr) return;

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
