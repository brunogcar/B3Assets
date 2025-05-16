//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////

function Import(){
  // Check if L2 has the expected colors
  if (!checkAutorizeScript()) {
    LogDebug("Import aborted: L2 does not have the correct background and font colors.", 'MIN');
    return;
  }

  const Source_Id = getConfigValue(SIR, 'Config');                               // SIR = Source ID
  if (!Source_Id) {
    LogDebug(`❌ ERROR IMPORT: Source ID is empty.`, 'MIN');
    return;
  }

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
      LogDebug(`❌ Invalid Import  Option: ${Option}`, 'MIN');
    }
  }
}

/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function import_Current(){
  LogDebug(`IMPORT: import_Current`, 'MIN');

  import_config();

  doImportProventos();
  doImportShares();

  doImportBasics();
  doImportFinancials();

  doCheckTriggers();
  update_form();

// doCleanZeros();

  LogDebug(`✅ SUCCESS IMPORT: Finished`, 'MIN');
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
    BLOCK, AFTER
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
  LogDebug(`IMPORT: ${ProvName}`, 'MIN');

  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  if (!Source_Id) {
    LogDebug(`❌ ERROR IMPORT: Source ID is empty.`, 'MIN');
    return;
    }

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

      LogDebug(`✅ SUCCESS IMPORT: ${ProvName}.`, 'MIN');
    }
    else
    {
      LogDebug(`❌ ERROR IMPORT: ${ProvName} - B3 != Proventos: doImportProv`, 'MIN');
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
      LogDebug(`Invalid update form value: ${Update_Form}`, 'MIN');
      break;
  }
}

/////////////////////////////////////////////////////////////////////Functions/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////Config/////////////////////////////////////////////////////////////////////

function import_config() {
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  if (!Source_Id) {
    LogDebug(`❌ ERROR IMPORT: Source ID is empty.`, 'MIN');
    return;
  }

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
  if (!Source_Id) {
    LogDebug(`❌ ERROR IMPORT: Source ID is empty.`, 'MIN');
    return;
    }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('DATA');         // Source Sheet
    var L1 = sheet_sr.getRange("L1").getValue();
    var L2 = sheet_sr.getRange("L2").getValue();
  const sheet_tr = fetchSheetByName('DATA');                                          // Target Sheet
  if (!sheet_tr) return;

    var SheetName = sheet_tr.getName()

  LogDebug(`IMPORT: Shares and FF`, 'MIN');

  if (!ErrorValues.includes(L1) && !ErrorValues.includes(L2))
  {
    var Data = sheet_sr.getRange("L1:L2").getValues();
    sheet_tr.getRange("L1:L2").setValues(Data);
  }
  else
  {
    LogDebug(`❌ ERROR IMPORT: ${SheetName} - ErrorValues - L1=${L1} or L2=${L2}: doImportShares`, 'MIN');
  }
  LogDebug(`✅ SUCCESS IMPORT: Shares and FF`, 'MIN');
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

  [AFTER]:      { flag: IAF },              // IBK = Import to BLOCK
};

function doImportBasic(SheetName) {
  LogDebug(`IMPORT: ${SheetName}`, 'MIN');

  const Source_Id = getConfigValue(SIR, 'Config');
  if (!Source_Id) {
    LogDebug('❌ ERROR IMPORT: Source ID. Aborting.', 'MIN');
    return;
  }

  const cfg = basicImportMap[SheetName];
  if (!cfg) {
    LogDebug(`🚩 ERROR IMPORT: ${SheetName} - No entry in basicImportMap: doImportShares`, 'MIN');
    return;
  }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);
  if (!sheet_sr) {
    LogDebug(`❌ ERROR IMPORT: Source sheet ${SheetName} not found in ${Source_Id}.`, 'MIN');
    return;
  }

  const sheet_tr = fetchSheetByName(SheetName);
  if (!sheet_tr) {
    LogDebug(`❌ ERROR IMPORT: Target sheet ${SheetName} not found.`, 'MIN');
    return;
  }

  const Import = getConfigValue(cfg.flag, 'Config');
  if (Import !== "TRUE") {
    LogDebug(`❌ ERROR IMPORT: ${SheetName} - IMPORT is set to FALSE.`, 'MIN');
    return;
  }

  const Check = sheet_sr.getRange("A5").getValue();
  if (Check === "") {
    LogDebug(`❌ ERROR IMPORT: ${SheetName} - A5 ${Check} is blank: doImportBasic.`, 'MIN');
    return;
  }

  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  // Copy body rows 5→LR, cols 1→LC
  const Data_Body = sheet_sr.getRange(5, 1, LR - 4, LC).getValues();
  sheet_tr.getRange(5, 1, LR - 4, LC).setValues(Data_Body);

  // Copy header row 1, cols 1→LC
  const Data_Header = sheet_sr.getRange(1, 1, 1, LC).getValues();
  sheet_tr.getRange(1, 1, 1, LC).setValues(Data_Header);

  LogDebug(`✅ SUCCESS IMPORT: ${SheetName}.`, 'MIN');
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


function doImportFinancial(SheetName) {
  LogDebug(`IMPORT: ${SheetName}`, 'MIN');

  const Source_Id  = getConfigValue(SIR, 'Config');
  if (!Source_Id) {
    LogDebug(`❌ ERROR IMPORT: Missing Config/Settings/Source ID. Aborting.`, 'MIN');
    return;
  }

  const cfg = financialImportMap[SheetName];
  if (!cfg) {
    LogDebug(`🚩 ERROR IMPORT: ${SheetName} - No entry in financialImportMap: doImportBasic`, 'MIN');
    return;
  }

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);
  if (!sheet_sr) {
    LogDebug(`❌ ERROR IMPORT: "${SheetName}" not found in source ${Source_Id}.`, 'MIN');
    return;
  }

  const sheet_tr = fetchSheetByName(SheetName);
  if (!sheet_tr) {
    LogDebug(`❌ ERROR IMPORT: Target sheet "${SheetName}" not found.`, 'MIN');
    return;
  }

  const Import = getConfigValue(cfg.flag, 'Config');
  if (Import !== "TRUE") {
    LogDebug(`❌ ERROR IMPORT: ${SheetName} - IMPORT is set to FALSE.`, 'MIN');
    return;
  }

  const Check = sheet_sr.getRange(cfg.checkCell).getValue();
  if (Check === "") {
    LogDebug(`❌ ERROR IMPORT: ${SheetName} - ${cfg.checkCell} is blank: doImportFinancial.`, 'MIN');
    return;
  }

  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();
  const width = LC - cfg.dataOffset.colTrim;
  const data = sheet_sr
    .getRange(1, cfg.dataOffset.colStart, LR, width)
    .getValues();

  sheet_tr.getRange(1, cfg.dataOffset.colStart, LR, width)
          .setValues(data);

  LogDebug(`✅ SUCCESS IMPORT: ${SheetName}.`, 'MIN');
}

/////////////////////////////////////////////////////////////////////IMPORT TEMPLATE/////////////////////////////////////////////////////////////////////
