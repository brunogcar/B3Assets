//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////

function Import(){
  // Check if L6 has the expected colors
  if (!checkAutorizeScript()) {
    LogDebug("Import aborted: L6 does not have the correct background and font colors.", 'MIN');
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
    // open the source spreadsheet & sheet
    const ss_s    = SpreadsheetApp.openById(Source_Id);
    const sheet_s = ss_s.getSheetByName('Index');                                // Source Sheet
    if (!sheet_s) {
      LogDebug(`❌ ERROR IMPORT: sheet_s not found in source.`, 'MIN');
      return;
    }

    // read the trigger cell
    const triggerCell = 'K1';
    const cellValue   = sheet_s.getRange(triggerCell).getDisplayValue();

    // define your “old” vs “new” markers
    const OLD_MARKER = '';
    const NEW_MARKER = 'Inflação';

    if (cellValue === NEW_MARKER) {
      import_Current();
    } else if (cellValue === OLD_MARKER) {
      import_Upgrade();
    } else {
      LogDebug(`❌ AUTO mode: unexpected value '${cellValue}' in ${triggerCell}. Aborting import.`, 'MIN');
      return;
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
      import_Upgrade();
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
  const SheetNames = [...SheetsBasic,...SheetsExtra];
  doImportGroup(SheetNames, doImportBasic, 'basic');
}

//-------------------------------------------------------------------FINANCIALS-------------------------------------------------------------------//

function doImportFinancials() {
  const SheetNames = SheetsFinancialFull;
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

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('DATA');         // Source Sheet
    var L5 = sheet_sr.getRange("L5").getValue();
    var L6 = sheet_sr.getRange("L6").getValue();
  const sheet_tr = fetchSheetByName('DATA');                                          // Target Sheet
  if (!sheet_tr) return;

    var SheetName = sheet_tr.getName()

  LogDebug(`IMPORT: Shares and FF`, 'MIN');

  if (!ErrorValues.includes(L5) && !ErrorValues.includes(L6))
  {
    var Data = sheet_sr.getRange("L5:L6").getValues();
    sheet_tr.getRange("L5:L6").setValues(Data);
  }
  else
  {
    LogDebug(`❌ ERROR IMPORT: ${SheetName} - ErrorValues - L5=${L5} or L6=${L6}: doImportShares`, 'MIN');
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

  const cfg = basicImportMap[SheetName];
  if (!cfg) {
    LogDebug(`🚩 ERROR IMPORT: ${SheetName} - No entry in basicImportMap: doImportBasic`, 'MIN');
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
