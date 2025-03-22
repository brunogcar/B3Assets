/////////////////////////////////////////////////////////////////////CHECK/////////////////////////////////////////////////////////////////////

function doCheckDATAS() 
{
  const SheetNames = [
    SWING_4, SWING_12, SWING_52,
    PROV, OPCOES, BTC, TERMO, FUND,
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doCheckDATA(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      Logger.error(`Error checking DATA for sheet ${SheetName}: ${error}`);
    }
  });
}

/////////////////////////////////////////////////////////////////////DO CHECK TEMPLATE/////////////////////////////////////////////////////////////////////

function doCheckDATA(SheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_s = ss.getSheetByName(SheetName); // Source sheet
  const sheet_d = ss.getSheetByName('DATA');    // DATA sheet
  const sheet_p = ss.getSheetByName(PROV);      // PROV sheet
  const sheet_o = ss.getSheetByName('OPT');     // OPT sheet
  const sheet_b = ss.getSheetByName(Balanco);   // Balanco sheet
  const sheet_r = ss.getSheetByName(Resultado); // Resultado sheet
  const sheet_f = ss.getSheetByName(Fluxo);     // Fluxo sheet
  const sheet_v = ss.getSheetByName(Valor);     // Valor sheet

  let Check;

  Logger.log(`CHECK Sheet: ${SheetName}`);

  switch (SheetName) {
//-------------------------------------------------------------------PROV-------------------------------------------------------------------//
    case PROV:
      Check = sheet_p.getRange("B3").getValue();
      break;

//-------------------------------------------------------------------OPCOES-------------------------------------------------------------------//
    case OPCOES:
      Check = sheet_o.getRange("B2").getValue();
      if (Check === '') {
        sheet_o.hideSheet();
        Logger.log(`HIDDEN:`, `OPT`);
      } else if (sheet_o.isSheetHidden()) {
        sheet_o.showSheet();
        Logger.log(`DISPLAYED:  ${SheetName}`);
      }
      break;
//-------------------------------------------------------------------SWING-------------------------------------------------------------------//
    case SWING_4:
    case SWING_12:
    case SWING_52:
      const sheet_c = ss.getSheetByName('Config');
      const Class = sheet_c.getRange(IST).getDisplayValue(); // IST = Is Stock?
      Check = Class === 'STOCK' ? sheet_d.getRange('B16').getValue() : 'TRUE';
      break;
//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
    case BTC:
      Check = sheet_d.getRange("B3").getValue();
      break;
//-------------------------------------------------------------------TERMO-------------------------------------------------------------------//
    case TERMO:
      Check = sheet_d.getRange("B24").getValue();
      break;
//-------------------------------------------------------------------FUTURE-------------------------------------------------------------------//
    case FUTURE:
      const futureChecks = ["B32", "B33", "B34"];
      for (let i = 0; i < futureChecks.length; i++) {
        Check = sheet_d.getRange(futureChecks[i]).getValue();
        if (!ErrorValues.includes(Check)) break;
      }
      break;

    case FUTURE_1:
      Check = sheet_d.getRange("B32").getValue();
      break;

    case FUTURE_2:
      Check = sheet_d.getRange("B33").getValue();
      break;

    case FUTURE_3:
      Check = sheet_d.getRange("B34").getValue();
      break;
//-------------------------------------------------------------------RIGHT-------------------------------------------------------------------//
    case RIGHT_1:
      Check = sheet_d.getRange("C38").getValue();
      break;

    case RIGHT_2:
      Check = sheet_d.getRange("C39").getValue();
      break;
//-------------------------------------------------------------------RECEIPT-------------------------------------------------------------------//
    case RECEIPT_9:
      Check = sheet_d.getRange("C44").getValue();
      break;

    case RECEIPT_10:
      Check = sheet_d.getRange("C45").getValue();
      break;
//-------------------------------------------------------------------WARRANT-------------------------------------------------------------------//
    case WARRANT_11:
      Check = sheet_d.getRange("C50").getValue();
      break;

    case WARRANT_12:
      Check = sheet_d.getRange("C51").getValue();
      break;

    case WARRANT_13:
      Check = sheet_d.getRange("C52").getValue();
      break;
//-------------------------------------------------------------------BLOCK-------------------------------------------------------------------//
    case BLOCK:
      const blockChecks = ["C56", "C57", "C58"];
      for (let i = 0; i < blockChecks.length; i++) {
        Check = sheet_d.getRange(blockChecks[i]).getValue();
        if (!ErrorValues.includes(Check)) break;
      }
      break;
//-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//
    case BLC:
      Check = sheet_b.getRange("B1").getValue();
      break;

    case DRE:
      Check = sheet_r.getRange("C1").getValue();
      break;

    case FLC:
      Check = sheet_f.getRange("C1").getValue();
      break;

    case DVA:
      Check = sheet_v.getRange("C1").getValue();
      break;
//-------------------------------------------------------------------DEFAULT-------------------------------------------------------------------//
    default:
      Check = 'FALSE';
      Logger.log(`Sheet Name ${SheetName} not recognized.`);
      break;
  }

  return processCheckDATA(sheet_s, SheetName, Check);
}

/////////////////////////////////////////////////////////////////////DO CHECK Process/////////////////////////////////////////////////////////////////////

function processCheckDATA(sheet_s, SheetName, Check) {
  const fixedSheets = [BLC, DRE, FLC, DVA];

  if (ErrorValues.includes(Check)) {
    if (fixedSheets.includes(SheetName)) {
      Logger.log(`DATA Check: FALSE for ${SheetName}`);
      return "FALSE";
    }
    if (!sheet_s.isSheetHidden()) {
      sheet_s.hideSheet();
      Logger.log(`Sheet ${SheetName} HIDDEN`);
    }
    Logger.log(`DATA Check: FALSE for ${SheetName}`);
    return "FALSE";
  }

  if (sheet_s.isSheetHidden()) {
    sheet_s.showSheet();
    Logger.log(`Sheet ${SheetName} DISPLAYED`);
  }

  Logger.log(`DATA Check: TRUE for ${SheetName}`);
  return "TRUE";
}

/////////////////////////////////////////////////////////////////////TRIM TEMPLATE/////////////////////////////////////////////////////////////////////

function doTrim() {
  const SheetNames = [
    SWING_4, SWING_12, SWING_52
  ];

  SheetNames.forEach(SheetName => 
  {
    try { doTrimSheet(SheetName); } 
    catch (error) { Logger.error(`Error saving sheet ${SheetName}: ${error}`); }
  });
}

function doTrimSheet(SheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_s = ss.getSheetByName(SheetName); // Target

  Logger.log(`TRIM: ${SheetName}`);

  if (!sheet_s) { 
    Logger.error(`Sheet ${SheetName} not found.`); 
    return; 
  }

  var LR = sheet_s.getLastRow();
  var LC = sheet_s.getLastColumn();

  switch (SheetName) {
    case SWING_4:
      if (LR > 126) {
        sheet_s.getRange(127, 1, LR - 126, LC).clearContent();
        Logger.log(`SUCCESS TRIM. Sheet: ${SheetName}.`);
        Logger.log(`Cleared data below row 126 in ${SheetName}.`);
      }
      break;

    case SWING_12:
      if (LR > 366) {
        sheet_s.getRange(367, 1, LR - 366, LC).clearContent();
        Logger.log(`SUCCESS TRIM. Sheet: ${SheetName}.`);
        Logger.log(`Cleared data below row 366 in ${SheetName}.`);
      }
      break;

    case SWING_52:
      Logger.log(`NOTHING TO TRIM. Sheet: ${SheetName}.`);
      break;

    default:
      Logger.log(`No specific logic defined  to Trim for ${SheetName}.`);
  }
}

/////////////////////////////////////////////////////////////////////Hide and Show Sheets/////////////////////////////////////////////////////////////////////

function doDisableSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_c = ss.getSheetByName('Config');
  const sheets = ss.getSheets();

  var Class = sheet_c.getRange(IST).getDisplayValue();                                                                 // IST = Is Stock?
  let SheetNames = [];

  switch (Class) {
    case 'STOCK':
      SheetNames = ['DATA', 'Prov_', 'FIBO', 'Cotações', 'UPDATE', 'Balanço', 'Balanço Ativo', 'Balanço Passivo', 'Resultado', 'Demonstração', 'Fluxo', 'Fluxo de Caixa', 'Valor', 'Demonstração do Valor Adicionado'];

      sheets.forEach(sheet => {
        if (!sheet.isSheetHidden() && SheetNames.includes(sheet.getName())) {
          sheet.hideSheet();
          Logger.log(`Sheet: ${sheet.getName()} HIDDEN`);
        }
      });
      break;

    case 'ADR':
      SheetNames = new Set(['Config', 'Settings', 'Index', 'Preço', 'FIBO', SWING_4, SWING_12, SWING_52, 'Cotações']);

      for (let i = sheets.length - 1; i >= 0; i--) {                                                                    // Reverse iteration to avoid index shifting
        const sheet = sheets[i];
        if (!SheetNames.has(sheet.getName())) {                                                                         // Delete all but SheetNames
          Logger.log(`Deleting sheet: ${sheet.getName()}`);
          ss.deleteSheet(sheet);
        }
      }
      break;

    case 'BDR':
    case 'ETF':
      SheetNames = new Set(['Config', 'Settings', 'Index', 'Prov', 'Prov_', 'Preço', 'FIBO', SWING_4, SWING_12, SWING_52, 'Cotações', 'DATA', 'OPT', 'Opções', 'BTC', 'Termo']);

      for (let i = sheets.length - 1; i >= 0; i--) {                                                                    // Reverse iteration to avoid index shifting
        const sheet = sheets[i];
        if (!SheetNames.has(sheet.getName())) {                                                                         // Delete all but SheetNames
          Logger.log(`Deleting sheet: ${sheet.getName()}`);
          ss.deleteSheet(sheet);
        }
      }
      break;
      
    default:
      Logger.log(`Class ${Class} not recognized. No sheets modified.`);
  }
  hideConfig();
}

/////////////////////////////////////////////////////////////////////HIDE CONFIG/////////////////////////////////////////////////////////////////////

function hideConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_s = ss.getSheetByName(`Settings`);                        // Source sheet
  const sheet_c = ss.getSheetByName(`Config`);                          // Config sheet

  var Hide_Config = sheet_c.getRange(HCR).getDisplayValue();            // HCR = Hide Config Range

  if (Hide_Config == "TRUE") {
    if (sheet_s && !sheet_s.isSheetHidden()) {
      sheet_s.hideSheet();
      Logger.log(`HIDDEN: ${sheet_s.getName()}`);
    }
    if (sheet_c && !sheet_c.isSheetHidden()) {
      sheet_c.hideSheet();
      Logger.log(`HIDDEN: ${sheet_c.getName()}`);
    }
  }
};

/////////////////////////////////////////////////////////////////////SAVE FUNCTIONS/////////////////////////////////////////////////////////////////////