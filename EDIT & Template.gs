/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doEditAll()
{
  doEditDatas();
  doEditSheets();
  doIsFormula();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

//-------------------------------------------------------------------SHEETS-------------------------------------------------------------------//

function doEditSheets() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  const totalSheets = SheetNames.length;
  let Count = 0;

  Logger.log(`Starting editing of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    Logger.log(`[${Count}/${totalSheets}] (${progress}%) Editing ${SheetName}...`);

    try {
      doEditSheet(SheetName);
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} edited successfully`);
    } catch (error) {
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error editing ${SheetName}: ${error}`);
    }
  });
  Logger.log(`Edit completed: ${Count} of ${totalSheets} sheets edited successfully`);
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doEditExtras() {
  const SheetNames = [FUTURE, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  const totalSheets = SheetNames.length;
  let Count = 0;

  Logger.log(`Starting editing of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    Logger.log(`[${Count}/${totalSheets}] (${progress}%) Editing ${SheetName}...`);

    try {
      doEditSheet(SheetName);
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} edited successfully`);
    } catch (error) {
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error editing ${SheetName}: ${error}`);
    }
  });
  Logger.log(`Edit completed: ${Count} of ${totalSheets} extra sheets edited successfully`);
}

//-------------------------------------------------------------------DATA-------------------------------------------------------------------//

function doEditDatas() {
  const SheetNames = [BLC, Balanco, DRE, Resultado, FLC, Fluxo, DVA, Valor];
  const totalSheets = SheetNames.length;
  let Count = 0;

  Logger.log(`Starting editing of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    Logger.log(`[${Count}/${totalSheets}] (${progress}%) Editing ${SheetName}...`);

    try {
      doEditData(SheetName);
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} edited successfully`);
    } catch (error) {
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error editing ${SheetName}: ${error}`);
    }
  });
  Logger.log(`Edit completed: ${Count} of ${totalSheets} data sheets edited successfully`);
}

/////////////////////////////////////////////////////////////////////SHEETS TEMPLATE/////////////////////////////////////////////////////////////////////

function doEditSheet(SheetName) {
  Logger.log(`EDIT: ${SheetName}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = fetchSheetByName(SheetName); // Source sheet
  if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditSheet from sheet_sr`); return; }

  
  const sheet_co = fetchSheetByName('Config');                                    // Config sheet
  const sheet_se = fetchSheetByName('Settings');
  if (!sheet_co || !sheet_se) return;

  Utilities.sleep(2500); // 2.5 secs pause

  let Edit;

  switch (SheetName) {
//-------------------------------------------------------------------Swing-------------------------------------------------------------------//
    case SWING_4:
    case SWING_12:
    case SWING_52:
      Edit = getConfigValue(DTR);                                                 // DTR = Edit to Swing
      var Class = sheet_co.getRange(IST).getDisplayValue();                       // IST = Is Stock? 
      var C2 = sheet_sr.getRange("C2").getValue();
      if (Class == 'STOCK') {
        if (C2 > 0) {
          processEditSheet(sheet_sr, SheetName, Edit);
        }
      }
      if (Class == 'BDR' || Class == 'ETF' || Class == 'ADR') {
        if (C2 > 0) {
          processEditSheet(sheet_sr, SheetName, Edit);
        } else {
          Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
        }
      }
      break;
      
//-------------------------------------------------------------------Opções-------------------------------------------------------------------//
    case OPCOES:
      Edit = getConfigValue(DOP);                                                 // DOP = Edit to Option
      var Call = sheet_sr.getRange("C2").getValue();
      var Put = sheet_sr.getRange("E2").getValue();
      if ((Call != 0 && Put != 0) &&
          (Call != "" && Put != "")) {
        processEditSheet(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
    case BTC:
      Edit = getConfigValue(DBT);                                                 // DBT = Edit to BTC
      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processEditSheet(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
//-------------------------------------------------------------------Termo-------------------------------------------------------------------//
    case TERMO:
      Edit = getConfigValue(DTE);                                                 // DTE = Edit to Termo
      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processEditSheet(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
//-------------------------------------------------------------------Fund-------------------------------------------------------------------//
    case FUND:
      Edit = getConfigValue(DFU);                                                 // DFU = Edit to Fund
      var B2 = sheet_sr.getRange("B2").getValue();
      if (!ErrorValues.includes(B2)) {
        processEditSheet(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
//-------------------------------------------------------------------Future-------------------------------------------------------------------//
    case FUTURE:
      Edit = getConfigValue(DFT);                                                 // DFT = Edit to Future
      var C2 = sheet_sr.getRange("C2").getValue();
      var E2 = sheet_sr.getRange("E2").getValue();
      var G2 = sheet_sr.getRange("G2").getValue();
      if ((!ErrorValues.includes(C2) || !ErrorValues.includes(E2) || !ErrorValues.includes(G2))) {
        processEditSheet(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
    // -------------------- Future variants -------------------- //
    case FUTURE_1:
    case FUTURE_2:
    case FUTURE_3:
      Edit = getConfigValue(DFT);                                                 // DFT = Edit to Future
      var C2 = sheet_sr.getRange("C2").getValue();
      if (!ErrorValues.includes(C2)) {
        processEditExtra(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
//-------------------------------------------------------------------Right-------------------------------------------------------------------//
    case RIGHT_1:
    case RIGHT_2:
      Edit = getConfigValue(DRT);                                                 // DRT = Edit to Right
      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processEditExtra(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
//-------------------------------------------------------------------Receipt-------------------------------------------------------------------//
    case RECEIPT_9:
    case RECEIPT_10:
      Edit = getConfigValue(DRC);                                                 // DRC = Edit to Receipt
      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processEditExtra(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
//-------------------------------------------------------------------Warrant-------------------------------------------------------------------//
    case WARRANT_11:
    case WARRANT_12:
    case WARRANT_13:
      Edit = getConfigValue(DWT);                                                 // DWT = Edit to Warrant
      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processEditExtra(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
//-------------------------------------------------------------------Block-------------------------------------------------------------------//
    case BLOCK:
      Edit = getConfigValue(DBK);                                                 // DBK = Edit to Block
      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processEditExtra(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;
      
    default:
      Logger.log(`ERROR EDIT: ${SheetName} - Unhandled sheet type in doEditSheet`);
      break;
  }
}

/////////////////////////////////////////////////////////////////////DATA TEMPLATE/////////////////////////////////////////////////////////////////////

// sheet_sr and sheet_tr are checked  inside the blocks

function doEditData(SheetName) {
    Logger.log(`EDIT: ${SheetName}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = fetchSheetByName(SheetName);                        // Source sheet
  if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

  const sheet_co = fetchSheetByName('Config');                         // Config sheet
  const sheet_se = fetchSheetByName('Settings');
  if (!sheet_co || !sheet_se) return;

  Utilities.sleep(2500); // 2,5 secs

  let Edit, Values_sr;

  switch (SheetName) {
//-------------------------------------------------------------------BLC-------------------------------------------------------------------//
    case BLC: {
      Edit = getConfigValue(DBL);                                                     // DBL = Edit to BLC

      const sheet_tr = fetchSheetByName(BLC);
      if (!sheet_tr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

      const sheet_sr = fetchSheetByName(Balanco);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:C1").getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S, B27_S] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
          ( B2_S != 0 && B2_S != "" ) && 
          ( B27_S != 0 && B27_S != "" ) )
      {
        processEditData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Edit);

        doEditData(Balanco);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditData`);
      }
    }
      break;
//-------------------------------------------------------------------Balanço-------------------------------------------------------------------//
    case Balanco: {
      Edit = getConfigValue(DBL);                                                     // DBL = Edit to BLC

      const sheet_sr = fetchSheetByName(Balanco);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:C1").getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S, B27_S] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
          ( B2_S != 0 && B2_S != "" ) && 
          ( B27_S != 0 && B27_S != "" ) )
      {
        processEditData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Edit);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditData`);
      }
    }
      break;
//-------------------------------------------------------------------DRE-------------------------------------------------------------------//
    case DRE: {
      Edit = getConfigValue(DDE);                                                     // DDE = Edit to DRE

      const sheet_tr = fetchSheetByName(DRE);
      if (!sheet_tr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

      const sheet_sr = fetchSheetByName(Resultado);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = 
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
 
      var [C4_S, C27_S] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());
      
      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
          ( C4_S != 0 && C4_S != "" ) && 
          ( C27_S != 0 && C27_S != "" ) )
      {
        processEditData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Edit);

        doEditData(Resultado);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditData`);
      }
    }
      break;
//-------------------------------------------------------------------Resultado-------------------------------------------------------------------//
    case Resultado: {
      Edit = getConfigValue(DDE);                                                     // DDE = Edit to DRE

      const sheet_sr = fetchSheetByName(Resultado);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] =
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
 
      var [C4_S, C27_S] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());
      
      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
          ( C4_S != "" ) && 
          ( C27_S != 0 && C27_S != "" ) )
      {
        processEditData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Edit);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditData`);
      }
    }
      break;
//-------------------------------------------------------------------FLC-------------------------------------------------------------------//
    case FLC: {
      Edit = getConfigValue(DFL);                                                     // DFL = Edit to FLC

      const sheet_tr = fetchSheetByName(FLC);
      if (!sheet_tr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

      const sheet_sr = fetchSheetByName(Fluxo);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] =
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S] = ["B2"].map(r => sheet_sr.getRange(r).getDisplayValue());
      
      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
          ( B2_S != 0 && B2_S != "" ) )
      {
        processEditData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Edit);

        doEditData(Fluxo);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditData`);
      }
    }
      break;
//-------------------------------------------------------------------Fluxo-------------------------------------------------------------------//
    case Fluxo: {
      Edit = getConfigValue(DFL);                                                     // DFL = Edit to FLC

      const sheet_sr = fetchSheetByName(Fluxo);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] =
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
      
      var [B2_S] = ["B2"].map(r => sheet_sr.getRange(r).getDisplayValue());
      
      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
          ( B2_S != 0 && B2_S != "" ) )
      {
        processEditData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Edit);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditData`);
      }
    }
      break;
//-------------------------------------------------------------------DVA-------------------------------------------------------------------//
    case DVA: {
      Edit = getConfigValue(DDV);                                                     // DDV = Edit to DVA

      const sheet_tr = fetchSheetByName(DVA);
      if (!sheet_tr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_tr`); return; }

      var LR = sheet_tr.getLastRow();
      var LC = sheet_tr.getLastColumn();

      var B = sheet_tr.getRange("B1:B" + LR).getValues().flat();

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] =
          Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

      const sheet_sr = fetchSheetByName(Valor);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] =
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
      
      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());
      
      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
          ( B2_S != "" ) ) {
        processEditData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Edit);
        doEditData(Valor);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditData`);
      }
    }
      break;
//-------------------------------------------------------------------Valor-------------------------------------------------------------------//
    case Valor: {
      Edit = getConfigValue(DDV);                                                     // DDV = Edit to DVA

      const sheet_sr = fetchSheetByName(Valor);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditData from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] =
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
      
      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());
      
      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
          ( C2_S != "" ) ) {
        processEditData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditData`);
      }
    }
      break;

    default: 
      Logger.log(`ERROR EDIT: ${SheetName} - Unhandled sheet type in doEditData`);
      break;
  }
}

/////////////////////////////////////////////////////////////////////EDIT TEMPLATE/////////////////////////////////////////////////////////////////////