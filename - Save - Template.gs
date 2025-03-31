function doSaveSheet(SheetName) {
  Logger.log(`SAVE: ${SheetName}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = fetchSheetByName(SheetName); // Source sheet
  if (!sheet_sr) {
    Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveSheet from sheet_sr`);
    return;
  }
  
  const sheet_co = fetchSheetByName('Config');     // Config sheet
  const sheet_se = fetchSheetByName('Settings');   // Settings sheet
  if (!sheet_co || !sheet_se) return;

  Utilities.sleep(2500); // 2.5 secs pause

  let Save, Edit;

  switch (SheetName) {
//-------------------------------------------------------------------Swing-------------------------------------------------------------------//
    case SWING_4:
    case SWING_12:
    case SWING_52:
      Save = getConfigValue(STR)                                                     // STR = Save to Swing
      Edit = getConfigValue(DTR)                                                     // DTR = Edit to Swing

      var Class = sheet_co.getRange(IST).getDisplayValue();                          // IST = Is Stock? 

      var B2 = sheet_sr.getRange("B2").getValue();
      var C2 = sheet_sr.getRange("C2").getValue();
      
      if (Class == 'STOCK') {
        if (B2 != 0 && C2 > 0) {
          processSaveSheet(sheet_sr, SheetName, Save, Edit);
          doTrimSheet(SheetName);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
        }
      }
      if (Class == 'BDR' || Class == 'ETF' || Class == 'ADR') {
        if (C2 > 0) {
          processSaveSheet(sheet_sr, SheetName, Save, Edit);
          doTrimSheet(SheetName);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
        }
      }
      break;
//-------------------------------------------------------------------Opções-------------------------------------------------------------------//
    case OPCOES:
      Save = getConfigValue(SOP)                                                     // SOP = Save to Option
      Edit = getConfigValue(DOP)                                                     // DOP = Edit to Option

      var [Call, Call_, Call_PM, Call_PM_, Put, Put_, Put_PM, Put_PM_, Diff, Diff_2] = 
        ["C2", "C3", "D2", "D3", "E2", "E3", "F2", "F3", "K3", "N3"].map(r => sheet_sr.getRange(r).getValue());

      if ((Call != 0 && Put != 0) &&
          (Call != '' && Put != '') &&
          (Call_PM != 0 || Put_PM != 0) &&
          (Call_PM != '' || Put_PM != '') &&
          (Diff != 0 || Diff_2 != 0)) {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
    case BTC:
      Save = getConfigValue(SBT)                                                     // SBT = Save to BTC
      Edit = getConfigValue(DBT)                                                     // DBT = Edit to BTC

      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
//-------------------------------------------------------------------Termo-------------------------------------------------------------------//
    case TERMO:
      Save = getConfigValue(STE)                                                     // STE = Save to Termo
      Edit = getConfigValue(DTE)                                                     // DTE = Edit to Termo

      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
//-------------------------------------------------------------------Fund-------------------------------------------------------------------//
    case FUND:
      Save = getConfigValue(SFU)                                                     // SFU = Save to Fund
      Edit = getConfigValue(DFU)                                                     // DFU = Edit to Fund

      var B2 = sheet_sr.getRange("B2").getValue();
      if (!ErrorValues.includes(B2)) {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
//-------------------------------------------------------------------FUTURE-------------------------------------------------------------------//
    case FUTURE:
      Save = getConfigValue(SFT)                                                     // SFT = Save to Future
      Edit = getConfigValue(DFT)                                                     // DFT = Edit to Future

      var C2 = sheet_sr.getRange("C2").getValue();
      var E2 = sheet_sr.getRange("E2").getValue();
      var G2 = sheet_sr.getRange("G2").getValue();
      // You mentioned possibly && instead of ||. Keep as needed.
      if ((!ErrorValues.includes(C2) || !ErrorValues.includes(E2) || !ErrorValues.includes(G2))) {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
    // ----- Future variants: FUTURE_1, FUTURE_2, FUTURE_3 -----
    case FUTURE_1:
    case FUTURE_2:
    case FUTURE_3:
      Save = getConfigValue(SFT)                                                     // SFT = Save to Future
      Edit = getConfigValue(DFT)                                                     // DFT = Edit to Future

      var C2 = sheet_sr.getRange("C2").getValue();
      if (!ErrorValues.includes(C2)) {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
//-------------------------------------------------------------------Right-------------------------------------------------------------------//
    case RIGHT_1:
    case RIGHT_2:
      Save = getConfigValue(SRT)                                                     // SRT = Save to Right
      Edit = getConfigValue(DRT)                                                     // DRT = Edit to Right

      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
//-------------------------------------------------------------------Receipt-------------------------------------------------------------------//
    case RECEIPT_9:
    case RECEIPT_10:
      Save = getConfigValue(SRC)                                                     // SRC = Save to Receipt
      Edit = getConfigValue(DRC)                                                     // DRC = Edit to Receipt

      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
//-------------------------------------------------------------------Warrant-------------------------------------------------------------------//
    case WARRANT_11:
    case WARRANT_12:
    case WARRANT_13:
      Save = getConfigValue(SWT)                                                     // SWT = Save to Warrant
      Edit = getConfigValue(DWT)                                                     // DWT = Edit to Warrant

      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
//-------------------------------------------------------------------Block-------------------------------------------------------------------//
    case BLOCK:
      Save = getConfigValue(SBK)                                                     // SBK = Save to Block
      Edit = getConfigValue(DBK)                                                     // DBK = Edit to Block

      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveSheet`);
      }
      break;
      
    default:
      Logger.log(`ERROR SAVE: ${SheetName} - Unhandled sheet type in doSaveSheet`);
      break;
  }
}

/////////////////////////////////////////////////////////////////////DATA TEMPLATE/////////////////////////////////////////////////////////////////////
// sheet_sr and sheet_tr are checked  inside the blocks

function doSaveData(SheetName) {
  Logger.log(`SAVE: ${SheetName}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');     // Config sheet
  const sheet_se = fetchSheetByName('Settings');
  const sheet_up = fetchSheetByName('UPDATE');
  if (!sheet_co || !sheet_se || !sheet_up) return;

  let Save, Edit;
  let sheet_tr, sheet_sr;

  switch (SheetName) {
    // -------------------------------------------------------------------BLC -------------------------------------------------------------------//
    case BLC:
      Save = getConfigValue(SBL)                                                     // SBL = Save to BLC
      Edit = getConfigValue(DBL)                                                     // DBL = Edit to BLC

      sheet_tr = fetchSheetByName(BLC);
      if (!sheet_tr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

      sheet_sr = fetchSheetByName(Balanco);
      if (!sheet_sr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_sr`); return; }

      var Values_sr = sheet_sr.getRange("B1:C1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S, B27_S] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      var CHECK1 = sheet_up.getRange("K3").getValue();
      var CHECK2 = sheet_up.getRange("K4").getValue();

      if (((CHECK1 >= 90 && CHECK1 <= 92) || (CHECK1 == 0 || CHECK1 > 40000)) && 
          ((CHECK2 >= 90 && CHECK2 <= 92) || (CHECK2 == 0 || CHECK1 > 40000))) {
        if ((New_S.valueOf() != "-" && New_S.valueOf() != "") &&
            (B2_S != 0 && B2_S != "") &&
            (B27_S != 0 && B27_S != "")) {
          processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit);
          doSaveData(Balanco);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveData`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData`);
      }
      break;
    // -------------------------------------------------------------------Balanço -------------------------------------------------------------------//
    case Balanco:
      Save = getConfigValue(SBL)                                                     // SBL = Save to BLC
      Edit = getConfigValue(DBL)                                                     // DBL = Edit to BLC

      sheet_sr = fetchSheetByName(Balanco);
      if (!sheet_sr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_sr`); return; }

      var Values_sr = sheet_sr.getRange("B1:C1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C4_S, C27_S] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if ((New_S.valueOf() != "-" && New_S.valueOf() != "") &&
          (C4_S != 0 && C4_S != "") &&
          (C27_S != 0 && C27_S != "")) {
        processSaveData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveData`);
      }
      break;
    // -------------------------------------------------------------------DRE -------------------------------------------------------------------//
    case DRE:
      Save = getConfigValue(SDE)                                                     // SDE = Save to DRE
      Edit = getConfigValue(DDE)                                                     // DDE = Edit to DRE

      sheet_tr = fetchSheetByName(DRE);
      if (!sheet_tr) {Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

      sheet_sr = fetchSheetByName(Resultado);
      if (!sheet_sr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_sr`); return; }

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C4_S, C27_S] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());
      var CHECK = sheet_up.getRange("K5").getValue();

      if (((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000))) {
        if ((New_S.valueOf() != "-" && New_S.valueOf() != "") &&
            (C4_S != 0 && C4_S != "") &&
            (C27_S != 0 && C27_S != "")) {
          processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit);
          doSaveData(Resultado);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveData`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData`);
      }
      break;
    // -------------------------------------------------------------------Resultado -------------------------------------------------------------------//
    case Resultado:
      Save = getConfigValue(SDE)                                                     // SDE = Save to DRE
      Edit = getConfigValue(DDE)                                                     // DDE = Edit to DRE

      sheet_sr = fetchSheetByName(Resultado);
      if (!sheet_sr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_sr`); return; }

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = 
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C4_S, C27_S] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());
      if (sheet_sr) {
        if ((New_S.valueOf() != "-" && New_S.valueOf() != "") &&
            (C4_S != "") &&
            (C27_S != 0 && C27_S != "")) {
          processSaveData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Save, Edit);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveData`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData`);
      }
      break;
    // -------------------------------------------------------------------FLC -------------------------------------------------------------------//
    case FLC:
      Save = getConfigValue(SFL)                                                     // SFL = Save to FLC
      Edit = getConfigValue(DFL)                                                     // DFL = Edit to FLC

      sheet_tr = fetchSheetByName(FLC);
      if (!sheet_tr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = 
          Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

      sheet_sr = fetchSheetByName(Fluxo);
      if (!sheet_sr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_sr`); return; }

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = 
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());
      var CHECK = sheet_up.getRange("K6").getValue();

      if ((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000)) {
        if ((New_S.valueOf() != "-" && New_S.valueOf() != "") &&
            (C2_S != 0 && C2_S !== "")) {
          processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit);
          doSaveData(Fluxo);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveData`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData`);
      }
      break;
    // -------------------------------------------------------------------Fluxo -------------------------------------------------------------------//
    case Fluxo:
      Save = getConfigValue(SFL)                                                     // SFL = Save to FLC
      Edit = getConfigValue(DFL)                                                     // DFL = Edit to FLC

      sheet_sr = fetchSheetByName(Fluxo);
      if (!sheet_sr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_sr`); return; }

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] =
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if ((New_S.valueOf() != "-" && New_S.valueOf() != "") &&
          (C2_S != 0 && C2_S !== "")) {
        processSaveData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveData`);
      }
      break;
    // -------------------------------------------------------------------DVA -------------------------------------------------------------------//
    case DVA:
      Save = getConfigValue(SDV)                                                     // SDV = Save to DVA
      Edit = getConfigValue(DDV)                                                     // DDV = Edit to DVA

      sheet_tr = fetchSheetByName(DVA);
      if (!sheet_tr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] =
          Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

      sheet_sr = fetchSheetByName(Valor);
      if (!sheet_sr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_sr`); return; }

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] =
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());
      var CHECK = sheet_up.getRange("K7").getValue();

      if ((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000)) {
        if ((New_S.valueOf() != "-" && New_S.valueOf() != "") &&
            (C2_S != 0 && C2_S !== "")) {
          processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit);
          doSaveData(Valor);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveData`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData`);
      }
      break;
    // -------------------------------------------------------------------Valor -------------------------------------------------------------------//
    case Valor:
      Save = getConfigValue(SDV)                                                     // SDV = Save to DVA
      Edit = getConfigValue(DDV)                                                     // DDV = Edit to DVA

      sheet_sr = fetchSheetByName(Valor);
      if (!sheet_sr) { Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveData from sheet_sr`); return; }

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] =
          Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if ((New_S.valueOf() != "-" && New_S.valueOf() != "") &&
          (C2_S != 0 && C2_S !== "")) {
        processSaveData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveData`);
      }
      break;
    default: 
      Logger.log(`ERROR SAVE: ${SheetName} - Unhandled sheet type in doSaveData`);
      break;
  }
}

/////////////////////////////////////////////////////////////////////OTHER/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////PROVENTOS TEMPLATE/////////////////////////////////////////////////////////////////////

function doProventos() 
{
  doCheckDATA(PROV);
  doGetProventos();
  doSaveProventos();
  doExportProventos();
}

function doSaveProventos() {
  const ProvNames = [
    { name: 'Proventos', checkCell: "B3", expectedValue: 'Proventos', sourceRange: "B3:H60", targetRange: "B3:H60" },
    { name: 'Subscrição', checkCell: "L3", expectedValue: 'Tipo', sourceRange: "L3:T60", targetRange: "L3:T60" },
    { name: 'Ativos', checkCell: "B64", expectedValue: 'Proventos', sourceRange: "B64:H200", targetRange: "B64:H200" },
    { name: 'Historico', checkCell: "L64", expectedValue: 'Tipo de Ativo', dynamicRange: true }
  ];

  ProvNames.forEach(config => {
    try { doSaveProv(config); } 
    catch (error) { Logger.error(`Error saving ${config.name}: ${error}`); }
  });
}

function doSaveProv(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = fetchSheetByName('Prov_');                                    // Source Sheet

  if (!sheet_sr) { Logger.log(`ERROR: Target sheet "Prov_" does not exist. Skipping operation.`); return; }

  const sheet_tr = fetchSheetByName('Prov');                                     // Target Sheet

  if (!sheet_tr) { Logger.log(`ERROR: Target sheet "Prov" does not exist. Skipping operation.`); return; }
  const checkValue = sheet_sr.getRange(config.checkCell).getDisplayValue().trim();
  
  if (checkValue === config.expectedValue) {
    let data;
    
    if (config.dynamicRange) {
      const lr = sheet_sr.getLastRow();
      const lc = sheet_sr.getLastColumn();
      const sourceRange = sheet_sr.getRange(64, 12, lr - 63, lc - 11);
      const targetRange = sheet_tr.getRange(64, 12, lr - 63, lc - 11);
      
      data = sourceRange.getValues();
      targetRange.clearContent(); // Clear target range before writing data
      targetRange.setValues(data);
    } else {
      const sourceRange = sheet_sr.getRange(config.sourceRange);
      const targetRange = sheet_tr.getRange(config.targetRange);
      
      data = sourceRange.getValues();
      targetRange.clearContent(); // Clear target range before writing data
      targetRange.setValues(data);
    }
    
    Logger.log(`SUCCESS SAVE: ${config.name}.`);
  } else {
    Logger.log(`ERROR SAVE: ${config.name}, ${config.checkCell} != ${config.expectedValue}`);
  }
}

function doGetProventos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                   // Config sheet
  const sheet_tr = fetchSheetByName('Prov_');

  if (!sheet_tr)  { Logger.log(`ERROR: Target sheet "Prov_" does not exist. Skipping operation.`); return; }

  const TKT = sheet_co.getRange(TKR).getValue();                                 // TKR = Ticket Range
  const ticker = TKT.substring(0, 4);
  const language = 'pt-br';

  const data = JSON.stringify({
    issuingCompany: ticker,
    language: language
  });

  const base64Params = Utilities.base64Encode(data);

  const url = `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/${base64Params}`;
  Logger.log(`URL: ${url}`);

  let responseText;
  try {
    const response = UrlFetchApp.fetch(url);
    responseText = response.getContentText().trim();
    Logger.log(`API Response: ${responseText}`);
  } catch (error) {
    Logger.log(`ERROR: Failed to fetch API response. ${error}`);
    return; // Exit if the API request fails
  }

  if (!responseText) { Logger.log("ERROR: Empty response from API."); }

  let content;
  try { content = JSON.parse(responseText); } 
  catch (error) { Logger.log(`ERROR: Failed to parse JSON response. ${error}`); }

  if (!content || !content[0]) { Logger.log("ERROR: No data returned from API."); }

  fillCashDividends(sheet_tr, content[0]?.cashDividends || []);
  fillStockDividends(sheet_tr, content[0]?.stockDividends || []);
  fillSubscriptions(sheet_tr, content[0]?.subscriptions || []);
}

// Fill Cash Dividends from B2 to B60
function fillCashDividends(sheet_tr, dividends) {
  const headerRange = "B3:H3";
  const startRow = 4;
  const maxRows = 57;

  sheet_tr.getRange("B2").setValue('Proventos em Dinheiro');
  sheet_tr.getRange("B3:H60").clearContent();

  const headers = ['Proventos', 'Código ISIN', 'Data de Aprovação', 'Última Data Com', 'Valor (R$)', 'Relacionado a', 'Data de Pagamento'];
  sheet_tr.getRange(headerRange).setValues([headers]);

  dividends.slice(0, maxRows).forEach((div, i) => {
    sheet_tr.getRange(startRow + i, 2, 1, 7).setValues([[ 
      div.label, div.isinCode, div.approvedOn, div.lastDatePrior, div.rate, div.relatedTo, div.paymentDate
    ]]);
  });
}

// Fill Stock Dividends from row 63
function fillStockDividends(sheet_tr, stockDividends) {
  const startRow = 63;
  const headerRange = `B${startRow + 1}:G${startRow + 1}`;

  sheet_tr.getRange(`B${startRow}:G${startRow + stockDividends.length + 1}`).clearContent();
  sheet_tr.getRange(`B${startRow}`).setValue("Dividendos em Ações");

  const headers = ['Proventos', 'Código ISIN', 'Data de Aprovação', 'Última Data Com', 'Fator', 'Ativo Emitido'];
  sheet_tr.getRange(headerRange).setValues([headers]);

  stockDividends.forEach((stockDiv, i) => {
    sheet_tr.getRange(startRow + 2 + i, 2, 1, 6).setValues([[ 
      stockDiv.label, stockDiv.isinCode, stockDiv.approvedOn, stockDiv.lastDatePrior, stockDiv.factor, stockDiv.assetIssued
    ]]);
  });
}

// Fill Subscriptions starting from column L, row 2
function fillSubscriptions(sheet_tr, subscriptions) {
  const headerRange = "L3:T3";
  const startRow = 4;

  sheet_tr.getRange("L2").setValue('Subscrições');
  sheet_tr.getRange("L3:T60").clearContent();

  const headers = ['Tipo', 'Código ISIN', 'Data de Aprovação', 'Última Data Com', 'Percentual (%)', 'Ativo Emitido', 'Preço Emissão (R$)', 'Período de Negociação', 'Data de Subscrição'];
  sheet_tr.getRange(headerRange).setValues([headers]);

  subscriptions.forEach((sub, i) => {
    sheet_tr.getRange(startRow + i, 12, 1, 9).setValues([[ 
      sub.label, sub.isinCode, sub.approvedOn, sub.lastDatePrior, sub.percentage, sub.assetIssued, sub.priceUnit, sub.tradingPeriod, sub.subscriptionDate
    ]]);
  });
}

/////////////////////////////////////////////////////////////////////CodeCVM/////////////////////////////////////////////////////////////////////

function doGetCodeCVM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_tr = fetchSheetByName('Info');                                    // Target sheet
 
  if (!sheet_tr) { Logger.log(`ERROR: Target sheet "Info" does not exist. Skipping operation.`); return; }

  const sheet_co = fetchSheetByName('Config');                                   // Config sheet
  const TKT = sheet_co.getRange(TKR).getValue();                                 // TKR = Ticket Range
  const ticker = TKT.substring(0, 4);
  const language = 'pt-br';

  const data = JSON.stringify({
    issuingCompany: ticker,
    language: language
  });

  const base64Params = Utilities.base64Encode(data);

  const url = `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/${base64Params}`;
  Logger.log("URL:", url);

  let responseText;
  try {
    const response = UrlFetchApp.fetch(url);
    responseText = response.getContentText().trim();
    Logger.log("API Response:", responseText);
  } 
  catch (error) {
    Logger.log(`ERROR: Failed to fetch API response. ${error}`);

    return; // Exit if the API request fails
  }

  if (!responseText) { Logger.log("ERROR: Empty response from API."); }

  let content;
  try { content = JSON.parse(responseText); } 
  catch (error) { Logger.log(`ERROR: Failed to parse JSON response. ${error}`); }

  if (!content || !content[0]) { Logger.log(`ERROR: No data returned from API.`); }

  const codeCVM = content[0]?.codeCVM || 'N/A';                                     // Default to 'N/A' if codeCVM is missing
  Logger.log(`Extracted codeCVM: ${codeCVM}`);

  // Write to the Info sheet
  sheet_tr.getRange("C3").setValue(codeCVM);
}

/////////////////////////////////////////////////////////////////////SAVE AND SHARES TEMPLATE/////////////////////////////////////////////////////////////////////

function doSaveShares() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = fetchSheetByName('DATA');

  if (!sheet_sr) { Logger.log(`ERROR: DATA sheet not found. Skipping shares save.`); return; }

  try {
    var M1 = sheet_sr.getRange("M1").getValue();
    var M2 = sheet_sr.getRange("M2").getValue();

    Logger.log(`SAVE: Shares and FF`);

    if (!isNaN(M1) && !isNaN(M2) && !ErrorValues.includes(M1) && !ErrorValues.includes(M2)) {
      M1 = Number(M1); // Convert to number if not already
      M2 = Number(M2);

      var Data = sheet_sr.getRange("M1:M2").getValues();
      sheet_sr.getRange("L1:L2").setValues(Data);
      Logger.log(`SUCCESS SAVE: Shares and FF`);
    } else { Logger.log(`ERROR SAVE: Invalid values in M1/M2`); }
  } 
  catch (error) { Logger.log(`ERROR in doSaveShares:`, error.message); }
}

/////////////////////////////////////////////////////////////////////SAVE TEMPLATE/////////////////////////////////////////////////////////////////////