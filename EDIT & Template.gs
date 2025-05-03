/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doEditAll()
{
  doEditBasics();
  doEditExtras();
  doEditFinancials();
  doIsFormula();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

function doEditGroup(SheetNames, editFunction, label) {
  _doGroup(SheetNames, editFunction, "Editing", "edited", label);
}

//-------------------------------------------------------------------BASICS-------------------------------------------------------------------//

function doEditBasics() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  doEditGroup(SheetNames, doEditSheet, 'basic');
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doEditExtras() {
  const SheetNames = [FUTURE, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  doEditGroup(SheetNames, doEditSheet, 'extra');
}

//-------------------------------------------------------------------FINANCIALS-------------------------------------------------------------------//

function doEditFinancials() {
  const SheetNames = [BLC, Balanco, DRE, Resultado, FLC, Fluxo, DVA, Valor];
  doEditGroup(SheetNames, doEditFinancial, 'financial');
}

/////////////////////////////////////////////////////////////////////BASIC TEMPLATE/////////////////////////////////////////////////////////////////////

function doEditBasic(SheetName) {
  Logger.log(`EDIT: ${SheetName}`);
  const sheet_sr = fetchSheetByName(SheetName); // Source sheet
  if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditSheet from sheet_sr`); return; }

  Utilities.sleep(2500); // 2.5 secs pause

  let Edit;

  switch (SheetName) {
//-------------------------------------------------------------------Swing-------------------------------------------------------------------//
    case SWING_4:
    case SWING_12:
    case SWING_52:
      Edit        = getConfigValue(DTR);                                          // DTR = Edit to Swing
      const Class = getConfigValue(IST, 'Config');                                // IST = Is Stock?

      var C2 = sheet_sr.getRange("C2").getValue();
      if (Class == 'STOCK') {
        if (C2 > 0) {
          processEditBasic(sheet_sr, SheetName, Edit);
        }
      }
      if (Class == 'BDR' || Class == 'ETF' || Class == 'ADR') {
        if (C2 > 0) {
          processEditBasic(sheet_sr, SheetName, Edit);
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
        processEditBasic(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;

//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
    case BTC:
      Edit = getConfigValue(DBT);                                                 // DBT = Edit to BTC

      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processEditBasic(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;

//-------------------------------------------------------------------Termo-------------------------------------------------------------------//
    case TERMO:
      Edit = getConfigValue(DTE);                                                 // DTE = Edit to Termo

      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        processEditBasic(sheet_sr, SheetName, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditSheet`);
      }
      break;

//-------------------------------------------------------------------Fund-------------------------------------------------------------------//
    case FUND:
      Edit = getConfigValue(DFU);                                                 // DFU = Edit to Fund

      var B2 = sheet_sr.getRange("B2").getValue();
      if (!ErrorValues.includes(B2)) {
        processEditBasic(sheet_sr, SheetName, Edit);
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
        processEditBasic(sheet_sr, SheetName, Edit);
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

/////////////////////////////////////////////////////////////////////FINANCIAL TEMPLATE/////////////////////////////////////////////////////////////////////

// sheet_sr and sheet_tr are checked  inside the blocks

function doEditFinancial(SheetName) {
    Logger.log(`EDIT: ${SheetName}`);
  const sheet_sr = fetchSheetByName(SheetName);
  if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

  Utilities.sleep(2500); // 2,5 secs

  let Edit, Values_sr;

  switch (SheetName) {
//-------------------------------------------------------------------BLC-------------------------------------------------------------------//
    case BLC: {
      Edit = getConfigValue(DBL);                                                     // DBL = Edit to BLC

      const sheet_tr = fetchSheetByName(BLC);
      if (!sheet_tr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [New_tr, Old_tr]  = doFinancialDateHelper(Values_tr);

      const sheet_sr = fetchSheetByName(Balanco);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:C1").getValues()[0];
      var [New_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [B2_sr, B27_sr] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_sr.valueOf() != "-" && New_sr.valueOf() != "" ) &&
          ( B2_sr != 0 && B2_sr != "" ) &&
          ( B27_sr != 0 && B27_sr != "" ) )
      {
        processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit);
        doEditFinancial(Balanco);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`);
      }
    }
      break;
//-------------------------------------------------------------------Balanço-------------------------------------------------------------------//
    case Balanco: {
      Edit = getConfigValue(DBL);                                                     // DBL = Edit to BLC

      const sheet_sr = fetchSheetByName(Balanco);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:C1").getValues()[0];
      var [New_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [B2_sr, B27_sr] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_sr.valueOf() != "-" && New_sr.valueOf() != "" ) &&
          ( B2_sr != 0 && B2_sr != "" ) &&
          ( B27_sr != 0 && B27_sr != "" ) )
      {
        processEditFinancial(sheet_sr, sheet_sr, '', '', New_sr, Old_sr, Edit);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`);
      }
    }
      break;
//-------------------------------------------------------------------DRE-------------------------------------------------------------------//
    case DRE: {
      Edit = getConfigValue(DDE);                                                     // DDE = Edit to DRE

      const sheet_tr = fetchSheetByName(DRE);
      if (!sheet_tr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [New_tr, Old_tr]  = doFinancialDateHelper(Values_tr);

      const sheet_sr = fetchSheetByName(Resultado);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C4_sr, C27_sr] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_sr.valueOf() != "-" && New_sr.valueOf() != "" ) &&
          ( C4_sr != 0 && C4_sr != "" ) &&
          ( C27_sr != 0 && C27_sr != "" ) )
      {
        processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit);
        doEditFinancial(Resultado);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`);
      }
    }
      break;
//-------------------------------------------------------------------Resultado-------------------------------------------------------------------//
    case Resultado: {
      Edit = getConfigValue(DDE);                                                     // DDE = Edit to DRE

      const sheet_sr = fetchSheetByName(Resultado);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C4_sr, C27_sr] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_sr.valueOf() != "-" && New_sr.valueOf() != "" ) &&
          ( C4_sr != "" ) &&
          ( C27_sr != 0 && C27_sr != "" ) )
      {
        processEditFinancial(sheet_sr, sheet_sr, '', '', New_sr, Old_sr, Edit);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`);
      }
    }
      break;
//-------------------------------------------------------------------FLC-------------------------------------------------------------------//
    case FLC: {
      Edit = getConfigValue(DFL);                                                     // DFL = Edit to FLC

      const sheet_tr = fetchSheetByName(FLC);
      if (!sheet_tr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [New_tr, Old_tr]  = doFinancialDateHelper(Values_tr);

      const sheet_sr = fetchSheetByName(Fluxo);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [B2_sr] = ["B2"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_sr.valueOf() != "-" && New_sr.valueOf() != "" ) &&
          ( B2_sr != 0 && B2_sr != "" ) )
      {
        processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit);
        doEditFinancial(Fluxo);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`);
      }
    }
      break;
//-------------------------------------------------------------------Fluxo-------------------------------------------------------------------//
    case Fluxo: {
      Edit = getConfigValue(DFL);                                                     // DFL = Edit to FLC

      const sheet_sr = fetchSheetByName(Fluxo);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [B2_sr] = ["B2"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_sr.valueOf() != "-" && New_sr.valueOf() != "" ) &&
          ( B2_sr != 0 && B2_sr != "" ) )
      {
        processEditFinancial(sheet_sr, sheet_sr, '', '', New_sr, Old_sr, Edit);
      }
      else
      {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`);
      }
    }
      break;
//-------------------------------------------------------------------DVA-------------------------------------------------------------------//
    case DVA: {
      Edit = getConfigValue(DDV);                                                     // DDV = Edit to DVA

      const sheet_tr = fetchSheetByName(DVA);
      if (!sheet_tr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_tr`); return; }

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [New_tr, Old_tr]  = doFinancialDateHelper(Values_tr);

      const sheet_sr = fetchSheetByName(Valor);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C2_sr] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_sr.valueOf() != "-" && New_sr.valueOf() != "" ) &&
          ( B2_sr != "" ) ) {
        processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit);
        doEditFinancial(Valor);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`);
      }
    }
      break;
//-------------------------------------------------------------------Valor-------------------------------------------------------------------//
    case Valor: {
      Edit = getConfigValue(DDV);                                                     // DDV = Edit to DVA

      const sheet_sr = fetchSheetByName(Valor);
      if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditFinancial from sheet_sr`); return; }

      Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C2_sr] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if( ( New_sr.valueOf() != "-" && New_sr.valueOf() != "" ) &&
          ( C2_sr != "" ) ) {
        processEditFinancial(sheet_sr, sheet_sr, '', '', New_sr, Old_sr, Edit);
      } else {
        Logger.log(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`);
      }
    }
      break;

    default:
      Logger.log(`ERROR EDIT: ${SheetName} - Unhandled sheet type in doEditFinancial`);
      break;
  }
}

/////////////////////////////////////////////////////////////////////EDIT TEMPLATE/////////////////////////////////////////////////////////////////////
