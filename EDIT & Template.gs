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
  const sheet_sr = fetchSheetByName(SheetName);
  if (!sheet_sr) { Logger.log(`ERROR EDIT: ${SheetName} - Does not exist on doEditSheet from sheet_sr`); return; }
  Utilities.sleep(2500);

  const editTable = [
    {
      names: [SWING_4, SWING_12, SWING_52],
      editKey: DTR,                                          // DTR = Edit to Swing
      cells: ['C2'],
      test: ([c2]) => {
        const cls = getConfigValue(IST, 'Config');       // IST = Is Stock?
        return c2 > 0 && ['STOCK','BDR','ETF','ADR'].includes(cls);
      },
      handler: processEditBasic
    },
    {
      names: [OPCOES],
      editKey: DOP,                                          // DOP = Edit to Option
      cells: ['C2','E2'],
      test: ([call, put]) => (call != 0 && put != 0 && call !== '' && put !== ''),
      handler: processEditBasic
    },
    {
      names: [BTC],
      editKey: DBT,                                          // DBT = Edit to BTC
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditBasic
    },
    {
      names: [TERMO],
      editKey: DTE,                                          // DTE = Edit to Termo
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditBasic
    },
    {
      names: [FUND],
      editKey: DFU,                                          // DFU = Edit to Fund
      cells: ['B2'],
      test: ([b2]) => !ErrorValues.includes(b2),
      handler: processEditBasic
    },
    {
      names: [FUTURE],
      editKey: DFT,                                          // DFT = Edit to Future
      cells: ['C2','E2','G2'],
      test: vals => vals.some(v => !ErrorValues.includes(v)),
      handler: processEditBasic
    },
    {
      names: [FUTURE_1, FUTURE_2, FUTURE_3],
      editKey: DFT,                                          // DFT = Edit to Future
      cells: ['C2'],
      test: ([c2]) => !ErrorValues.includes(c2),
      handler: processEditExtra
    },
    {
      names: [RIGHT_1, RIGHT_2],
      editKey: DRT,                                          // DRT = Edit to Right
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [RECEIPT_9, RECEIPT_10],
      editKey: DRC,                                          // DRC = Edit to Receipt
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [WARRANT_11, WARRANT_12, WARRANT_13],
      editKey: DWT,                                          // DWT = Edit to Warrant
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [BLOCK],
      editKey: DBK,                                          // DBK = Edit to Block
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    }
  ];

  const config = editTable.find(cfg => cfg.names.includes(SheetName));
  if (!config) { Logger.log(`ERROR EDIT: ${SheetName} - Unhandled sheet type in doEditSheet`); return; }

  const editValue = getConfigValue(config.editKey);
  const values = config.cells.map(a1 => sheet_sr.getRange(a1).getValue());

  if (config.test(values)) {
    config.handler(sheet_sr, SheetName, editValue);
  } else { Logger.log(`ERROR EDIT: ${SheetName} - Conditions aren’t met in doEditSheet`); }
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
