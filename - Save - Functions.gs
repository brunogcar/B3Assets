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
      Logger.error(`Error checking DATA for sheet ${SheetName}:`, error);
    }
  });
}

/////////////////////////////////////////////////////////////////////DO CHECK TEMPLATE/////////////////////////////////////////////////////////////////////

function doCheckDATA(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_s = ss.getSheetByName(SheetName);                    // Source sheet

  var sheet_p = ss.getSheetByName(PROV);
  var sheet_o = ss.getSheetByName('OPT');
  var sheet_d = ss.getSheetByName('DATA');
  var sheet_b = ss.getSheetByName(Balanco);
  var sheet_r = ss.getSheetByName(Resultado);
  var sheet_f = ss.getSheetByName(Fluxo);
  var sheet_v = ss.getSheetByName(Valor);

  let Check;

  Logger.log('CHECK Sheet:', SheetName);


//-------------------------------------------------------------------BTC-------------------------------------------------------------------//

  if (SheetName === PROV)
  {
    Check = sheet_p.getRange('B3').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }
//-------------------------------------------------------------------OPCOES-------------------------------------------------------------------//

  if (SheetName === OPCOES)
  {
    Check = sheet_o.getRange('B2').getValue();

    if (Check === '') 
    {
      sheet_o.hideSheet();
      Logger.log('HIDDEN:', 'OPT');
    }
    if (Check !== '') 
    {
      if (sheet_o.isSheetHidden()) 
      {
        sheet_o.showSheet();
        Logger.log('DISPLAYED:', SheetName);
      }
    }
    return processCheckDATA(sheet_s, SheetName, Check);
  }

//-------------------------------------------------------------------SWING-------------------------------------------------------------------//

  if (SheetName === SWING_4 || SheetName === SWING_12 || SheetName === SWING_52)
  {
    const sheet_c = ss.getSheetByName('Config');
    var Class = sheet_c.getRange(IST).getDisplayValue();                             // IST = Is Stock? 

    if (Class == 'STOCK') 
    {
      Check = sheet_d.getRange('B16').getValue();
    }
    else
    {
      Check = 'TRUE'
    }

    return processCheckDATA(sheet_s, SheetName, Check);
  }

//-------------------------------------------------------------------BTC-------------------------------------------------------------------//

  if (SheetName === BTC)
  {
    Check = sheet_d.getRange('B3').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

//-------------------------------------------------------------------TERMO-------------------------------------------------------------------//

  if (SheetName === TERMO)
  {
    Check = sheet_d.getRange('B24').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

//-------------------------------------------------------------------FUTURE-------------------------------------------------------------------//

  if (SheetName === FUTURE) 
  {
    let Checks = ['B32', 'B33', 'B34'];

    for (let i = 0; i < Checks.length; i++) 
    {
      Check = sheet_d.getRange(Checks[i]).getValue();
      if (!ErrorValues.includes(Check)) 
      {
        return processCheckDATA(sheet_s, SheetName, Check);
      }
    }
  }

  if (SheetName === FUTURE_1)
  {
    Check = sheet_d.getRange('B32').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === FUTURE_2)
  {
    Check = sheet_d.getRange('B33').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === FUTURE_3)
  {
    Check = sheet_d.getRange('B34').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

//-------------------------------------------------------------------RIGHT-------------------------------------------------------------------//

  if (SheetName === RIGHT_1)
  {
    Check = sheet_d.getRange('C38').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === RIGHT_2)
  {
    Check = sheet_d.getRange('C39').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

//-------------------------------------------------------------------RECEIPT-------------------------------------------------------------------//

  if (SheetName === RECEIPT_9)
  {
    Check = sheet_d.getRange('C44').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === RECEIPT_10)
  {
    Check = sheet_d.getRange('C45').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

//-------------------------------------------------------------------WARRANT-------------------------------------------------------------------//

  if (SheetName === WARRANT_11)
  {
    Check = sheet_d.getRange('C50').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === WARRANT_12)
  {
    Check = sheet_d.getRange('C51').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === WARRANT_13)
  {
    Check = sheet_d.getRange('C52').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

//-------------------------------------------------------------------BLOCK-------------------------------------------------------------------//

  if (SheetName === BLOCK) 
  {
    let Checks = ['C56', 'C57', 'C58'];

    for (let i = 0; i < Checks.length; i++) 
    {
      Check = sheet_d.getRange(Checks[i]).getValue();
      if (!ErrorValues.includes(Check)) 
//      if (ErrorValues.includes(Check)) 
      {
        return processCheckDATA(sheet_s, SheetName, Check);
      }
    }
  }

//-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//

  if (SheetName === BLC)
  {
    Check = sheet_b.getRange('B1').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === DRE)
  {
    Check = sheet_r.getRange('C1').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === FLC)
  {
    Check = sheet_f.getRange('C1').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }

  if (SheetName === DVA)
  {
    Check = sheet_v.getRange('C1').getValue();

    return processCheckDATA(sheet_s, SheetName, Check);
  }
}

/////////////////////////////////////////////////////////////////////DO CHECK Process/////////////////////////////////////////////////////////////////////

function processCheckDATA(sheet_s, SheetName, Check) 
{
  if (!ErrorValues.includes(Check)) 
  {
    if (sheet_s.isSheetHidden()) 
    {
      sheet_s.showSheet();
      Logger.log('DISPLAYED:', SheetName);
    }
    Logger.log("DATA Check: TRUE");
    return "TRUE";
  }
  else
  {
    if (SheetName === BLC || SheetName === DRE || SheetName === FLC || SheetName === DVA) 
    {
      Logger.log("DATA Check: FALSE");
      return "FALSE"; // Add a default return value if the conditions are not met
    }
    else
    {
      if (!sheet_s.isSheetHidden()) 
      {
        sheet_s.hideSheet();
        Logger.log('HIDDEN:', SheetName);
      }
      Logger.log("DATA Check: FALSE");
      return "FALSE"; // Add a default return value if the conditions are not met
    }
  }
  // Add a default return value in case none of the conditions are met
  return "FALSE";
}

/////////////////////////////////////////////////////////////////////TRIM TEMPLATE/////////////////////////////////////////////////////////////////////

function doTrim() 
{
  const SheetNames = [
    SWING_4, SWING_12, SWING_52
  ];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doTrimSheet(SheetName);
    } 
    catch (error) 
    {
      Logger.error(`Error saving sheet ${SheetName}:`, error);
    }
  });
}

function doTrimSheet(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_s = ss.getSheetByName(SheetName); // Target

  Logger.log('TRIM:', SheetName);

  if (!sheet_s) 
  {
    Logger.error(`Sheet ${SheetName} not found.`);
    return;
  }

  var LR = sheet_s.getLastRow();
  var LC = sheet_s.getLastColumn();

  if (SheetName === SWING_4) 
  {
    if (LR > 126) 
    {
      sheet_s.getRange(127, 1, LR - 126, LC).clearContent();
      Logger.log(`SUCCESS TRIM. Sheet: ${SheetName}.`);
      Logger.log(`Cleared data below row 126 in ${SheetName}.`);
    }
  } 
  else if (SheetName === SWING_12) 
  {
    if (LR > 366) 
    {
      sheet_s.getRange(367, 1, LR - 366, LC).clearContent();
      Logger.log(`SUCCESS TRIM. Sheet: ${SheetName}.`);
      Logger.log(`Cleared data below row 366 in ${SheetName}.`);
    }
  } 
  else if (SheetName === SWING_52) 
  {
      Logger.log(`NOTHING TO TRIM. Sheet: ${SheetName}.`);
  } 
  else 
  {
    // Default logic for other sheets
    Logger.log(`No specific logic defined for ${SheetName}.`);
  }

}

/////////////////////////////////////////////////////////////////////Hide and Show Sheets/////////////////////////////////////////////////////////////////////

function doDisableSheets() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_c = ss.getSheetByName('Config');
  const sheets = ss.getSheets();

  var Class = sheet_c.getRange(IST).getDisplayValue();                                                                 // IST = Is Stock?

  if (Class === 'STOCK') 
  {
    var SheetNames = ['DATA', 'Prov_', 'FIBO', 'Cotações', 'UPDATE', 'Balanço', 'Balanço Ativo', 'Balanço Passivo', 'Resultado', 'Demonstração', 'Fluxo', 'Fluxo de Caixa', 'Valor', 'Demonstração do Valor Adicionado'];

    for (var i = 0; i < sheets.length; i++) 
    {
      const sheet = sheets[i];
      if (sheet && SheetNames.indexOf(sheet.getName()) !== -1 && !sheet.isSheetHidden()) {
        sheet.hideSheet();
        Logger.log('Sheet:', sheet.getName(), 'HIDDEN');
      }
    }
  } 
  else if (Class === 'ADR') 
  {
    var SheetNames = ['Config', 'Settings', 'Index', 'Preço', 'FIBO', SWING_4, SWING_12, SWING_52, 'Cotações'];

    for (var i = sheets.length - 1; i >= 0; i--) 
    {                                                                                                                 // Reverse iteration to avoid index shifting
      const sheet = sheets[i];
      if (sheet && SheetNames.indexOf(sheet.getName()) === -1) 
      {                                                                                                               // Delete all but SheetNames
        Logger.log('Deleting sheet:', sheet.getName());
        ss.deleteSheet(sheet);
      }
    }
  } 
  else if (Class === 'BDR' || Class === 'ETF') 
  {
    var SheetNames = ['Config', 'Settings', 'Index', 'Prov', 'Prov_', 'Preço', 'FIBO', SWING_4, SWING_12, SWING_52, 'Cotações', 'DATA', 'OPT', 'Opções', 'BTC', 'Termo'];

    for (var i = sheets.length - 1; i >= 0; i--) 
    {                                                                                                                 // Reverse iteration to avoid index shifting
      const sheet = sheets[i];
      if (sheet && SheetNames.indexOf(sheet.getName()) === -1) 
      { // Delete all but SheetNames
        Logger.log('Deleting sheet:', sheet.getName());
        ss.deleteSheet(sheet);
      }
    }
  }
  hideConfig();
}

/////////////////////////////////////////////////////////////////////HIDE CONFIG/////////////////////////////////////////////////////////////////////

function hideConfig()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_s = ss.getSheetByName('Settings');                        // Source sheet
  const sheet_c = ss.getSheetByName('Config');                          // Config sheet

  var Hide_Config = sheet_c.getRange(HCR).getDisplayValue();                       // HCR = Hide Config Range

  if ( Hide_Config == "TRUE")
  {
    if (sheet_s && !sheet_s.isSheetHidden())
    {
      sheet_s.hideSheet();
      Logger.log('HIDDEN:', sheet_s.getName()); 
    }
    if (sheet_c && !sheet_c.isSheetHidden())
    {
      sheet_c.hideSheet();
      Logger.log('HIDDEN:', sheet_c.getName()); 
    }
  }
};

/////////////////////////////////////////////////////////////////////SAVE FUNCTIONS/////////////////////////////////////////////////////////////////////