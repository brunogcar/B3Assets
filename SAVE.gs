/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doSaveAll()
{
  Logger.log(SNAME(2));

//-------------------------------------------------------------------DATA-------------------------------------------------------------------//

  var sheetsToCheck = [BLC, DRE, FLC, DVA];
  var sheetsToSave = [];

  sheetsToCheck.forEach(function(SheetName)
  {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
    if (sheet)
    {
      var availableData = doCheckDATA(SheetName);
      if (availableData === "TRUE")
      {
        sheetsToSave.push(SheetName);
      }
    }
    else
    {
      Logger.log(`ERROR SAVE: ${SheetName} - Does not exist`);
    }
  });

  if (sheetsToSave.length > 0)
  {
    SpreadsheetApp.flush();

    sheetsToSave.forEach(function(SheetName)
    {
      doSaveData(SheetName);
    });
  }
  else
  {
    Logger.log(`No valid data found. Skipping save operation.`);
  }

  doSaveShares();
  doProventos();

//-------------------------------------------------------------------SHEETS-------------------------------------------------------------------//

  var sheetsToCheck = [OPCOES, BTC, TERMO];
  var sheetsToSave = [];

  sheetsToCheck.forEach(function(SheetName)
  {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
    if (sheet)
    {
      var availableData = doCheckDATA(SheetName);
      if (availableData === "TRUE")
      {
        sheetsToSave.push(SheetName);
      }
    }
    else
    {
      Logger.log(`ERROR SAVE: ${SheetName} - Does not exist`);
    }
  });

  if (sheetsToSave.length > 0)
  {
    SpreadsheetApp.flush();

    sheetsToSave.forEach(function(SheetName)
    {
      doSaveSheet(SheetName);
    });
  }
  else
  {
    Logger.log(`No valid data found. Skipping save operation.`);
  }

  doSaveSheet(FUND);

  SpreadsheetApp.flush();      // to wait for data load to save proventos and swings
  
  doSaveSheet(SWING_4);        // out of order, to give time to get data
  doSaveSheet(SWING_12);       // out of order, to give time to get data
  doSaveSheet(SWING_52);       // out of order, to give time to get data

//-------------------------------------------------------------------EXTRA-------------------------------------------------------------------//

  var sheetsToCheck = [
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ];
  var sheetsToSave = [];

  sheetsToCheck.forEach(function(SheetName)
  {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
    if (sheet)
    {
      var availableData = doCheckDATA(SheetName);
      if (availableData === "TRUE")
      {
        sheetsToSave.push(SheetName);
      }
    }
    else
    {
      Logger.log(`ERROR SAVE: ${SheetName} - Does not exist`);
    }
  });

  if (sheetsToSave.length > 0) 
  {
    SpreadsheetApp.flush();

    sheetsToSave.forEach(function(SheetName)
    {
      doSaveSheet(SheetName);
    });
  }
  else 
  {
    Logger.log(`No valid data found. Skipping save operation.`);
  }

//  doExportAll();             // sheets are exported individually

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
};

/////////////////////////////////////////////////////////////////////Individual/////////////////////////////////////////////////////////////////////

//-------------------------------------------------------------------SHEETS-------------------------------------------------------------------//

function doSaveAllSheets()
{
  Logger.log(SNAME(2));

  doSaveShares();
  doProventos();

  var sheetsToCheck = [OPCOES, BTC, TERMO];
  var sheetsToSave = [];

  sheetsToCheck.forEach(function(SheetName)
  {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
    if (sheet)
    {
      var availableData = doCheckDATA(SheetName);
      if (availableData === "TRUE")
      {
        sheetsToSave.push(SheetName);
      }
    }
    else
    {
      Logger.log(`ERROR SAVE: ${SheetName} - Does not exist`);
    }
  });

  if (sheetsToSave.length > 0)
  {
    SpreadsheetApp.flush();

    sheetsToSave.forEach(function(SheetName)
    {
      doSaveSheet(SheetName);
    });
  }
  else 
  {
    Logger.log(`No valid data found. Skipping save operation.`);
  }

  doSaveSheet(FUND);

  doSaveSheet(SWING_4);        // out of order, to give time to get data
  doSaveSheet(SWING_12);       // out of order, to give time to get data
  doSaveSheet(SWING_52);       // out of order, to give time to get data

  doExportExtras();
  doExportDatas();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doSaveAllExtras()
{
  Logger.log(SNAME(2));

  var sheetsToCheck = [
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ];
  var sheetsToSave = [];

  sheetsToCheck.forEach(function(SheetName)
  {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
    if (sheet)
    {
      var availableData = doCheckDATA(SheetName);
      if (availableData === "TRUE")
      {
        sheetsToSave.push(SheetName);
      }
    }
    else
    {
      Logger.log(`ERROR SAVE: ${SheetName} - Does not exist`);
    }
  });

  if (sheetsToSave.length > 0)
  {
    SpreadsheetApp.flush();

    sheetsToSave.forEach(function(SheetName)
    {
      doSaveSheet(SheetName);
    });
  } 
  else 
  {
    Logger.log(`No valid data found. Skipping save operation.`);
  }

  doSaveShares();
  doProventos();

  doExportSheets();
  doExportDatas();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}

//-------------------------------------------------------------------DATAS-------------------------------------------------------------------//

function doSaveAllDatas()
{
  Logger.log(SNAME(2));

  var sheetsToCheck = [BLC, DRE, FLC, DVA];
  var sheetsToSave = [];

  sheetsToCheck.forEach(function(SheetName)
  {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
    if (sheet)
    {
      var availableData = doCheckDATA(SheetName);
      if (availableData === "TRUE")
      {
        sheetsToSave.push(SheetName);
      }
    }
    else
    {
      Logger.log(`ERROR SAVE: ${SheetName} - Does not exist`);
    }
  });

  if (sheetsToSave.length > 0)
  {
    SpreadsheetApp.flush();

    sheetsToSave.forEach(function(SheetName)
    {
      doSaveData(SheetName);
    });
  }
  else
  {
    Logger.log(`No valid data found. Skipping save operation.`);
  }

  doSaveShares();
  doProventos();

  doExportSheets();
  doExportExtras();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

  // could add addicional checks with && "and" || "or"
  // can get SheetName from ss.getName() as well

/////////////////////////////////////////////////////////////////////SHEETS/////////////////////////////////////////////////////////////////////

function doSaveSheets() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  const totalSheets = SheetNames.length;
  let Count = 0;

  Logger.log(`Starting save of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    Logger.log(`[${Count}/${totalSheets}] (${progress}%) saving ${SheetName}...`);

    try {
      doSaveSheet(SheetName);
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} saved successfully`);
    } catch (error) {
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error saving ${SheetName}: ${error}`);
    }
  });
  Logger.log(`Save completed: ${Count} of ${totalSheets} sheets saved successfully`);
}

/////////////////////////////////////////////////////////////////////EXTRAS/////////////////////////////////////////////////////////////////////

function doSaveExtras() {
  const SheetNames = [FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  const totalSheets = SheetNames.length;
  let Count = 0;

  Logger.log(`Starting save of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    Logger.log(`[${Count}/${totalSheets}] (${progress}%) saving ${SheetName}...`);

    try {
      doSaveSheet(SheetName);
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} saved successfully`);
    } catch (error) {
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error saving ${SheetName}: ${error}`);
    }
  });
  Logger.log(`Save completed: ${Count} of ${totalSheets} sheets saved successfully`);
}

/////////////////////////////////////////////////////////////////////DATAS/////////////////////////////////////////////////////////////////////

function doSaveDatas()
{
  const SheetNames = [BLC, DRE, FLC, DVA];                  //Balanço, Resultado, Fluxo and Valor are saved after parent SheetNames

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_u = ss.getSheetByName(`UPDATE`);                         // UPDATE sheet

    var ACTV = sheet_u.getRange(`B3`).getValue();
    var SOMA = sheet_u.getRange(`K8`).getValue();

  if (!ACTV || (ACTV && ((SOMA >= 450 && SOMA <= 460) || (SOMA == 0 || SOMA > 125000)))) // 200.000 if all new
  {
    SheetNames.forEach(SheetName =>
    {
      try
      {
        doSaveData(SheetName);
      }
      catch (error)
      {
        Logger.error(`Error saving sheet ${SheetName}:`, error);
      }
    });
  }
}

/////////////////////////////////////////////////////////////////////SAVE/////////////////////////////////////////////////////////////////////