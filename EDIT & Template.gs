/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doEditAll()
{
  doEditDatas();

  doEditSheets();

  doIsFormula();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

function doEditSheets() 
{
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  
  SheetNames.forEach(SheetName => 
  {
    try 
    {
    doEditSheet(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      console.error(`Error editing sheet ${SheetName}:`, error);
    }
  });
}

function doEditExtras() 
{
  const SheetNames = [FUTURE, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  
  SheetNames.forEach(SheetName => 
  {
    try 
    {
    doEditSheet(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      console.error(`Error editing sheet ${SheetName}:`, error);
    }
  });
}

function doEditDatas() 
{
  const SheetNames = [BLC, Balanco, DRE, Resultado, FLC, Fluxo, DVA, Valor];
  
  SheetNames.forEach(SheetName => 
  {
    try 
    {
    doEditData(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      console.error(`Error editing sheet ${SheetName}:`, error);
    }
  });
}

/////////////////////////////////////////////////////////////////////SHEETS TEMPLATE/////////////////////////////////////////////////////////////////////

function doEditSheet(SheetName) 
{
  console.log('EDIT:', SheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = getSheetnameByName(SheetName);                        // Source sheet
  if (!sheet_sr) { console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditSheet from sheet_sr'); return; }

  const sheet_co = getSheetnameByName('Config');                         // Config sheet
  const sheet_se = getSheetnameByName('Settings');
  if (!sheet_co || !sheet_se) return;

  Utilities.sleep(2500); // 2,5 secs

  let Edit;

//-------------------------------------------------------------------Swing-------------------------------------------------------------------//

  if (SheetName === SWING_4 || SheetName === SWING_12 || SheetName === SWING_52)
  {
    Edit = getConfigValue(DTR)                                                     // DTR = Edit to Swing

    var Class = sheet_co.getRange(IST).getDisplayValue();                       // IST = Is Stock? 

    var C2 = sheet_sr.getRange('C2').getValue();  

    if (Class == 'STOCK') 
    {
      if( C2 > 0 )
      {
        processEditSheet(sheet_sr, SheetName, Edit) 
      }
    }
    if (Class == 'BDR' || Class == 'ETF' || Class == 'ADR') 
    {
      if( C2 > 0 )
      {
        processEditSheet(sheet_sr, SheetName, Edit) 
      }
      else
      {
        console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
      }
    }
  }

//-------------------------------------------------------------------Opções-------------------------------------------------------------------//

  if (SheetName === OPCOES) 
  {
    Edit = getConfigValue(DOP)                                                     // DOP = Edit to Option

    var Call = sheet_sr.getRange('C2').getValue();
    var Put = sheet_sr.getRange('E2').getValue();  

    if( ( Call != 0 && Put != 0 ) &&
        ( Call != "" && Put != "" ) )
    { 
      processEditSheet(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
 
  if (SheetName === BTC)
  {
    Edit = getConfigValue(DBT)                                                     // DBT = Edit to BTC

    var D2 = sheet_sr.getRange('D2').getValue();

    if( !ErrorValues.includes(D2) )
    {
      processEditSheet(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

//-------------------------------------------------------------------Termo-------------------------------------------------------------------//
 
  if (SheetName === TERMO)
  {
    Edit = getConfigValue(DTE)                                                     // DTE = Edit to Termo

    var D2 = sheet_sr.getRange('D2').getValue();

    if( !ErrorValues.includes(D2) )
    {
      processEditSheet(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

//-------------------------------------------------------------------Fund-------------------------------------------------------------------//
 
  if (SheetName === FUND) //those can be merged
  {
    Edit = getConfigValue(DFU)                                                     // DFU = Edit to Fund

    var B2 = sheet_sr.getRange('B2').getValue();

    if( !ErrorValues.includes(B2) )
    {
      processEditSheet(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

//-------------------------------------------------------------------Future-------------------------------------------------------------------//
 
  if (SheetName === FUTURE)
  {
    Edit = getConfigValue(DFT)                                                     // DFT = Edit to Future

    var C2 = sheet_sr.getRange('C2').getValue();
    var E2 = sheet_sr.getRange('E2').getValue();
    var G2 = sheet_sr.getRange('G2').getValue();

    if( ( !ErrorValues.includes(C2) || !ErrorValues.includes(E2) || !ErrorValues.includes(G2) ) )
    {
      processEditSheet(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

  if (SheetName === FUTURE_1 || SheetName === FUTURE_2 || SheetName === FUTURE_3)
  {
    Edit = getConfigValue(DFT)                                                     // DFT = Edit to Future

    var C2 = sheet_sr.getRange('C2').getValue();

    if( !ErrorValues.includes(C2) )
    {
      processEditExtra(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

//-------------------------------------------------------------------Right-------------------------------------------------------------------//

  if (SheetName === RIGHT_1 || SheetName === RIGHT_2)
  {
    Edit = getConfigValue(DRT)                                                     // DRT = Edit to Right

    var D2 = sheet_sr.getRange('D2').getValue();

    if( !ErrorValues.includes(D2) )
    {
      processEditExtra(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

//-------------------------------------------------------------------Receipt-------------------------------------------------------------------//

  if (SheetName === RECEIPT_9 || SheetName === RECEIPT_10)
  {
    Edit = getConfigValue(DRC)                                                     // DRC = Edit to Receipt

    var D2 = sheet_sr.getRange('D2').getValue();

    if( !ErrorValues.includes(D2) )
    {
      processEditExtra(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

//-------------------------------------------------------------------Warrant-------------------------------------------------------------------//

  if (SheetName === WARRANT_11 || SheetName === WARRANT_12 || SheetName === WARRANT_13)
  {
    Edit = getConfigValue(DWT)                                                     // DWT = Edit to Warrant

    var D2 = sheet_sr.getRange('D2').getValue();

    if( !ErrorValues.includes(D2) )
    {
      processEditExtra(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }

//-------------------------------------------------------------------Block-------------------------------------------------------------------//

  if (SheetName === BLOCK)
  {
    Edit = getConfigValue(DBK)                                                     // DBK = Edit to Block

    var D2 = sheet_sr.getRange('D2').getValue();

    if( !ErrorValues.includes(D2) )
    {
      processEditExtra(sheet_sr, SheetName, Edit) 
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditSheet');
    }
  }
}

/////////////////////////////////////////////////////////////////////DATA TEMPLATE/////////////////////////////////////////////////////////////////////

// sheet_sr is checked  inside the blocks

function doEditData(SheetName) 
{
  console.log('EDIT:', SheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = getSheetnameByName('Config');                         // Config sheet
  const sheet_se = getSheetnameByName('Settings');
  if (!sheet_co || !sheet_se) return;

  let Edit;

//-------------------------------------------------------------------BLC-------------------------------------------------------------------//

  if (SheetName === BLC) 
  {
    Edit = getConfigValue(DBL)                                                     // DBL = Edit to BLC

    const sheet_tr = getSheetnameByName(BLC);
    if (!sheet_tr) return;

      var Values_tr = sheet_tr.getRange('B1:C1').getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

    const sheet_sr = getSheetnameByName(Balanco);
    if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange('B1:C1').getValues()[0];

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
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditData');
    }
  }

//-------------------------------------------------------------------Balanço-------------------------------------------------------------------//

  else if (SheetName === Balanco) 
  {
    Edit = getConfigValue(DBL)                                                     // DBL = Edit to BLC

    const sheet_sr = getSheetnameByName(Balanco);
    if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange('B1:C1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S, B27_S] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
        ( B2_S != 0 && B2_S != "" ) && 
        ( B27_S != 0 && B27_S != "" ) )
    {
      processEditData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Edit); // Omit New and Old here
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditData');
    }
  }

//-------------------------------------------------------------------DRE-------------------------------------------------------------------//

  else if (SheetName === DRE) 
  {
    Edit = getConfigValue(DDE)                                                     // DDE = Edit to DRE

    const sheet_tr = getSheetnameByName(DRE);
    if (!sheet_tr) return;

      var Values_tr = sheet_tr.getRange('B1:C1').getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

    const sheet_sr = getSheetnameByName(Resultado);
    if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
 
      var [B4_S, B27_S] = ["B4", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
        ( B4_S != "" ) && 
        ( B27_S != 0 && B27_S != "" ) )
    {
      processEditData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Edit);

      doEditData(Resultado);
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditData');
    }
  }

//-------------------------------------------------------------------Resultado-------------------------------------------------------------------//

  else if (SheetName === Resultado) 
  {
    Edit = getConfigValue(DDE)                                                     // DDE = Edit to DRE

    const sheet_sr = getSheetnameByName(Resultado);
    if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
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
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditData');
    }
  }

//-------------------------------------------------------------------FLC-------------------------------------------------------------------//

  else if (SheetName === FLC) 
  {
    Edit = getConfigValue(DFL)                                                     // DFL = Edit to FLC

    const sheet_tr = getSheetnameByName(FLC);
    if (!sheet_tr) return;

      var Values_tr = sheet_tr.getRange('B1:C1').getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

    const sheet_sr = getSheetnameByName(Fluxo);
    if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
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
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditData');
    }
  }

//-------------------------------------------------------------------Fluxo-------------------------------------------------------------------//

  else if (SheetName === Fluxo) 
  {
    Edit = getConfigValue(DFL)                                                     // DFL = Edit to FLC

    const sheet_sr = getSheetnameByName(Fluxo);
    if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
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
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditData');
    }
  }

//-------------------------------------------------------------------DVA-------------------------------------------------------------------//

  else if (SheetName === DVA) 
  {
    Edit = getConfigValue(DDV)                                                     // DDV = Edit to DVA

    const sheet_tr = getSheetnameByName(DVA);
    if (!sheet_tr) return;

      var LR = sheet_tr.getLastRow();
      var LC = sheet_tr.getLastColumn();

      var B = sheet_tr.getRange("B1:B" + LR).getValues().flat();

      var Values_tr = sheet_tr.getRange('B1:C1').getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

    const sheet_sr = getSheetnameByName(Valor);
    if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S] = ["B2"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
        ( B2_S != "" ) )
    {
      processEditData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Edit);

      doEditData(Valor);
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditData');
    }
  }

//-------------------------------------------------------------------Valor-------------------------------------------------------------------//

  else if (SheetName === Valor) 
  {
    Edit = getConfigValue(DDV)                                                     // DDV = Edit to DVA

    const sheet_sr = getSheetnameByName(Valor);
    if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
 
      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) && 
        ( C2_S != "" ) )
    {
      processEditData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Edit);
    }
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Conditions arent met on doEditData');
    }
  }
}

/////////////////////////////////////////////////////////////////////EDIT/////////////////////////////////////////////////////////////////////