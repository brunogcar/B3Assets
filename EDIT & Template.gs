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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = ss.getSheetByName(SheetName);
  const sheet_co = ss.getSheetByName('Config');                                        // Config Sheet
  const sheet_se = ss.getSheetByName('Settings');

  Utilities.sleep(2500); // 2,5 secs

  console.log('EDIT:', SheetName);

  let Edit;
  let Value_se = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback

  if (sheet_sr)
  {

//-------------------------------------------------------------------Swing-------------------------------------------------------------------//

    if (SheetName === SWING_4 || SheetName === SWING_12 || SheetName === SWING_52)
    {
      Value_se = sheet_se.getRange(DTR).getDisplayValue().trim();                 // DTR = Edit to Swing

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DTR).getDisplayValue().trim()                         // Use Config value if Settings has "DEFAULT"
        : Value_se;                                                               // Use Settings value

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
      Value_se = sheet_se.getRange(DOP).getDisplayValue().trim();                 // DOP = Edit to Option 

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DOP).getDisplayValue().trim()
        : Value_se;     

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
      Value_se = sheet_se.getRange(DBT).getDisplayValue().trim();                 // DBT = Edit to BTC

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DBT).getDisplayValue().trim()
        : Value_se; 

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
      Value_se = sheet_se.getRange(DTE).getDisplayValue().trim();                 // DTE = Edit to Termo

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DTE).getDisplayValue().trim()
        : Value_se; 

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
      Value_se = sheet_se.getRange(DFU).getDisplayValue().trim();                 // DFU = Edit to Fund

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DFU).getDisplayValue().trim()
        : Value_se; 

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
      Value_se = sheet_se.getRange(DFT).getDisplayValue().trim();                 // DFT = Edit to Future

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DFT).getDisplayValue().trim()
        : Value_se; 

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
      Value_se = sheet_se.getRange(DFT).getDisplayValue().trim();                 // DFT = Edit to Future

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DFT).getDisplayValue().trim()
        : Value_se; 

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
      Value_se = sheet_se.getRange(DRT).getDisplayValue().trim();                 // DTE = Edit to Right

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DRT).getDisplayValue().trim()
        : Value_se; 

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
      Value_se = sheet_se.getRange(DRC).getDisplayValue().trim();                 // DRC = Edit to Receipt

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DRC).getDisplayValue().trim()
        : Value_se; 

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
      Value_se = sheet_se.getRange(DWT).getDisplayValue().trim();                 // DWT = Edit to Warrant

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DWT).getDisplayValue().trim()
        : Value_se; 

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
      Value_se = sheet_se.getRange(DBK).getDisplayValue().trim();                 // DBK = Edit to Block

      Edit = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(DBK).getDisplayValue().trim()
        : Value_se; 

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
  else
  {
    console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditSheet');
  }
}

/////////////////////////////////////////////////////////////////////DATA TEMPLATE/////////////////////////////////////////////////////////////////////

// sheet_sr is checked  inside the blocks

function doEditData(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet_co = ss.getSheetByName('Config');                                        // sheet_co = config spreadsheet
  const sheet_se = ss.getSheetByName('Settings');

  console.log('EDIT:', SheetName);

  let Edit;
  let Value_se = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback

//-------------------------------------------------------------------BLC-------------------------------------------------------------------//

  if (SheetName === BLC) 
  {
    Value_se = sheet_se.getRange(DBL).getDisplayValue().trim();                 // DBL = Edit to BLC

    Edit = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(DBL).getDisplayValue().trim()                         // Use Config value if Settings has "DEFAULT"
      : Value_se;                                                               // Use Settings value

    const sheet_tr = ss.getSheetByName(BLC);

      var Values_tr = sheet_tr.getRange('B1:C1').getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

    const sheet_sr = ss.getSheetByName(Balanco);

      var Values_sr = sheet_sr.getRange('B1:C1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S, B27_S] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
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
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditData');
    }
  }

//-------------------------------------------------------------------Balanço-------------------------------------------------------------------//

  else if (SheetName === Balanco) 
  {
    Value_se = sheet_se.getRange(DBL).getDisplayValue().trim();                 // DBL = Edit to BLC

    Edit = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(DBL).getDisplayValue().trim()
      : Value_se;

    const sheet_sr = ss.getSheetByName(Balanco);

      var Values_sr = sheet_sr.getRange('B1:C1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S, B27_S] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
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
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditData');
    }
  }

//-------------------------------------------------------------------DRE-------------------------------------------------------------------//

  else if (SheetName === DRE) 
  {
    Value_se = sheet_se.getRange(DDE).getDisplayValue().trim();                 // DDE = Edit to DRE

    Edit = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(DDE).getDisplayValue().trim()
      : Value_se;

    const sheet_tr = ss.getSheetByName(DRE);

      var Values_tr = sheet_tr.getRange('B1:C1').getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

    const sheet_sr = ss.getSheetByName(Resultado);

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
 
      var [B4_S, B27_S] = ["B4", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
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
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditData');
    }
  }

//-------------------------------------------------------------------Resultado-------------------------------------------------------------------//

  else if (SheetName === Resultado) 
  {
    Value_se = sheet_se.getRange(DDE).getDisplayValue().trim();                 // DDE = Edit to DRE

    Edit = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(DDE).getDisplayValue().trim()
      : Value_se;

    const sheet_sr = ss.getSheetByName(Resultado);

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
 
      var [C4_S, C27_S] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
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
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditData');
    }
  }

//-------------------------------------------------------------------FLC-------------------------------------------------------------------//

  else if (SheetName === FLC) 
  {
    Value_se = sheet_se.getRange(DFL).getDisplayValue().trim();                 // DFL = Edit to FLC

    Edit = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(DFL).getDisplayValue().trim()
      : Value_se;

    const sheet_tr = ss.getSheetByName(FLC);

      var Values_tr = sheet_tr.getRange('B1:C1').getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

    const sheet_sr = ss.getSheetByName(Fluxo);

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S] = ["B2"].map(r => sheet_sr.getRange(r).getDisplayValue());


    if (sheet_sr)
    {
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
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditData');
    }
  }

//-------------------------------------------------------------------Fluxo-------------------------------------------------------------------//

  else if (SheetName === Fluxo) 
  {
    Value_se = sheet_se.getRange(DFL).getDisplayValue().trim();                 // DFL = Edit to FLC

    Edit = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(DFL).getDisplayValue().trim()
      : Value_se;

    const sheet_sr = ss.getSheetByName(Fluxo);

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S] = ["B2"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
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
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditData');
    }
  }

//-------------------------------------------------------------------DVA-------------------------------------------------------------------//

  else if (SheetName === DVA) 
  {
    Value_se = sheet_se.getRange(DDV).getDisplayValue().trim();                 // DDV = Edit to DVA

    Edit = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(DDV).getDisplayValue().trim()
      : Value_se;

    const sheet_tr = ss.getSheetByName(DVA);

      var LR = sheet_tr.getLastRow();
      var LC = sheet_tr.getLastColumn();

      var B = sheet_tr.getRange("B1:B" + LR).getValues().flat();

      var Values_tr = sheet_tr.getRange('B1:C1').getValues()[0];

      var [[new_T_D, new_T_M, new_T_Y], [old_T_D, old_T_M, old_T_Y]] = Values_tr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_T = new_T_D && new_T_M && new_T_Y ? new Date(new_T_Y, new_T_M - 1, new_T_D).getTime() : "";
      var Old_T = old_T_D && old_T_M && old_T_Y ? new Date(old_T_Y, old_T_M - 1, old_T_D).getTime() : "";

    const sheet_sr = ss.getSheetByName(Valor);

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [B2_S] = ["B2"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
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
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditData');
    }
  }

//-------------------------------------------------------------------Valor-------------------------------------------------------------------//

  else if (SheetName === Valor) 
  {
    Value_se = sheet_se.getRange(DDV).getDisplayValue().trim();                 // DDV = Edit to DVA

    Edit = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(DDV).getDisplayValue().trim()
      : Value_se;

    const sheet_sr = ss.getSheetByName(Valor);

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";
 
      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
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
    else
    {
      console.log('ERROR EDIT:', SheetName, 'Does not exist on doEditData');
    }
  }
}

/////////////////////////////////////////////////////////////////////EDIT/////////////////////////////////////////////////////////////////////