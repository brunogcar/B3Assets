/////////////////////////////////////////////////////////////////////SAVE TEMPLATE/////////////////////////////////////////////////////////////////////

function doSaveSheet(SheetName)
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = ss.getSheetByName(SheetName);                        // Source sheet
  const sheet_co = ss.getSheetByName('Config');                         // Config sheet
  const sheet_se = ss.getSheetByName('Settings');

  Utilities.sleep(2500); // 2,5 secs

  console.log('SAVE:', SheetName);

  let Save;
  let Value_se_sa = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback
  let Edit;
  let Value_se_ed = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback

  if (sheet_sr)
  {
//-------------------------------------------------------------------Swing-------------------------------------------------------------------//

    if (SheetName === SWING_4 || SheetName === SWING_12 || SheetName === SWING_52)
    {
      Value_se_sa = sheet_se.getRange(STR).getDisplayValue().trim();                 // STR = Save to Swing
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(STR).getDisplayValue().trim()                            // Use Config value if Settings has "DEFAULT"
        : Value_se_sa;

      Value_se_ed = sheet_se.getRange(DTR).getDisplayValue().trim();                 // DTR = Edit to Swing
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DTR).getDisplayValue().trim()                            // Use Config value if Settings has "DEFAULT"
        : Value_se_ed;

      var Class = sheet_co.getRange(IST).getDisplayValue();                          // IST = Is Stock? 

      var B2 = sheet_sr.getRange('B2').getValue();
      var C2 = sheet_sr.getRange('C2').getValue();

      if (Class == 'STOCK') 
      {
        if( B2 != 0 && C2 > 0 )
        {
          processSaveSheet(sheet_sr, SheetName, Save, Edit);
          doTrimSheet(SheetName);
        }
        else
        {
          console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
        }
      }
      if (Class == 'BDR' || Class == 'ETF' || Class == 'ADR') 
      {
        if( C2 > 0 )
        {
          processSaveSheet(sheet_sr, SheetName, Save, Edit);
          doTrimSheet(SheetName);
        }
        else
        {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
        }
      }
    }

//-------------------------------------------------------------------Opções-------------------------------------------------------------------//

     if (SheetName === OPCOES)
    {
      Value_se_sa = sheet_se.getRange(SOP).getDisplayValue().trim();                 // SOP = Save to Option
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SOP).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DOP).getDisplayValue().trim();                 // DOP = Edit to Option
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DOP).getDisplayValue().trim()
        : Value_se_ed;   

      var [Call, Call_, Call_PM, Call_PM_,Put, Put_, Put_PM, Put_PM_,Diff,Diff_2] = ["C2", "C3", "D2", "D3","E2", "E3", "F2", "F3",'K3','N3'].map(r => sheet_sr.getRange(r).getValue());

      if( ( Call != 0 && Put != 0 ) &&
          ( Call != "" && Put != "" ) &&
          ( Call_PM != 0 || Put_PM != 0 ) &&
          ( Call_PM != "" || Put_PM != "" ) &&
          ( Diff != 0 || Diff_2 != 0 ) )
      {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

//-------------------------------------------------------------------BTC-------------------------------------------------------------------//

    if (SheetName === BTC)
    {
      Value_se_sa = sheet_se.getRange(SBT).getDisplayValue().trim();                 // SBT = Save to BTC
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SBT).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DBT).getDisplayValue().trim();                 // DBT = Edit to BTC
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DBT).getDisplayValue().trim()
        : Value_se_ed;   

      var D2 = sheet_sr.getRange('D2').getValue();

      if( !ErrorValues.includes(D2) )
      {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

//-------------------------------------------------------------------Termo-------------------------------------------------------------------//

    if (SheetName === TERMO)
    {
      Value_se_sa = sheet_se.getRange(STE).getDisplayValue().trim();                 // STE = Save to Termo
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(STE).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DTE).getDisplayValue().trim();                 // DTE = Edit to Termo
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DTE).getDisplayValue().trim()
        : Value_se_ed;   

      var D2 = sheet_sr.getRange('D2').getValue();

      if( !ErrorValues.includes(D2) )
      {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

//-------------------------------------------------------------------Fund-------------------------------------------------------------------//

    if (SheetName === FUND)
    {
      Value_se_sa = sheet_se.getRange(SFU).getDisplayValue().trim();                 // SFU = Save to Fund
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SFU).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DFU).getDisplayValue().trim();                 // DFU = Edit to Fund
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DFU).getDisplayValue().trim()
        : Value_se_ed;   

      var B2 = sheet_sr.getRange('B2').getValue();

      if( ( !ErrorValues.includes(B2) ) )
      {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

//-------------------------------------------------------------------FUTURE-------------------------------------------------------------------//

    if (SheetName === FUTURE)
    {
      Value_se_sa = sheet_se.getRange(SFT).getDisplayValue().trim();                 // SFT = Save to Future
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SFT).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DFT).getDisplayValue().trim();                 // DFT = Edit to Future
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DFT).getDisplayValue().trim()
        : Value_se_ed;   

      var C2 = sheet_sr.getRange('C2').getValue();
      var E2 = sheet_sr.getRange('E2').getValue();
      var G2 = sheet_sr.getRange('G2').getValue();

      if( ( !ErrorValues.includes(C2) || !ErrorValues.includes(E2) || !ErrorValues.includes(G2) ) )  // possibly && instead of ||
      {
        processSaveSheet(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

    if (SheetName === FUTURE_1 || SheetName === FUTURE_2 || SheetName === FUTURE_3)
    {
      Value_se_sa = sheet_se.getRange(SFT).getDisplayValue().trim();                 // SFT = Save to Future
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SFT).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DFT).getDisplayValue().trim();                 // DFT = Edit to Future
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DFT).getDisplayValue().trim()
        : Value_se_ed; 

      var C2 = sheet_sr.getRange('C2').getValue();

      if( !ErrorValues.includes(C2) )
      {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

//-------------------------------------------------------------------Right-------------------------------------------------------------------//

    if (SheetName === RIGHT_1 || SheetName === RIGHT_2)
    {
      Value_se_sa = sheet_se.getRange(SRT).getDisplayValue().trim();                 // SRT = Save to Right
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SRT).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DRT).getDisplayValue().trim();                 // DTE = Edit to Right
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DRT).getDisplayValue().trim()
        : Value_se_ed;   

      var D2 = sheet_sr.getRange('D2').getValue();

      if( !ErrorValues.includes(D2) )
      {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

//-------------------------------------------------------------------Receipt-------------------------------------------------------------------//

    if (SheetName === RECEIPT_9 || SheetName === RECEIPT_10)
    {
      Value_se_sa = sheet_se.getRange(SRC).getDisplayValue().trim();                 // SRC = Save to Receipt
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SRC).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DRC).getDisplayValue().trim();                 // DRC = Edit to Receipt
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DRC).getDisplayValue().trim()
        : Value_se_ed;   

      var D2 = sheet_sr.getRange('D2').getValue();

      if( !ErrorValues.includes(D2) )
      {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

//-------------------------------------------------------------------Warrant-------------------------------------------------------------------//

    if (SheetName === WARRANT_11 || SheetName === WARRANT_12 || SheetName === WARRANT_13)
    {
      Value_se_sa = sheet_se.getRange(SWT).getDisplayValue().trim();                 // SWT = Save to Warrant
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SWT).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DWT).getDisplayValue().trim();                 // DWT = Edit to Warrant
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DWT).getDisplayValue().trim()
        : Value_se_ed;   

      var D2 = sheet_sr.getRange('D2').getValue();

      if( !ErrorValues.includes(D2) )
      {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }

//-------------------------------------------------------------------Block-------------------------------------------------------------------//

    if (SheetName === BLOCK)
    {
      Value_se_sa = sheet_se.getRange(SBK).getDisplayValue().trim();                 // SBK = Save to Block
      Save = (Value_se_sa === "DEFAULT") 
        ? sheet_co.getRange(SBK).getDisplayValue().trim()
        : Value_se_sa;   

      Value_se_ed = sheet_se.getRange(DBK).getDisplayValue().trim();                 // DBK = Edit to Block
      Edit = (Value_se_ed === "DEFAULT") 
        ? sheet_co.getRange(DBK).getDisplayValue().trim()
        : Value_se_ed;   

      var D2 = sheet_sr.getRange('D2').getValue();

      if( !ErrorValues.includes(D2) )
      {
        processSaveExtra(sheet_sr, SheetName, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveSheet');
      }
    }
  }
  else
  {
    console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveSheet');
  }
}

/////////////////////////////////////////////////////////////////////DATA TEMPLATE/////////////////////////////////////////////////////////////////////

// sheet_sr is checked  inside the blocks

function doSaveData(SheetName)
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config');
  const sheet_se = ss.getSheetByName('Settings');
  const sheet_up = ss.getSheetByName('UPDATE');

  console.log('SAVE:', SheetName);

  let Save;
  let Value_se_sa = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback
  let Edit;
  let Value_se_ed = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback

//-------------------------------------------------------------------BLC-------------------------------------------------------------------//

  if (SheetName === BLC)
  {
    Value_se_sa = sheet_se.getRange(SBL).getDisplayValue().trim();                 // SBL = Save to BLC
    Save = (Value_se_sa === "DEFAULT") 
      ? sheet_co.getRange(SBL).getDisplayValue().trim()                            // Use Config value if Settings has "DEFAULT"
      : Value_se_sa;

    Value_se_ed = sheet_se.getRange(DBL).getDisplayValue().trim();                 // DBL = Edit to BLC
    Edit = (Value_se_ed === "DEFAULT") 
      ? sheet_co.getRange(DBL).getDisplayValue().trim()                            // Use Config value if Settings has "DEFAULT"
      : Value_se_ed;

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

    var CHECK1 = sheet_up.getRange('K3').getValue();
    var CHECK2 = sheet_up.getRange('K4').getValue();

    if (sheet_sr)
    {
      if (((CHECK1 >= 90 && CHECK1 <= 92) || (CHECK1 == 0 || CHECK1 > 40000)) && 
          ((CHECK2 >= 90 && CHECK2 <= 92) || (CHECK2 == 0 || CHECK1 > 40000)))
      {
        if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) &&
            ( B2_S != 0 && B2_S != "" ) &&
            ( B27_S != 0 && B27_S != "" ) )
        {
          processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit);

          doSaveData(Balanco);
        }
        else
        {
          console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveData');
        }
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveData');
      }

    }
    else
    {
      console.log('ERROR SAVE:', SheetName, CHECK1, CHECK2, 'Are not in CHECK Range on doSaveData');
    }
  }

//-------------------------------------------------------------------Balanço-------------------------------------------------------------------//

  else if (SheetName === Balanco)
  {
    Value_se_sa = sheet_se.getRange(SBL).getDisplayValue().trim();                 // SBL = Save to BLC
    Save = (Value_se_sa === "DEFAULT") 
      ? sheet_co.getRange(SBL).getDisplayValue().trim()
      : Value_se_sa;

    Value_se_ed = sheet_se.getRange(DBL).getDisplayValue().trim();                 // DBL = Edit to BLC
    Edit = (Value_se_ed === "DEFAULT") 
      ? sheet_co.getRange(DBL).getDisplayValue().trim()
      : Value_se_ed;

    const sheet_sr = ss.getSheetByName(Balanco);

      var Values_sr = sheet_sr.getRange('B1:C1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C4_S, C27_S] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) &&
          ( C4_S != 0 && C4_S != "" ) &&
          ( C27_S != 0 && C27_S != "" ) )
      {
        processSaveData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveData');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveData');
    }
  }

//-------------------------------------------------------------------DRE-------------------------------------------------------------------//

  else if (SheetName === DRE)
  {
    Value_se_sa = sheet_se.getRange(SDE).getDisplayValue().trim();                 // SDE = Save to DRE
    Save = (Value_se_sa === "DEFAULT") 
      ? sheet_co.getRange(SDE).getDisplayValue().trim()
      : Value_se_sa;

    Value_se_ed = sheet_se.getRange(DDE).getDisplayValue().trim();                 // DDE = Edit to DRE
    Edit = (Value_se_ed === "DEFAULT") 
      ? sheet_co.getRange(DDE).getDisplayValue().trim()
      : Value_se_ed;

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

    var [C4_S, C27_S] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());

    var CHECK = sheet_up.getRange('K5').getValue();

    if (sheet_sr)
    {
      if ((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000))
      {
        if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) &&
          ( C4_S != 0 && C4_S != "" ) &&
          ( C27_S != 0 && C27_S != "" ) )
        {
          processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit);

          doSaveData(Resultado);
        }
        else
        {
          console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveData');
        }
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveData');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, CHECK, 'Is not in CHECK Range on doSaveData');
    }
  }

//-------------------------------------------------------------------Resultado-------------------------------------------------------------------//

  else if (SheetName === Resultado)
  {
    Value_se_sa = sheet_se.getRange(SDE).getDisplayValue().trim();                 // SDE = Save to DRE
    Save = (Value_se_sa === "DEFAULT") 
      ? sheet_co.getRange(SDE).getDisplayValue().trim()
      : Value_se_sa;

    Value_se_ed = sheet_se.getRange(DDE).getDisplayValue().trim();                 // DDE = Edit to DRE
    Edit = (Value_se_ed === "DEFAULT") 
      ? sheet_co.getRange(DDE).getDisplayValue().trim()
      : Value_se_ed;

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
        processSaveData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveData');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveData');
    }
  }

//-------------------------------------------------------------------FLC-------------------------------------------------------------------//

  else if (SheetName === FLC)
  {
    Value_se_sa = sheet_se.getRange(SFL).getDisplayValue().trim();                 // SFL = Save to FLC
    Save = (Value_se_sa === "DEFAULT") 
      ? sheet_co.getRange(SFL).getDisplayValue().trim()
      : Value_se_sa;

    Value_se_ed = sheet_se.getRange(DFL).getDisplayValue().trim();                 // DFL = Edit to FLC
    Edit = (Value_se_ed === "DEFAULT") 
      ? sheet_co.getRange(DFL).getDisplayValue().trim()
      : Value_se_ed;

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

      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

    var CHECK = sheet_up.getRange('K6').getValue();

    if (sheet_sr)
    {
      if ((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000))
      {
        if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) &&
          ( C2_S != 0 && C2_S !== "" ) )
        {
          processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit);

          doSaveData(Fluxo);
        }
        else
        {
          console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveData');
        }
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveData');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, CHECK, 'Is not in CHECK Range on doSaveData');
    }
  }

//-------------------------------------------------------------------Fluxo-------------------------------------------------------------------//

  else if (SheetName === Fluxo)
  {
    Value_se_sa = sheet_se.getRange(SFL).getDisplayValue().trim();                 // SFL = Save to FLC
    Save = (Value_se_sa === "DEFAULT") 
      ? sheet_co.getRange(SFL).getDisplayValue().trim()
      : Value_se_sa;

    Value_se_ed = sheet_se.getRange(DFL).getDisplayValue().trim();                 // DFL = Edit to FLC
    Edit = (Value_se_ed === "DEFAULT") 
      ? sheet_co.getRange(DFL).getDisplayValue().trim()
      : Value_se_ed;

    const sheet_sr = ss.getSheetByName(Fluxo);

      var Values_sr = sheet_sr.getRange('B1:D1').getValues()[0];

      var [[new_S_D, new_S_M, new_S_Y], [temp_S_D, temp_S_M, temp_S_Y], [old_S_D, old_S_M, old_S_Y]] = Values_sr.map(v => v ? v.split("/") : Array(3).fill(""));
      var New_S = new_S_D && new_S_M && new_S_Y ? new Date(new_S_Y, new_S_M - 1, new_S_D).getTime() : "";
      var temp_S = temp_S_D && temp_S_M && temp_S_Y ? new Date(temp_S_Y, temp_S_M - 1, temp_S_D).getTime() : "";
      var Old_S = old_S_D && old_S_M && old_S_Y ? new Date(old_S_Y, old_S_M - 1, old_S_D).getTime() : "";

      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

    if (sheet_sr)
    {
      if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) &&
          ( C2_S != 0 && C2_S !== "" ) )
      {
        processSaveData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveData');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveData');
    }
  }

//-------------------------------------------------------------------DVA-------------------------------------------------------------------//

  else if (SheetName === DVA)
  {
    Value_se_sa = sheet_se.getRange(SDV).getDisplayValue().trim();                 // SDV = Save to DVA
    Save = (Value_se_sa === "DEFAULT") 
      ? sheet_co.getRange(SDV).getDisplayValue().trim()
      : Value_se_sa;

    Value_se_ed = sheet_se.getRange(DDV).getDisplayValue().trim();                 // DDV = Edit to DVA
    Edit = (Value_se_ed === "DEFAULT") 
      ? sheet_co.getRange(DDV).getDisplayValue().trim()
      : Value_se_ed;

    const sheet_tr = ss.getSheetByName(DVA);

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

      var [C2_S] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

    var CHECK = sheet_up.getRange('K7').getValue();

    if (sheet_sr)
    {
      if ((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000))
      {
        if( ( New_S.valueOf() != "-" && New_S.valueOf() != "" ) &&
            ( C2_S != 0 && C2_S !== "" ) )
        {
          processSaveData(sheet_tr, sheet_sr, New_T, Old_T, New_S, Old_S, Save, Edit);

          doSaveData(Valor);
        }
        else
        {
          console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveData');
        }
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveData');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, CHECK, 'Is not in CHECK Range on doSaveData');
    }
  }

//-------------------------------------------------------------------Valor-------------------------------------------------------------------//

  else if (SheetName === Valor)
  {
    Value_se_sa = sheet_se.getRange(SDV).getDisplayValue().trim();                 // SDV = Save to DVA
    Save = (Value_se_sa === "DEFAULT") 
      ? sheet_co.getRange(SDV).getDisplayValue().trim()
      : Value_se_sa;

    Value_se_ed = sheet_se.getRange(DDV).getDisplayValue().trim();                 // DDV = Edit to DVA
    Edit = (Value_se_ed === "DEFAULT") 
      ? sheet_co.getRange(DDV).getDisplayValue().trim()
      : Value_se_ed;

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
          ( C2_S != 0 && C2_S !== "" ) )
      {
        processSaveData(sheet_sr, sheet_sr, '', '', New_S, Old_S, Save, Edit);
      }
      else
      {
        console.log('ERROR SAVE:', SheetName, 'Conditions arent met on doSaveData');
      }
    }
    else
    {
      console.log('ERROR SAVE:', SheetName, 'Does not exist on doSaveData');
    }
  }
}

/////////////////////////////////////////////////////////////////////OTHER/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////PROVENTOS TEMPLATE/////////////////////////////////////////////////////////////////////

function doProventos() 
{
  doCheckDATA(PROV);
  doGetProventos();
  doSaveProventos();
}

function doSaveProventos() 
{
  const ProvNames = [
    { name: 'Proventos', checkCell: 'B3', expectedValue: 'Proventos', sourceRange: 'B3:H60', targetRange: 'B3:H60' },
    { name: 'Subscrição', checkCell: 'L3', expectedValue: 'Tipo', sourceRange: 'L3:T60', targetRange: 'L3:T60' },
    { name: 'Ativos', checkCell: 'B64', expectedValue: 'Proventos', sourceRange: 'B64:H200', targetRange: 'B64:H200' },
    { name: 'Historico', checkCell: 'L64', expectedValue: 'Tipo de Ativo', dynamicRange: true }
  ];

  ProvNames.forEach(config => {
    try {
      doSaveProv(config);
    } catch (error) {
      console.error(`Error saving ${config.name}:`, error);
    }
  });
}

function doSaveProv(config) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = ss.getSheetByName('Prov_');  // Source Sheet

  if (!sheet_sr) 
  {
    console.log('ERROR: Target sheet "Prov_" does not exist. Skipping operation.');
    return; // Exit the function
  }

  const sheet_tr = ss.getSheetByName('Prov');   // Target Sheet

  if (!sheet_tr) 
  {
    console.log('ERROR: Target sheet "Prov" does not exist. Skipping operation.');
    return; // Exit the function
  }
  const checkValue = sheet_sr.getRange(config.checkCell).getDisplayValue().trim();
  
  if (checkValue === config.expectedValue) 
  {
    let data;
    
    if (config.dynamicRange) 
    {
      const lr = sheet_sr.getLastRow();
      const lc = sheet_sr.getLastColumn();
      const sourceRange = sheet_sr.getRange(64, 12, lr - 63, lc - 11);
      const targetRange = sheet_tr.getRange(64, 12, lr - 63, lc - 11);
      
      data = sourceRange.getValues();
      targetRange.clearContent(); // Clear target range before writing data
      targetRange.setValues(data);
    } 
    else 
    {
      const sourceRange = sheet_sr.getRange(config.sourceRange);
      const targetRange = sheet_tr.getRange(config.targetRange);
      
      data = sourceRange.getValues();
      targetRange.clearContent(); // Clear target range before writing data
      targetRange.setValues(data);
    }
    
    console.log(`SUCCESS SAVE: ${config.name}.`);
  } 
  else 
  {
    console.log(`ERROR SAVE: ${config.name}, ${config.checkCell} != ${config.expectedValue}`);
  }
}

function doGetProventos() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config'); // Config sheet
  const sheet_tr = ss.getSheetByName('Prov_');

  if (!sheet_tr) 
  {
    console.log('ERROR: Target sheet "Prov_" does not exist. Skipping operation.');
    return; // Exit the function
  }

  const TKT = sheet_co.getRange(TKR).getValue(); // TKR = Ticket Range
  const ticker = TKT.substring(0, 4);
  const language = 'pt-br';

  const data = JSON.stringify({
    issuingCompany: ticker,
    language: language
  });

  const base64Params = Utilities.base64Encode(data);

  const url = `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/${base64Params}`;
  console.log("URL:", url);

  let responseText;
  try 
  {
    const response = UrlFetchApp.fetch(url);
    responseText = response.getContentText().trim();
    console.log("API Response:", responseText);
  } catch (error) {
    console.log("ERROR: Failed to fetch API response.", error);
    return; // Exit if the API request fails
  }

  if (!responseText) 
  {
    console.log("ERROR: Empty response from API.");
    return; // Exit if the response is empty
  }

  let content;
  try 
  {
    content = JSON.parse(responseText);
  } 
  catch (error) 
  {
    console.log("ERROR: Failed to parse JSON response.", error);
    return; // Exit if parsing fails
  }

  if (!content || !content[0]) 
  {
    console.log("ERROR: No data returned from API.");
    return; // Exit if no data is available
  }

  fillCashDividends(sheet_tr, content[0]?.cashDividends || []);
  fillStockDividends(sheet_tr, content[0]?.stockDividends || []);
  fillSubscriptions(sheet_tr, content[0]?.subscriptions || []);
}

// Fill Cash Dividends from B2 to B60
function fillCashDividends(sheet_tr, dividends) 
{
  const headerRange = "B3:H3";
  const startRow = 4;
  const maxRows = 57;

  sheet_tr.getRange("B2").setValue("Proventos em Dinheiro");
  sheet_tr.getRange("B3:H60").clearContent();

  const headers = ["Proventos", "Código ISIN", "Data de Aprovação", "Última Data Com", "Valor (R$)", "Relacionado a", "Data de Pagamento"];
  sheet_tr.getRange(headerRange).setValues([headers]);

  dividends.slice(0, maxRows).forEach((div, i) => {
    sheet_tr.getRange(startRow + i, 2, 1, 7).setValues([[ 
      div.label, div.isinCode, div.approvedOn, div.lastDatePrior, div.rate, div.relatedTo, div.paymentDate
    ]]);
  });
}

// Fill Stock Dividends from row 63
function fillStockDividends(sheet_tr, stockDividends) 
{
  const startRow = 63;
  const headerRange = `B${startRow + 1}:G${startRow + 1}`;

  sheet_tr.getRange(`B${startRow}:G${startRow + stockDividends.length + 1}`).clearContent();
  sheet_tr.getRange(`B${startRow}`).setValue("Dividendos em Ações");

  const headers = ["Proventos", "Código ISIN", "Data de Aprovação", "Última Data Com", "Fator", "Ativo Emitido"];
  sheet_tr.getRange(headerRange).setValues([headers]);

  stockDividends.forEach((stockDiv, i) => {
    sheet_tr.getRange(startRow + 2 + i, 2, 1, 6).setValues([[ 
      stockDiv.label, stockDiv.isinCode, stockDiv.approvedOn, stockDiv.lastDatePrior, stockDiv.factor, stockDiv.assetIssued
    ]]);
  });
}

// Fill Subscriptions starting from column L, row 2
function fillSubscriptions(sheet_tr, subscriptions) 
{
  const headerRange = "L3:T3";
  const startRow = 4;

  sheet_tr.getRange("L2").setValue("Subscrições");
  sheet_tr.getRange("L3:T60").clearContent();

  const headers = ["Tipo", "Código ISIN", "Data de Aprovação", "Última Data Com", "Percentual (%)", "Ativo Emitido", "Preço Emissão (R$)", "Período de Negociação", "Data de Subscrição"];
  sheet_tr.getRange(headerRange).setValues([headers]);

  subscriptions.forEach((sub, i) => {
    sheet_tr.getRange(startRow + i, 12, 1, 9).setValues([[ 
      sub.label, sub.isinCode, sub.approvedOn, sub.lastDatePrior, sub.percentage, sub.assetIssued, sub.priceUnit, sub.tradingPeriod, sub.subscriptionDate
    ]]);
  });
}

/////////////////////////////////////////////////////////////////////CodeCVM/////////////////////////////////////////////////////////////////////

function doGetCodeCVM() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_tr = ss.getSheetByName('Info'); // Target sheet

  if (!sheet_tr) 
  {
    console.log('ERROR: Target sheet "Info" does not exist. Skipping operation.');
    return; // Exit the function if the target sheet does not exist
  }

  const sheet_co = ss.getSheetByName('Config'); // Config sheet
  const TKT = sheet_co.getRange(TKR).getValue(); // TKR = Ticket Range
  const ticker = TKT.substring(0, 4);
  const language = 'pt-br';

  const data = JSON.stringify({
    issuingCompany: ticker,
    language: language
  });

  const base64Params = Utilities.base64Encode(data);

  const url = `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/${base64Params}`;
  console.log("URL:", url);

  let responseText;
  try 
  {
    const response = UrlFetchApp.fetch(url);
    responseText = response.getContentText().trim();
    console.log("API Response:", responseText);
  } 
  catch (error) 
  {
    console.log("ERROR: Failed to fetch API response.", error);
    return; // Exit if the API request fails
  }

  if (!responseText) 
  {
    console.log("ERROR: Empty response from API.");
    return; // Exit if the response is empty
  }

  let content;
  try 
  {
    content = JSON.parse(responseText);
  } 
  catch (error) 
  {
    console.log("ERROR: Failed to parse JSON response.", error);
    return; // Exit if parsing fails
  }

  if (!content || !content[0]) 
  {
    console.log("ERROR: No data returned from API.");
    return; // Exit if no data is available
  }

  const codeCVM = content[0]?.codeCVM || 'N/A'; // Default to 'N/A' if codeCVM is missing
  console.log("Extracted codeCVM:", codeCVM);

  // Write to the Info sheet
  sheet_tr.getRange('C3').setValue(codeCVM); // Set data
}

/////////////////////////////////////////////////////////////////////SAVE AND SHARES TEMPLATE/////////////////////////////////////////////////////////////////////

function doSaveShares() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = ss.getSheetByName('DATA');

  var SheetName = sheet_sr.getName();

  var M1 = sheet_sr.getRange('M1').getValue();
  var M2 = sheet_sr.getRange('M2').getValue();

  console.log('SAVE: Shares and FF');

  // Validate if M1 and M2 are numeric or can be converted to numbers
  if (!isNaN(M1) && !isNaN(M2) && !ErrorValues.includes(M1) && !ErrorValues.includes(M2)) 
  {
    M1 = Number(M1); // Convert to number if not already
    M2 = Number(M2);

    if (M1 !== 0 && M2 !== 0) 
    {
      var Data = sheet_sr.getRange("M1:M2").getValues();
      sheet_sr.getRange("L1:L2").setValues(Data);

      console.log(`SUCCESS SAVE: Shares and FF`);
    }
    else
    {
       console.log('ERROR SAVE:', SheetName, 'M1 or M2 is 0');
    }
  }
  else
  {
    console.log('ERROR SAVE:', SheetName, 'Invalid values in M1 or M2');
  }
}

/////////////////////////////////////////////////////////////////////COTACOES TEMPLATE/////////////////////////////////////////////////////////////////////

function doSaveCotacoes()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_sr = ss.getSheetByName('Cotações');

  var afRange = sheet_sr.getRange("A2:F" + sheet_sr.getLastRow());
  var nsRange = sheet_sr.getRange("N2:S" + sheet_sr.getLastRow());
  var afValues = afRange.getValues();
  var nsValues = nsRange.getValues();

  var minRows = Math.min(afValues.length, nsValues.length);


   var N2 = sheet_sr.getRange("N2").getDisplayValue();

   if (!ErrorValues.includes(N2))
   {
      afRange.setValues(nsValues);
   }

  // Clear any extra rows in afRange
  if (afValues.length > nsValues.length) 
  {
    var numRowsToDelete = afValues.length - nsValues.length;
    var startRowToDelete = nsValues.length + 2;
    sheet_sr.getRange(startRowToDelete, 1, numRowsToDelete, afValues[0].length).clearContent();
  }
}

/////////////////////////////////////////////////////////////////////SAVE TEMPLATE/////////////////////////////////////////////////////////////////////