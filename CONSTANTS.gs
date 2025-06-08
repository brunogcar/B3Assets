/////////////////////////////////////////////////////////////////////ERROR/////////////////////////////////////////////////////////////////////

  const ErrorValues = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#ERROR!', 'Loading!', '', '-', null];

/////////////////////////////////////////////////////////////////////SheetNames/////////////////////////////////////////////////////////////////////

  const PROV       = 'Prov';

  const SWING_4    = 'Swing_4';
  const SWING_12   = 'Swing_12';
  const SWING_52   = 'Swing_52';

  const OPCOES     = 'Opções';
  const BTC        = 'BTC';
  const TERMO      = 'Termo';
  const FUND       = 'Fund';

  const FUTURE     = 'Future';
  const FUTURE_1   = 'FUT+1';
  const FUTURE_2   = 'FUT+2';
  const FUTURE_3   = 'FUT+3';

  const RIGHT_1    = 'DRT_1';
  const RIGHT_2    = 'DRT_2';
  const RECEIPT_9  = 'RCB_9';
  const RECEIPT_10 = 'RCB_10';
  const WARRANT_11 = 'GAR_11';
  const WARRANT_12 = 'GAR_12';
  const WARRANT_13 = 'GAR_13';
  const BLOCK      = 'BLOCK';
  const AFTER      = 'After';

  const BLC        = 'BLC';
  const Balanco    = 'Balanço';
  const DRE        = 'DRE';
  const Resultado  = 'Resultado';
  const FLC        = 'FLC';
  const Fluxo      = 'Fluxo';
  const DVA        = 'DVA';
  const Valor      = 'Valor';

  const SheetsBasic         = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, AFTER, FUND];
  const SheetsExtra         = [FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  const SheetsFinancial     = [BLC, DRE, FLC, DVA];
  const SheetsFinancialFull = [BLC, Balanco, DRE, Resultado, FLC, Fluxo, DVA, Valor];

  //“all the non-financial” sheets
  const SheetsBasicAndExtra = SheetsBasic.concat(SheetsExtra);                //just in case if i want to replace [...SheetsBasic,...SheetsExtra]

/////////////////////////////////////////////////////////////////////Ranges/////////////////////////////////////////////////////////////////////

  const TKR = "B3"     // TKR = Ticker                     // Tab: Config

  const PRV = "B3:H60" // PRV = Provento Range             // Tab: Prov

  const TIR = "C3:D4"  // TIR = Tab Info Range             // Tab: Info

  const SIR = "D3"     // SIR = Source ID                  // Tab: Config
  const IDR = "D10"    // IDR = ID Sheet                   // Tab: Config
  const EPR = "D13"    // EPR = Exportable?                // Tab: Config
  const EXR = "D16"    // EXR = Exported?                  // Tab: Config
  const FOR = "D19"    // FOR = Formula?                   // Tab: Config
  const TDR = "D22"    // TDR = Target ID                  // Tab: Config
  const DIR = "D25"    // DIR = DATA Source ID             // Tab: Config

  const ICR = "F13"    // ICR = Sheet ID Check             // Tab: Config
  const IER = "F16"    // IER = ID Exported?               // Tab: Config

  const OPR = "L3"     // OPR = Option                     // Tab: Config
  const UFR = "L6"     // UFR = Update Form                // Tab: Config
  const HCR = "L9"     // HCR = Hide Config                // Tab: Config
  const DBG = "L12"    // DBG = Debug Mode                 // Tab: Config
  const IST = "L18"    // IST = Is Stock?                  // Tab: Config
  const TGR = "L21"    // TGR = Number of Triggers         // Tab: Config

  const TG1 = "N21"    // TG1 = Sheet Trigger Event        // Tab: Config
  const TG2 = "N24"    // TG2 = Data Trigger Event         // Tab: Config
  const TG3 = "N27"    // TG3 = Extra Trigger Event        // Tab: Config
  const TG4 = "N30"    // TG4 = Settings Trigger Event     // Tab: Config
  const TG5 = "N33"    // TG5 = SaveAll Trigger Event      // Tab: Config

  const COR = "I4:J7"  // COR = Config Options Range       // Tab: Config

/////////////////////////////////////////////////////////////////////Export/////////////////////////////////////////////////////////////////////

  const ETR = "P4"     // ETR = Export to Swing            // Tab: Config
  const EOP = "P6"     // EOP = Export to Option           // Tab: Config
  const EBT = "P8"     // EBT = Export to BTC              // Tab: Config
  const ETE = "P10"    // ETE = Export to Termo            // Tab: Config
  const EFU = "P12"    // EFU = Export to Fund             // Tab: Config

  const EBL = "P15"    // EBL = Export to BLC              // Tab: Config
  const EDR = "P17"    // EDR = Export to DRE              // Tab: Config
  const EFL = "P19"    // EFL = Export to FLC              // Tab: Config
  const EDV = "P21"    // EDV = Export to DVA              // Tab: Config

  const ETF = "P24"    // ETF = Export to Future           // Tab: Config
  const ERT = "P26"    // ERT = Export to Right            // Tab: Config
  const EWT = "P28"    // EWT = Export to Warrant          // Tab: Config
  const ERC = "P30"    // ERC = Export to Receipt          // Tab: Config
  const EBK = "P32"    // EBK = Export to Block            // Tab: Config
  const EAF = "P34"    // EAF = Export to After            // Tab: Config, Settings

/////////////////////////////////////////////////////////////////////Import/////////////////////////////////////////////////////////////////////

  const ITR = "R4"     // ITR = Import to Swing            // Tab: Config, Settings
  const IOP = "R6"     // IOP = Import to Option           // Tab: Config, Settings
  const IBT = "R8"     // IBT = Import to BTC              // Tab: Config, Settings
  const ITE = "R10"    // ITE = Import to Termo            // Tab: Config, Settings
  const IFU = "R12"    // IFU = Import to Fund             // Tab: Config, Settings

  const IBL = "R15"    // IBL = Import to BLC / Balanco    // Tab: Config, Settings
  const IDE = "R17"    // IDE = Import to DRE / Resultado  // Tab: Config, Settings
  const IFL = "R19"    // IFL = Import to FLC / Fluxo      // Tab: Config, Settings
  const IDV = "R21"    // IDV = Import to DVA / Valor      // Tab: Config, Settings

  const IFT = "R24"    // IFT = Import to Future           // Tab: Config, Settings
  const IRT = "R26"    // IRT = Import to Right            // Tab: Config, Settings
  const IWT = "R28"    // IWT = Import to Warrant          // Tab: Config, Settings
  const IRC = "R30"    // IRC = Import to Receipt          // Tab: Config, Settings
  const IBK = "R32"    // IBK = Import to Block            // Tab: Config, Settings
  const IAF = "R34"    // IAF = Import to After            // Tab: Config, Settings

/////////////////////////////////////////////////////////////////////Save/////////////////////////////////////////////////////////////////////

  const STR = "T4"     // STR = Save to Swing              // Tab: Config, Settings
  const SOP = "T6"     // SOP = Save to Option             // Tab: Config, Settings
  const SBT = "T8"     // SBT = Save to BTC                // Tab: Config, Settings
  const STE = "T10"    // STE = Save to Termo              // Tab: Config, Settings
  const SFU = "T12"    // SFU = Save to Fund               // Tab: Config, Settings

  const SBL = "T15"    // SBL = Save to BLC                // Tab: Config, Settings
  const SDE = "T17"    // SDE = Save to DRE                // Tab: Config, Settings
  const SFL = "T19"    // SFL = Save to FLC                // Tab: Config, Settings
  const SDV = "T21"    // SDV = Save to DVA                // Tab: Config, Settings

  const SFT = "T24"    // SFT = Save to Future             // Tab: Config, Settings
  const SRT = "T26"    // SRT = Save to Right              // Tab: Config, Settings
  const SWT = "T28"    // SWT = Save to Warrant            // Tab: Config, Settings
  const SRC = "T30"    // SRC = Save to Receipt            // Tab: Config, Settings
  const SBK = "T32"    // SBK = Save to Block              // Tab: Config, Settings
  const SAF = "T34"    // SAF = Save to After              // Tab: Config, Settings

/////////////////////////////////////////////////////////////////////Edit/////////////////////////////////////////////////////////////////////

  const DTR = "V4"     // DTR = Edit to Swing              // Tab: Config, Settings
  const DOP = "V6"     // DOP = Edit to Option             // Tab: Config, Settings
  const DBT = "V8"     // DBT = Edit to BTC                // Tab: Config, Settings
  const DTE = "V10"    // DTE = Edit to Termo              // Tab: Config, Settings
  const DFU = "V12"    // DFU = Edit to Fund               // Tab: Config, Settings

  const DBL = "V15"    // DBL = Edit to BLC                // Tab: Config, Settings
  const DDE = "V17"    // DDE = Edit to DRE                // Tab: Config, Settings
  const DFL = "V19"    // DFL = Edit to FLC                // Tab: Config, Settings
  const DDV = "V21"    // DDV = Edit to DVA                // Tab: Config, Settings

  const DFT = "V24"    // DFT = Edit to Future             // Tab: Config, Settings
  const DRT = "V26"    // DRT = Edit to Right              // Tab: Config, Settings
  const DWT = "V28"    // DWT = Edit to Warrant            // Tab: Config, Settings
  const DRC = "V30"    // DRC = Edit to Receipt            // Tab: Config, Settings
  const DBK = "V32"    // DBK = Edit to Block              // Tab: Config, Settings
  const DAF = "V34"    // DAF = Edit to After              // Tab: Config, Settings

/////////////////////////////////////////////////////////////////////Hide/////////////////////////////////////////////////////////////////////

  const HTR = "X4"     // HTR = Hide to Swing              // Tab: Config, Settings
  const HOP = "X6"     // HOP = Hide to Option             // Tab: Config, Settings
  const HBT = "X8"     // HBT = Hide to BTC                // Tab: Config, Settings
  const HTE = "X10"    // HTE = Hide to Termo              // Tab: Config, Settings
  const HFU = "X12"    // HFU = Hide to Fund               // Tab: Config, Settings

  const HBL = "X15"    // HBL = Hide to BLC                // Tab: Config, Settings
  const HDE = "X17"    // HDE = Hide to DRE                // Tab: Config, Settings
  const HFL = "X19"    // HFL = Hide to FLC                // Tab: Config, Settings
  const HDV = "X21"    // HDV = Hide to DVA                // Tab: Config, Settings

  const HFT = "X24"    // HFT = Hide to Future             // Tab: Config, Settings
  const HRT = "X26"    // HRT = Hide to Right              // Tab: Config, Settings
  const HWT = "X28"    // HWT = Hide to Warrant            // Tab: Config, Settings
  const HRC = "X30"    // HRC = Hide to Receipt            // Tab: Config, Settings
  const HBK = "X32"    // HBK = Hide to Block              // Tab: Config, Settings
  const HAF = "X34"    // HAF = Hide to After              // Tab: Config, Settings

/////////////////////////////////////////////////////////////////////Settings/////////////////////////////////////////////////////////////////////

  const ACT = "B3"     // SET = Settings                   // Tab: Settings
  const TRU = "F3"     // TRU = True                       // Tab: Settings
  const SAV = "K3"     // SAV = Save                       // Tab: Settings
  const IND = "K10"    // IND = Individual                 // Tab: Settings
  const EXT = "F10"    // EXT = Extra                      // Tab: Settings

  const MIN = "B10"    // MIN = Min Value To Exp           // Tab: Settings
  const MAX = "B12"    // MAX = Max Value To Exp           // Tab: Settings

/////////////////////////////////////////////////////////////////////CONSTANT/////////////////////////////////////////////////////////////////////
