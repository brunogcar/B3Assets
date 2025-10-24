// function to add custom menu to spreadsheet

function onOpen()
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')

  /////////////////////////////////////////////////////////////////////SAVE/////////////////////////////////////////////////////////////////////

    .addSubMenu
    (ui.createMenu('Save')
      .addSubMenu
      (ui.createMenu('Save All')
        .addItem('Save All', 'doSaveAll')
        .addSeparator()
        .addItem('Save Basics',  'doSaveAllBasics')
        .addItem('Save Extras',  'doSaveAllExtras')
        .addItem('Save Financials',   'doSaveAllFinancials')
      )
      .addSeparator()
      .addItem('Save Proventos', 'doSaveProventos')
      .addSeparator()
      .addItem('Save Grafics (Swing - Opções - BTC - Termo - Fund)', 'doSaveBasics')
      .addSeparator()
      .addSubMenu
      (ui.createMenu('Save Swing')
        .addItem('Save Swing', 'doSaveSWING')
        .addSeparator()
        .addItem('Save Swing 4',   'menuSaveSwing_4')
        .addItem('Save Swing 12',  'menuSaveSwing_12')
        .addItem('Save Swing 52',  'menuSaveSwing_52')
      )
      .addItem('Save Opções',    'menuSaveOpcoes')
      .addItem('Save BTC',       'menuSaveBTC')
      .addItem('Save Termo',     'menuSaveTermo')
      .addItem('Save Future',    'menuSaveFuture')
      .addItem('Save BLOCK',     'menuSaveBLOCK')
      .addItem('Save Fund',      'menuSaveFund')
      .addSeparator()
      .addItem('Save Extra (RIGH - RECEIPT - WARRANT - After)', 'doSaveExtras')
      .addSeparator()
      .addItem('Save RIGHT',     'doSaveRIGHT')
      .addItem('Save RECEIPT',   'doSaveRECEIPT')
      .addItem('Save WARRANT',  'doSaveWARRANT')
      .addItem('Save After',       'menuSaveAFTER')
      .addSeparator()
      .addItem('Save Financials (BLC - DRE - FLC - DVA)','doSaveFinancials')
      .addSeparator()
      .addSubMenu
      (ui.createMenu('BLC')
        .addItem('Save BLC'    ,  'menuSaveBLC')
        .addItem('Save Balanço',  'menuSaveBalanco')
      )
      .addSubMenu
      (ui.createMenu('DRE')
        .addItem('Save DRE',      'menuSaveDRE')
        .addItem('Save Resultado','menuSaveResultado')
      )
      .addSubMenu
      (ui.createMenu('FLC')
        .addItem('Save FLC',      'menuSaveFLC')
        .addItem('Save Fluxo',    'menuSaveFluxo')
      )
      .addSubMenu
      (ui.createMenu('DVA')
        .addItem('Save DVA',      'menuSaveDVA')
        .addItem('Save Valor',    'menuSaveValor')
      )
    )
    .addSeparator()

/////////////////////////////////////////////////////////////////////EDIT/////////////////////////////////////////////////////////////////////

    .addSubMenu
    (ui.createMenu('Edit')
      .addItem('Edit ALL','doEditAll')
      .addSeparator()
      .addItem('Edit Grafics (Swing - Opções - BTC - Termo - Futuro - Fund)', 'doEditBasics')
      .addSeparator()
      .addSubMenu
      (ui.createMenu('Edit Swing')
        .addItem('Edit Swing', 'doEditSWING')
        .addSeparator()
        .addItem('Edit Swing 4',    'menuEditSwing_4')
        .addItem('Edit Swing 12',   'menuEditSwing_12')
        .addItem('Edit Swing 52',   'menuEditSwing_52')
      )
      .addItem('Edit Opções',     'menuEditOpcoes')
      .addItem('Edit BTC',        'menuEditBTC')
      .addItem('Edit Termo',      'menuEditTermo')
      .addItem('Edit Future',     'menuEditFuture')
      .addItem('Edit Fund',       'menuEditFund')
      .addSeparator()
      .addItem('Edit Balanço (BLC - DRE - FLC - DVA)','doEditFinancials')
      .addSeparator()
      .addSubMenu
      (ui.createMenu('BLC')
        .addItem('Edit BLC',      'menuEditBLC')
        .addItem('Edit Balanço',  'menuEditBalanco')
      )
      .addSubMenu
      (ui.createMenu('DRE')
        .addItem('Edit DRE',      'menuEditDRE')
        .addItem('Edit Resultado','menuEditResultado')
      )
      .addSubMenu
      (ui.createMenu('FLC')
        .addItem('Edit FLC',      'menuEditFLC')
        .addItem('Edit Fluxo',    'menuEditFluxo')
      )
      .addSubMenu
      (ui.createMenu('DVA')
        .addItem('Edit DVA',      'menuEditDVA')
        .addItem('Edit Valor',    'menuEditValor')
      )
    )
    .addSeparator()

/////////////////////////////////////////////////////////////////////COPY/////////////////////////////////////////////////////////////////////

    .addSubMenu
    (ui.createMenu('Copy')
      .addItem('Copy Swing 4',    'menuCopySwing_4')
      .addItem('Copy Swing 12',   'menuCopySwing_12')
      .addItem('Copy Swing 52',   'menuCopySwing_52')
      .addItem('Copy Opções',     'menuCopyOpcoes')
      .addItem('Copy BTC',        'menuCopyBTC')
      .addItem('Copy Termo',      'menuCopyTermo')
      .addItem('Copy Fund',       'menuCopyFund')
      .addSeparator()
      .addSubMenu
      (ui.createMenu('BLC')
        .addItem('Copy BLC',      'menuCopyBLC')
        .addItem('Copy Balanço',  'menuCopyBalanco')
      )
      .addSubMenu
      (ui.createMenu('DRE')
        .addItem('Copy DRE',      'menuCopyDRE')
        .addItem('Copy Resultado','menuCopyResultado')
      )
      .addSubMenu
      (ui.createMenu('FLC')
        .addItem('Copy FLC',      'menuCopyFLC')
        .addItem('Copy Fluxo',    'menuCopyFluxo')
      )
    .addSubMenu
      (ui.createMenu('DVA')
        .addItem('Copy DVA',      'menuCopyDVA')
        .addItem('Copy Valor',    'menuCopyValor')
      )
    )
    .addSeparator()

/////////////////////////////////////////////////////////////////////CLEAR / CLEAN/////////////////////////////////////////////////////////////////////

    .addSubMenu
    (ui.createMenu('Clear / Clean')

/////////////////////////////////////////////////////////////////////CLEAR/////////////////////////////////////////////////////////////////////

      .addSubMenu
      (ui.createMenu('Clear')
        .addItem('Clear ALL(Grafics - Balanço)', 'doClearAll')
        .addSeparator()
        .addItem('Clear Grafics (Swing - Opções - BTC - Termo - Futuro - Fund)', 'doClearBasics')
        .addSeparator()
        .addItem('Clear Swing 4',  'menuClearSwing_4')
        .addItem('Clear Swing 12', 'menuClearSwing_12')
        .addItem('Clear Swing 52', 'menuClearSwing_52')
        .addItem('Clear Opções',   'menuClearOpcoes')
        .addItem('Clear BTC',      'menuClearBTC')
        .addItem('Clear Termo',    'menuClearTermo')
        .addItem('Clear Future',   'menuClearFuture')
        .addItem('Clear Fund',     'menuClearFund')
        .addSeparator()
        .addItem('Clear Balanço (BLC - DRE - FLC - DVA)','doClearFinancials')
        .addSeparator()
        .addItem('Clear BLC',      'menuClearBLC')
        .addItem('Clear DRE',      'menuClearDRE')
        .addItem('Clear FLC',      'menuClearFLC')
        .addItem('Clear DVA',      'menuClearDVA')
      )
      .addSeparator()

/////////////////////////////////////////////////////////////////////CLEAN/////////////////////////////////////////////////////////////////////

      .addSubMenu
      (ui.createMenu('Clean "0" and "-"')
        .addItem('Clean Grafics (Swing - Opções - BTC - Termo - Futuro - Fund)', 'doCleanBasics')
        .addSeparator()
        .addItem('Clean Swing 4',  'menuCleanSwing_4')
        .addItem('Clean Swing 12', 'menuCleanSwing_12')
        .addItem('Clean Swing 52', 'menuCleanSwing_52')
        .addItem('Clean Opções',   'menuCleanOpcoes')
        .addItem('Clean BTC',      'menuCleanBTC')
        .addItem('Clean Termo',    'menuCleanTermo')
        .addItem('Clean Future',   'menuCleanFuture')
        .addItem('Clean Fund',     'menuCleanFund')
      )
    )
    .addSeparator()


/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////

    .addItem('IMPORT', 'Import')
//    .addSubMenu
//    (ui.createMenu('IMPORT')
//       .addItem('IMPORT', 'import_Current')
//       .addItem('IMPORT 12x to 13', 'import_12x_to_13')
//    )
    .addSeparator()

/////////////////////////////////////////////////////////////////////EXPORT/////////////////////////////////////////////////////////////////////

    .addSubMenu
    (ui.createMenu('EXPORT')
    .addSubMenu
      (ui.createMenu('Info')
        .addItem('Export Relação', 'doExportInfo')
        .addItem('Export Sheet ID','setSheetID')
      )
    .addSubMenu
      (ui.createMenu('DATA')
        .addItem('Export ALL (Graphics - Balanço)', 'doExportAll')
        .addSeparator()
        .addItem('Export Proventos', 'doExportProventos')
        .addSeparator()
        .addItem('Export Graphics (Swing - Options - BTC - Future - Fund)', 'doExportBasics')
        .addSeparator()
        .addSubMenu
        (ui.createMenu('Export Swing')
          .addItem('Export Swing', 'doExportSWING')
          .addSeparator()
          .addItem('Export Swing 4',   'menuExportSwing_4')
          .addItem('Export Swing 12',  'menuExportSwing_12')
          .addItem('Export Swing 52',  'menuExportSwing_52')
        )
        .addItem('Export Opções',    'menuExportOpcoes')
        .addItem('Export BTC',       'menuExportBTC')
        .addItem('Export Termo',     'menuExportTermo')
        .addItem('Export Future',    'menuExportFuture')
        .addItem('Export Fund',      'menuExportFund')
        .addItem('Export After',     'menuExportAfter')
        .addSeparator()
        .addItem('Export Balanço (BLC - DRE - FLC - DVA)', 'doExportFinancials')
        .addSeparator()
        .addItem('Export BLC',       'menuExportBLC')
        .addItem('Export DRE',       'menuExportDRE')
        .addItem('Export FLC',       'menuExportFLC')
        .addItem('Export DVA',       'menuExportDVA')
      )
    )
    .addSeparator()

/////////////////////////////////////////////////////////////////////FIX/////////////////////////////////////////////////////////////////////

    .addSubMenu
    (ui.createMenu('Split or Inplit')
      .addItem('Fix Split', 'fixSplit')
      .addSeparator()
      .addItem('Fix Swing 4',        'fixSWING_4Split')
      .addItem('Fix Swing 12',       'fixSWING_12Split')
      .addItem('Fix Swing 52',       'fixSWING_52Split')
      .addItem('Fix Opções',         'fixOptionsSplit')
      .addItem('Fix BTC',            'fixBTCSplit')
      .addItem('Fix Termo',          'fixTermoSplit')
      .addItem('Fix After',          'fixAfterSplit')
      .addItem('Fix Future',         'fixFutureSplit')
      .addItem('Fix Futures Plus',   'fixFUTPlusSplits')
      .addItem('Fix Fund',           'fixFundSplit')
      .addItem('Fix Extra',          'fixEXTRASplits')
    )
    .addSeparator()
    .addSubMenu
    (ui.createMenu('Retire/Delete')
      .addItem('Apagar Tabela',      'doDelete')
    .addSubMenu
      (ui.createMenu('Arquivar')
        .addItem('Arquivar Dados',     'doRetire')
      )
      .addItem('Revogar Acesso',      'revokeOwnAccess')
    )

/////////////////////////////////////////////////////////////////////Retire/Delete/////////////////////////////////////////////////////////////////////

    .addSeparator()
    .addSubMenu
    (ui.createMenu('OTHER')
      .addItem('Pegar Codigo CVM',      'doGetCodeCVM')
    )

    .addToUi();
};

/////////////////////////////////////////////////////////////////////SAVE/////////////////////////////////////////////////////////////////////

function menuSaveSwing_4()    { doSaveBasic(SWING_4);}
function menuSaveSwing_12()   { doSaveBasic(SWING_12);}
function menuSaveSwing_52()   { doSaveBasic(SWING_52);}
function menuSaveOpcoes()     { doSaveBasic(OPCOES);}
function menuSaveBTC()        { doSaveBasic(BTC);}
function menuSaveTermo()      { doSaveBasic(TERMO);}
function menuSaveFuture()     { doSaveBasic(FUTURE);}
function menuSaveBLOCK()      { doSaveBasic(BLOCK);}
function menuSaveFund()       { doSaveBasic(FUND);}

function menuSaveDRT_1()      { doSaveBasic(RIGHT_1);}
function menuSaveDRT_2()      { doSaveBasic(RIGHT_2);}
function menuSaveRCB_9()      { doSaveBasic(RECEIPT_9);}
function menuSaveRCB_10()     { doSaveBasic(RECEIPT_10);}
function menuSaveGAR_11()     { doSaveBasic(WARRANT_11);}
function menuSaveGAR_12()     { doSaveBasic(WARRANT_12);}
function menuSaveGAR_13()     { doSaveBasic(WARRANT_13);}
function menuSaveAFTER()      { doSaveBasic(AFTER);}

function menuSaveBLC()        { doSaveFinancial(BLC);}
function menuSaveBalanco()    { doSaveFinancial(Balanco);}
function menuSaveDRE()        { doSaveFinancial(DRE);}
function menuSaveResultado()  { doSaveFinancial(Resultado);}
function menuSaveFLC()        { doSaveFinancial(FLC);}
function menuSaveFluxo()      { doSaveFinancial(Fluxo);}
function menuSaveDVA()        { doSaveFinancial(DVA);}
function menuSaveValor()      { doSaveFinancial(Valor);}

/////////////////////////////////////////////////////////////////////EDIT/////////////////////////////////////////////////////////////////////

function menuEditSwing_4()    { doEditBasic(SWING_4);}
function menuEditSwing_12()   { doEditBasic(SWING_12);}
function menuEditSwing_52()   { doEditBasic(SWING_52);}
function menuEditOpcoes()     { doEditBasic(OPCOES);}
function menuEditBTC()        { doEditBasic(BTC);}
function menuEditTermo()      { doEditBasic(TERMO);}
function menuEditFuture()     { doEditBasic(FUTURE);}
function menuEditFund()       { doEditBasic(FUND);}


function menuEditBLC()        { doEditFinancial(BLC);}
function menuEditBalanco()    { doEditFinancial(Balanco);}
function menuEditDRE()        { doEditFinancial(DRE);}
function menuEditResultado()  { doEditFinancial(Resultado);}
function menuEditFLC()        { doEditFinancial(FLC);}
function menuEditFluxo()      { doEditFinancial(Fluxo);}
function menuEditDVA()        { doEditFinancial(DVA);}
function menuEditValor()      { doEditFinancial(Valor);}

/////////////////////////////////////////////////////////////////////COPY/////////////////////////////////////////////////////////////////////

function menuCopySwing_4()    { doCopyBasic(SWING_4);}
function menuCopySwing_12()   { doCopyBasic(SWING_12);}
function menuCopySwing_52()   { doCopyBasic(SWING_52);}
function menuCopyOpcoes()     { doCopyBasic(OPCOES);}
function menuCopyBTC()        { doCopyBasic(BTC);}
function menuCopyTermo()      { doCopyBasic(TERMO);}
function menuCopyFuture()     { doCopyBasic(FUTURE);}
function menuCopyFund()       { doCopyBasic(FUND);}

function menuCopyBLC()        { doCopyFinancial(BLC);}
function menuCopyBalanco()    { doCopyFinancial(Balanco);}
function menuCopyDRE()        { doCopyFinancial(DRE);}
function menuCopyResultado()  { doCopyFinancial(Resultado);}
function menuCopyFLC()        { doCopyFinancial(FLC);}
function menuCopyFluxo()      { doCopyFinancial(Fluxo);}
function menuCopyDVA()        { doCopyFinancial(DVA);}
function menuCopyValor()      { doCopyFinancial(Valor);}

/////////////////////////////////////////////////////////////////////CLEAN/////////////////////////////////////////////////////////////////////

function menuClearSwing_4()   { doClearBasic(SWING_4);}
function menuClearSwing_12()  { doClearBasic(SWING_12);}
function menuClearSwing_52()  { doClearBasic(SWING_52);}
function menuClearOpcoes()    { doClearBasic(OPCOES);}
function menuClearBTC()       { doClearBasic(BTC);}
function menuClearTermo()     { doClearBasic(TERMO);}
function menuClearFuture()    { doClearBasic(FUTURE);}
function menuClearFund()      { doClearBasic(FUND);}

function menuClearBLC()       { doClearFinancial(BLC);}
function menuClearDRE()       { doClearFinancial(DRE);}
function menuClearFLC()       { doClearFinancial(FLC);}
function menuClearDVA()       { doClearFinancial(DVA);}

/////////////////////////////////////////////////////////////////////CLEAN/////////////////////////////////////////////////////////////////////

function menuCleanSwing_4()   { doCleanBasic(SWING_4);}
function menuCleanSwing_12()  { doCleanBasic(SWING_12);}
function menuCleanSwing_52()  { doCleanBasic(SWING_52);}
function menuCleanOpcoes()    { doCleanBasic(OPCOES);}
function menuCleanBTC()       { doCleanBasic(BTC);}
function menuCleanTermo()     { doCleanBasic(TERMO);}
function menuCleanFuture()    { doCleanBasic(FUTURE);}
function menuCleanFund()      { doCleanBasic(FUND);}

/////////////////////////////////////////////////////////////////////EXPORT/////////////////////////////////////////////////////////////////////

function menuExportSwing_4()  { doExportBasic(SWING_4);}
function menuExportSwing_12() { doExportBasic(SWING_12);}
function menuExportSwing_52() { doExportBasic(SWING_52);}
function menuExportOpcoes()   { doExportBasic(OPCOES);}
function menuExportBTC()      { doExportBasic(BTC);}
function menuExportTermo()    { doExportBasic(TERMO);}
function menuExportFuture()   { doExportBasic(FUTURE);}
function menuExportFund()     { doExportBasic(FUND);}
function menuExportAfter()    { doExportBasic(AFTER);}

function menuExportBLC()      { doExportFinancial(BLC);}
function menuExportDRE()      { doExportFinancial(DRE);}
function menuExportFLC()      { doExportFinancial(FLC);}
function menuExportDVA()      { doExportFinancial(DVA);}

/////////////////////////////////////////////////////////////////////OTHER/////////////////////////////////////////////////////////////////////

function doSaveSWING()        { doSaveBasic(SWING_4); doSaveBasic(SWING_12); doSaveBasic(SWING_52);}
function doEditSWING()        { doEditBasic(SWING_4); doEditBasic(SWING_12); doEditBasic(SWING_52);}
function doExportSWING()      { doExportBasic(SWING_4); doExportBasic(SWING_12); doExportBasic(SWING_52);}
function doSaveRIGHT()        { doSaveBasic(RIGHT_1); doSaveBasic(RIGHT_2);}
function doSaveRECEIPT()      { doSaveBasic(RECEIPT_9); doSaveBasic(RECEIPT_10);}
function doSaveWARRANT()      { doSaveBasic(WARRANT_11); doSaveBasic(WARRANT_12); doSaveBasic(WARRANT_13);}

/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////
