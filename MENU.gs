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
        .addItem('Save Sheets',  'doSaveAllSheets')
        .addItem('Save Extras',  'doSaveAllExtras')
        .addItem('Save Datas',   'doSaveAllDatas')
      )
      .addSeparator()
      .addItem('Save Proventos', 'doSaveProventos')
      .addSeparator()
      .addItem('Save Grafics (Swing - Opções - BTC - Termo - Fund)', 'doSaveSheets')
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
      .addItem('Save Fund',      'menuSaveFund')
      .addSeparator()
      .addItem('Save Balanço (BLC - DRE - FLC - DVA)','doSaveDatas')
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
      .addItem('Edit Grafics (Swing - Opções - BTC - Termo - Futuro - Fund)', 'doEditSheets')
      .addSeparator()
      .addItem('Edit Swing 4',    'menuEditSwing_4')
      .addItem('Edit Swing 12',   'menuEditSwing_12')
      .addItem('Edit Swing 52',   'menuEditSwing_52')
      .addItem('Edit Opções',     'menuEditOpcoes')
      .addItem('Edit BTC',        'menuEditBTC')
      .addItem('Edit Termo',      'menuEditTermo')
      .addItem('Edit Future',     'menuEditFuture')
      .addItem('Edit Fund',       'menuEditFund')
      .addSeparator()
      .addItem('Edit Balanço (BLC - DRE - FLC - DVA)','doEditDatas')
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
        .addItem('Clear Grafics (Swing - Opções - BTC - Termo - Futuro - Fund)', 'doClearSheets')
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
        .addItem('Clear Balanço (BLC - DRE - FLC - DVA)','doClearDatas')
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
        .addItem('Clean Grafics (Swing - Opções - BTC - Termo - Futuro - Fund)', 'doCleanSheets')
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
        .addItem('Export Graphics (Swing - Options - BTC - Future - Fund)', 'doExportSheets')
        .addSeparator()
        .addItem('Export Swing 4',   'menuExportSwing_4')
        .addItem('Export Swing 12',  'menuExportSwing_12')
        .addItem('Export Swing 52',  'menuExportSwing_52')
        .addItem('Export Opções',    'menuExportOpcoes')
        .addItem('Export BTC',       'menuExportBTC')
        .addItem('Export Termo',     'menuExportTermo')
        .addItem('Export Future',    'menuExportFuture')
        .addItem('Export Fund',      'menuExportFund')
        .addSeparator()
        .addItem('Export Balanço (BLC - DRE - FLC - DVA)', 'doExportDatas')
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

function menuSaveSwing_4()    { doSaveSheet(SWING_4);}
function menuSaveSwing_12()   { doSaveSheet(SWING_12);}
function menuSaveSwing_52()   { doSaveSheet(SWING_52);}
function menuSaveOpcoes()     { doSaveSheet(OPCOES);}
function menuSaveBTC()        { doSaveSheet(BTC);}
function menuSaveTermo()      { doSaveSheet(TERMO);}
function menuSaveFuture()     { doSaveSheet(FUTURE);}
function menuSaveFund()       { doSaveSheet(FUND);}

function menuSaveBLC()        { doSaveData(BLC);}
function menuSaveBalanco()    { doSaveData(Balanco);}
function menuSaveDRE()        { doSaveData(DRE);}
function menuSaveResultado()  { doSaveData(Resultado);}
function menuSaveFLC()        { doSaveData(FLC);}
function menuSaveFluxo()      { doSaveData(Fluxo);}
function menuSaveDVA()        { doSaveData(DVA);}
function menuSaveValor()      { doSaveData(Valor);}

/////////////////////////////////////////////////////////////////////EDIT/////////////////////////////////////////////////////////////////////

function menuEditSwing_4()    { doEditSheet(SWING_4);}
function menuEditSwing_12()   { doEditSheet(SWING_12);}
function menuEditSwing_52()   { doEditSheet(SWING_52);}
function menuEditOpcoes()     { doEditSheet(OPCOES);}
function menuEditBTC()        { doEditSheet(BTC);}
function menuEditTermo()      { doEditSheet(TERMO);}
function menuEditFuture()     { doEditSheet(FUTURE);}
function menuEditFund()       { doEditSheet(FUND);}


function menuEditBLC()        {doEditData(BLC);}
function menuEditBalanco()    {doEditData(Balanco);}
function menuEditDRE()        {doEditData(DRE);}
function menuEditResultado()  {doEditData(Resultado);}
function menuEditFLC()        {doEditData(FLC);}
function menuEditFluxo()      {doEditData(Fluxo);}
function menuEditDVA()        {doEditData(DVA);}
function menuEditValor()      {doEditData(Valor);}

/////////////////////////////////////////////////////////////////////COPY/////////////////////////////////////////////////////////////////////

function menuCopySwing_4()    { doCopySheet(SWING_4);}
function menuCopySwing_12()   { doCopySheet(SWING_12);}
function menuCopySwing_52()   { doCopySheet(SWING_52);}
function menuCopyOpcoes()     { doCopySheet(OPCOES);}
function menuCopyBTC()        { doCopySheet(BTC);}
function menuCopyTermo()      { doCopySheet(TERMO);}
function menuCopyFuture()     { doCopySheet(FUTURE);}
function menuCopyFund()       { doCopySheet(FUND);}

function menuCopyBLC()        { doCopyData(BLC);}
function menuCopyBalanco()    { doCopyData(Balanco);}
function menuCopyDRE()        { doCopyData(DRE);}
function menuCopyResultado()  { doCopyData(Resultado);}
function menuCopyFLC()        { doCopyData(FLC);}
function menuCopyFluxo()      { doCopyData(Fluxo);}
function menuCopyDVA()        { doCopyData(DVA);}
function menuCopyValor()      { doCopyData(Valor);}

/////////////////////////////////////////////////////////////////////CLEAN/////////////////////////////////////////////////////////////////////

function menuClearSwing_4()   { doClearSheet(SWING_4);}
function menuClearSwing_12()  { doClearSheet(SWING_12);}
function menuClearSwing_52()  { doClearSheet(SWING_52);}
function menuClearOpcoes()    { doClearSheet(OPCOES);}
function menuClearBTC()       { doClearSheet(BTC);}
function menuClearTermo()     { doClearSheet(TERMO);}
function menuClearFuture()    { doClearSheet(FUTURE);}
function menuClearFund()      { doClearSheet(FUND);}

function menuClearBLC()       { doClearData(BLC);}
function menuClearDRE()       { doClearData(DRE);}
function menuClearFLC()       { doClearData(FLC);}
function menuClearDVA()       { doClearData(DVA);}

/////////////////////////////////////////////////////////////////////CLEAN/////////////////////////////////////////////////////////////////////

function menuCleanSwing_4()   { doCleanSheet(SWING_4);}
function menuCleanSwing_12()  { doCleanSheet(SWING_12);}
function menuCleanSwing_52()  { doCleanSheet(SWING_52);}
function menuCleanOpcoes()    { doCleanSheet(OPCOES);}
function menuCleanBTC()       { doCleanSheet(BTC);}
function menuCleanTermo()     { doCleanSheet(TERMO);}
function menuCleanFuture()    { doCleanSheet(FUTURE);}
function menuCleanFund()      { doCleanSheet(FUND);}

/////////////////////////////////////////////////////////////////////EXPORT/////////////////////////////////////////////////////////////////////

function menuExportSwing_4()  { doExportSheet(SWING_4);}
function menuExportSwing_12() { doExportSheet(SWING_12);}
function menuExportSwing_52() { doExportSheet(SWING_52);}
function menuExportOpcoes()   { doExportSheet(OPCOES);}
function menuExportBTC()      { doExportSheet(BTC);}
function menuExportTermo()    { doExportSheet(TERMO);}
function menuExportFuture()   { doExportSheet(FUTURE);}
function menuExportFund()     { doExportSheet(FUND);}

function menuExportBLC()      { doExportData(BLC);}
function menuExportDRE()      { doExportData(DRE);}
function menuExportFLC()      { doExportData(FLC);}
function menuExportDVA()      { doExportData(DVA);}

/////////////////////////////////////////////////////////////////////OTHER/////////////////////////////////////////////////////////////////////

function doSaveSWING()      { doSaveSheet(SWING_4); doSaveSheet(SWING_12); doSaveSheet(SWING_52);}

/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////