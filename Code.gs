function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    
    let result;
    switch(action) {
      case 'save_spk_ctl'      : result = saveSPK_CTL(payload.data);    break;
      case 'update_spk_status' : result = updateSPKStatus(payload.data); break;
      case 'get_spk_list'      : result = getSPKList(payload.filter);    break;
      case 'get_gr_form_data'  : result = getGRFormData();               break;
      case 'save_gr'           : result = saveGR(payload.data);          break;
      default: throw new Error('Unknown action: ' + action);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ok: true, data: result}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var view = (e && e.parameter && e.parameter.v) ? String(e.parameter.v).toLowerCase() : '';
  
  var viewConfig = {
    'shr': { page: 'page_board_shr', title: 'Board Mesin SHR', icon: 'ti-scissors' },
    'ctl': { page: 'page_board_ctl', title: 'Board Mesin CTL', icon: 'ti-cut' },
    'slt': { page: 'page_board_slt', title: 'Board Mesin SLT', icon: 'ti-slice' }
  };

  // Standalone guide page (no shell)
  if (view === 'panduan_live_stok') {
    return HtmlService.createHtmlOutputFromFile('page_panduan_live_stok')
      .setTitle('Panduan Live Inventory — COS FC')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  if (viewConfig[view]) {
    var tpl = HtmlService.createTemplateFromFile('index_board');
    tpl.targetPage = viewConfig[view].page;
    tpl.pageTitle  = viewConfig[view].title;
    tpl.pageIcon   = viewConfig[view].icon;
    return tpl.evaluate()
      .setTitle(viewConfig[view].title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Default: index utama (full COS app dengan sidebar)
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('FC Core Operations System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const SHEET_NAMES = {
  SPK       : 'SPK',
  STOK_COIL : 'Stok_Coil',
  STOK_SHEET: 'Stok_Sheet',
  STOK_WIP  : 'Stok_WIP',
  STOK_FG   : 'Stok_FG',
  TRACE_LOG : 'Trace_Log',
  M_CONFIG  : 'M_Config',
  REKAP_ICT : 'Rekap_ICT',
  PENERIMAAN: 'Penerimaan_Material',
  SO        : 'SO',
  M_ITEM    : 'M_ITEM',
  M_SUPP    : 'M_SUPP'
};

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();	
  return ss.getSheetByName(name);
}

function getPageContent(pageName) {
  try {
    return HtmlService.createHtmlOutputFromFile(pageName).getContent();
  } catch(e) {
    return '<div style="padding:40px;text-align:center;color:#888">Halaman <b>' + pageName + '</b> belum dibuat.</div>';
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}