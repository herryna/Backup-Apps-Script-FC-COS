// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - PURCHASE ORDER (PO) CONTROL SERVICE
// =========================================================================

function getPOInitData() {
  try {
    return JSON.stringify({
      pool: getOpenPOPool(),
      dash: getPODashboardData(1, null),
      vendors: getVendorList(),
      items: getItemList()
    });
  } catch (e) { return JSON.stringify({ error: e.message }); }
}

// UPDATED: Terima param filter (status + search) dari frontend
function getPODashboardPage(page, filter) {
  try { return JSON.stringify(getPODashboardData(page, filter)); } 
  catch (e) { return JSON.stringify({ error: e.message }); }
}

// ─────────────────────────────────────────────────────────────────────────
// AMBIL MASTER DATA UNTUK AUTOCOMPLETE SEARCH
// ─────────────────────────────────────────────────────────────────────────
function getVendorList() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('M_VENDOR');
  if(!sh) return [];
  var data = sh.getDataRange().getValues();
  var vendors = [];
  for(var i = 1; i < data.length; i++) {
    if(data[i][0]) vendors.push({ code: String(data[i][0]).trim(), name: String(data[i][1]).trim() });
  }
  return vendors;
}

function getItemList() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('M_ITEM');
  if(!sh) return [];
  var data = sh.getDataRange().getValues();
  var items = [];
  
  for(var i = 1; i < data.length; i++) {
    if(data[i][0]) {
      items.push({ 
        code: String(data[i][0]).trim(),                 // Index 0: Item_Code
        desc: String(data[i][2]).trim(),                 // Index 2: Description
        uom: String(data[i][3]).trim(),                  // Index 3: Unit of Measure
        wg: parseFloat(data[i][10]) || 0                 // Index 10: Wg/Pce FC
      });
    }
  }
  return items;
}

// ─────────────────────────────────────────────────────────────────────────
// LEGACY: getOpenPOPool — tidak dipanggil dari frontend baru,
// tapi dipertahankan untuk backward compatibility.
// ─────────────────────────────────────────────────────────────────────────
function getOpenPOPool() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var poSheet = ss.getSheetByName('PO');
  if (!poSheet) return [];
  
  var poData = poSheet.getDataRange().getValues();
  if (poData.length <= 1) return [];
  
  var headers = poData[0].map(function(h) { return String(h).trim(); });
  var pool = [];
  var today = new Date();
  today.setHours(0,0,0,0);

  for (var i = 1; i < poData.length; i++) {
    var status = String(poData[i][headers.indexOf('STATUS')]).trim().toUpperCase();
    if (status !== 'OPEN') continue;

    var poNo = String(poData[i][headers.indexOf('PO_No')]).trim();
    if (!poNo || poNo === "") continue;

    var schDate = poData[i][headers.indexOf('Schedule')];
    var urgency = 'GREEN'; 
    var daysDiff = null;

    if (schDate instanceof Date) {
      var targetDate = new Date(schDate);
      targetDate.setHours(0,0,0,0);
      daysDiff = Math.round((targetDate - today) / (1000 * 60 * 60 * 24));
      if (daysDiff < 0) urgency = 'RED'; 
      else if (daysDiff <= 2) urgency = 'YELLOW'; 
    }

    var schDateStr = schDate instanceof Date ? Utilities.formatDate(schDate, Session.getScriptTimeZone(), 'dd MMM yy') : '-';

    pool.push({
      po_no: poNo,
      vendor: String(poData[i][headers.indexOf('Vendor')]).trim(),
      schedule_str: schDateStr,
      item_code: String(poData[i][headers.indexOf('Item_Code')]).trim(),
      po_kg: parseFloat(poData[i][headers.indexOf('PO_KG')]) || 0,
      gr_kg: parseFloat(poData[i][headers.indexOf('GR_KG')]) || 0,
      bl_kg: parseFloat(poData[i][headers.indexOf('BL_KG')]) || 0,
      owner: String(poData[i][headers.indexOf('Owner')]).trim(),
      urgency: urgency,
      days_diff: daysDiff
    });
  }

  pool.sort(function(a, b) {
    var score = { 'RED': 3, 'YELLOW': 2, 'GREEN': 1 };
    return score[b.urgency] - score[a.urgency];
  });

  return pool;
}

// ─────────────────────────────────────────────────────────────────────────
// AMBIL DATA LIST PO — GROUPED BY PO_No, DENGAN FILTER STATUS & SEARCH
//
// filter: { status: 'ALL'|'OPEN'|'OVERDUE'|'CLOSED', search: '<keyword>' }
//
// Return shape:
// {
//   data: [
//     {
//       po_no, po_date, vendor, schedule, owner,
//       overall_status: 'OPEN'|'OVERDUE'|'CLOSED',
//       total_items, total_q, total_kg,
//       items: [{ item_code, description, uom, wg_pce,
//                 po_q, po_kg, gr_q, gr_kg, bl_q, bl_kg,
//                 target_loc, status }]
//     }
//   ],
//   totalPages, currentPage, totalItems
// }
// ─────────────────────────────────────────────────────────────────────────
function getPODashboardData(page, filter) {
  var pageSize = 30;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('PO');
  if (!sh) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };

  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };

  var h = data[0].map(function(x) { return String(x).trim(); });
  var iPO   = h.indexOf('PO_No');
  var iDate = h.indexOf('PO_Date');
  var iVen  = h.indexOf('Vendor');
  var iSch  = h.indexOf('Schedule');
  var iCode = h.indexOf('Item_Code');
  var iDesc = h.indexOf('Description');
  var iUom  = h.indexOf('UoM');
  var iWg   = h.indexOf('Wg/Pce');
  var iPoQ  = h.indexOf('PO_Q');
  var iPoKg = h.indexOf('PO_KG');
  var iGrQ  = h.indexOf('GR_Q');
  var iGrKg = h.indexOf('GR_KG');
  var iBlQ  = h.indexOf('BL_Q');
  var iBlKg = h.indexOf('BL_KG');
  var iSts  = h.indexOf('STATUS');
  var iOwn  = h.indexOf('Owner');
  var iTgt  = h.indexOf('Target_Loc');

  var tz = Session.getScriptTimeZone();
  var today = new Date();
  today.setHours(0,0,0,0);

  // Group by PO_No
  var group = {};
  var order = [];

  for (var i = 1; i < data.length; i++) {
    var poNo = String(data[i][iPO] || '').trim();
    if (!poNo) continue;

    var itmStatusRaw = String(data[i][iSts] || '').trim().toUpperCase();
    var schDate = data[i][iSch];
    var schStr = schDate instanceof Date ? Utilities.formatDate(schDate, tz, 'dd MMM yy') : '-';

    // Escalate ke OVERDUE bila schedule sudah lewat dan masih OPEN
    var itmOverall = itmStatusRaw;
    if (itmStatusRaw === 'OPEN' && schDate instanceof Date) {
      var target = new Date(schDate); target.setHours(0,0,0,0);
      if (target < today) itmOverall = 'OVERDUE';
    }

    if (!group[poNo]) {
      var rawTgl = data[i][iDate];
      var tglStr = rawTgl instanceof Date ? Utilities.formatDate(rawTgl, tz, 'dd MMM yy') : String(rawTgl);
      group[poNo] = {
        po_no: poNo,
        po_date: tglStr,
        po_date_raw: rawTgl instanceof Date ? rawTgl.getTime() : 0,
        vendor: String(data[i][iVen] || '').trim(),
        schedule: schStr,
        schedule_raw: schDate instanceof Date ? schDate.getTime() : 0,
        owner: String(data[i][iOwn] || '').trim(),
        items: [],
        overall_status: 'CLOSED',
        total_items: 0,
        total_q: 0,
        total_kg: 0
      };
      order.push(poNo);
    }

    var gp = group[poNo];
    var itmPoQ  = parseInt(data[i][iPoQ]) || 0;
    var itmPoKg = parseFloat(data[i][iPoKg]) || 0;

    gp.items.push({
      item_code: String(data[i][iCode] || '').trim(),
      description: String(data[i][iDesc] || '').trim(),
      uom: iUom !== -1 ? String(data[i][iUom] || '').trim() : '',
      wg_pce: parseFloat(data[i][iWg]) || 0,
      po_q: itmPoQ,
      po_kg: itmPoKg,
      gr_q: parseInt(data[i][iGrQ]) || 0,
      gr_kg: parseFloat(data[i][iGrKg]) || 0,
      bl_q: parseInt(data[i][iBlQ]) || 0,
      bl_kg: parseFloat(data[i][iBlKg]) || 0,
      target_loc: iTgt !== -1 ? String(data[i][iTgt] || '').trim() : '',
      status: itmOverall
    });

    gp.total_items++;
    gp.total_q  += itmPoQ;
    gp.total_kg += itmPoKg;

    // Escalate overall status: OVERDUE > OPEN > CLOSED
    if (itmOverall === 'OVERDUE') gp.overall_status = 'OVERDUE';
    else if (itmOverall === 'OPEN' && gp.overall_status !== 'OVERDUE') gp.overall_status = 'OPEN';
  }

  // Sort by PO_Date DESC (paling baru di atas)
  var allData = order.map(function(k){ return group[k]; });
  allData.sort(function(a,b){ return b.po_date_raw - a.po_date_raw; });

  // Apply filter
  filter = filter || {};
  var fStatus = String(filter.status || '').toUpperCase();
  var fSearch = String(filter.search || '').toLowerCase().trim();

  if (fStatus && fStatus !== 'ALL' && fStatus !== 'SEMUA') {
    allData = allData.filter(function(p){ return p.overall_status === fStatus; });
  }

  if (fSearch) {
    allData = allData.filter(function(p){
      if (p.po_no.toLowerCase().indexOf(fSearch) !== -1) return true;
      if (p.vendor.toLowerCase().indexOf(fSearch) !== -1) return true;
      for (var k = 0; k < p.items.length; k++) {
        if (p.items[k].item_code.toLowerCase().indexOf(fSearch) !== -1) return true;
        if (p.items[k].description.toLowerCase().indexOf(fSearch) !== -1) return true;
      }
      return false;
    });
  }

  var total = allData.length;
  var totalPages = Math.ceil(total / pageSize) || 1;
  var p = Math.max(1, Math.min(page || 1, totalPages));
  var start = (p - 1) * pageSize;
  var paged = allData.slice(start, start + pageSize);

  return { data: paged, totalPages: totalPages, currentPage: p, totalItems: total };
}

// ─────────────────────────────────────────────────────────────────────────
// SIMPAN INPUT PO BARU (DENGAN TARGET_LOC)
// ─────────────────────────────────────────────────────────────────────────
function savePurchaseOrder(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var poSheet = ss.getSheetByName('PO');
    if (!poSheet) throw new Error("Sheet dengan nama 'PO' tidak ditemukan!");
    
    var info = payload.info;
    var items = payload.items; 
    
    var poData = poSheet.getDataRange().getValues();
    var poHeaders = poData[0].map(function(h) { return String(h).trim(); });
    
    var poColIdx = poHeaders.indexOf('PO_No');
    var actualLastRow = 1;
    if (poColIdx !== -1) {
      for (var d = poData.length - 1; d >= 1; d--) {
        if (poData[d][poColIdx] !== "" && poData[d][poColIdx] !== null && poData[d][poColIdx] !== undefined) {
          actualLastRow = d + 1;
          break;
        }
      }
    }
    
    var startWritingRow = actualLastRow + 1;
    var datePOObj = new Date(info.po_date);
    var dateSchObj = new Date(info.schedule);
    
    for (var i = 0; i < items.length; i++) {
      var itm = items[i];
      
      var rowMap = {
        'PO_Date': datePOObj,
        'PO_No': String(info.po_no).trim().toUpperCase(),
        'Vendor': String(info.vendor).trim().toUpperCase(),
        'Schedule': dateSchObj,
        'Item_Code': String(itm.item_code).trim(),
        'Description': String(itm.description).trim(),
        'UoM': String(itm.uom).trim(),
        'Wg/Pce': parseFloat(itm.wg_pce) || 0,
        'PO_Q': parseInt(itm.po_q) || 0,
        'PO_KG': parseFloat(itm.po_kg) || 0,
        'Target_Loc': String(itm.target_loc).trim(),         
        'Owner': String(info.owner).trim().toUpperCase() 
      };
      
      poHeaders.forEach(function(headerName, colIndex) {
        if (rowMap[headerName] !== undefined && rowMap[headerName] !== "") {
          poSheet.getRange(startWritingRow, colIndex + 1).setValue(rowMap[headerName]);
        }
      });
      startWritingRow++;
    }
    
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, total_items: items.length });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}