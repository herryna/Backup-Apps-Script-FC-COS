// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - SO SERVICE
// Updated: SO_Period support + Close SO with password
// =========================================================================

function getSOInitData() {
  try {
    var result = { 
      form: JSON.parse(getSOFormData()), 
      dash: getSODashboardData(1, null) 
    };
    return JSON.stringify(result);
  } catch (e) { return JSON.stringify({ error: e.message }); }
}

function getSODashboardPage(page, filter) {
  try { return JSON.stringify(getSODashboardData(page, filter)); } 
  catch (e) { return JSON.stringify({ error: e.message }); }
}

function getSOFormData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var items = [], customers = [];
  
  var itemSheet = ss.getSheetByName('M_ITEM');
  if (itemSheet) {
    var iData = itemSheet.getDataRange().getValues();
    var iHeaders = iData[0].map(function(h) { return String(h).trim(); });
    for (var i = 1; i < iData.length; i++) {
      var rowObj = {}; iHeaders.forEach(function(h, idx) { rowObj[h] = iData[i][idx]; });
      if (rowObj['Item_Code']) {
        items.push({
          item_code: String(rowObj['Item_Code']).trim(),
          description: String(rowObj['Description']).trim(),
          spec: String(rowObj['Spec']||'').trim(),
          t: rowObj['T']||0, p: rowObj['P']||0, l: rowObj['L']||0,
          uom: String(rowObj['Unit of Measure']||rowObj['UoM_MC']||'Pcs').trim(),
          wg_fc: parseFloat(rowObj['Wg/Pce FC'] || rowObj['Weight_FC'] || 0) || 0,
          wg_cust: parseFloat(rowObj['Wg/Pce Cust'] || 0) || 0
        });
      }
    }
  }
  
  var custSheet = ss.getSheetByName('M_CUST');
  if (custSheet) {
    var cData = custSheet.getDataRange().getValues();
    var cHeaders = cData[0].map(function(h) { return String(h).trim(); });
    var codeIdx = cHeaders.indexOf('C_CODE');
    var initIdx = cHeaders.indexOf('C_INITIAL');
    var nameIdx = cHeaders.indexOf('CUSTOMER_NAME');
    
    for (var j = 1; j < cData.length; j++) {
      if (cData[j][codeIdx]) {
        customers.push({
          code: String(cData[j][codeIdx]).trim(),
          initial: String(cData[j][initIdx] || '').trim(),
          nama: String(cData[j][nameIdx]).trim()
        });
      }
    }
  }
  return JSON.stringify({ items: items, customers: customers });
}

// ─────────────────────────────────────────────────────────────────────────
// GET SO DASHBOARD — GROUPED BY SO_No, DENGAN FILTER STATUS & SEARCH
//
// filter: { status: 'ALL'|'OPEN'|'OVERDUE'|'OVER'|'CLOSED', search: '<keyword>' }
//
// Return shape:
// {
//   data: [
//     {
//       so_no, so_date, so_date_raw, so_period, cust, schedule, schedule_raw,
//       owner_used, overall_status, close_date, close_reason,
//       total_items, total_q, total_kg, total_delv_q, total_bl_q,
//       items: [{ item_code, description, spec, t, p, l, uom, so_q, so_kg,
//                 delv_q, delv_kg, bl_q, bl_kg, note, ref_spk, status }]
//     }
//   ],
//   totalPages, currentPage, totalItems
// }
// ─────────────────────────────────────────────────────────────────────────
function getSODashboardData(page, filter) {
  var pageSize = 30;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SO');
  if (!sh) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };

  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };

  var h = data[0].map(function(x) { return String(x).trim(); });
  
  // Load M_ITEM TYPE map untuk enrichment (Sheet/Part/Coil detection)
  var itemTypeMap = {};
  var itemSheet = ss.getSheetByName('M_ITEM');
  if (itemSheet) {
    var iData = itemSheet.getDataRange().getValues();
    var iH = iData[0].map(function(x){ return String(x).trim(); });
    var iCodeIdx = iH.indexOf('Item_Code');
    var iTypeIdx = iH.indexOf('TYPE');
    if (iCodeIdx >= 0 && iTypeIdx >= 0) {
      for (var m = 1; m < iData.length; m++) {
        if (iData[m][iCodeIdx]) {
          itemTypeMap[String(iData[m][iCodeIdx]).trim()] = String(iData[m][iTypeIdx] || '').trim();
        }
      }
    }
  }
  var iDate     = h.indexOf('SO_DATE');
  var iPeriod   = h.indexOf('SO_Period');
  var iSoNo     = h.indexOf('SO_No');
  var iCust     = h.indexOf('Cust');
  var iSch      = h.indexOf('SCHEDULE_DATE');
  var iCode     = h.indexOf('Item_Code');
  var iDesc     = h.indexOf('Description');
  var iSpec     = h.indexOf('Spec');
  var iT        = h.indexOf('T');
  var iP        = h.indexOf('P');
  var iL        = h.indexOf('L');
  var iUom      = h.indexOf('UoM');
  var iSoQ      = h.indexOf('SO_Q');
  var iSoKg     = h.indexOf('SO_KG');
  var iDelvQ    = h.indexOf('Delv_Q');
  var iDelvKg   = h.indexOf('Delv_KG');
  var iBlQ      = h.indexOf('BL_Q');
  var iBlKg     = h.indexOf('BL_KG');
  var iSts      = h.indexOf('STATUS');
  var iRefSpk   = h.indexOf('Ref_SPK');
  var iNote     = h.indexOf('Note');
  var iOwnUsed  = h.indexOf('Owner_Used');
  var iCloseFlag= h.indexOf('Close_Flag');
  var iCloseDt  = h.indexOf('Close_Date');
  var iCloseRsn = h.indexOf('Close_Reason');

  var tz = Session.getScriptTimeZone();
  var today = new Date();
  today.setHours(0,0,0,0);

  // Group by SO_No
  var group = {};
  var order = [];

  for (var i = 1; i < data.length; i++) {
    var soNo = String(data[i][iSoNo] || '').trim();
    if (!soNo) continue;

    // Normalize legacy 'DONE' → 'CLOSED' (untuk transisi selama formula belum di-update)
    var itmStatusRaw = String(data[i][iSts] || '').trim().toUpperCase();
    if (itmStatusRaw === 'DONE') itmStatusRaw = 'CLOSED';

    var schDate = data[i][iSch];
    var schStr = schDate instanceof Date ? Utilities.formatDate(schDate, tz, 'dd MMM yy') : '-';

    // Escalate ke OVERDUE bila schedule sudah lewat dan status masih OPEN
    var itmOverall = itmStatusRaw;
    if (itmStatusRaw === 'OPEN' && schDate instanceof Date) {
      var target = new Date(schDate); target.setHours(0,0,0,0);
      if (target < today) itmOverall = 'OVERDUE';
    }

    if (!group[soNo]) {
      var rawTgl = data[i][iDate];
      var tglStr = rawTgl instanceof Date ? Utilities.formatDate(rawTgl, tz, 'dd MMM yy') : String(rawTgl);
      var rawCloseDt = iCloseDt >= 0 ? data[i][iCloseDt] : '';
      var closeDtStr = rawCloseDt instanceof Date ? Utilities.formatDate(rawCloseDt, tz, 'dd MMM yy') : String(rawCloseDt || '');
      
      group[soNo] = {
        so_no: soNo,
        so_date: tglStr,
        so_date_raw: rawTgl instanceof Date ? rawTgl.toISOString().split('T')[0] : '',
        so_date_sort: rawTgl instanceof Date ? rawTgl.getTime() : 0,
        so_period: iPeriod >= 0 ? String(data[i][iPeriod] || '').trim() : '',
        cust: String(data[i][iCust] || '').trim(),
        schedule: schStr,
        schedule_raw: schDate instanceof Date ? schDate.toISOString().split('T')[0] : '',
        owner_used: iOwnUsed >= 0 ? String(data[i][iOwnUsed] || '').trim() : '',
        close_date: closeDtStr,
        close_reason: iCloseRsn >= 0 ? String(data[i][iCloseRsn] || '').trim() : '',
        items: [],
        overall_status: 'CLOSED',
        total_items: 0,
        total_q: 0,
        total_kg: 0,
        total_delv_q: 0,
        total_bl_q: 0
      };
      order.push(soNo);
    }

    var gp = group[soNo];
    var itmSoQ  = parseInt(data[i][iSoQ]) || 0;
    var itmSoKg = parseFloat(data[i][iSoKg]) || 0;
    var itmDelvQ = parseInt(data[i][iDelvQ]) || 0;
    var itmBlQ  = parseInt(data[i][iBlQ]) || 0;

    var itmCode = String(data[i][iCode] || '').trim();
    gp.items.push({
      item_code:   itmCode,
      description: String(data[i][iDesc] || '').trim(),
      spec:        iSpec >= 0 ? String(data[i][iSpec] || '').trim() : '',
      t:           data[i][iT] || '',
      p:           data[i][iP] || '',
      l:           data[i][iL] || '',
      uom:         iUom >= 0 ? String(data[i][iUom] || '').trim() : 'Pcs',
      type:        itemTypeMap[itmCode] || '',
      so_q:        itmSoQ,
      so_kg:       itmSoKg,
      delv_q:      itmDelvQ,
      delv_kg:     iDelvKg >= 0 ? parseFloat(data[i][iDelvKg]) || 0 : 0,
      bl_q:        itmBlQ,
      bl_kg:       iBlKg >= 0 ? parseFloat(data[i][iBlKg]) || 0 : 0,
      note:        iNote >= 0 ? String(data[i][iNote] || '').trim() : '',
      ref_spk:     iRefSpk >= 0 ? String(data[i][iRefSpk] || '').trim() : '',
      status:      itmOverall
    });

    gp.total_items++;
    gp.total_q      += itmSoQ;
    gp.total_kg     += itmSoKg;
    gp.total_delv_q += itmDelvQ;
    gp.total_bl_q   += itmBlQ;

    // Escalate overall status: OVERDUE > OVER > OPEN > CLOSED
    if (itmOverall === 'OVERDUE') {
      gp.overall_status = 'OVERDUE';
    } else if (itmOverall === 'OVER' && gp.overall_status !== 'OVERDUE') {
      gp.overall_status = 'OVER';
    } else if (itmOverall === 'OPEN' && gp.overall_status !== 'OVERDUE' && gp.overall_status !== 'OVER') {
      gp.overall_status = 'OPEN';
    }
  }

  // Sort by SO_Date DESC (paling baru di atas), tie-break by so_no ASC
  var allData = order.map(function(k){ return group[k]; });
  allData.sort(function(a,b){ 
    if (b.so_date_sort !== a.so_date_sort) return b.so_date_sort - a.so_date_sort;
    return a.so_no.localeCompare(b.so_no);
  });

  // Apply filter
  filter = filter || {};
  var fStatus = String(filter.status || '').toUpperCase();
  var fSearch = String(filter.search || '').toLowerCase().trim();

  if (fStatus && fStatus !== 'ALL' && fStatus !== 'SEMUA') {
    allData = allData.filter(function(p){ return p.overall_status === fStatus; });
  }

  if (fSearch) {
    allData = allData.filter(function(p){
      if (p.so_no.toLowerCase().indexOf(fSearch) !== -1) return true;
      if (p.cust.toLowerCase().indexOf(fSearch) !== -1) return true;
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

  // Clean sort field before return (frontend doesn't need it)
  paged.forEach(function(g){ delete g.so_date_sort; });

  return { data: paged, totalPages: totalPages, currentPage: p, totalItems: total };
}

function saveSalesOrder(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SO');
    if (!sheet) throw new Error("Sheet dengan nama 'SO' tidak ditemukan!");
    
    var head = payload.header; var list = payload.items;
    
    // Validasi SO_Period wajib + format YYYY-MM
    if (!head.so_period || !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(head.so_period).trim())) {
      throw new Error("Periode SO wajib diisi dengan format YYYY-MM (contoh: 2026-07).");
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    
    var soIdx = headers.indexOf('SO_No');
    var newSoNo = String(head.so_no).trim().toLowerCase();
    
    // SOLUSI BUG ARRAY FORMULA: Cari manual baris terakhir dari kolom SO_No
    var actualLastRow = 1; 
    if (soIdx !== -1) {
      for (var d = data.length - 1; d >= 1; d--) {
        if (data[d][soIdx] !== "" && data[d][soIdx] !== null && data[d][soIdx] !== undefined) {
          actualLastRow = d + 1;
          break;
        }
      }
      
      // VALIDASI ANTI-DUPLIKAT
      for (var row = 1; row < actualLastRow; row++) {
        if (String(data[row][soIdx]).trim().toLowerCase() === newSoNo) {
          throw new Error("No. SO [" + head.so_no + "] SUDAH ADA di database! Gunakan nomor urut lain.");
        }
      }
    }
    
    var startRow = actualLastRow + 1;
    
    // Guard: Wg/Pce Cust wajib > 0 (auto-fetch dari M_ITEM). 
    // Kalau item belum ada di Master Data, block save.
    for (var v = 0; v < list.length; v++) {
      var wgCustCheck = parseFloat(list[v].wg_cust || 0);
      if (!wgCustCheck || wgCustCheck <= 0) {
        throw new Error("Item [" + list[v].item_code + "] belum ada 'Wg/Pce Cust' di M_ITEM. Lengkapi Master Data dulu sebelum simpan SO.");
      }
    }
    
    for (var i = 0; i < list.length; i++) {
      var item = list[i];
      
      var rowMap = {
        'SO_DATE':       new Date(head.so_date),
        'SO_Period':     String(head.so_period).trim(),
        'SO_No':         String(head.so_no).trim(),
        'Cust':          String(head.cust_nama).trim(),
        'SCHEDULE_DATE': new Date(head.schedule_date),
        'Item_Code':     String(item.item_code).trim(),
        'Description':   String(item.description).trim(),
        'Spec':          String(item.spec).trim(),
        'T':             parseFloat(item.t),
        'P':             parseFloat(item.p),
        'L':             (function(v){
                            if (v === '' || v === null || v === undefined) return '';
                            var n = parseFloat(v);
                            return isNaN(n) ? String(v).trim() : n;
                         })(item.l),
        'UoM':           String(item.uom).trim(),
        'Wg/Pce FC':     parseFloat(item.wg_fc || 0),
        'Wg/Pce Cust':   parseFloat(item.wg_cust || 0),
        'SO_Q':          parseInt(item.so_q),
        'Note':          String(item.note || '').trim(),
        'Close_Flag':    false  // Default false saat create
      };
      
      var targetRow = startRow + i;
      headers.forEach(function(h, index) {
        if (rowMap[h] !== undefined) {
          sheet.getRange(targetRow, index + 1).setValue(rowMap[h]);
        }
      });
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, count: list.length, so_no: head.so_no });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}

function updateSingleSO(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('SO');
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(x) { return String(x).trim(); });
    
    var soIdx = h.indexOf('SO_No');
    var itemIdx = h.indexOf('Item_Code');
    
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][soIdx]).trim() === String(payload.so_no).trim() && 
          String(data[i][itemIdx]).trim() === String(payload.item_code).trim()) {
        targetRow = i + 1; break;
      }
    }
    
    if (targetRow === -1) throw new Error("Data Item SO tidak ditemukan di database!");
    
    if (payload.qty !== undefined && payload.qty !== "" && h.indexOf('SO_Q') >= 0) {
      sh.getRange(targetRow, h.indexOf('SO_Q') + 1).setValue(parseInt(payload.qty));
    }
    if (payload.note !== undefined && h.indexOf('Note') >= 0) {
      sh.getRange(targetRow, h.indexOf('Note') + 1).setValue(String(payload.note));
    }
    if (payload.so_period !== undefined && payload.so_period !== "" && h.indexOf('SO_Period') >= 0) {
      // Validasi format
      if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(payload.so_period).trim())) {
        throw new Error("Format SO_Period salah. Gunakan YYYY-MM (contoh: 2026-07).");
      }
      sh.getRange(targetRow, h.indexOf('SO_Period') + 1).setValue(String(payload.so_period).trim());
    }
    
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}

// =========================================================================
// CLOSE SO — Manual close oleh PPIC dengan password verify
// Payload: { so_no, reason, password }
// =========================================================================
function closeSalesOrder(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    // 1. Validate password
    var storedPwd = PropertiesService.getScriptProperties().getProperty('SO_CLOSE_PASSWORD');
    if (!storedPwd) throw new Error("Password Close SO belum di-setup di Script Properties.");
    if (String(payload.password || '').trim() !== storedPwd) {
      throw new Error("Password salah. Konfirmasi password sebelum close SO.");
    }
    
    // 2. Validate inputs
    var soNo = String(payload.so_no || '').trim();
    if (!soNo) throw new Error("No. SO wajib diisi.");
    var reason = String(payload.reason || '').trim();
    if (!reason) throw new Error("Alasan close wajib diisi.");
    
    // 3. Find all rows with this SO_No (multi-item SO)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('SO');
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(x) { return String(x).trim(); });
    
    var soIdx          = h.indexOf('SO_No');
    var closeFlagIdx   = h.indexOf('Close_Flag');
    var closeDtIdx     = h.indexOf('Close_Date');
    var closeReasonIdx = h.indexOf('Close_Reason');
    
    if (closeFlagIdx < 0)   throw new Error("Kolom 'Close_Flag' tidak ditemukan di sheet SO.");
    if (closeDtIdx < 0)     throw new Error("Kolom 'Close_Date' tidak ditemukan di sheet SO.");
    if (closeReasonIdx < 0) throw new Error("Kolom 'Close_Reason' tidak ditemukan di sheet SO.");
    
    var now = new Date();
    var affectedRows = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][soIdx]).trim() === soNo) {
        sh.getRange(i + 1, closeFlagIdx + 1).setValue(true);
        sh.getRange(i + 1, closeDtIdx + 1).setValue(now);
        sh.getRange(i + 1, closeReasonIdx + 1).setValue(reason);
        affectedRows++;
      }
    }
    
    if (affectedRows === 0) throw new Error("SO [" + soNo + "] tidak ditemukan di database.");
    
    SpreadsheetApp.flush();
    return JSON.stringify({ 
      success: true, 
      so_no: soNo, 
      affected_items: affectedRows,
      close_date: Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm')
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}

// =========================================================================
// REOPEN SO — Undo close (kalau salah close). Tetap butuh password.
// Payload: { so_no, password }
// =========================================================================
function reopenSalesOrder(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var storedPwd = PropertiesService.getScriptProperties().getProperty('SO_CLOSE_PASSWORD');
    if (!storedPwd) throw new Error("Password belum di-setup.");
    if (String(payload.password || '').trim() !== storedPwd) {
      throw new Error("Password salah.");
    }
    
    var soNo = String(payload.so_no || '').trim();
    if (!soNo) throw new Error("No. SO wajib diisi.");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('SO');
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(x) { return String(x).trim(); });
    
    var soIdx          = h.indexOf('SO_No');
    var closeFlagIdx   = h.indexOf('Close_Flag');
    var closeDtIdx     = h.indexOf('Close_Date');
    var closeReasonIdx = h.indexOf('Close_Reason');
    
    var affectedRows = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][soIdx]).trim() === soNo) {
        sh.getRange(i + 1, closeFlagIdx + 1).setValue(false);
        sh.getRange(i + 1, closeDtIdx + 1).setValue('');
        sh.getRange(i + 1, closeReasonIdx + 1).setValue('');
        affectedRows++;
      }
    }
    
    if (affectedRows === 0) throw new Error("SO [" + soNo + "] tidak ditemukan.");
    
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, so_no: soNo, affected_items: affectedRows });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}

// =========================================================================
// SPRINT B: TAMBAH ITEM BARU KE SO EXISTING (Append row baru dg SO_No sama)
// Payload: { so_no, item: { item_code, description, spec, t, p, l, uom,
//                            wg_fc, wg_cust, so_q, note } }
// Guard: SO belum fully CLOSED, wg_cust > 0
// =========================================================================
function addItemToExistingSO(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var soNo = String(payload.so_no || '').trim();
    var item = payload.item || {};
    if (!soNo) throw new Error("No. SO wajib diisi.");
    if (!item.item_code) throw new Error("Item Code wajib diisi.");
    if (!item.so_q || parseInt(item.so_q) <= 0) throw new Error("Qty Order wajib > 0.");
    
    var wgCust = parseFloat(item.wg_cust || 0);
    if (!wgCust || wgCust <= 0) {
      throw new Error("Item [" + item.item_code + "] belum ada 'Wg/Pce Cust' di M_ITEM. Lengkapi Master Data dulu.");
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('SO');
    if (!sh) throw new Error("Sheet 'SO' tidak ditemukan.");
    
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(x){ return String(x).trim(); });
    
    var iSoNo      = h.indexOf('SO_No');
    var iDate      = h.indexOf('SO_DATE');
    var iPeriod    = h.indexOf('SO_Period');
    var iCust      = h.indexOf('Cust');
    var iSch       = h.indexOf('SCHEDULE_DATE');
    var iCloseFlag = h.indexOf('Close_Flag');
    
    // Cari sample row untuk copy header (SO_DATE, Cust, Periode, Schedule)
    // + guard: SO belum fully CLOSED + cari posisi append
    var sampleRow = null;
    var isFullyClosed = true;
    var actualLastRow = 1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][iSoNo] !== "" && data[i][iSoNo] !== null && data[i][iSoNo] !== undefined) {
        actualLastRow = i + 1;
      }
      if (String(data[i][iSoNo]).trim() === soNo) {
        if (!sampleRow) sampleRow = data[i];
        if (data[i][iCloseFlag] !== true) isFullyClosed = false;
      }
    }
    
    if (!sampleRow) throw new Error("SO [" + soNo + "] tidak ditemukan di database.");
    if (isFullyClosed) throw new Error("SO [" + soNo + "] sudah CLOSED — tidak bisa tambah item. Reopen dulu jika ingin diedit.");
    
    var newRow = actualLastRow + 1;
    var rowMap = {
      'SO_DATE':       sampleRow[iDate],
      'SO_Period':     sampleRow[iPeriod],
      'SO_No':         soNo,
      'Cust':          sampleRow[iCust],
      'SCHEDULE_DATE': sampleRow[iSch],
      'Item_Code':     String(item.item_code).trim(),
      'Description':   String(item.description || '').trim(),
      'Spec':          String(item.spec || '').trim(),
      'T':             parseFloat(item.t) || '',
      'P':             parseFloat(item.p) || '',
      'L':             (function(v){
                          if (v === '' || v === null || v === undefined) return '';
                          var n = parseFloat(v);
                          return isNaN(n) ? String(v).trim() : n;
                       })(item.l),
      'UoM':           String(item.uom || 'Pcs').trim(),
      'Wg/Pce FC':     parseFloat(item.wg_fc || 0),
      'Wg/Pce Cust':   wgCust,
      'SO_Q':          parseInt(item.so_q),
      'Note':          String(item.note || '').trim(),
      'Close_Flag':    false
    };
    
    h.forEach(function(hd, idx) {
      if (rowMap[hd] !== undefined) {
        sh.getRange(newRow, idx + 1).setValue(rowMap[hd]);
      }
    });
    
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, so_no: soNo, item_code: item.item_code });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}

// =========================================================================
// SPRINT B: UPDATE QTY UNTUK 1 ITEM DI SO EXISTING
// Payload: { so_no, item_code, new_qty }
// Guard: SO belum fully CLOSED, new_qty >= Delv_Q per item
// =========================================================================
function updateSOItemQty(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var soNo = String(payload.so_no || '').trim();
    var itemCode = String(payload.item_code || '').trim();
    var newQty = parseInt(payload.new_qty);
    
    if (!soNo || !itemCode) throw new Error("No. SO dan Item Code wajib diisi.");
    if (isNaN(newQty) || newQty <= 0) throw new Error("Qty baru harus > 0.");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('SO');
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(x){ return String(x).trim(); });
    
    var iSoNo      = h.indexOf('SO_No');
    var iItemCode  = h.indexOf('Item_Code');
    var iSoQ       = h.indexOf('SO_Q');
    var iDelvQ     = h.indexOf('Delv_Q');
    var iCloseFlag = h.indexOf('Close_Flag');
    
    var isFullyClosed = true;
    var targetRow = -1;
    var targetDelvQ = 0;
    var hasMatch = false;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iSoNo]).trim() === soNo) {
        hasMatch = true;
        if (data[i][iCloseFlag] !== true) isFullyClosed = false;
        if (String(data[i][iItemCode]).trim() === itemCode && targetRow === -1) {
          targetRow = i + 1;
          targetDelvQ = parseInt(data[i][iDelvQ]) || 0;
        }
      }
    }
    
    if (!hasMatch) throw new Error("SO [" + soNo + "] tidak ditemukan.");
    if (isFullyClosed) throw new Error("SO [" + soNo + "] sudah CLOSED — tidak bisa edit qty.");
    if (targetRow === -1) throw new Error("Item [" + itemCode + "] tidak ditemukan di SO [" + soNo + "].");
    if (newQty < targetDelvQ) {
      throw new Error("Qty baru (" + newQty + ") tidak boleh < Qty yang sudah dikirim (" + targetDelvQ + ").");
    }
    
    sh.getRange(targetRow, iSoQ + 1).setValue(newQty);
    SpreadsheetApp.flush();
    
    return JSON.stringify({ success: true, so_no: soNo, item_code: itemCode, new_qty: newQty });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}

// =========================================================================
// SPRINT B: UPDATE NOTE UNTUK 1 ITEM DI SO EXISTING
// Payload: { so_no, item_code, new_note }
// Guard: SO belum fully CLOSED
// =========================================================================
function updateSOItemNote(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var soNo = String(payload.so_no || '').trim();
    var itemCode = String(payload.item_code || '').trim();
    var newNote = String(payload.new_note || '').trim();
    
    if (!soNo || !itemCode) throw new Error("No. SO dan Item Code wajib diisi.");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('SO');
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(x){ return String(x).trim(); });
    
    var iSoNo      = h.indexOf('SO_No');
    var iItemCode  = h.indexOf('Item_Code');
    var iNote      = h.indexOf('Note');
    var iCloseFlag = h.indexOf('Close_Flag');
    
    var isFullyClosed = true;
    var targetRow = -1;
    var hasMatch = false;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iSoNo]).trim() === soNo) {
        hasMatch = true;
        if (data[i][iCloseFlag] !== true) isFullyClosed = false;
        if (String(data[i][iItemCode]).trim() === itemCode && targetRow === -1) {
          targetRow = i + 1;
        }
      }
    }
    
    if (!hasMatch) throw new Error("SO [" + soNo + "] tidak ditemukan.");
    if (isFullyClosed) throw new Error("SO [" + soNo + "] sudah CLOSED — tidak bisa edit note.");
    if (targetRow === -1) throw new Error("Item [" + itemCode + "] tidak ditemukan di SO [" + soNo + "].");
    
    sh.getRange(targetRow, iNote + 1).setValue(newNote);
    SpreadsheetApp.flush();
    
    return JSON.stringify({ success: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}