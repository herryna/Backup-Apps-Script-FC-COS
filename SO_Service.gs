// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - SO SERVICE
// Updated: SO_Period support + Close SO with password
// =========================================================================

function getSOInitData() {
  try {
    var result = { form: JSON.parse(getSOFormData()), dash: getSODashboardData(1) };
    return JSON.stringify(result);
  } catch (e) { return JSON.stringify({ error: e.message }); }
}

function getSODashboardPage(page) {
  try { return JSON.stringify(getSODashboardData(page)); } 
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
          wg_fc: rowObj['Wg/Pce FC'] || rowObj['Weight_FC'] || 0
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

function getSODashboardData(page) {
  var pageSize = 50;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SO');
  if (!sh) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };

  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };

  var h = data[0].map(function(x) { return String(x).trim(); });
  var allData = [];

  var dtIdx       = h.indexOf('SO_DATE');
  var periodIdx   = h.indexOf('SO_Period');
  var soIdx       = h.indexOf('SO_No');
  var custIdx     = h.indexOf('Cust');
  var schIdx      = h.indexOf('SCHEDULE_DATE');
  var itemIdx     = h.indexOf('Item_Code');
  var descIdx     = h.indexOf('Description');
  var qtyIdx      = h.indexOf('SO_Q');
  var kgIdx       = h.indexOf('SO_KG'); 
  var delvQIdx    = h.indexOf('Delv_Q');
  var statusIdx   = h.indexOf('STATUS');
  var noteIdx     = h.indexOf('Note');
  var closeFlagIdx= h.indexOf('Close_Flag');
  var closeDtIdx  = h.indexOf('Close_Date');
  var closeRsnIdx = h.indexOf('Close_Reason');

  for (var i = 1; i < data.length; i++) {
    if (!data[i][soIdx]) continue; 
    
    var rawDt = data[i][dtIdx];
    var dtStr = rawDt instanceof Date ? Utilities.formatDate(rawDt, Session.getScriptTimeZone(), 'dd MMM yy') : String(rawDt);
    var rawSch = data[i][schIdx];
    var schStr = rawSch instanceof Date ? Utilities.formatDate(rawSch, Session.getScriptTimeZone(), 'dd MMM yy') : String(rawSch);
    var rawCloseDt = closeDtIdx >= 0 ? data[i][closeDtIdx] : '';
    var closeDtStr = rawCloseDt instanceof Date ? Utilities.formatDate(rawCloseDt, Session.getScriptTimeZone(), 'dd MMM yy') : String(rawCloseDt || '');

    var rawKg = parseFloat(data[i][kgIdx]) || 0;
    var roundedKg = parseFloat(rawKg.toFixed(2));

    allData.push({
      so_date:        dtStr,
      so_date_raw:    rawDt instanceof Date ? rawDt.toISOString().split('T')[0] : '',
      so_period:      periodIdx >= 0 ? String(data[i][periodIdx] || '').trim() : '',
      so_no:          String(data[i][soIdx]).trim(),
      cust:           String(data[i][custIdx]).trim(),
      schedule:       schStr,
      schedule_raw:   rawSch instanceof Date ? rawSch.toISOString().split('T')[0] : '',
      item_code:      String(data[i][itemIdx]).trim(),
      desc:           String(data[i][descIdx]).trim(),
      qty:            data[i][qtyIdx] || 0,
      kg:             roundedKg, 
      delv_q:         delvQIdx >= 0 ? (data[i][delvQIdx] || 0) : 0,
      status:         statusIdx >= 0 ? String(data[i][statusIdx] || '').trim().toUpperCase() : '',
      note:           data[i][noteIdx] || '',
      close_flag:     closeFlagIdx >= 0 ? (data[i][closeFlagIdx] === true) : false,
      close_date:     closeDtStr,
      close_reason:   closeRsnIdx >= 0 ? String(data[i][closeRsnIdx] || '').trim() : '',
      sort_val:       rawDt instanceof Date ? rawDt.getTime() : 0
    });
  }

  allData.sort(function(a, b) { return b.sort_val - a.sort_val; });
  var total = allData.length;
  var totalPages = Math.ceil(total / pageSize) || 1;
  var p = Math.max(1, Math.min(page || 1, totalPages));
  var start = (p - 1) * pageSize;
  var paged = allData.slice(start, start + pageSize);

  paged.sort(function(a, b) {
    if (a.sort_val === b.sort_val) return a.so_no.localeCompare(b.so_no);
    return a.sort_val - b.sort_val;
  });

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
        'T':             parseFloat(item.t), 'P': parseFloat(item.p), 'L': parseFloat(item.l),
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