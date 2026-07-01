// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - DELIVERY & LOAD PLANNING SERVICE
// =========================================================================

function getDeliveryInitData() {
  try {
    return JSON.stringify({
      pool: getReadyToShipPool(),
      dash: getDeliveryDashboardData(1)
    });
  } catch (e) { return JSON.stringify({ error: e.message }); }
}

function getDeliveryDashboardPage(page) {
  try { return JSON.stringify(getDeliveryDashboardData(page)); } 
  catch (e) { return JSON.stringify({ error: e.message }); }
}

// ─────────────────────────────────────────────────────────────────────────
// 1. AMBIL POOL BARANG READY TO SHIP (DENGAN INDIKATOR JADWAL SO)
// ─────────────────────────────────────────────────────────────────────────
function getReadyToShipPool() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fgSheet = ss.getSheetByName('Stok_FG'); 
  var soSheet = ss.getSheetByName('SO');
  
  if (!fgSheet) return [];
  
  var fgData = fgSheet.getDataRange().getValues();
  if (fgData.length <= 1) return [];
  
  var fgHeaders = fgData[0].map(function(h) { return String(h).trim(); });
  var bchIdx = fgHeaders.indexOf('Batch_ID');
  var itemIdx = fgHeaders.indexOf('Item_Code');
  var descIdx = fgHeaders.indexOf('Description');
  
  // Header Spesifik
  var soIdx = fgHeaders.indexOf('SO_Ref'); 
  var qtyBalIdx = fgHeaders.indexOf('Qty_Avail'); 
  var kgBalIdx = fgHeaders.indexOf('KG_Avail'); 
  var statusIdx = fgHeaders.indexOf('Delv_Status'); 
  var spkIdx = fgHeaders.indexOf('SPK_Ref'); 
  var ownIdx = fgHeaders.indexOf('Owner');
  var ownUsedIdx = fgHeaders.indexOf('Owner_Used');

  // Tarik referensi Schedule Date & Cust dari Sheet SO untuk mapping urgensi
  var soMap = {};
  if (soSheet) {
    var soData = soSheet.getDataRange().getValues();
    var soHeaders = soData[0].map(function(h) { return String(h).trim(); });
    var sNoIdx = soHeaders.indexOf('SO_No');
    var sSchIdx = soHeaders.indexOf('SCHEDULE_DATE');
    var sCustIdx = soHeaders.indexOf('Cust');
    var sOwnIdx = soHeaders.indexOf('Owner');
    
    for (var s = 1; s < soData.length; s++) {
      var soNoKey = String(soData[s][sNoIdx]).trim();
      if (soNoKey && !soMap[soNoKey]) {
        soMap[soNoKey] = {
          schedule: soData[s][sSchIdx],
          cust: soData[s][sCustIdx] || 'UNKNOWN',
          owner: soData[s][sOwnIdx] || 'FCM'
        };
      }
    }
  }

  var pool = [];
  var today = new Date();
  today.setHours(0,0,0,0);

  for (var i = 1; i < fgData.length; i++) {
    // 1. Cek Status Pengiriman
    var stat = String(fgData[i][statusIdx]).trim().toUpperCase();
    if (stat !== 'READY' && stat !== 'AVAILABLE' && stat !== 'GUDANG') continue;
    
    // 2. POKAYOKE: Cek Qty & KG Avail (Jika 0 atau minus, abaikan)
    var qtyAktual = parseInt(fgData[i][qtyBalIdx]) || 0;
    var kgAktual = parseFloat(fgData[i][kgBalIdx]) || 0;
    if (qtyAktual <= 0 || kgAktual <= 0) continue;

    // 3. POKAYOKE: Cek Keberadaan SO_Ref (Jika kosong, abaikan / bukan barang siap kirim ke Cust)
    var refSo = String(fgData[i][soIdx]).trim();
    if (!refSo || refSo === "") continue;

    var soInfo = soMap[refSo] || { schedule: '', cust: 'TIDAK DITEMUKAN DI DATA SO', owner: 'FCM' };
    
    // Hitung Urgensi Lampu Lalu Lintas
    var urgency = 'GREEN'; 
    var daysDiff = null;
    if (soInfo.schedule instanceof Date) {
      var targetDate = new Date(soInfo.schedule);
      targetDate.setHours(0,0,0,0);
      daysDiff = Math.round((targetDate - today) / (1000 * 60 * 60 * 24));
      
      if (daysDiff < 0) urgency = 'RED'; 
      else if (daysDiff <= 1) urgency = 'YELLOW'; 
    }

    pool.push({
      batch_id: String(fgData[i][bchIdx]).trim(),
      item_code: String(fgData[i][itemIdx]).trim(),
      description: String(fgData[i][descIdx]).trim(),
      so_no: refSo,
      uom: 'Pcs',
      qty: qtyAktual, 
      kg: kgAktual,   
      cust: soInfo.cust,
      // Owner delivery = Owner_Used (pemilik setelah ICT) → fallback ke Owner asli → fallback SO
      owner: (ownUsedIdx >= 0 ? String(fgData[i][ownUsedIdx]).trim() : '') ||
             String(fgData[i][ownIdx]).trim() ||
             soInfo.owner,
      spk_ref: String(fgData[i][spkIdx] || '').trim(),
      schedule_str: soInfo.schedule instanceof Date ? Utilities.formatDate(soInfo.schedule, Session.getScriptTimeZone(), 'dd MMM yy') : '-',
      urgency: urgency,
      days_diff: daysDiff
    });
  }

  // Sort prioritas urgensi: RED -> YELLOW -> GREEN
  pool.sort(function(a, b) {
    var score = { 'RED': 3, 'YELLOW': 2, 'GREEN': 1 };
    return score[b.urgency] - score[a.urgency];
  });

  return pool;
}

// ─────────────────────────────────────────────────────────────────────────
// 2. AMBIL DATA RIWAYAT SURAT JALAN (DASHBOARD LOG)
// ─────────────────────────────────────────────────────────────────────────
function getDeliveryDashboardData(page) {
  var pageSize = 50;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('DELV');
  if (!sh) return { data: [], totalPages: 1, currentPage: 1 };

  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { data: [], totalPages: 1, currentPage: 1 };

  var h = data[0].map(function(x) { return String(x).trim(); });
  var allData = [];

  var tglIdx = h.indexOf('Tanggal');
  var sjIdx = h.indexOf('SJ_No');
  var soIdx = h.indexOf('SO_No');
  var custIdx = h.indexOf('Cust');
  var itemIdx = h.indexOf('Item_Code');
  var qtyIdx = h.indexOf('Delv_Q');
  var kgIdx = h.indexOf('Delv_KG');
  var driverIdx = h.indexOf('Driver');
  var armadaIdx = h.indexOf('No_Armada');

  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][sjIdx]) continue;
    
    var rawTgl = data[i][tglIdx];
    var tglStr = rawTgl instanceof Date ? Utilities.formatDate(rawTgl, Session.getScriptTimeZone(), 'dd MMM yy') : String(rawTgl);

    allData.push({
      tanggal: tglStr,
      sj_no: String(data[i][sjIdx]).trim(),
      so_no: String(data[i][soIdx]).trim(),
      cust: String(data[i][custIdx]).trim(),
      item_code: String(data[i][itemIdx]).trim(),
      qty: data[i][qtyIdx] || 0,
      kg: parseFloat(data[i][kgIdx]) || 0,
      driver: String(data[i][driverIdx] || '').trim(),
      armada: String(data[i][armadaIdx] || '').trim()
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
// 3. PROSES SIMPAN LOAD PLAN & AUTO SPLIT SURAT JALAN PER SO
// ─────────────────────────────────────────────────────────────────────────
function saveLoadPlanDelivery(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var delvSheet = ss.getSheetByName('DELV');
    if (!delvSheet) throw new Error("Sheet dengan nama 'DELV' tidak ditemukan!");
    
    var info = payload.info;
    var cartItems = payload.items; 
    
    var delvData = delvSheet.getDataRange().getValues();
    var delvHeaders = delvData[0].map(function(h) { return String(h).trim(); });
    
    // Grouping by SO
    var groupedBySo = {};
    cartItems.forEach(function(item) {
      if (!groupedBySo[item.so_no]) groupedBySo[item.so_no] = [];
      groupedBySo[item.so_no].push(item);
    });

    // ⚡ UBAH MENJADI ARRAY DAN URUTKAN BERDASARKAN NAMA CUSTOMER
    var sortedGroups = [];
    for (var soNo in groupedBySo) {
      sortedGroups.push({
        so_no: soNo,
        cust: groupedBySo[soNo][0].cust,
        items: groupedBySo[soNo]
      });
    }
    
    // Proses penyusunan (Sort) alfabetis berdasarkan nama Customer
    sortedGroups.sort(function(a, b) {
      return String(a.cust).localeCompare(String(b.cust));
    });
    
    // Setup Logika Auto-Numbering SJ [YY]1-[NNNN]
    var dateObj = new Date(info.tanggal);
    var yy = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yy');
    var divCode = "1"; // 1 = SSC
    var prefixSJ = "SJ" + yy + divCode + "-"; 
    
    var lastNumber = 0;
    var sjColIdx = delvHeaders.indexOf('SJ_No');
    var actualLastRow = 1;
    
    // Cari Baris Fisik Terakhir sekaligus mencari Nomor SJ Tertinggi di tahun/divisi ini
    if (sjColIdx !== -1) {
      for (var d = 1; d < delvData.length; d++) {
        var existingSJ = String(delvData[d][sjColIdx]).trim();
        if (existingSJ !== "") actualLastRow = d + 1; // Deteksi baris aktual
        
        if (existingSJ.indexOf(prefixSJ) === 0) {
          var numPart = parseInt(existingSJ.substring(prefixSJ.length), 10);
          if (!isNaN(numPart) && numPart > lastNumber) {
            lastNumber = numPart;
          }
        }
      }
    }
    
    var startWritingRow = actualLastRow + 1;
    
    // Iterasi Eksekusi Penulisan (Sudah terurut per Customer)
    for (var g = 0; g < sortedGroups.length; g++) {
      var group = sortedGroups[g];
      
      // Generate Nomor Baru Berurutan
      lastNumber++;
      var runningStr = ("0000" + lastNumber).slice(-4); 
      var generatedSJ = prefixSJ + runningStr; 
      
      for (var i = 0; i < group.items.length; i++) {
        var itm = group.items[i];
        
        var rowMap = {
          'Tanggal': dateObj,
          'Item_Code': String(itm.item_code).trim(),
          'Description': String(itm.description).trim(),
          'SO_No': String(itm.so_no).trim(),
          'SJ_No': generatedSJ,
          'UoM': String(itm.uom).trim(),
          'Delv_Q': parseInt(itm.qty),
          'Delv_KG': parseFloat(itm.kg),
          'Cust': String(itm.cust).trim(),
          'Owner': String(itm.owner).trim(),
          'Spk_ref': String(itm.spk_ref).trim(),
          'Batch_ID': String(itm.batch_id).trim(),
          'No_Armada': String(info.no_armada).trim().toUpperCase(),
          'Driver': String(info.driver).trim().toUpperCase(),
          'Note': String(info.note || '').trim(),
          'Tgl_Kembali': "",      
          'Diterima_Finance': ""  
        };
        
        delvHeaders.forEach(function(headerName, colIndex) {
          if (rowMap[headerName] !== undefined) {
            delvSheet.getRange(startWritingRow, colIndex + 1).setValue(rowMap[headerName]);
          }
        });
        startWritingRow++;
      }
    }
    
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, total_sj: sortedGroups.length });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}