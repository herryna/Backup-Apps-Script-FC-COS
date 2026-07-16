// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - GR SERVICE (SPA DASHBOARD VERSION FINAL)
// =========================================================================

function getActualLastRow(sheet) {
  var colA = sheet.getRange("A:A").getValues();
  for (var i = colA.length - 1; i >= 0; i--) { 
    if (colA[i][0] !== "") return i + 1; 
  }
  return 1;
}

function getBaseBatchId_GR(tglObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('M_Config');
  if (!cfg) throw new Error('Sheet M_Config tidak ditemukan!');

  var tz = Session.getScriptTimeZone();
  var yy = Utilities.formatDate(tglObj, tz, 'yy');
  var configKey = 'LAST_GR_' + yy;   // ex: LAST_GR_26

  var data = cfg.getDataRange().getValues();
  var rowIdx = -1;
  var currentSeq = 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === configKey) {
      rowIdx = i + 1;
      currentSeq = parseInt(data[i][1], 10) || 0;
      break;
    }
  }

  var nextSeq = currentSeq + 1;
  if (rowIdx === -1) {
    // Auto-create key untuk tahun baru (rollover)
    cfg.appendRow([configKey, nextSeq]);
  } else {
    cfg.getRange(rowIdx, 2).setValue(nextSeq);
  }

  return 'GR' + yy + String(nextSeq).padStart(4, '0');
}

function getSuffix(index) {
  var suffix = ''; 
  var q = index;
  do { 
    var r = q % 26; 
    suffix = String.fromCharCode(65 + r) + suffix; 
    q = Math.floor(q / 26) - 1; 
  } while (q >= 0);
  return '-' + suffix;
}

function formatTglLabel(dateObj) {
  var mth = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agt", "Sep", "Okt", "Nov", "Des"];
  return String(dateObj.getDate()).padStart(2, '0') + ' ' + mth[dateObj.getMonth()] + ' ' + String(dateObj.getFullYear()).slice(-2);
}

// ─────────────────────────────────────────────────────────────────────────
// HELPER: Vendor Name/Code Resolver
// Baca M_VENDOR (kolom A=Code, kolom B=Nama) → return map untuk konversi
// nama full atau code (case-insensitive) menjadi Code.
// ─────────────────────────────────────────────────────────────────────────
function _buildVendorNameToCodeMap() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('M_VENDOR');
  var map = {};
  if (!sh) return map;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][0] || '').trim();
    var name = String(data[i][1] || '').trim().toUpperCase();
    if (code) {
      if (name) map[name] = code;             // nama full → code
      map[code.toUpperCase()] = code;         // code (case-insensitive) → code
    }
  }
  return map;
}

function _resolveVendorCode(rawValue, vMap) {
  var v = String(rawValue || '').trim().toUpperCase();
  if (!v) return '-';
  return vMap[v] || rawValue; // fallback ke raw kalau tidak ditemukan (biar user aware)
}

// ─────────────────────────────────────────────────────────────────────────
// 1. GET INITIAL DATA (MENARIK DATA PO YANG OPEN SAJA)
// ─────────────────────────────────────────────────────────────────────────
function getInitData(dateFrom, dateTo) {
  try {
    var result = { form: getGRFormData(), dash: getGRDashboardData(1, dateFrom, dateTo) };
    return JSON.stringify(result);
  } catch (e) { return JSON.stringify({ error: e.message }); }
}

function getGRDashboardPage(page, dateFrom, dateTo) {
  try { return JSON.stringify(getGRDashboardData(page, dateFrom, dateTo)); } 
  catch (e) { return JSON.stringify({ error: e.message }); }
}

function getGRFormData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suppliers = [];
  var pos = {};
  var itemMap = {}; 

  // Ambil Data Supplier
  var suppSheet = ss.getSheetByName('M_SUPP');
  if (suppSheet) {
    var sData = suppSheet.getDataRange().getValues();
    for (var j = 1; j < sData.length; j++) {
      if (sData[j][0]) suppliers.push({ code: String(sData[j][0]).trim(), nama: String(sData[j][1]).trim() });
    }
  }

  // Ambil Master Item (Untuk menyedot Spec, T, P, L murni dari Master)
  var itemSheet = ss.getSheetByName('M_ITEM');
  if (itemSheet) {
    var iData = itemSheet.getDataRange().getValues();
    var iH = iData[0].map(function(h) { return String(h).trim(); });
    for (var i = 1; i < iData.length; i++) {
      if (iData[i][0]) {
        itemMap[String(iData[i][0]).trim()] = {
          spec: String(iData[i][iH.indexOf('Spec')] || '').trim(),
          t: iData[i][iH.indexOf('T')] || 0,
          p: iData[i][iH.indexOf('P')] || 0,
          l: iData[i][iH.indexOf('L')] || 0
        };
      }
    }
  }

  // Ambil Daftar PO yang OPEN beserta variabel PO Total Awal untuk Toleransi
  var poSheet = ss.getSheetByName('PO');
  if (poSheet) {
    var poData = poSheet.getDataRange().getValues();
    var poH = poData[0].map(function(h) { return String(h).trim(); });
    
    for (var p = 1; p < poData.length; p++) {
      var stat = String(poData[p][poH.indexOf('STATUS')]).trim().toUpperCase();
      var noPO = String(poData[p][poH.indexOf('PO_No')]).trim();
      var balQty = parseInt(poData[p][poH.indexOf('BL_Q')]) || 0;
      var balKg = parseFloat(poData[p][poH.indexOf('BL_KG')]) || 0;
      
      if (stat === 'OPEN' && noPO && (balQty > 0 || balKg > 0)) {
        if (!pos[noPO]) {
          pos[noPO] = { po_no: noPO, vendor: String(poData[p][poH.indexOf('Vendor')]).trim(), owner: String(poData[p][poH.indexOf('Owner')]).trim(), items: [] };
        }

        var iCode = String(poData[p][poH.indexOf('Item_Code')]).trim();
        var iRef = itemMap[iCode] || { spec: '-', t: 0, p: 0, l: 0 }; 

        pos[noPO].items.push({
          item_code: iCode,
          description: String(poData[p][poH.indexOf('Description')]).trim(),
          uom: String(poData[p][poH.indexOf('UoM')]).trim(),
          wg_pce: parseFloat(poData[p][poH.indexOf('Wg/Pce')]) || 0,
          target_loc: String(poData[p][poH.indexOf('Target_Loc')]).trim(),
          spec: iRef.spec, thick: iRef.t, lebar: iRef.p, l_dim: iRef.l,
          bal_qty: balQty,
          bal_kg: balKg,
          po_qty: parseInt(poData[p][poH.indexOf('PO_Q')]) || 0,   // DITAMBAHKAN UNTUK TOLERANSI
          po_kg: parseFloat(poData[p][poH.indexOf('PO_KG')]) || 0 // DITAMBAHKAN UNTUK TOLERANSI
        });
      }
    }
  }
  
  var poArray = Object.keys(pos).map(function(k) { return pos[k]; });
  return { suppliers: suppliers, pos: poArray };
}

// ─────────────────────────────────────────────────────────────────────────
// 2. GET DASHBOARD DATA (MENAMPILKAN RIWAYAT)
// ─────────────────────────────────────────────────────────────────────────
function getGRDashboardData(page, dateFrom, dateTo) {
  var pageSize = 50;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allData = [];

  // ── Parse date range filter (opsional; null/empty = tidak filter) ──
  var dateFromObj = null, dateToObj = null;
  if (dateFrom) {
    var pf = String(dateFrom).split('-');
    if (pf.length === 3) {
      var yf = parseInt(pf[0], 10), mf = parseInt(pf[1], 10) - 1, df = parseInt(pf[2], 10);
      if (!isNaN(yf) && !isNaN(mf) && !isNaN(df)) dateFromObj = new Date(yf, mf, df, 0, 0, 0);
    }
  }
  if (dateTo) {
    var pt = String(dateTo).split('-');
    if (pt.length === 3) {
      var yt = parseInt(pt[0], 10), mt = parseInt(pt[1], 10) - 1, dt = parseInt(pt[2], 10);
      if (!isNaN(yt) && !isNaN(mt) && !isNaN(dt)) dateToObj = new Date(yt, mt, dt, 23, 59, 59);
    }
  }

  // 1. Tarik Master Item untuk mendapatkan dimensi (T, P, L)
  var itemMap = {};
  var itemSheet = ss.getSheetByName('M_ITEM');
  if (itemSheet) {
    var iData = itemSheet.getDataRange().getValues();
    var iH = iData[0].map(function(h) { return String(h).trim(); });
    var iCodeIdx = 0; // Asumsi Item Code ada di kolom pertama
    var specIdx = iH.indexOf('Spec');
    var tIdx = iH.indexOf('T');
    var pIdx = iH.indexOf('P');
    var lIdx = iH.indexOf('L');
    var uomIdx = iH.indexOf('Unit of Measure');
    if (uomIdx === -1) uomIdx = iH.indexOf('UoM'); // fallback
    
    for (var m = 1; m < iData.length; m++) {
      if (iData[m][iCodeIdx]) {
        itemMap[String(iData[m][iCodeIdx]).trim()] = {
          spec: specIdx >= 0 ? String(iData[m][specIdx] || '').trim() : '',
          uom: uomIdx >= 0 ? String(iData[m][uomIdx] || '').trim() : '',
          t: tIdx >= 0 ? (parseFloat(iData[m][tIdx]) || 0) : 0,
          p: pIdx >= 0 ? (parseFloat(iData[m][pIdx]) || 0) : 0,
          l: lIdx >= 0 ? (parseFloat(iData[m][lIdx]) || 0) : 0
        };
      }
    }
  }

  // 2. Tarik log riwayat dari Sheet GR
  var sh = ss.getSheetByName('GR');
  if (!sh) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };
  
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };
  
  var h = data[0].map(function(x){ return String(x).trim(); });
  var bIdx = h.indexOf('Batch_ID');
  var tIdx = h.indexOf('Tgl_Masuk'); 
  var iIdx = h.indexOf('Item_Code');
  var doIdx = h.indexOf('No_DO');
  var poIdx = h.indexOf('No_PO'); 
  var sIdx = h.indexOf('Supplier'); 
  var qIdx = h.indexOf('Qty_In');
  if (qIdx === -1) qIdx = h.indexOf('QTY_In');
  var kIdx = h.indexOf('KG_In');
  var descIdx = h.indexOf('Description'); 
  var cIdx = h.indexOf('No_Coil');
  var ownIdx = h.indexOf('Owner');
  var locIdx = h.indexOf('Target_Loc');
  var noteIdx = h.indexOf('NOTE');                  // 🟢 T1.6
  if (noteIdx === -1) noteIdx = h.indexOf('Note');  // fallback nama header
  // 🟢 Snapshot Spec/T/P/L dari sheet GR (kalau ada) — prioritas snapshot → fallback itemMap
  var specGRIdx = h.indexOf('Spec');
  var tGRIdx = h.indexOf('T');
  var pGRIdx = h.indexOf('P');
  var lGRIdx = h.indexOf('L');
  // 🟢 Vendor code resolver map (Nama Full → V001)
  var vendorMap = _buildVendorNameToCodeMap();

  for (var i = 1; i < data.length; i++) {
    if (!data[i][bIdx]) continue;
    
    var rawTgl = tIdx >= 0 ? data[i][tIdx] : new Date();
    var tglStr = "-"; var sortVal = 0;
    if (rawTgl instanceof Date) { 
      tglStr = formatTglLabel(rawTgl); 
      sortVal = rawTgl.getTime();
    }

    // ── Apply date range filter ──
    if (dateFromObj || dateToObj) {
      if (!(rawTgl instanceof Date)) continue;
      if (dateFromObj && rawTgl < dateFromObj) continue;
      if (dateToObj   && rawTgl > dateToObj)   continue;
    }

    var targetLoc = locIdx >= 0 ? String(data[i][locIdx]).trim() : '';
    var jenis = (targetLoc === 'Stok_Coil') ? 'COIL' : 'LEMBARAN';
    var itemCode = iIdx >= 0 ? String(data[i][iIdx]).trim() : '';
    
    // 3. Resolve Spec/UoM/dimensi — snapshot GR dulu, fallback itemMap
    var itmRef = itemMap[itemCode] || { spec:'', uom:'', t: 0, p: 0, l: 0 };
    var snapSpec = specGRIdx >= 0 ? String(data[i][specGRIdx] || '').trim() : '';
    var snapT = tGRIdx >= 0 ? (parseFloat(data[i][tGRIdx]) || 0) : 0;
    var snapP = pGRIdx >= 0 ? (parseFloat(data[i][pGRIdx]) || 0) : 0;
    var snapL = lGRIdx >= 0 ? (parseFloat(data[i][lGRIdx]) || 0) : 0;
    var finalSpec = snapSpec || itmRef.spec || '';
    var finalT = snapT || itmRef.t || 0;
    var finalP = snapP || itmRef.p || 0;
    var finalL = snapL || itmRef.l || 0;
    var dimStr = [finalT, finalP, finalL].filter(function(v){ return v > 0; }).join(' x ');
    
    // 4. Resolve Vendor Code (kolom Supplier bisa berisi nama full → convert ke V001)
    var rawSupplier = sIdx >= 0 ? String(data[i][sIdx] || '').trim() : '';
    var supplierCode = rawSupplier ? _resolveVendorCode(rawSupplier, vendorMap) : '-';
    
    allData.push({
      batch_id: String(data[i][bIdx]).trim(), 
      tgl_str: tglStr, 
      sort_val: sortVal, 
      jenis: jenis, 
      item_code: itemCode, 
      desc: descIdx >= 0 ? data[i][descIdx] : '-',
      no_do: doIdx >= 0 ? data[i][doIdx] : '-', 
      no_po: poIdx >= 0 ? data[i][poIdx] : '-', 
      supplier: supplierCode, // 🟢 sekarang code (V001), bukan nama full
      qty: qIdx >= 0 ? data[i][qIdx] : null, 
      kg: kIdx >= 0 ? data[i][kIdx] : 0, 
      uom: itmRef.uom || 'Pcs', // 🟢 fix: uom untuk label re-print
      no_coil: cIdx >= 0 ? data[i][cIdx] : '-', 
      owner: ownIdx >= 0 ? data[i][ownIdx] : '',
      note: noteIdx >= 0 ? String(data[i][noteIdx] || '') : '',   // 🟢 T1.6
      sheetName: 'GR',
      spec: finalSpec, // 🟢 untuk Re-Print label
      dim: dimStr,     // 🟢 untuk Re-Print label
      t: finalT,
      p: finalP,
      l: finalL
    });
  }

  allData.sort(function(a, b) { 
    return b.sort_val === a.sort_val ? b.batch_id.localeCompare(a.batch_id) : b.sort_val - a.sort_val; 
  });
  
  var total = allData.length;

  // ── Kalau ada date filter aktif → return SEMUA (no server pagination).
  //    Frontend akan handle chip filter + search + pagination client-side.
  if (dateFromObj || dateToObj) {
    return { data: allData, totalPages: 1, currentPage: 1, totalItems: total };
  }

  // Legacy behavior (tanpa date filter): paginate seperti sebelumnya
  var pageSize = 50;
  var totalPages = Math.ceil(total / pageSize) || 1;
  var p = Math.max(1, Math.min(page || 1, totalPages));
  var start = (p - 1) * pageSize; 
  var paged = allData.slice(start, start + pageSize).reverse();
  
  return { data: paged, totalPages: totalPages, currentPage: p, totalItems: total };
}

// ─────────────────────────────────────────────────────────────────────────
// 3. RUTING PENYIMPANAN GR (LOGIKA 3 ARAH: GR + COIL/SHEET)
// ─────────────────────────────────────────────────────────────────────────
/* ─── 1. REPLACE function saveGR ──────────────────────────────────────── */
function saveGR(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);

  try {
    var head = payload.header; var list = payload.items;
    if (!head.no_do || !head.no_po || !head.supplier || !head.tgl_masuk || !head.owner) throw new Error('Header Dokumen tidak lengkap.');
    if (!list || list.length === 0) throw new Error('Daftar item kosong.');

    var tgl = new Date(head.tgl_masuk);
    var baseBatchId = getBaseBatchId_GR(tgl);
    var savedCount = 0; var labelsToPrint = [];
    var vendorMap = _buildVendorNameToCodeMap(); // 🟢 untuk convert nama full → code

    for (var i = 0; i < list.length; i++) {
      var itemData = list[i];
      var finalBatchId = baseBatchId + getSuffix(savedCount);
      itemData.no_do = head.no_do; itemData.no_po = head.no_po; itemData.supplier = head.supplier; itemData.owner = head.owner;

      // 🟢 T1.5: GR hanya tulis ke sheet GR — Stok_Coil/Sheet derive via formula
      _saveGR_Log(itemData, finalBatchId, tgl);
      savedCount++;

      var lblDim = [itemData.thick, itemData.lebar, itemData.l_dim].filter(Boolean).join(' x ');
      var isCoilLbl = String(itemData.target_loc).trim() === 'Stok_Coil';
      var tipeLabel = isCoilLbl ? 'COIL' : 'LEMBARAN';
      var qtyKgStr = isCoilLbl
        ? (itemData.kg_in + ' Kg')
        : (itemData.qty_in + ' ' + (itemData.uom||'Pcs') + ' / ' + itemData.kg_in + ' Kg');
      var vendorCode = _resolveVendorCode(itemData.supplier, vendorMap); // 🟢 convert ke V001
      labelsToPrint.push({
        jenis: tipeLabel, spec: itemData.spec || '-', dim: lblDim ? lblDim + ' mm' : '-',
        qty_kg: qtyKgStr,
        tgl_gr: formatTglLabel(tgl), no_gr: finalBatchId,
        vendor: vendorCode || '-',
        no_coil: itemData.no_coil || '-'
      });
    }

    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, count: savedCount, labels: labelsToPrint });
  } catch (e) { return JSON.stringify({ error: e.message }); } finally { lock.releaseLock(); }
}


// =========================================================================
// MAPPING PENYIMPANAN POKAYOKE (TANPA KOLOM SPEC/T/P/L)
// =========================================================================

/* ─── 2. REPLACE function _saveGR_Log ──────────────────────────────────── */
function _saveGR_Log(data, batchId, tgl) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GR'); if(!sheet) return;
  var headers = sheet.getDataRange().getValues()[0].map(function(h) { return String(h).trim(); });
  var rowMap = {
    'Batch_ID'   : batchId,
    'Tgl_Masuk'  : tgl,
    'Item_Code'  : data.item_code,
    'Description': data.description,
    'Spec'       : data.spec || '',                   // 🟢 T1.5: snapshot dari payload
    'T'          : parseFloat(data.thick)  || 0,      // 🟢 T1.5
    'P'          : parseFloat(data.lebar)  || 0,      // 🟢 T1.5
    'L'          : parseFloat(data.l_dim)  || 0,      // 🟢 T1.5
    'Supplier'   : data.supplier,
    'No_Coil'    : data.no_coil || '',
    // Coil = material kontinyu → QTY = KG (1:1). Sheet = diskrit → QTY = lembar
    'QTY_In'     : (String(data.target_loc).trim() === 'Stok_Coil')
                     ? parseFloat(data.kg_in) || 0
                     : parseInt(data.qty_in)  || 0,
    'KG_In'      : parseFloat(data.kg_in)  || 0,
    'Owner'      : data.owner,
    'No_DO'      : data.no_do,
    'No_PO'      : data.no_po,
    'Target_Loc' : data.target_loc,
    'NOTE'       : String(data.note || '').trim(),   // 🟢 T1.6: fix — sebelumnya di-drop
    'Note'       : String(data.note || '').trim()    // fallback kalau header di-label 'Note' bukan 'NOTE'
  };
  var targetRow = getActualLastRow(sheet) + 1;
  // Tulis 1× setValues (lebih cepat dari setValue per cell)
  var rowArr = headers.map(function(h) { return rowMap[h] !== undefined ? rowMap[h] : ''; });
  sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowArr]);
}

// ─────────────────────────────────────────────────────────────────────────
// 4. UPDATE SINGLE GR (Mode Edit di Dashboard)
// ─────────────────────────────────────────────────────────────────────────
/* ─── 3. REPLACE function updateSingleGR ───────────────────────────────── */
function updateSingleGR(payload) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('GR');
    if (!sh) throw new Error('Sheet GR tidak ditemukan.');

    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(x){ return String(x).trim(); });
    var bIdx = h.indexOf('Batch_ID');

    var targetRow = -1;
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][bIdx]).trim() === String(payload.batch_id).trim()) {
        targetRow = j + 1;
        break;
      }
    }
    if (targetRow === -1) throw new Error('Batch_ID tidak ditemukan: ' + payload.batch_id);

    // 🟢 T1.5: cukup update GR sheet. Stok_Coil/Sheet formula auto-refresh.
    if (payload.kg !== undefined && payload.kg !== null && payload.kg !== "") {
      var kgCol = h.indexOf('KG_In');
      if (kgCol >= 0) sh.getRange(targetRow, kgCol + 1).setValue(parseFloat(payload.kg));
    }
    if (payload.qty !== undefined && payload.qty !== null && payload.qty !== "") {
      var qtyCol = h.indexOf('Qty_In');
      if (qtyCol === -1) qtyCol = h.indexOf('QTY_In');
      if (qtyCol >= 0) sh.getRange(targetRow, qtyCol + 1).setValue(parseInt(payload.qty));
    }
    // 🟢 T1.6: handle note update (undefined = skip, empty string = clear allowed)
    if (payload.note !== undefined && payload.note !== null) {
      var noteCol = h.indexOf('NOTE');
      if (noteCol === -1) noteCol = h.indexOf('Note');
      if (noteCol >= 0) sh.getRange(targetRow, noteCol + 1).setValue(String(payload.note));
    }

    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, batch_id: payload.batch_id });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally { lock.releaseLock(); }
}
