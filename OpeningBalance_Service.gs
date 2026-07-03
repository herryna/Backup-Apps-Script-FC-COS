// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - OPENING BALANCE SERVICE
// -------------------------------------------------------------------------
// Modul input Saldo Awal Stok (cutoff-based, bypass PO/GR normal).
// Semua data ditulis ke sheet GR dengan prefix batch khusus:
//   OPN-CL-YYMM-####  → Coil    (Target_Loc: Stok_Coil)
//   OPN-SH-YYMM-####  → Sheet   (Target_Loc: Stok_Sheet)
//   OPN-WP-YYMM-####  → WIP     (Target_Loc: Stok_Sheet)  ← fisik di cust/stamping
//
// Field GR yang di-auto-fill khusus opening:
//   Supplier = 'OPENING'
//   No_DO    = 'OPENING'
//   No_PO    = 'OPENING'
// Sisa field mengikuti input user.
//
// Owner di-set PER ROW (dropdown FC/DRC per item).
// Toggle mode Coil/Sheet/WIP di-set PER FORM (1 sesi = 1 jenis).
// =========================================================================


// -- Constants ------------------------------------------------------------
var OPN_PREFIX = {
  'COIL' : 'OPN-CL',
  'SHEET': 'OPN-SH',
  'WIP'  : 'OPN-WP'
};

var OPN_TARGET_LOC = {
  'COIL' : 'Stok_Coil',
  'SHEET': 'Stok_Sheet',
  'WIP'  : 'Stok_Sheet'
};


// =========================================================================
// UTILITIES
// =========================================================================
function opnGetActualLastRow(sheet) {
  var colA = sheet.getRange("A:A").getValues();
  for (var i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] !== "") return i + 1;
  }
  return 1;
}

function opnFormatTglLabel(dateObj) {
  var mth = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agt","Sep","Okt","Nov","Des"];
  return String(dateObj.getDate()).padStart(2, '0') + ' ' + mth[dateObj.getMonth()] + ' ' + String(dateObj.getFullYear()).slice(-2);
}


/* -------------------------------------------------------------------------
 * generateOpeningBatchId(mode, tglObj)
 * mode   : 'COIL' | 'SHEET' | 'WIP'
 * tglObj : Date object (menentukan YYMM di batch)
 * Format : OPN-CL-2607-0001, OPN-SH-2607-0001, OPN-WP-2607-0001
 * Counter di-store di SYS_Sequence dengan key = prefix (OPN-CL/OPN-SH/OPN-WP)
 * ------------------------------------------------------------------------- */
function generateOpeningBatchId(mode, tglObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSeq = ss.getSheetByName('SYS_Sequence');
  if (!sheetSeq) {
    sheetSeq = ss.insertSheet('SYS_Sequence');
    sheetSeq.appendRow(['Tipe_Doc', 'Bulan_Tahun', 'Last_Seq']);
    sheetSeq.hideSheet();
  }

  var prefix = OPN_PREFIX[mode];
  if (!prefix) throw new Error('Mode opening tidak valid: ' + mode);

  var tz   = Session.getScriptTimeZone();
  var yymm = Utilities.formatDate(tglObj, tz, 'yyMM'); // ex: '2607'

  var data = sheetSeq.getDataRange().getValues();
  var rowIndex = -1, currentSeq = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === prefix &&
        String(data[i][1]).trim() === yymm) {
      rowIndex   = i + 1;
      currentSeq = parseInt(data[i][2], 10) || 0;
      break;
    }
  }

  var nextSeq = currentSeq + 1;
  if (rowIndex === -1) {
    sheetSeq.appendRow([prefix, "'" + yymm, nextSeq]);
  } else {
    sheetSeq.getRange(rowIndex, 3).setValue(nextSeq);
  }

  return prefix + '-' + yymm + '-' + String(nextSeq).padStart(4, '0');
}


// =========================================================================
// 1. INIT DATA (dipanggil saat halaman dibuka)
// =========================================================================
function getOpeningInitData() {
  try {
    var result = {
      form: getOpeningFormData(),
      dash: getOpeningDashboardData(1)
    };
    return JSON.stringify(result);
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function getOpeningDashboardPage(page) {
  try {
    return JSON.stringify(getOpeningDashboardData(page));
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}


// -- Master data untuk form (item list dari M_ITEM) -----------------------
function getOpeningFormData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var items = [];

  var itemSheet = ss.getSheetByName('M_ITEM');
  if (itemSheet) {
    var iData = itemSheet.getDataRange().getValues();
    var iH    = iData[0].map(function(h){ return String(h).trim(); });
    var idxCode = iH.indexOf('Item_Code');
    var idxDesc = iH.indexOf('Description');
    var idxSpec = iH.indexOf('Spec');
    var idxT    = iH.indexOf('T');
    var idxP    = iH.indexOf('P');
    var idxL    = iH.indexOf('L');
    var idxUom  = iH.indexOf('Unit of Measure');
    var idxWg   = iH.indexOf('Wg/Pce FC');

    for (var i = 1; i < iData.length; i++) {
      if (!iData[i][idxCode]) continue;
      items.push({
        item_code   : String(iData[i][idxCode]).trim(),
        description : idxDesc >= 0 ? String(iData[i][idxDesc]).trim() : '',
        spec        : idxSpec >= 0 ? String(iData[i][idxSpec]).trim() : '',
        t           : idxT    >= 0 ? (parseFloat(iData[i][idxT]) || 0) : 0,
        p           : idxP    >= 0 ? (parseFloat(iData[i][idxP]) || 0) : 0,
        l           : idxL    >= 0 ? (parseFloat(iData[i][idxL]) || 0) : 0,
        uom         : idxUom  >= 0 ? String(iData[i][idxUom]).trim() : 'Pcs',
        wg_pce      : idxWg   >= 0 ? (parseFloat(iData[i][idxWg]) || 0) : 0
      });
    }
  }

  return { items: items };
}


// =========================================================================
// 2. DASHBOARD (list batch opening yang sudah di-input)
// =========================================================================
function getOpeningDashboardData(page) {
  var pageSize = 50;
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sh       = ss.getSheetByName('GR');
  if (!sh) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };

  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };

  var h = data[0].map(function(x){ return String(x).trim(); });
  var bIdx    = h.indexOf('Batch_ID');
  var tIdx    = h.indexOf('Tgl_Masuk');
  var iIdx    = h.indexOf('Item_Code');
  var descIdx = h.indexOf('Description');
  var specIdx = h.indexOf('Spec');
  var tDimIdx = h.indexOf('T');
  var pDimIdx = h.indexOf('P');
  var lDimIdx = h.indexOf('L');
  var coilIdx = h.indexOf('No_Coil');
  var qIdx    = h.indexOf('QTY_In'); if (qIdx === -1) qIdx = h.indexOf('Qty_In');
  var kIdx    = h.indexOf('KG_In');
  var ownIdx  = h.indexOf('Owner');
  var locIdx  = h.indexOf('Target_Loc');

  var out = [];
  for (var i = 1; i < data.length; i++) {
    var bid = String(data[i][bIdx] || '').trim();
    if (!bid) continue;
    // Filter: hanya batch opening
    if (bid.indexOf('OPN-') !== 0) continue;

    var mode = 'SHEET';
    if      (bid.indexOf('OPN-CL-') === 0) mode = 'COIL';
    else if (bid.indexOf('OPN-WP-') === 0) mode = 'WIP';

    var rawTgl = tIdx >= 0 ? data[i][tIdx] : new Date();
    var tglStr = '-'; var sortVal = 0;
    if (rawTgl instanceof Date) {
      tglStr  = opnFormatTglLabel(rawTgl);
      sortVal = rawTgl.getTime();
    }

    out.push({
      batch_id    : bid,
      tgl_str     : tglStr,
      sort_val    : sortVal,
      mode        : mode,
      item_code   : iIdx    >= 0 ? String(data[i][iIdx]).trim() : '',
      description : descIdx >= 0 ? String(data[i][descIdx]).trim() : '',
      spec        : specIdx >= 0 ? String(data[i][specIdx]).trim() : '',
      t           : tDimIdx >= 0 ? (parseFloat(data[i][tDimIdx]) || 0) : 0,
      p           : pDimIdx >= 0 ? (parseFloat(data[i][pDimIdx]) || 0) : 0,
      l           : lDimIdx >= 0 ? (parseFloat(data[i][lDimIdx]) || 0) : 0,
      no_coil     : coilIdx >= 0 ? String(data[i][coilIdx]).trim() : '',
      qty         : qIdx    >= 0 ? (parseFloat(data[i][qIdx]) || 0) : 0,
      kg          : kIdx    >= 0 ? (parseFloat(data[i][kIdx]) || 0) : 0,
      owner       : ownIdx  >= 0 ? String(data[i][ownIdx]).trim() : '',
      target_loc  : locIdx  >= 0 ? String(data[i][locIdx]).trim() : ''
    });
  }

  // Sort: batch terbaru di atas
  out.sort(function(a, b){
    return b.sort_val === a.sort_val
      ? b.batch_id.localeCompare(a.batch_id)
      : b.sort_val - a.sort_val;
  });

  var total      = out.length;
  var totalPages = Math.ceil(total / pageSize) || 1;
  var p          = Math.max(1, Math.min(page || 1, totalPages));
  var start      = (p - 1) * pageSize;
  var paged      = out.slice(start, start + pageSize);

  return { data: paged, totalPages: totalPages, currentPage: p, totalItems: total };
}


// =========================================================================
// 3. SAVE (multi-row → sheet GR sekali flush)
// =========================================================================
function saveOpeningBalance(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var head = payload.header;
    var list = payload.items;

    if (!head.mode)      throw new Error('Mode Coil/Sheet/WIP wajib dipilih.');
    if (!head.tgl_masuk) throw new Error('Tanggal Masuk wajib diisi.');
    if (!list || list.length === 0) throw new Error('Daftar item kosong.');

    var mode = String(head.mode).toUpperCase();
    if (!OPN_PREFIX[mode]) throw new Error('Mode tidak valid: ' + mode);
    var targetLoc = OPN_TARGET_LOC[mode];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('GR');
    if (!sh) throw new Error('Sheet GR tidak ditemukan.');

    var headers = sh.getDataRange().getValues()[0].map(function(h){ return String(h).trim(); });
    var tgl     = new Date(head.tgl_masuk);

    // -- Ambil master item untuk enrichment dimensi -----------------------
    var itemMap = {};
    var itemSheet = ss.getSheetByName('M_ITEM');
    if (itemSheet) {
      var iData = itemSheet.getDataRange().getValues();
      var iH    = iData[0].map(function(h){ return String(h).trim(); });
      var idxCode = iH.indexOf('Item_Code');
      var idxDesc = iH.indexOf('Description');
      var idxSpec = iH.indexOf('Spec');
      var idxT    = iH.indexOf('T');
      var idxP    = iH.indexOf('P');
      var idxL    = iH.indexOf('L');
      for (var i = 1; i < iData.length; i++) {
        if (!iData[i][idxCode]) continue;
        itemMap[String(iData[i][idxCode]).trim()] = {
          description: idxDesc >= 0 ? String(iData[i][idxDesc]).trim() : '',
          spec       : idxSpec >= 0 ? String(iData[i][idxSpec]).trim() : '',
          t          : idxT    >= 0 ? (parseFloat(iData[i][idxT]) || 0) : 0,
          p          : idxP    >= 0 ? (parseFloat(iData[i][idxP]) || 0) : 0,
          l          : idxL    >= 0 ? (parseFloat(iData[i][idxL]) || 0) : 0
        };
      }
    }

    var savedCount   = 0;
    var labelsToPrint = [];
    var rows         = [];

    for (var r = 0; r < list.length; r++) {
      var it = list[r];

      // -- Validasi ------------------------------------------------------
      if (!it.item_code) throw new Error('Item Code kosong di baris ' + (r + 1));
      if (!it.no_coil)   throw new Error('No Coil / Batch Lama kosong di baris ' + (r + 1));
      if (!it.owner)     throw new Error('Owner kosong di baris ' + (r + 1));

      var qtyIn = parseFloat(it.qty_in) || 0;
      var kgIn  = parseFloat(it.kg_in)  || 0;
      if (kgIn <= 0) throw new Error('KG kosong/0 di baris ' + (r + 1));
      if (mode !== 'COIL' && qtyIn <= 0) throw new Error('Qty (lembar) kosong/0 di baris ' + (r + 1));

      var owner = String(it.owner).trim().toUpperCase();
      if (owner !== 'FC' && owner !== 'DRC') {
        throw new Error('Owner harus FC atau DRC di baris ' + (r + 1) + ' (dapat: ' + it.owner + ')');
      }

      // -- Enrichment dari M_ITEM ---------------------------------------
      var iRef = itemMap[String(it.item_code).trim()] || { description:'', spec:'', t:0, p:0, l:0 };

      // -- Generate Batch ID (satu batch per row) -----------------------
      var batchId = generateOpeningBatchId(mode, tgl);

      // -- Aturan Coil: Qty = KG (material kontinyu, 1:1 dengan KG) -----
      var finalQty = (mode === 'COIL') ? kgIn : qtyIn;

      var rowMap = {
        'Batch_ID'   : batchId,
        'Tgl_Masuk'  : tgl,
        'Item_Code'  : String(it.item_code).trim(),
        'Description': it.description || iRef.description,
        'Spec'       : it.spec || iRef.spec,
        'T'          : parseFloat(it.t) || iRef.t || 0,
        'P'          : parseFloat(it.p) || iRef.p || 0,
        'L'          : parseFloat(it.l) || iRef.l || 0,
        'Supplier'   : 'OPENING',
        'No_Coil'    : String(it.no_coil).trim(),
        'QTY_In'     : finalQty,
        'KG_In'      : kgIn,
        'Owner'      : owner,
        'No_DO'      : 'OPENING',
        'No_PO'      : 'OPENING',
        'Target_Loc' : targetLoc
      };

      var rowArr = headers.map(function(h){ return rowMap[h] !== undefined ? rowMap[h] : ''; });
      rows.push(rowArr);

      // -- Data label untuk cetak ---------------------------------------
      var lblDim = [rowMap['T'], rowMap['P'], rowMap['L']].filter(function(x){ return x > 0; }).join('x');
      labelsToPrint.push({
        jenis     : mode,
        item_code : rowMap['Item_Code'],
        description: rowMap['Description'],
        spec      : rowMap['Spec'] || '-',
        dim       : lblDim ? (lblDim + ' mm') : '-',
        qty_kg    : (mode === 'COIL' ? '' : (qtyIn + ' Pcs / ')) + kgIn + ' Kg',
        tgl_gr    : opnFormatTglLabel(tgl),
        no_gr     : batchId,
        no_coil   : rowMap['No_Coil'],
        owner     : rowMap['Owner']
      });

      savedCount++;
    }

    // -- Bulk write ke GR sekali flush (jauh lebih cepat) -----------------
    var startRow = opnGetActualLastRow(sh) + 1;
    sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    SpreadsheetApp.flush();

    return JSON.stringify({ success: true, count: savedCount, labels: labelsToPrint });

  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally {
    lock.releaseLock();
  }
}


// =========================================================================
// 4. UPDATE SINGLE (edit qty/kg/no_coil/owner untuk batch yg belum tersentuh SPK)
// =========================================================================
function updateSingleOpening(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var batchId = String(payload.batch_id || '').trim();
    if (!batchId) throw new Error('Batch_ID kosong.');
    if (batchId.indexOf('OPN-') !== 0) {
      throw new Error('Batch ini bukan Opening Balance. Edit tidak diizinkan lewat modul ini.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('GR');
    if (!sh) throw new Error('Sheet GR tidak ditemukan.');

    var data = sh.getDataRange().getValues();
    var h    = data[0].map(function(x){ return String(x).trim(); });
    var bIdx = h.indexOf('Batch_ID');

    var targetRow = -1;
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][bIdx]).trim() === batchId) { targetRow = j + 1; break; }
    }
    if (targetRow === -1) throw new Error('Batch_ID tidak ditemukan: ' + batchId);

    // -- Guard: kalau sudah dipakai SPK, tolak edit ----------------------
    var loc = String(data[targetRow - 1][h.indexOf('Target_Loc')]).trim();
    if (isBatchUsedInSpk(batchId, loc)) {
      throw new Error('Batch sudah dipakai di SPK. Edit tidak diperbolehkan (cancel SPK terkait dulu).');
    }

    var updates = [];

    // -- KG_In -----------------------------------------------------------
    if (payload.kg !== undefined && payload.kg !== null && payload.kg !== '') {
      var kgCol = h.indexOf('KG_In');
      if (kgCol >= 0) {
        var newKg = parseFloat(payload.kg);
        if (newKg <= 0) throw new Error('KG harus > 0.');
        sh.getRange(targetRow, kgCol + 1).setValue(newKg);
        updates.push('KG_In=' + newKg);

        // Kalau Coil, QTY_In selalu = KG_In (rule 1:1)
        if (loc === 'Stok_Coil') {
          var qtyCol1 = h.indexOf('QTY_In'); if (qtyCol1 === -1) qtyCol1 = h.indexOf('Qty_In');
          if (qtyCol1 >= 0) {
            sh.getRange(targetRow, qtyCol1 + 1).setValue(newKg);
            updates.push('QTY_In=' + newKg + ' (auto-sync KG)');
          }
        }
      }
    }

    // -- QTY_In (hanya untuk Sheet/WIP; Coil ignore, karena auto-sync) --
    if (payload.qty !== undefined && payload.qty !== null && payload.qty !== '' && loc !== 'Stok_Coil') {
      var qtyCol = h.indexOf('QTY_In'); if (qtyCol === -1) qtyCol = h.indexOf('Qty_In');
      if (qtyCol >= 0) {
        var newQty = parseFloat(payload.qty);
        if (newQty <= 0) throw new Error('Qty harus > 0.');
        sh.getRange(targetRow, qtyCol + 1).setValue(newQty);
        updates.push('QTY_In=' + newQty);
      }
    }

    // -- No_Coil ---------------------------------------------------------
    if (payload.no_coil !== undefined && payload.no_coil !== null) {
      var cCol = h.indexOf('No_Coil');
      if (cCol >= 0) {
        sh.getRange(targetRow, cCol + 1).setValue(String(payload.no_coil).trim());
        updates.push('No_Coil=' + payload.no_coil);
      }
    }

    // -- Owner -----------------------------------------------------------
    if (payload.owner !== undefined && payload.owner !== null && payload.owner !== '') {
      var newOwner = String(payload.owner).trim().toUpperCase();
      if (newOwner !== 'FC' && newOwner !== 'DRC') {
        throw new Error('Owner harus FC atau DRC.');
      }
      var oCol = h.indexOf('Owner');
      if (oCol >= 0) {
        sh.getRange(targetRow, oCol + 1).setValue(newOwner);
        updates.push('Owner=' + newOwner);
      }
    }

    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, batch_id: batchId, updates: updates });

  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally {
    lock.releaseLock();
  }
}


// =========================================================================
// 5. DELETE (hapus batch opening yg salah input, hanya kalau belum tersentuh SPK)
// =========================================================================
function deleteOpeningBatch(batchId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    batchId = String(batchId || '').trim();
    if (!batchId) throw new Error('Batch_ID kosong.');
    if (batchId.indexOf('OPN-') !== 0) {
      throw new Error('Batch ini bukan Opening Balance, tidak bisa dihapus lewat modul ini.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('GR');
    if (!sh) throw new Error('Sheet GR tidak ditemukan.');

    var data = sh.getDataRange().getValues();
    var h    = data[0].map(function(x){ return String(x).trim(); });
    var bIdx   = h.indexOf('Batch_ID');
    var locIdx = h.indexOf('Target_Loc');

    var targetRow = -1;
    var loc       = '';
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][bIdx]).trim() === batchId) {
        targetRow = j + 1;
        loc       = String(data[j][locIdx] || '').trim();
        break;
      }
    }
    if (targetRow === -1) throw new Error('Batch tidak ditemukan: ' + batchId);

    if (isBatchUsedInSpk(batchId, loc)) {
      throw new Error('Batch sudah dipakai di SPK, tidak bisa dihapus.');
    }

    sh.deleteRow(targetRow);
    return JSON.stringify({ success: true, batch_id: batchId });

  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally {
    lock.releaseLock();
  }
}


// =========================================================================
// 6. GET LABEL DATA (untuk reprint label single batch)
// =========================================================================
function getOpeningLabelData(batchId) {
  try {
    batchId = String(batchId || '').trim();
    if (!batchId) throw new Error('Batch_ID kosong.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('GR');
    if (!sh) throw new Error('Sheet GR tidak ditemukan.');

    var data = sh.getDataRange().getValues();
    var h    = data[0].map(function(x){ return String(x).trim(); });
    var bIdx = h.indexOf('Batch_ID');

    for (var j = 1; j < data.length; j++) {
      if (String(data[j][bIdx]).trim() !== batchId) continue;

      var row  = data[j];
      var mode = 'SHEET';
      if      (batchId.indexOf('OPN-CL-') === 0) mode = 'COIL';
      else if (batchId.indexOf('OPN-WP-') === 0) mode = 'WIP';

      var t = parseFloat(row[h.indexOf('T')]) || 0;
      var p = parseFloat(row[h.indexOf('P')]) || 0;
      var l = parseFloat(row[h.indexOf('L')]) || 0;
      var lblDim = [t, p, l].filter(function(x){ return x > 0; }).join('x');

      var qtyIdx = h.indexOf('QTY_In'); if (qtyIdx === -1) qtyIdx = h.indexOf('Qty_In');
      var qty    = parseFloat(row[qtyIdx]) || 0;
      var kg     = parseFloat(row[h.indexOf('KG_In')]) || 0;
      var tglObj = row[h.indexOf('Tgl_Masuk')];

      return JSON.stringify({
        success: true,
        label: {
          jenis      : mode,
          item_code  : String(row[h.indexOf('Item_Code')]).trim(),
          description: String(row[h.indexOf('Description')] || '').trim(),
          spec       : String(row[h.indexOf('Spec')] || '').trim() || '-',
          dim        : lblDim ? (lblDim + ' mm') : '-',
          qty_kg     : (mode === 'COIL' ? '' : (qty + ' Pcs / ')) + kg + ' Kg',
          tgl_gr     : tglObj instanceof Date ? opnFormatTglLabel(tglObj) : '-',
          no_gr      : batchId,
          no_coil    : String(row[h.indexOf('No_Coil')] || '').trim(),
          owner      : String(row[h.indexOf('Owner')] || '').trim()
        }
      });
    }

    throw new Error('Batch tidak ditemukan: ' + batchId);

  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}


// =========================================================================
// 7. HELPER — cek apakah suatu Batch_ID sudah dipakai di SPK non-cancelled
// =========================================================================
function isBatchUsedInSpk(batchId, targetLoc) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spk = ss.getSheetByName('SPK');
  if (!spk) return false;

  var data = spk.getDataRange().getValues();
  if (data.length <= 1) return false;

  var h = data[0].map(function(x){ return String(x).trim(); });
  // Untuk Coil (Stok_Coil):  SPK.Parent_SPK berisi source Batch_ID  (kolom D)
  // Untuk Sheet/WIP (Stok_Sheet): SPK.Batch_ID berisi source Batch_ID (kolom AM)
  var parentIdx   = h.indexOf('Parent_SPK');
  var batchIdxSpk = h.indexOf('Batch_ID');
  var statusIdx   = h.indexOf('Status');

  for (var i = 1; i < data.length; i++) {
    var stat = String(data[i][statusIdx] || '').trim().toUpperCase();
    if (stat === 'CANCELLED') continue;

    if (targetLoc === 'Stok_Coil') {
      if (parentIdx >= 0 && String(data[i][parentIdx]).trim() === batchId) return true;
    } else {
      // Stok_Sheet (SHEET dan WIP mode sama-sama target Stok_Sheet)
      if (batchIdxSpk >= 0 && String(data[i][batchIdxSpk]).trim() === batchId) return true;
    }
  }
  return false;
}