// =========================================================================
// STOCK SERVICE — FC COS
// Menulis ke: Stok_Sheet, Stok_WIP, Stok_FG, Trace_Log, Rekap_ICT
// Stok_Coil HANYA ditulis oleh GR_Service
// Qty_Keep (Stok_Sheet) & KG_Keep (Stok_Coil) = formula SUMIFS dari SPK
// =========================================================================

/* =========================================================================
 * GENERATE BATCH ID dari M_Config counter
 * type: SHT | SLT | WIP | FCL | FSH
 * Format hasil: SHT-2600001 | WIP-2600001 | FC-CL-2600001 | FC-SH-2600001
 * ========================================================================= */
/* =========================================================================
 * 🟢 FIX — generateBatchId (tambah dukungan FSL → FC-SL)
 *
 * Perubahan vs versi asli:
 *  - keyMap : tambah 'FSL':'LAST_FSL'
 *  - return : tambah handling type 'FSL' → 'FC-SL-' + yy + seq
 *
 * Prefix lain tidak berubah.
 * ========================================================================= */
function generateBatchId(type) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = ss.getSheetByName(SHEET_NAMES.M_CONFIG);
  if (!cfgSheet) throw new Error('Sheet M_Config tidak ditemukan!');

  var data     = cfgSheet.getDataRange().getValues();
  var keyMap   = { 'SHT':'LAST_SHT','SLT':'LAST_SLT','WIP':'LAST_WIP',
                   'FCL':'LAST_FCL','FSH':'LAST_FSH','FSL':'LAST_FSL' };
  var configKey = keyMap[type];
  if (!configKey) throw new Error('Tipe Batch ID tidak valid: ' + type);

  var rowIdx = -1, curVal = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === configKey) {
      rowIdx = i + 1;
      curVal = parseInt(data[i][1]) || 0;
      break;
    }
  }
  if (rowIdx === -1) throw new Error('Key ' + configKey + ' tidak ada di M_Config!');

  var newVal = curVal + 1;
  var tz     = Session.getScriptTimeZone();
  var yy     = Utilities.formatDate(new Date(), tz, 'yy');
  var seq    = String(newVal).padStart(7, '0').slice(2); // 5 digit

  cfgSheet.getRange(rowIdx, 2).setValue(newVal);

  if (type === 'FCL') return 'FC-CL-' + yy + seq;
  if (type === 'FSH') return 'FC-SH-' + yy + seq;
  if (type === 'FSL') return 'FC-SL-' + yy + seq;
  return type + '-' + yy + seq;
}

/* =========================================================================
 * 🟢 FIX — getFgBatchType (Machine-oriented, bukan Source-oriented)
 *
 * Prefix Batch FG ditentukan oleh MESIN PRODUSEN / bentuk fisik output,
 * bukan asal source batch:
 *
 *   SLT-OUT → FSL (FC-SL)  → strip slitting
 *   SHR-OUT → FSH (FC-SH)  → sheet/lembaran (apapun source-nya)
 *   CTL-OUT → FSH (FC-SH)  → sheet (jarang langsung FG, tapi bentuk = sheet)
 *
 *   FCL (FC-CL) → TIDAK di-return otomatis di sini.
 *     Dipakai khusus saat mother-coil utuh dijual via SPK Allocated,
 *     panggil generateBatchId('FCL') langsung dari modul tersebut.
 *
 * Parameter:
 *   sourceBatch : batch asal (untuk fallback inferensi)
 *   spkType     : 'CTL-OUT' | 'SHR-OUT' | 'SLT-OUT'  (optional, prioritas tinggi)
 *
 * Backward compat: caller lama yang panggil tanpa spkType tetap dapat
 * hasil "FSH" (bentuk sheet) — aman dari bug FCL-everywhere.
 * ========================================================================= */
function getFgBatchType(sourceBatch, spkType) {
  // Prioritas: lihat mesin produsen via spkType
  if (spkType) {
    var t = String(spkType).toUpperCase();
    if (t === 'SLT-OUT') return 'FSL';
    if (t === 'SHR-OUT') return 'FSH';
    if (t === 'CTL-OUT') return 'FSH'; // CTL → FG langsung = sheet
  }
  // Fallback: inferensi dari source (backward compat)
  if (!sourceBatch) return 'FSH';
  var s = String(sourceBatch).toUpperCase();
  if (s.indexOf('SLT-') === 0) return 'FSL';
  return 'FSH';
}

/* =========================================================================
 * GET ROOT BATCH
 * Trace Source_Batch sampai batch Level 0 (dari vendor = GR-YYNNNN-X)
 * ========================================================================= */
function getRootBatch(sourceBatch) {
  if (!sourceBatch) return '';
  var s = String(sourceBatch).trim();
  if (s.indexOf('GR-') === 0) return s;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TRACE_LOG);
  if (!sheet) return s;
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(x){ return String(x).trim(); });
  var iB    = hdr.indexOf('Batch_ID');
  var iRoot = hdr.indexOf('Root_Batch');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iB]).trim() === s) {
      var root = String(data[i][iRoot] || '').trim();
      return root || s;
    }
  }
  return s;
}

/* =========================================================================
 * GET SUPPLIER INFO dari root batch (Stok_Coil atau Stok_Sheet)
 * ========================================================================= */
/* ─── 1. REPLACE function getCoilSupplierInfo ──────────────────────────── */
function getCoilSupplierInfo(rootBatch) {
  if (!rootBatch) return { supplier: '', no_po: '', no_do: '' };

  // 🟢 T1.5: baca GR sheet langsung (single source of truth untuk supplier info)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('GR');
  if (!sh) return { supplier: '', no_po: '', no_do: '' };

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { supplier: '', no_po: '', no_do: '' };

  var hdr = data[0].map(function(h){ return String(h).trim(); });
  var iB  = hdr.indexOf('Batch_ID');
  var iSp = hdr.indexOf('Supplier');
  var iPo = hdr.indexOf('No_PO');
  var iDo = hdr.indexOf('No_DO');
  if (iB < 0) return { supplier: '', no_po: '', no_do: '' };

  var key = String(rootBatch).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iB]).trim() === key) {
      return {
        supplier : iSp > -1 ? String(data[i][iSp] || '').trim() : '',
        no_po    : iPo > -1 ? String(data[i][iPo] || '').trim() : '',
        no_do    : iDo > -1 ? String(data[i][iDo] || '').trim() : ''
      };
    }
  }
  return { supplier: '', no_po: '', no_do: '' };
}


/* =========================================================================
 * HELPER: Cari baris kosong pertama berdasarkan kolom Batch_ID / SPK_No
 * 🟢 OPTIMASI v2 — Baca SATU kolom saja (bukan getDataRange seluruh sheet)
 *    Pada sheet 30+ kolom dengan ribuan baris, ini ~10-30x lebih cepat.
 *    Aman untuk Stok_Sheet / Stok_WIP / Stok_FG yang punya ARRAYFORMULA spill
 *    di kolom lain — kita hanya scan kolom Batch_ID / SPK_No.
 * ========================================================================= */
function findWriteRow(sheet, colName) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1) return 2;

  // Baca header row saja
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                 .map(function(h){ return String(h).trim(); });
  var iCol = hdr.indexOf(colName);
  if (iCol < 0) return lastRow + 1; // fallback

  if (lastRow < 2) return 2;

  // Baca cuma 1 kolom (Batch_ID atau SPK_No) — jauh lebih ringan
  var colVals = sheet.getRange(2, iCol + 1, lastRow - 1, 1).getValues();
  for (var r = 0; r < colVals.length; r++) {
    if (!colVals[r][0] || String(colVals[r][0]).trim() === '') return r + 2;
  }
  return colVals.length + 2;
}


/* =========================================================================
 * TULIS KE TRACE_LOG
 * ========================================================================= */
function writeTraceLog(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TRACE_LOG);
  if (!sheet) throw new Error('Sheet Trace_Log tidak ditemukan!');

  var hdr  = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
                  .map(function(h){ return String(h).trim(); });
  var chain = p.batch_id;
  if (p.root_batch && p.root_batch !== p.batch_id) {
    chain = p.root_batch;
    if (p.source_batch && p.source_batch !== p.root_batch)
      chain += ' → ' + p.source_batch;
    chain += ' → ' + p.batch_id;
  }

  var row = hdr.map(function(h) {
    switch(h) {
      case 'Batch_ID'    : return p.batch_id      || '';
      case 'Tgl_Buat'   : return p.tgl_buat       || new Date();
      case 'Level'       : return p.level          || 0;
      case 'Type'        : return p.type           || '';
      case 'Source_Batch': return p.source_batch   || '';
      case 'Root_Batch'  : return p.root_batch     || '';
      case 'SPK_Ref'     : return p.spk_ref        || '';
      case 'GR_Ref'      : return p.gr_ref         || '';
      case 'Item_Code'   : return p.item_code      || '';
      case 'Description' : return p.description    || '';
      case 'Spec'        : return p.spec           || '';
      case 'T'           : return p.t              || '';
      case 'P'           : return p.p              || '';
      case 'L_dim'       : return p.l_dim          || '';
      case 'Qty'         : return p.qty            || 0;
      case 'KG'          : return p.kg             || 0;
      case 'Operator'    : return p.operator       || '';
      case 'MC_No'       : return p.mc_no          || '';
      case 'Tgl_Prod'    : return p.tgl_prod       || new Date();
      case 'Supplier'    : return p.supplier       || '';
      case 'No_PO'       : return p.no_po          || '';
      case 'No_DO'       : return p.no_do          || '';
      case 'Owner'       : return p.owner          || '';
      case 'Owner_Used'  : return p.owner_used     || '';
      case 'Full_Chain'  : return chain;
      default            : return '';
    }
  });
  sheet.getRange(findWriteRow(sheet,'Batch_ID'), 1, 1, hdr.length).setValues([row]);
}

/* =========================================================================
 * TULIS KE REKAP_ICT (cross-billing: Owner ≠ Owner_Used)
 * ========================================================================= */
function writeRekapICT(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.REKAP_ICT);
  if (!sheet) throw new Error('Sheet Rekap_ICT tidak ditemukan!');

  var hdr = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
                 .map(function(h){ return String(h).trim(); });
  var row = hdr.map(function(h) {
    switch(h) {
      case 'Tgl_Transfer': return p.tgl          || new Date();
      case 'SPK_No'      : return p.spk_no       || '';
      case 'Item_Code'   : return p.item_code    || '';
      case 'Description' : return p.description  || '';
      case 'Dari_Owner'  : return p.dari_owner   || '';
      case 'Ke_Owner'    : return p.ke_owner     || '';
      case 'Qty_Sht'     : return p.qty          || 0;
      case 'Qty_KG'      : return p.kg           || 0;
      default            : return '';
    }
  });
  sheet.getRange(findWriteRow(sheet,'SPK_No'), 1, 1, hdr.length).setValues([row]);
}

/* =========================================================================
 * VALIDATE COIL AVAILABILITY (sebelum save SPK CTL)
 * Cek KG_Avail dari formula sheet
 * ========================================================================= */
/* =========================================================================
 * VALIDATE COIL AVAILABILITY — dengan support excludeSpkNo (edit mode)
 *   - batchId        : Batch ID coil yang mau dipakai
 *   - kgNeeded       : kg yang dibutuhkan untuk SPK ini
 *   - excludeSpkNo   : (opsional) SPK No yang sedang diedit; konsumsinya dianggap "kembali"
 *                      saat hitung available (karena akan diganti dengan nilai baru)
 * ========================================================================= */
function validateCoilAvailability(batchId, kgNeeded, excludeSpkNo) {
  var sheet = getSheet(SHEET_NAMES.STOK_COIL);
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h){ return String(h).trim(); });
  var colB  = hdr.indexOf('Batch_ID');
  var colA  = hdr.indexOf('KG_Avail');
  if (colA === -1) return true; // formula sheet tidak bisa cek

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colB]).trim() === String(batchId).trim()) {
      var avail = parseFloat(data[i][colA]) || 0;

      // Edit mode: tambah balik konsumsi SPK yang sedang diedit
      var ownKg = 0;
      if (excludeSpkNo) {
        var spkSheet  = getSheet("SPK");
        var spkData   = spkSheet.getDataRange().getValues();
        var spkHdr    = spkData[0].map(function(h){ return String(h).trim(); });
        var iSpkNo    = spkHdr.indexOf("SPK_No");
        var iKgTarget = spkHdr.indexOf("KG_Target");
        var iStatus   = spkHdr.indexOf("Status");

        for (var j = 1; j < spkData.length; j++) {
          if (String(spkData[j][iSpkNo]).trim() === String(excludeSpkNo).trim() &&
              String(spkData[j][iStatus]).toUpperCase() !== 'CANCELLED') {
            ownKg = parseFloat(spkData[j][iKgTarget]) || 0;
            break;
          }
        }
      }

      var effectiveAvail = avail + ownKg;
      if (effectiveAvail < kgNeeded - 0.001) {
        throw new Error('❌ Stok coil tidak cukup.\nAvail: '
          + effectiveAvail.toFixed(2) + ' kg, Dibutuhkan: ' + kgNeeded.toFixed(2) + ' kg');
      }
      return true;
    }
  }
  throw new Error('Batch tidak ditemukan di Stok_Coil: ' + batchId);
}

/* =========================================================================
 * VALIDATE SHEET AVAILABILITY — untuk SPK SHR standalone
 *   - batchId        : Batch ID sheet (dari Stok_Sheet)
 *   - qtyNeeded      : qty sheets yang dibutuhkan
 *   - kgNeeded       : kg yang dibutuhkan
 *   - excludeSpkNo   : (opsional) SPK SHR yang sedang diedit
 * ========================================================================= */
function validateSheetAvailability(batchId, qtyNeeded, kgNeeded, excludeSpkNo) {
  var sheet = getSheet(SHEET_NAMES.STOK_SHEET);
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h){ return String(h).trim(); });
  var colB  = hdr.indexOf('Batch_ID');
  var colQA = hdr.indexOf('Qty_Avail');
  var colKA = hdr.indexOf('KG_Avail');
  if (colKA === -1) colKA = hdr.indexOf('Kg_Avail');

  if (colQA === -1 && colKA === -1) return true; // tidak bisa cek

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colB]).trim() === String(batchId).trim()) {
      var qtyAvail = colQA >= 0 ? (parseFloat(data[i][colQA]) || 0) : Infinity;
      var kgAvail  = colKA >= 0 ? (parseFloat(data[i][colKA]) || 0) : Infinity;

      // Edit mode: tambah balik konsumsi SPK yang sedang diedit
      var ownQty = 0, ownKg = 0;
      if (excludeSpkNo) {
        var spkSheet  = getSheet("SPK");
        var spkData   = spkSheet.getDataRange().getValues();
        var spkHdr    = spkData[0].map(function(h){ return String(h).trim(); });
        var iSpkNo    = spkHdr.indexOf("SPK_No");
        var iQtyTgt   = spkHdr.indexOf("Qty_Target");
        var iKgTgt    = spkHdr.indexOf("KG_Target");
        var iStatus   = spkHdr.indexOf("Status");

        for (var j = 1; j < spkData.length; j++) {
          if (String(spkData[j][iSpkNo]).trim() === String(excludeSpkNo).trim() &&
              String(spkData[j][iStatus]).toUpperCase() !== 'CANCELLED') {
            ownQty = parseFloat(spkData[j][iQtyTgt]) || 0;
            ownKg  = parseFloat(spkData[j][iKgTgt]) || 0;
            break;
          }
        }
      }

      var effQty = qtyAvail + ownQty;
      var effKg  = kgAvail + ownKg;

      if (effQty < qtyNeeded - 0.001) {
        throw new Error('❌ Stok sheet tidak cukup.\nAvail: '
          + effQty + ' sht, Dibutuhkan: ' + qtyNeeded + ' sht');
      }
      if (effKg < kgNeeded - 0.001) {
        throw new Error('❌ Stok sheet tidak cukup.\nAvail: '
          + effKg.toFixed(2) + ' kg, Dibutuhkan: ' + kgNeeded.toFixed(2) + ' kg');
      }
      return true;
    }
  }
  throw new Error('Batch tidak ditemukan di Stok_Sheet: ' + batchId);
}

/* =========================================================================
 * UPDATE SO SPK_Q (dipanggil saat SPK dibuat)
 * ========================================================================= */
function updateSO_SPK_Q(soNo, qtyPlan, kgPlan) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SO');
  if (!sheet) return;

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var colSo = -1, colSpkQ = -1, colSpkKg = -1;

  for (var j = 0; j < headers.length; j++) {
    var h = String(headers[j]).trim().toUpperCase().replace(/[^A-Z0-9]/g,'');
    if (h === 'SONO' || h === 'SO')       colSo    = j;
    if (h === 'SPKQ' || h === 'SPKQTY')  colSpkQ  = j;
    if (h === 'SPKKG'|| h === 'SPKQKG')  colSpkKg = j;
  }
  if (colSo === -1) return;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colSo]).trim().toUpperCase() === String(soNo).trim().toUpperCase()) {
      var rowNum = i + 1;
      if (colSpkQ !== -1)
        sheet.getRange(rowNum, colSpkQ+1).setValue((parseFloat(data[i][colSpkQ])||0) + parseFloat(qtyPlan));
      if (colSpkKg !== -1 && colSpkKg !== colSpkQ)
        sheet.getRange(rowNum, colSpkKg+1).setValue((parseFloat(data[i][colSpkKg])||0) + parseFloat(kgPlan));
      return;
    }
  }
}