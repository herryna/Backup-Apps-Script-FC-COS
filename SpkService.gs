// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - SPK SERVICE (DIET VERSION)
// =========================================================================

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error("Sheet '" + name + "' tidak ditemukan!");
  return sheet;
}

/* =========================================================================
 * HELPER: Safe append row untuk sheet SPK
 *   - appendRow() unreliable kalau ada ARRAYFORMULA yang spill di kolom mana pun
 *     (getLastRow ikut hitung row spillover sehingga data baru ter-append jauh
 *     di bawah). Helper ini scan kolom A (SPK_No) pakai TextFinder untuk cari
 *     baris data terakhir yang sebenarnya, lalu tulis via setValues.
 * ========================================================================= */
function _findLastSpkRow(sheet) {
  var finder = sheet.getRange("A:A").createTextFinder(".+").useRegularExpression(true);
  var matches = finder.findAll();
  return matches.length > 0 ? matches[matches.length - 1].getRow() : 1;
}

function _appendRowSafe(sheet, rowArray) {
  var targetRow = _findLastSpkRow(sheet) + 1;
  sheet.getRange(targetRow, 1, 1, rowArray.length).setValues([rowArray]);
  return targetRow;
}
/* =========================================================================
 * 🚩 SESI 2 — SAVE AS DRAFT SUPPORT
 * Helper: setelah semua rows di-tulis, flip Status → 'DRAFT' untuk semua
 * SPK_No yang baru di-generate. Dipanggil dari saveSPK_CTL kalau data.is_draft.
 * ========================================================================= */
function _updateStatusToDraft(spkSheet, spkNos) {
  if (!spkNos || !spkNos.length) return 0;
  var lastRow = spkSheet.getLastRow();
  var lastCol = spkSheet.getLastColumn();
  if (lastRow < 2) return 0;
  var vals    = spkSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var hdr     = vals[0].map(function(h){ return String(h).trim(); });
  var iSpk    = hdr.indexOf('SPK_No');
  var iStatus = hdr.indexOf('Status');
  if (iSpk < 0 || iStatus < 0) return 0;
  var spkSet = {};
  spkNos.forEach(function(s){ spkSet[String(s).trim()] = true; });
  var updated = 0;
  for (var r = 1; r < vals.length; r++) {
    var no = String(vals[r][iSpk]||'').trim();
    if (spkSet[no]) {
      spkSheet.getRange(r+1, iStatus+1).setValue('DRAFT');
      updated++;
    }
  }
  return updated;
}
/* =========================================================================
 * 1. SAVE SPK CTL (Fokus Menulis Antrean SPK, Stok Diambil Alih Rumus)
 * ========================================================================= */
/* =========================================================================
 * 1. SAVE SPK CTL
 * ========================================================================= */
/* =========================================================================
 * 🟢 T1 PATCH — saveSPK_CTL (Batch_ID inheritance untuk traceability)
 *
 * Perubahan vs versi asli:
 *  - CTL-HEADER : tambah field 'Batch_ID' = data.batch_id (coil GR)
 *  - CTL-OUT    : tambah field 'Batch_ID' = data.batch_id (inherit coil)
 *  - SHR-HEADER : tambah field 'Batch_ID' = ''  (akan diisi saat CTL-OUT DONE
 *                 dengan batch SHT yang baru di-generate — propagate otomatis)
 *  - SHR-OUT    : tambah field 'Batch_ID' = ''  (sama, propagate dari CTL-OUT)
 *
 * Logic CT, header rollup, dan field lain TIDAK BERUBAH.
 * ========================================================================= */
function saveSPK_CTL(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const spkSheet  = getSheet("SPK");
    const allData   = spkSheet.getDataRange().getValues();
    const headers   = allData[0].map(function(h) { return String(h).trim(); });
    const timestamp = new Date();
    const spkNo     = data.spk_no || getNextSpkNo('CTL', allData, headers);
    const generatedSpks = [];

    function buildRow(rowData) {
      return headers.map(function(h) { return rowData[h] !== undefined ? rowData[h] : ''; });
    }

    if (typeof validateCoilAvailability === 'function') {
      validateCoilAvailability(data.batch_id, data.weight_kg);
    }

    // ── LANGKAH 1: Pre-kalkulasi CT semua OUT sebelum menulis baris apapun ──
    var outCtlCtList = data.out_ctl.map(function(out) {
      return hitungRencanaDurasi(out.item_code, out.qty_plan, out.qty_plan_kg);
    });

    var shrCtMatrix = data.out_ctl.map(function(out) {
      if (!out.req_shr || out.req_shr.length === 0) return [];
      return out.req_shr.map(function(shr) {
        return hitungRencanaDurasi(shr.item_code, shr.qty_plan, shr.qty_plan_kg);
      });
    });

    // ── LANGKAH 2: Hitung agregat CTL-HEADER ──
    var coilCt        = hitungRencanaDurasi(data.coil_item_code, data.qty_input, data.weight_kg);
    var hdrPlanSetup  = coilCt.planSetup;
    var hdrPlanRun    = 0;
    var hdrTotalDurasi = hdrPlanSetup;
    outCtlCtList.forEach(function(ct) {
      hdrPlanRun     += ct.planRun;
      hdrTotalDurasi += ct.planSetup + ct.planRun;
    });

    // ── BARIS 1: CTL-HEADER ──
    var headerData = {
      'SPK_No'             : spkNo,
      'SPK_Type'           : 'CTL-HEADER',
      'Parent_SPK'         : data.batch_id,
      'Batch_ID'           : data.batch_id,            // 🟢 T1: batch coil source
      'Tgl_Buat'           : timestamp,
      'Priority'           : data.priority || 'Normal',
      'Source_Loc'         : data.source_loc,
      'Item_Code'          : data.coil_item_code || '',
      'Input_Spec'         : data.coil_desc || (data.thick + '×' + data.width),
      // CTL-HEADER konsumsi coil → Qty_Target = KG_Target (coil = material kontinyu, dihitung berat)
      'Qty_Target'         : parseFloat(data.weight_kg) || 0,
      'KG_Target'          : parseFloat(data.weight_kg) || 0,
      'MC_No'              : data.machine,
      'Status'             : 'ANTRIAN',
      'OP'                 : data.op,
      'Created_By'         : data.created_by,
      'Owner'              : data.coil_owner || 'FC',
      'Owner_Used'         : data.coil_owner || 'FC',
      'NOTE'               : data.note || '',
      'Plan_Setup_Menit'   : hdrPlanSetup,
      'Plan_Run_Menit'     : hdrPlanRun,
      'Total_Durasi_Menit' : hdrTotalDurasi,
      'Is_Habis'           : data.is_habis === true,
      'T'                  : parseFloat(data.thick) || 0,
      'Coil_Avail_Snapshot': parseFloat(data.coil_avail_kg) || 0
    };
    _appendRowSafe(spkSheet, buildRow(headerData));
    generatedSpks.push(spkNo);

    var yyNow        = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yy');
    var iSpkColCTL   = headers.indexOf('SPK_No');
    var reShrCTL     = new RegExp('^SHR-' + yyNow + '(\\d{4})$');
    var maxShrSeqCTL = 0;
    for (var rr = 1; rr < allData.length; rr++) {
      var mShr = String(allData[rr][iSpkColCTL] || '').trim().match(reShrCTL);
      if (mShr) { var ss = parseInt(mShr[1], 10); if (ss > maxShrSeqCTL) maxShrSeqCTL = ss; }
    }
    var currentShrSeq = maxShrSeqCTL + 1;

    // ── BARIS CTL-OUT + SHR (per item out_ctl) ──
    data.out_ctl.forEach(function(out, idx) {
      var outId     = spkNo + '-' + String(idx + 1).padStart(2, '0');
      var ctOut     = outCtlCtList[idx];
      var shrCtList = shrCtMatrix[idx];

      // ── BARIS CTL-OUT ──
      var outData = {
        'SPK_No'             : outId,
        'SPK_Type'           : 'CTL-OUT',
        'SO_Ref'             : out.so_ref      || '',
        'Cust'               : out.cust         || '',
        'Parent_SPK'         : spkNo,
        'Batch_ID'           : data.batch_id,           // 🟢 T1: inherit batch coil
        'Tgl_Buat'           : timestamp,
        'Priority'           : data.priority,
        'Source_Loc'         : data.machine,
        'Item_Code'          : out.item_code   || '',
        'Input_Spec'         : out.l            || '',
        'Target_Loc'         : out.target_loc  || '',
        'MC_No'              : data.machine,
        'BQ'                 : 1,
        'Qty_Target'         : out.qty_plan,
        'KG_Target'          : out.qty_plan_kg,
        'Status'             : 'ANTRIAN',
        'OP'                 : data.op,
        'Created_By'         : data.created_by,
        'Owner'              : out.owner      || 'FC',
        'Owner_Used'         : out.owner_used || 'FC',
        'Plan_Setup_Menit'   : ctOut.planSetup,
        'Plan_Run_Menit'     : ctOut.planRun,
        'Total_Durasi_Menit' : 0
      };
      _appendRowSafe(spkSheet, buildRow(outData));
      generatedSpks.push(outId);

      // ── SHR dari CTL (jika ada req_shr) ──
      if (out.req_shr && out.req_shr.length > 0) {
        var shrSpkNo = 'SHR-' + yyNow + String(currentShrSeq).padStart(4, '0');
        currentShrSeq++;

        var shrHdrPlanRun     = 0;
        var shrHdrTotalDurasi = 0;
        shrCtList.forEach(function(ct) {
          shrHdrPlanRun     += ct.planRun;
          shrHdrTotalDurasi += ct.planSetup + ct.planRun;
        });

        // ── BARIS SHR-HEADER ──
        // 🟢 Step 3.B — Resolve MC_No dari payload (semua OUT 1 HEADER pakai mesin sama)
        var shrHeaderMc = (out.req_shr[0] && out.req_shr[0].machine) || 'SHR-01';
        var shrHeaderData = {
          'SPK_No'             : shrSpkNo,
          'SPK_Type'           : 'SHR-HEADER',
          'Parent_SPK'         : outId,
          'Batch_ID'           : '',                       // 🟢 T1: kosong, akan di-propagate dari CTL-OUT DONE
          'Tgl_Buat'           : timestamp,
          'Priority'           : data.priority,
          'Source_Loc'         : out.target_loc || '',
          'Item_Code'          : out.item_code  || '',
          'Input_Spec'         : out.l           || '',
          'Qty_Target'         : out.qty_plan,
          'KG_Target'          : out.qty_plan_kg,
          'MC_No'              : shrHeaderMc,
          'Status'             : 'ANTRIAN',
          'Created_By'         : data.created_by,
          'Owner'              : data.coil_owner || 'FC',
          'Owner_Used'         : data.coil_owner || 'FC',
          'Plan_Setup_Menit'   : 0,
          'Plan_Run_Menit'     : shrHdrPlanRun,
          'Total_Durasi_Menit' : shrHdrTotalDurasi,
          'T'                  : parseFloat(data.thick) || 0
        };
        _appendRowSafe(spkSheet, buildRow(shrHeaderData));
        generatedSpks.push(shrSpkNo);

        // ── BARIS SHR-OUT ──
        out.req_shr.forEach(function(shr, sIdx) {
          var shrOutNo  = shrSpkNo + '-' + String(sIdx + 1).padStart(2, '0');
          var ctShrOut  = shrCtList[sIdx];

          var shrOutData = {
            'SPK_No'             : shrOutNo,
            'SPK_Type'           : 'SHR-OUT',
            'SO_Ref'             : shr.so_no      || '',
            'Cust'               : shr.cust        || '',
            'Parent_SPK'         : shrSpkNo,
            'Batch_ID'           : '',                       // 🟢 T1: kosong, akan di-propagate dari CTL-OUT DONE
            'Tgl_Buat'           : timestamp,
            'Priority'           : data.priority,
            'Source_Loc'         : shr.machine || 'SHR-01',
            'Item_Code'          : shr.item_code  || '',
            'Input_Spec'         : shr.l           || '',
            'Target_Loc'         : shr.target_loc || '',
            'MC_No'              : shr.machine || 'SHR-01',
            'BQ'                 : shr.cut_p * shr.cut_l,
            'Qty_Target'         : shr.qty_plan,
            'KG_Target'          : shr.qty_plan_kg,
            'Status'             : 'ANTRIAN',
            'Created_By'         : data.created_by,
            'NOTE'               : 'Cut_P:' + shr.cut_p + ' Cut_L:' + shr.cut_l + ' ' + (shr.proc_type || ''),
            'Owner'              : data.coil_owner || 'FC',
            'Owner_Used'         : shr.owner       || 'FC',
            'Plan_Setup_Menit'   : ctShrOut.planSetup,
            'Plan_Run_Menit'     : ctShrOut.planRun,
            'Total_Durasi_Menit' : 0,
            'T'                  : parseFloat(data.thick) || 0
          };
          _appendRowSafe(spkSheet, buildRow(shrOutData));
          generatedSpks.push(shrOutNo);
        });

        // 🟢 Step 3.B — STAGE 2 chain (kalau ada req_shr_stage2)
        if (out.req_shr_stage2) {
          var s2 = out.req_shr_stage2;
          var stage1Out01SpkNo = shrSpkNo + '-01';  // first SHR-OUT Stage 1
          var stage1Out01      = out.req_shr[0];     // data Req SHR 1

          // SHR-HEADER Stage 2
          var shrSpkNo2 = 'SHR-' + yyNow + String(currentShrSeq).padStart(4, '0');
          currentShrSeq++;

          var shr2HeaderData = {
            'SPK_No'             : shrSpkNo2,
            'SPK_Type'           : 'SHR-HEADER',
            'Parent_SPK'         : stage1Out01SpkNo,            // chain ke Stage 1 OUT
            'Batch_ID'           : '',
            'Tgl_Buat'           : timestamp,
            'Priority'           : data.priority,
            'Source_Loc'         : stage1Out01.target_loc || '',// output Stage 1 = input Stage 2
            'Item_Code'          : stage1Out01.item_code   || '',
            'Input_Spec'         : stage1Out01.l            || '',
            'Qty_Target'         : stage1Out01.qty_plan,
            'KG_Target'          : stage1Out01.qty_plan_kg,
            'MC_No'              : s2.machine || 'SHR-03',
            'Status'             : 'ANTRIAN',
            'Created_By'         : data.created_by,
            'Owner'              : data.coil_owner || 'FC',
            'Owner_Used'         : s2.owner || data.coil_owner || 'FC',
            'Plan_Setup_Menit'   : 0,
            'Plan_Run_Menit'     : 0,
            'Total_Durasi_Menit' : 0,
            'T'                  : parseFloat(data.thick) || 0,
            'NOTE'               : 'STAGE2-PARENT'
          };
          _appendRowSafe(spkSheet, buildRow(shr2HeaderData));
          generatedSpks.push(shrSpkNo2);

          // SHR-OUT Stage 2
          var shr2OutNo = shrSpkNo2 + '-01';
          var shr2OutData = {
            'SPK_No'             : shr2OutNo,
            'SPK_Type'           : 'SHR-OUT',
            'SO_Ref'             : s2.so_no || stage1Out01.so_no || '',
            'Cust'               : s2.cust  || stage1Out01.cust  || '',
            'Parent_SPK'         : shrSpkNo2,
            'Batch_ID'           : '',
            'Tgl_Buat'           : timestamp,
            'Priority'           : data.priority,
            'Source_Loc'         : s2.machine || 'SHR-03',
            'Item_Code'          : s2.item_code  || '',
            'Input_Spec'         : s2.l           || '',
            'Target_Loc'         : s2.target_loc || '',
            'MC_No'              : s2.machine || 'SHR-03',
            'BQ'                 : (s2.cut_p || 1) * (s2.cut_l || 1),
            'Qty_Target'         : s2.qty_plan,
            'KG_Target'          : s2.qty_plan_kg,
            'Status'             : 'ANTRIAN',
            'Created_By'         : data.created_by,
            'NOTE'               : 'Cut_P:' + s2.cut_p + ' Cut_L:' + s2.cut_l + ' Stage2',
            'Owner'              : data.coil_owner || 'FC',
            'Owner_Used'         : s2.owner || 'FC',
            'Plan_Setup_Menit'   : 0,
            'Plan_Run_Menit'     : 0,
            'Total_Durasi_Menit' : 0,
            'T'                  : parseFloat(data.thick) || 0
          };
          _appendRowSafe(spkSheet, buildRow(shr2OutData));
          generatedSpks.push(shr2OutNo);
        }
      }
    });

    SpreadsheetApp.flush();

    // 🚩 Sesi 2: Save as Draft — flip Status jadi 'DRAFT' semua rows yg baru dibuat
    // Skip scheduling untuk draft (baseline_mulai/plan_seq di-generate nanti saat promote)
    var isDraft = (data.is_draft === true);
    if (isDraft) {
      _updateStatusToDraft(spkSheet, generatedSpks);
      SpreadsheetApp.flush();
    } else {
      if (typeof kalkulasiEstimasiWaktu === 'function') kalkulasiEstimasiWaktu();
    }

    return { success: true, spk_no: spkNo, count_out: data.out_ctl.length, generated: generatedSpks, is_draft: isDraft };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * saveSPK SHR Final — OWNER FIX
 * Perubahan:
 *   - Owner HEADER & OUT = data.batch_owner (dari batch Stok_Sheet)
 *   - Owner_Used OUT     = shr.owner_used   (dari toggle UsedBy)
 * ========================================================================= */
/* =========================================================================
 * 🟢 T1 PATCH — saveSPK_SHR_Final (Batch_ID inheritance untuk traceability)
 *
 * Perubahan vs versi asli:
 *  - SHR-HEADER : tambah field 'Batch_ID' = data.batch_id (sheet batch source)
 *  - SHR-OUT    : tambah field 'Batch_ID' = data.batch_id (inherit dari header)
 *
 * Logic CT dan field lain TIDAK BERUBAH.
 *
 * Catatan: ini untuk SHR STANDALONE (input dari Stok_Sheet).
 * SHR turunan CTL di-generate dari saveSPK_CTL (patch T1a).
 * ========================================================================= */
function saveSPK_SHR_Final(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    // ✅ Validasi stok sheet/WIP sebelum lanjut — teruskan source_loc supaya
    //    validator tahu harus cek Stok_Sheet atau Stok_WIP (fallback 2-arah).
    if (typeof validateSheetAvailability === 'function') {
      validateSheetAvailability(data.batch_id, data.qty_input, data.kg_input, null, data.source_loc);
    }
    const spkSheet   = getSheet("SPK");
    if (!spkSheet) throw new Error("Sheet 'SPK' tidak ditemukan!");
    const allDataSHR = spkSheet.getDataRange().getValues();
    const headers    = allDataSHR[0].map(function(h) { return String(h).trim(); });
    const timestamp  = new Date();
    const spkNo      = data.spk_no || getNextSpkNo('SHR', allDataSHR, headers);

    const batchOwner = String(data.batch_owner || data.owner || 'FC').trim().toUpperCase();

    function buildRow(rowData) {
      return headers.map(function(h) { return rowData[h] !== undefined ? rowData[h] : ''; });
    }

    // LANGKAH 1: Pre-kalkulasi CT tiap SHR-OUT
    var outCtList = data.out_shr.map(function(shr) {
      return hitungRencanaDurasi(shr.item_code, shr.qty_plan, shr.kg_plan);
    });

    // LANGKAH 2: Hitung agregat SHR-HEADER
    var hdrPlanRun     = 0;
    var hdrTotalDurasi = 0;
    outCtList.forEach(function(ct) {
      hdrPlanRun     += ct.planRun;
      hdrTotalDurasi += ct.planSetup + ct.planRun;
    });

    // LANGKAH 3: Tulis SHR-HEADER
    var headerData = {
      'SPK_No'             : spkNo,
      'SPK_Type'           : 'SHR-HEADER',
      'Parent_SPK'         : data.batch_id,
      'Batch_ID'           : data.batch_id,        // 🟢 T1: batch sheet source (standalone)
      'Tgl_Buat'           : timestamp,
      'Priority'           : data.priority,
      'Source_Loc'         : data.source_loc,
      'Item_Code'          : data.item_input,
      'Input_Spec'         : data.spec_input,
      'Qty_Target'         : data.qty_input,
      'KG_Target'          : data.kg_input,
      'MC_No'              : data.machine,
      'Status'             : 'ANTRIAN',
      'OP'                 : data.op,
      'Owner'              : batchOwner,
      'Owner_Used'         : batchOwner,
      'Created_By'         : 'Admin',
      'Plan_Setup_Menit'   : 0,
      'Plan_Run_Menit'     : hdrPlanRun,
      'Total_Durasi_Menit' : hdrTotalDurasi,
      'T'                  : parseFloat(data.thick) || 0
    };
    _appendRowSafe(spkSheet, buildRow(headerData));

    // LANGKAH 4: Tulis SHR-OUT
    data.out_shr.forEach(function(shr, idx) {
      var outId  = spkNo + '-' + String(idx + 1).padStart(2, '0');
      var ctOut  = outCtList[idx];
      var ownerUsed = String(shr.owner_used || batchOwner).trim().toUpperCase();

      var outData = {
        'SPK_No'             : outId,
        'SPK_Type'           : 'SHR-OUT',
        'SO_Ref'             : shr.so_no       || '',
        'Cust'               : shr.cust         || '',
        'Parent_SPK'         : spkNo,
        'Batch_ID'           : data.batch_id,        // 🟢 T1: inherit dari header (sheet batch)
        'Tgl_Buat'           : timestamp,
        'Priority'           : data.priority,
        'Source_Loc'         : data.machine,
        'Item_Code'          : shr.item_code,
        'Input_Spec'         : shr.description,
        'Target_Loc'         : shr.target_loc,
        'MC_No'              : data.machine,
        'BQ'                 : shr.cut_p * shr.cut_l,
        'Qty_Target'         : shr.qty_plan,
        'KG_Target'          : shr.kg_plan,
        'Status'             : 'ANTRIAN',
        'OP'                 : data.op,
        'Owner'              : batchOwner,
        'Owner_Used'         : ownerUsed,
        'Created_By'         : 'Admin',
        'NOTE'               : 'Cut_P:' + shr.cut_p + ' Cut_L:' + shr.cut_l,
        'Plan_Setup_Menit'   : ctOut.planSetup,
        'Plan_Run_Menit'     : ctOut.planRun,
        'Total_Durasi_Menit' : 0,
        'T'                  : parseFloat(data.thick) || 0
      };
      _appendRowSafe(spkSheet, buildRow(outData));
    });

    SpreadsheetApp.flush();
    if (typeof kalkulasiEstimasiWaktu === 'function') kalkulasiEstimasiWaktu();

    return { success: true, spk_no: spkNo, count_out: data.out_shr.length };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * 3. SAVE BOARD ACTUAL DATA (Operator Klik Selesai & Otomatis Roll-Up ke Atas)
 * ========================================================================= */
function saveBoardActualData(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheet     = ss.getSheetByName("SPK");
    let rows        = sheet.getDataRange().getValues();
    const headers   = rows[0].map(function(h) { return String(h).trim(); });
    const timestamp = new Date();
    
    const iSpk    = headers.indexOf("SPK_No");
    const iType   = headers.indexOf("SPK_Type");
    const iParent = headers.indexOf("Parent_SPK");
    const iQtyAct = headers.indexOf("Qty_Actual");
    const iKgAct  = headers.indexOf("KG_Actual");
    const iQtyNg  = headers.indexOf("Qty_NG");
    const iKgNg   = headers.indexOf("KG_NG");
    const iStatus = headers.indexOf("Status");
    const iSelesai= headers.indexOf("Selesai_DT");
    const iItem   = headers.indexOf("Item_Code");
    const iQtyTgt = headers.indexOf("Qty_Target");
    const iMC     = headers.indexOf("MC_No");
    const iOP     = headers.indexOf("OP");
    const iOwner  = headers.indexOf("Owner");
    const iOwnerU = headers.indexOf("Owner_Used");
    const iTgtLoc = headers.indexOf("Target_Loc");
    const iSoRef  = headers.indexOf("SO_Ref");
    const iCust   = headers.indexOf("Cust");
    const iBatch  = headers.indexOf("Batch_ID");
    const iSpec   = headers.indexOf("Input_Spec");

    let childRowIdx  = -1;
    let parentSpkNo  = "";
    let childRowData = null;
    
    // 1. Simpan Aktual pada Baris Anak (CTL-OUT / SHR-OUT)
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][iSpk]).trim() === payload.spk_no.trim()
          && (rows[i][iType] === 'CTL-OUT' || rows[i][iType] === 'SHR-OUT')) {
        childRowIdx  = i + 1;
        parentSpkNo  = String(rows[i][iParent]).trim();
        childRowData = rows[i];
        
        sheet.getRange(childRowIdx, iQtyAct + 1).setValue(Math.round(payload.qty_actual));
        sheet.getRange(childRowIdx, iKgAct  + 1).setValue(Math.round(payload.kg_actual));
        if (payload.qty_ng !== undefined && payload.qty_ng !== null) {
          sheet.getRange(childRowIdx, iQtyNg + 1).setValue(Math.round(payload.qty_ng));
          rows[i][iQtyNg] = Math.round(payload.qty_ng);
        }
        if (payload.kg_ng !== undefined && payload.kg_ng !== null) {
          sheet.getRange(childRowIdx, iKgNg + 1).setValue(Math.round(payload.kg_ng));
          rows[i][iKgNg] = Math.round(payload.kg_ng);
        }
        sheet.getRange(childRowIdx, iStatus + 1).setValue("DONE");                         
        sheet.getRange(childRowIdx, iSelesai+ 1).setValue(timestamp);

        // Update rows di memori agar roll-up HEADER baca nilai terbaru
        rows[i][iQtyAct] = Math.round(payload.qty_actual);
        rows[i][iKgAct]  = Math.round(payload.kg_actual);
        rows[i][iStatus] = "DONE";
        rows[i][iSelesai]= timestamp;
        break;
      }
    }
    
    // 2. Update HEADER: rollup + Is_Habis logic + NG items processing
    let parentRowIdx  = -1;
    let parentQtyTgt  = 0;
    let parentKgTgt   = 0;
    let totalKgActOut = 0;
    let headerIsHabis  = false;
    let parentMcNo     = '';
    let parentOwner    = '';
    let parentItemCode = '';
    let parentSpec     = '';
    const iIsHabis     = headers.indexOf("Is_Habis");
    const iKgTgt       = headers.indexOf("KG_Target");

    if (parentSpkNo) {
      for (let i = 1; i < rows.length; i++) {
        const rowSpk    = String(rows[i][iSpk]    || '').trim();
        const rowParent = String(rows[i][iParent] || '').trim();
        const rowType   = String(rows[i][iType]   || '');
        const rowStatus = String(rows[i][iStatus] || '').toUpperCase();

        if (rowSpk === parentSpkNo) {
          parentRowIdx     = i + 1;
          parentQtyTgt     = Number(rows[i][iQtyTgt]) || 0;
          parentKgTgt      = Number(rows[i][iKgTgt])  || 0;
          parentMcNo       = String(rows[i][iMC]      || '').trim();
          parentOwner      = String(rows[i][iOwner]   || '').trim();
          parentItemCode   = String(rows[i][iItem]    || '').trim();
          headerIsHabis    = (rows[i][iIsHabis] === true || String(rows[i][iIsHabis]).toUpperCase() === 'TRUE');
        }

        if (rowParent === parentSpkNo
            && (rowType === 'SHR-OUT' || rowType === 'CTL-OUT')
            && rowStatus !== 'CANCELLED') {
          totalKgActOut += Number(rows[i][iKgAct]) || 0;
        }
      }
    }

    // 2.b NG Items Processing (CTL + SHR, kalau ada payload.ng_items)
    // 🟢 BATCH 2 — Extend NG handling ke SHR. NG type dinamis (CTL-NG / SHR-NG).
    const isCtlHeader = parentSpkNo && String(parentSpkNo).indexOf('CTL-') === 0;
    const isShrHeader = parentSpkNo && String(parentSpkNo).indexOf('SHR-') === 0;
    const isAnyHeader = isCtlHeader || isShrHeader;
    const ngType      = isCtlHeader ? 'CTL-NG' : (isShrHeader ? 'SHR-NG' : '');

    let createdNgList  = [];
    let totalNgQty     = 0;
    let totalNgKg      = 0;

    if (isAnyHeader && payload.ng_items && Array.isArray(payload.ng_items) && payload.ng_items.length > 0) {
      // Lookup Spec dari M_ITEM via parentItemCode
      if (parentItemCode) {
        var miSheet2 = ss.getSheetByName('M_ITEM');
        if (miSheet2) {
          var miData2 = miSheet2.getDataRange().getValues();
          var miHdr2  = miData2[0].map(function(h){ return String(h).trim(); });
          var miIc2   = miHdr2.indexOf('Item_Code');
          var miSp2   = miHdr2.indexOf('Spec');
          for (var m2 = 1; m2 < miData2.length; m2++) {
            if (String(miData2[m2][miIc2] || '').trim() === parentItemCode) {
              parentSpec = String(miData2[m2][miSp2] || '').trim();
              break;
            }
          }
        }
      }

      let maxNgSuffix = 0;
      // Cari max suffix NG existing untuk parent ini (filter by ngType)
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][iType] === ngType && String(rows[i][iParent]).trim() === parentSpkNo) {
          const m = String(rows[i][iSpk]).match(/-NG(\d+)$/);
          if (m) {
            const n = parseInt(m[1], 10) || 0;
            if (n > maxNgSuffix) maxNgSuffix = n;
          }
        }
      }

      // Append 1 NG row per ukuran NG
      payload.ng_items.forEach(function(ng, ngIdx) {
        try {
          const ngT = parseFloat(ng.t)     || 0;
          const ngP = parseFloat(ng.p)     || 0;
          const ngL = parseFloat(ng.l)     || 0;
          const ngQ = parseInt(ng.qty, 10) || 0;

          Logger.log('NG iter ' + ngIdx + ' (' + ngType + '): T=' + ngT + ' P=' + ngP + ' L=' + ngL + ' Q=' + ngQ);

          if (ngT <= 0 || ngP <= 0 || ngL <= 0 || ngQ <= 0) {
            Logger.log('NG iter ' + ngIdx + ' SKIPPED — invalid dims');
            return;
          }

          maxNgSuffix++;
          const ngSpkNo = parentSpkNo + '-NG' + String(maxNgSuffix).padStart(2, '0');
          const ngKg    = Math.round(ngT * ngP * ngL * ngQ * 7.85 / 1000000);

          const ngData = {
            'SPK_No'     : ngSpkNo,
            'SPK_Type'   : ngType,             // 🟢 dynamic: CTL-NG atau SHR-NG
            'Parent_SPK' : parentSpkNo,
            'Tgl_Buat'   : timestamp,
            'Item_Code'  : parentItemCode,
            'Input_Spec' : ngT + ' x ' + ngP + ' x ' + ngL,
            'Qty_Target' : ngQ,
            'KG_Target'  : ngKg,
            'Qty_Actual' : ngQ,
            'KG_Actual'  : ngKg,
            'MC_No'      : parentMcNo,
            'Owner'      : parentOwner,
            'Owner_Used' : parentOwner,
            'Status'     : 'DONE',
            'Target_Loc' : 'Stok_NG',
            'Selesai_DT' : timestamp,
            'Created_By' : payload.created_by || 'operator',
            'T'          : ngT,
            'P'          : ngP,
            'L'          : ngL
          };
          const ngRow = headers.map(function(h) { return ngData[h] !== undefined ? ngData[h] : ''; });
          _appendRowSafe(sheet, ngRow);
          SpreadsheetApp.flush();

          createdNgList.push({
            spk_no    : ngSpkNo,
            item_code : parentItemCode,
            spec      : parentSpec,
            t         : ngT,
            p         : ngP,
            l         : ngL,
            qty       : ngQ,
            kg        : ngKg,
            owner     : parentOwner,
            mc_no     : parentMcNo
          });

          totalNgQty += ngQ;
          totalNgKg  += ngKg;

          Logger.log('NG iter ' + ngIdx + ' OK — appended ' + ngSpkNo);
        } catch (ngErr) {
          Logger.log('NG iter ' + ngIdx + ' FAILED: ' + ngErr.toString() + ' | stack: ' + (ngErr.stack || ''));
        }
      });
    }

    // 2.c Update HEADER (rollup KG_Actual + NG totals)
    if (parentRowIdx !== -1) {
      // KG_Actual logic — compute dulu
      let finalKgAct;
      if (isCtlHeader) {
        // CTL: Is_Habis lock ke plan, kalau tidak habis = sum OUT + NG
        finalKgAct = headerIsHabis ? Math.round(parentKgTgt) : Math.round(totalKgActOut + totalNgKg);
      } else if (isShrHeader) {
        // 🟢 BATCH 2 — SHR: sum DONE OUT + NG (mirror CTL behavior, tanpa Is_Habis lock)
        finalKgAct = Math.round(totalKgActOut + totalNgKg);
      } else {
        finalKgAct = Math.round(totalKgActOut);
      }
      sheet.getRange(parentRowIdx, iKgAct + 1).setValue(finalKgAct);

      // CTL-HEADER konsumsi coil → Qty_Actual = KG_Actual (1:1). SHR/lainnya pakai parentQtyTgt.
      var qtyActFinal = isCtlHeader ? finalKgAct : parentQtyTgt;
      sheet.getRange(parentRowIdx, iQtyAct + 1).setValue(qtyActFinal);

      // 🟢 BATCH 2 — Update aggregate NG totals di HEADER (CTL + SHR)
      if (isAnyHeader && (totalNgQty > 0 || totalNgKg > 0)) {
        const curHeaderQtyNg = parseFloat(rows[parentRowIdx - 1][iQtyNg]) || 0;
        const curHeaderKgNg  = parseFloat(rows[parentRowIdx - 1][iKgNg])  || 0;
        sheet.getRange(parentRowIdx, iQtyNg + 1).setValue(curHeaderQtyNg + totalNgQty);
        sheet.getRange(parentRowIdx, iKgNg + 1).setValue(Math.round(curHeaderKgNg + totalNgKg));
      }
    }
    
    SpreadsheetApp.flush();

    // 3. Penulisan ke Stok & Trace_Log setelah DONE
    if (childRowData && payload.qty_actual > 0) {
      const spkType   = String(childRowData[iType]   || '');
      const targetLoc = String(childRowData[iTgtLoc] || '').trim();
      const owner     = String(childRowData[iOwner]  || '').trim();
      const ownerUsed = String(childRowData[iOwnerU] || '').trim();
      const itemCode  = String(childRowData[iItem]   || '').trim();
      const soRef     = String(childRowData[iSoRef]  || '').trim();
      const cust      = String(childRowData[iCust]   || '').trim();
      const mcNo      = String(childRowData[iMC]     || '').trim();
      const op        = String(childRowData[iOP]     || '').trim();
      const srcBatch  = iBatch > -1 ? String(childRowData[iBatch] || '').trim() : '';
      const qtyAct    = Math.round(payload.qty_actual);
      const kgAct     = Math.round(payload.kg_actual);

      // Ambil desc, spec, T, P, L dari M_ITEM
      var tDim = '', pDim = '', lDim = '', desc = '', specCode = '', uom = 'Sht';
      var miSheet = ss.getSheetByName('M_ITEM');
      if (miSheet) {
        var miData = miSheet.getDataRange().getValues();
        var miHdr  = miData[0].map(function(h){ return String(h).trim(); });
        var miIc   = miHdr.indexOf('Item_Code');
        var miDesc = miHdr.indexOf('Description');
        var miSpec = miHdr.indexOf('Spec');
        var miT    = miHdr.indexOf('T');
        var miP    = miHdr.indexOf('P');
        var miL    = miHdr.indexOf('L');
        var miUom  = miHdr.indexOf('UoM_MC');
        for (var m = 1; m < miData.length; m++) {
          if (String(miData[m][miIc] || '').trim() === itemCode) {
            desc     = String(miData[m][miDesc] || '').trim();
            specCode = String(miData[m][miSpec] || '').trim();
            tDim     = String(miData[m][miT]    || '').trim();
            pDim     = String(miData[m][miP]    || '').trim();
            lDim     = String(miData[m][miL]    || '').trim();
            uom      = String(miData[m][miUom]  || 'Sht').trim();
            break;
          }
        }
      }

      var rootBatch    = getRootBatch(srcBatch) || srcBatch;
      var supplierInfo = getCoilSupplierInfo(rootBatch);

      // CTL-OUT DONE → tulis Stok_Sheet
      if (spkType === 'CTL-OUT') {
        var newBatchSHT = generateBatchId('SHT');
        writeStokSheet({
          batch_id    : newBatchSHT,
          tgl_masuk   : timestamp,
          item_code   : itemCode,
          description : desc,
          qty_in      : qtyAct,
          kg_in       : kgAct,
          owner       : owner,
          no_do       : supplierInfo.no_do,
          no_po       : supplierInfo.no_po,
          supplier    : supplierInfo.supplier,
          source_batch: srcBatch || rootBatch
        });
        writeTraceLog({
          batch_id    : newBatchSHT,
          tgl_buat    : timestamp,
          level       : 1,
          type        : 'SHEET',
          source_batch: srcBatch || rootBatch,
          root_batch  : rootBatch,
          spk_ref     : payload.spk_no,
          gr_ref      : rootBatch,
          item_code   : itemCode,
          description : desc,
          spec        : specCode,
          t           : tDim,
          p           : pDim,
          l_dim       : lDim,
          qty         : qtyAct,
          kg          : kgAct,
          operator    : op,
          mc_no       : mcNo,
          tgl_prod    : timestamp,
          supplier    : supplierInfo.supplier,
          no_po       : supplierInfo.no_po,
          no_do       : supplierInfo.no_do,
          owner       : owner,
          owner_used  : ownerUsed
        });
      }

      // SHR-OUT DONE → tulis Stok_WIP atau Stok_FG atau Stok_Sheet
      if (spkType === 'SHR-OUT') {
        var isWIP   = (targetLoc === 'WIP_Cust' || targetLoc === 'WIP_Stamping');
        var isFG    = (targetLoc === 'FG_Cust'  || targetLoc === 'FG_RM_Stamping');
        var isSht   = (targetLoc === 'Stok_Sheet');
        var newBatch = '';

        if (isWIP) {
          newBatch = generateBatchId('WIP');
          writeStokWIP({
            batch_id    : newBatch,
            tgl_masuk   : timestamp,
            spk_ref     : payload.spk_no,
            source_batch: srcBatch,
            root_batch  : rootBatch,
            so_ref      : soRef,
            stp_ref     : '',
            cust        : cust,
            item_code   : itemCode,
            description : desc,
            spec        : specCode,
            t           : tDim,
            p           : pDim,
            l_dim       : lDim,
            uom         : uom,
            qty_in      : qtyAct,
            kg_in       : kgAct,
            owner       : owner,
            owner_used  : ownerUsed
          });
        }

        if (isFG) {
          var fgType = getFgBatchType(srcBatch);
          newBatch   = generateBatchId(fgType);
          writeStokFG({
            batch_id    : newBatch,
            tgl_output  : timestamp,
            spk_ref     : payload.spk_no,
            so_ref      : soRef,
            item_code   : itemCode,
            description : desc,
            spec        : specCode,
            t           : tDim,
            p           : pDim,
            l_dim       : lDim,
            uom         : uom,
            qty_in      : qtyAct,
            kg_in       : kgAct,
            owner       : owner,
            owner_used  : ownerUsed,
            target_loc  : targetLoc
          });
        }

        // SHR-OUT → Stok_Sheet (sisa potongan / trimming kembali ke stok lembaran)
        if (isSht) {
          newBatch = generateBatchId('SHT');
          writeStokSheet({
            batch_id    : newBatch,
            tgl_masuk   : timestamp,
            item_code   : itemCode,
            description : desc,
            qty_in      : qtyAct,
            kg_in       : kgAct,
            owner       : owner,
            no_do       : supplierInfo.no_do,
            no_po       : supplierInfo.no_po,
            supplier    : supplierInfo.supplier,
            source_batch: srcBatch || rootBatch
          });
        }

        if (newBatch) {
          writeTraceLog({
            batch_id    : newBatch,
            tgl_buat    : timestamp,
            level       : 2,
            type        : isWIP ? 'WIP' : isSht ? 'SHEET' : (targetLoc === 'FG_Cust' ? 'FGC' : 'FGS'),
            source_batch: srcBatch,
            root_batch  : rootBatch,
            spk_ref     : payload.spk_no,
            gr_ref      : rootBatch,
            item_code   : itemCode,
            description : desc,
            spec        : specCode,
            t           : tDim,
            p           : pDim,
            l_dim       : lDim,
            qty         : qtyAct,
            kg          : kgAct,
            operator    : op,
            mc_no       : mcNo,
            tgl_prod    : timestamp,
            supplier    : supplierInfo.supplier,
            no_po       : supplierInfo.no_po,
            no_do       : supplierInfo.no_do,
            owner       : owner,
            owner_used  : ownerUsed
          });
        }

        // Rekap ICT: Owner ≠ Owner_Used DAN Target_Loc = FG
        // 🔒 GUARD: hanya trigger cross-billing kalau hasilnya sudah jadi FG (FG_Cust / FG_RM_Stamping)
        if (owner && ownerUsed && owner !== ownerUsed && isFG) {
          writeRekapICT({
            tgl        : timestamp,
            spk_no     : payload.spk_no,
            item_code  : itemCode,
            description: desc,
            dari_owner : owner,
            ke_owner   : ownerUsed,
            qty        : qtyAct,
            kg         : kgAct
          });
        }
      }
    }

    return { success: true, ng_created: createdNgList };
  } catch (e) {
    Logger.log('saveBoardActualData error: ' + e.toString());
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * 4. ENGINE EVALUASI SELESAI & ESTIMASI (JADWAL & UTILS)
 * ========================================================================= */
function checkHeaderFinishedStatus(parentSpkNo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SPK");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(function(h) { return String(h).trim(); });

  // 🟢 BUG-FIX #9 — Pakai indexOf untuk semua kolom, jangan hardcode
  const iSpk       = headers.indexOf("SPK_No");
  const iType      = headers.indexOf("SPK_Type");
  const iParent    = headers.indexOf("Parent_SPK");
  const iStatus    = headers.indexOf("Status");
  const statusCol  = iStatus + 1;
  const selesaiCol = headers.indexOf("Selesai_DT") + 1;

  const isShr            = String(parentSpkNo).toUpperCase().indexOf("SHR-") === 0;
  const targetHeaderType = isShr ? 'SHR-HEADER' : 'CTL-HEADER';
  const targetOutType    = isShr ? 'SHR-OUT'    : 'CTL-OUT';

  let anyChildRunning = false;
  let headerRowIndex  = -1;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][iType] === targetHeaderType && String(rows[i][iSpk]).trim() === parentSpkNo.trim()) {
      headerRowIndex = i + 1;
    }
    if (rows[i][iType] === targetOutType && String(rows[i][iParent]).trim() === parentSpkNo.trim()) {
      const childStatus = String(rows[i][iStatus] || '').toUpperCase();
      if (childStatus !== 'DONE' && childStatus !== 'CANCELLED') {
        anyChildRunning = true;
      }
    }
  }

  if (!anyChildRunning && headerRowIndex > -1) {
    sheet.getRange(headerRowIndex, statusCol).setValue("DONE");
    sheet.getRange(headerRowIndex, selesaiCol).setValue(new Date());
    SpreadsheetApp.flush();
    return false;
  }
  return true;
}

function updateStatusToRunning(spkNo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SPK");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(function(h) { return String(h).trim(); });

  // 🟢 BUG-FIX #9 — Pakai indexOf untuk semua kolom, jangan hardcode
  const iSpk    = headers.indexOf("SPK_No");
  const iType   = headers.indexOf("SPK_Type");
  const iParent = headers.indexOf("Parent_SPK");
  const iStatus = headers.indexOf("Status");
  const statusCol = iStatus + 1;
  const mulaiCol  = headers.indexOf("Mulai_DT") + 1;

  const now = new Date();

  // Tentukan tipe HEADER dan OUT berdasarkan prefix SPK
  const isCtl      = String(spkNo).toUpperCase().indexOf("SHR-") !== 0;
  const headerType = isCtl ? 'CTL-HEADER' : 'SHR-HEADER';
  const outType    = isCtl ? 'CTL-OUT'    : 'SHR-OUT';

  for (let i = 1; i < rows.length; i++) {
    const rowSpk    = String(rows[i][iSpk]).trim();
    const rowType   = rows[i][iType];
    const rowParent = String(rows[i][iParent] || '').trim();

    // Update HEADER
    if (rowSpk === spkNo.trim() && rowType === headerType) {
      sheet.getRange(i + 1, statusCol).setValue("RUNNING");
      sheet.getRange(i + 1, mulaiCol).setValue(now);
    }

    // Cascade ke semua OUT anak langsung
    if (rowParent === spkNo.trim() && rowType === outType) {
      sheet.getRange(i + 1, statusCol).setValue("RUNNING");
      sheet.getRange(i + 1, mulaiCol).setValue(now);
    }
  }
  SpreadsheetApp.flush();
  return { success: true };
}

function cancelSpkBackend(spkNo) {
  return cascadeCancelSpk(spkNo);
}

function fetchBoardQueueData(machineNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spkSheet = ss.getSheetByName("SPK");
  const stokSheet = ss.getSheetByName("Stok_Coil");
  if (!spkSheet) return [];
  
  const spkRows = spkSheet.getDataRange().getValues();
  const spkHeaders = spkRows[0].map(function(h) { return String(h).trim(); });
  
  const idxType = spkHeaders.indexOf("SPK_Type");
  const idxMc = spkHeaders.indexOf("MC_No");
  const idxStatus = spkHeaders.indexOf("Status");
  const idxParent = spkHeaders.indexOf("Parent_SPK");

  let stokMap = {};
  if (stokSheet) {
    const stkRows = stokSheet.getDataRange().getValues();
    const stkHeaders = stkRows[0].map(function(h) { return String(h).trim(); });
    const idxBatch = stkHeaders.indexOf("Batch_ID");
    const idxKgIn = stkHeaders.indexOf("KG_In");
    const idxKgDone = stkHeaders.indexOf("KG_Done");
    const idxSpec = stkHeaders.indexOf("Spec");
    
    if (idxBatch > -1) {
      for(let i = 1; i < stkRows.length; i++) {
        let kgIn = parseFloat(stkRows[i][idxKgIn]) || 0;
        let kgDone = parseFloat(stkRows[i][idxKgDone]) || 0;
        stokMap[String(stkRows[i][idxBatch]).trim()] = {
           kg_fisik: kgIn - kgDone,
           spec: idxSpec > -1 ? String(stkRows[i][idxSpec]).trim() : ""
        };
      }
    }
  }

  let queueList = [];
  for (let i = 1; i < spkRows.length; i++) {
    if (spkRows[i][idxType] === 'CTL-HEADER' && 
        String(spkRows[i][idxMc]).trim().toUpperCase() === machineNo.toUpperCase() && 
        (spkRows[i][idxStatus] === 'ANTRIAN' || spkRows[i][idxStatus] === 'RUNNING')) {
          
      let rowData = {};
      spkHeaders.forEach(function(h, idx) {
        let val = spkRows[i][idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
        }
        rowData[h] = val;
      });
      const coilData = stokMap[String(spkRows[i][idxParent]).trim()] || {kg_fisik: 0, spec: ""};
      rowData['KG_Fisik'] = coilData.kg_fisik; 
      rowData['Mat_Spec'] = coilData.spec; 
      queueList.push(rowData);
    }
  }
  return queueList.sort(function(a, b) {
    const prioWeight = { 'Urgent': 3, 'High': 2, 'Normal': 1 };
    return (prioWeight[b.Priority] || 1) - (prioWeight[a.Priority] || 1);
  });
}

/**
 * 🟢 Fetch SPK HEADER yang DONE hari ini, untuk mesin tertentu.
 * Dipakai di board untuk counter "Selesai Hari Ini".
 */
function fetchTodayDoneJobs(machineNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spkSheet = ss.getSheetByName("SPK");
  if (!spkSheet) return [];
  
  const spkRows = spkSheet.getDataRange().getValues();
  const spkHeaders = spkRows[0].map(function(h) { return String(h).trim(); });
  
  const idxType    = spkHeaders.indexOf("SPK_Type");
  const idxMc      = spkHeaders.indexOf("MC_No");
  const idxStatus  = spkHeaders.indexOf("Status");
  const idxSelesai = spkHeaders.indexOf("Selesai_DT");
  
  if (idxType < 0 || idxMc < 0 || idxStatus < 0 || idxSelesai < 0) return [];
  
  // Range hari ini (jam 00:00 - 23:59)
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrowStart = new Date(todayStart.getTime() + 24*60*60*1000);
  
  const isShr = String(machineNo).toUpperCase().indexOf("SHR") === 0;
  const headerType = isShr ? 'SHR-HEADER' : 'CTL-HEADER';
  
  let doneList = [];
  for (let i = 1; i < spkRows.length; i++) {
    if (spkRows[i][idxType] !== headerType) continue;
    if (String(spkRows[i][idxMc]).trim().toUpperCase() !== machineNo.toUpperCase()) continue;
    if (spkRows[i][idxStatus] !== 'DONE') continue;
    
    var selesaiDt = spkRows[i][idxSelesai];
    if (!(selesaiDt instanceof Date)) continue;
    if (selesaiDt < todayStart || selesaiDt >= tomorrowStart) continue;
    
    let rowData = {};
    spkHeaders.forEach(function(h, idx) {
      let val = spkRows[i][idx];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
      }
      rowData[h] = val;
    });
    doneList.push(rowData);
  }
  
  // Sort by Selesai_DT terbaru dulu
  return doneList.sort(function(a, b) {
    return String(b.Selesai_DT || '').localeCompare(String(a.Selesai_DT || ''));
  });
}

/* =========================================================================
 * BUG FIX #2 — fetchBoardQueueDataSHR
 * Sebelumnya: SHR-HEADER DONE-hari-ini masih masuk _queueList → bisa di-pick
 * lagi & menyebabkan tampilan board reset ke MULAI setelah submit.
 * Solusi: filter hanya ANTRIAN / RUNNING (sama pattern dengan CTL).
 * DONE-hari-ini sudah dihandle terpisah oleh fetchTodayDoneJobs.
 * ========================================================================= */
function fetchBoardQueueDataSHR(machineNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SPK");
  if (!sheet) throw new Error("Sheet SPK not found");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(function(h) { return String(h).trim(); });
  const idxType = headers.indexOf("SPK_Type");
  const idxMc = headers.indexOf("MC_No");
  const idxStatus = headers.indexOf("Status");

  let queueList = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idxType] === 'SHR-HEADER' &&
        String(rows[i][idxMc]).trim().toUpperCase() === machineNo.toUpperCase() &&
        (rows[i][idxStatus] === 'ANTRIAN' || rows[i][idxStatus] === 'RUNNING')) {
      let rowData = {};
      headers.forEach(function(h, idx) {
        let val = rows[i][idx];
        if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
        rowData[h] = val;
      });
      queueList.push(rowData);
    }
  }
  return queueList.sort(function(a, b) {
    const prioWeight = { 'Urgent': 3, 'High': 2, 'Normal': 1 };
    return (prioWeight[b.Priority] || 1) - (prioWeight[a.Priority] || 1);
  });
}


function fetchBoardChildData(parentSpkNo, spkType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SPK");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(function(h) { return String(h).trim(); });
  let children = [];
  for (let i = 1; i < rows.length; i++) {
    let rowData = {};
    headers.forEach(function(h, idx) { 
      let val = rows[i][idx];
      if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
      rowData[h] = val;
    });
    if (rowData['SPK_Type'] === spkType && String(rowData['Parent_SPK']).trim() === parentSpkNo.trim()) {
      const tQty = parseFloat(rowData['Qty_Target']) || 1;
      const tKg = parseFloat(rowData['KG_Target']) || 0;
      rowData['BQ'] = tKg / tQty; 
      children.push(rowData);
    }
  }
  return children;
}

function getNoCoilByBatch(batchId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Stok_Coil");

  if (!sh) return ''; // ✅ FIX: jangan crash

  const data = sh.getDataRange().getValues();
  const header = data[0];

  const idxBatch = header.indexOf("Batch_ID");
  const idxNoCoil = header.indexOf("No_Coil");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxBatch]).trim() === String(batchId).trim()) {
      return data[i][idxNoCoil];
    }
  }

  return '';
}


function getNoCoilMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Stok_Coil");
  if (!sheet) throw new Error("Sheet not found: Stok_Coil");

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};

  var headers = data[0];
  var iBatch = headers.indexOf("Batch_ID");
  var iNoCoil = headers.indexOf("No_Coil");

  var map = {};

  for (var i = 1; i < data.length; i++) {
    var batch = String(data[i][iBatch] || '').trim();
    var coil  = String(data[i][iNoCoil] || '').trim();
    if (batch) map[batch] = coil;
  }

  return map;
}


function getLiveBoardDataBackend() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("SPK");
    if (!sheet) throw new Error("Sheet SPK not found");
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0].map(function(h) { return String(h || '').trim(); });

    // ====== INDEX KOLOM ======
    var iSpk        = headers.indexOf("SPK_No");
    var iType       = headers.indexOf("SPK_Type");
    var iParent     = headers.indexOf("Parent_SPK");
    var iDate       = headers.indexOf("Tgl_Buat")  !== -1 ? headers.indexOf("Tgl_Buat")  : headers.indexOf("Date");
    var iMachine    = headers.indexOf("MC_No")     !== -1 ? headers.indexOf("MC_No")     : headers.indexOf("Machine");
    var iSpecIn     = headers.indexOf("Input_Spec");
    var iQtyTgt     = headers.indexOf("Qty_Tar")   !== -1 ? headers.indexOf("Qty_Tar")   : headers.indexOf("Qty_Target");
    var iQtyAct     = headers.indexOf("Qty_Act")   !== -1 ? headers.indexOf("Qty_Act")   : headers.indexOf("Qty_Actual");
    var iKgTgt      = headers.indexOf("KG_Tar")    !== -1 ? headers.indexOf("KG_Tar")    : headers.indexOf("KG_Target");
    var iKgAct      = headers.indexOf("KG_Act")    !== -1 ? headers.indexOf("KG_Act")    : headers.indexOf("KG_Actual");
    var iCust       = headers.indexOf("Cust")      !== -1 ? headers.indexOf("Cust")      : headers.indexOf("Customer");
    var iStatus     = headers.indexOf("Status");
    var iEstMulai   = headers.indexOf("Estimasi_Jam_Mulai");
    var iEstSelesai = headers.indexOf("Estimasi_Jam_Selesai");
    // ⭐ TAMBAHAN UNTUK PRINT SPK
    var iSoRef      = headers.indexOf("SO_Ref");
    var iSrcLoc     = headers.indexOf("Source_Loc");
    var iTgtLoc     = headers.indexOf("Target_Loc");
    var iOp         = headers.indexOf("OP");
    var iNote       = headers.indexOf("NOTE")      !== -1 ? headers.indexOf("NOTE")      : headers.indexOf("Note");
    var iLeader     = headers.indexOf("Leader");
    var coilMap     = getNoCoilMap();

    // ====== HELPER: parsing Cut_P / Cut_L dari NOTE ======
    function parseCut(note, key) {
      if (!note) return 1;
      var m = String(note).match(new RegExp(key + ':\\s*(\\d+)', 'i'));
      return m ? parseInt(m[1], 10) : 1;
    }

    var parentList = [];
    var parentMap = {};
    var childrenMap = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var spkNo = String(row[iSpk] || '').trim();
      if (!spkNo) continue;

      var spkType   = String(row[iType] || '').toUpperCase();
      var parentSpk = String(row[iParent] || '').trim();
      var statusStr = row[iStatus] ? String(row[iStatus]).toUpperCase() : 'ANTRIAN';
      var dateStr = '';
      var rawDate = row[iDate];
      if (rawDate instanceof Date) {
        dateStr = rawDate.getDate().toString().padStart(2, '0') + "/" + (rawDate.getMonth() + 1).toString().padStart(2, '0');
      } else if (rawDate) {
        dateStr = String(rawDate).substring(0, 10);
      }

      // Ambil field tambahan (sekali, dipakai parent & child)
      var soRef   = iSoRef   !== -1 ? String(row[iSoRef]   || '').trim() : '';
      var srcLoc  = iSrcLoc  !== -1 ? String(row[iSrcLoc]  || '').trim() : '';
      var tgtLoc  = iTgtLoc  !== -1 ? String(row[iTgtLoc]  || '').trim() : '';
      var opName  = iOp      !== -1 ? String(row[iOp]      || '').trim() : '';
      var noteVal = iNote    !== -1 ? String(row[iNote]    || '').trim() : '';
      var leader  = iLeader  !== -1 ? String(row[iLeader]  || '').trim() : '';
      var mcNo    = iMachine !== -1 ? String(row[iMachine] || '').trim() : '';

      var isChild = (parentSpk !== '' && spkNo.indexOf(parentSpk) !== -1 && spkNo !== parentSpk) 
                    || spkType.indexOf('OUT') !== -1 
                    || spkNo.split('-').length > 2;

      if (isChild) {
        if (!childrenMap[parentSpk]) childrenMap[parentSpk] = [];
        childrenMap[parentSpk].push({
          spk_no     : spkNo,
          cust       : iCust !== -1 ? String(row[iCust] || '').trim() : '--',
          parent_spk : parentSpk,
          spec_size  : String(row[iSpecIn] || '').trim(),
          qty_tgt    : parseFloat(row[iQtyTgt]) || 0,
          qty_act    : parseFloat(row[iQtyAct]) || 0,
          kg_tgt     : iKgTgt !== -1 ? parseFloat(row[iKgTgt]) || 0 : 0,
          kg_act     : iKgAct !== -1 ? parseFloat(row[iKgAct]) || 0 : 0,
          status     : statusStr,
          est_mulai  : iEstMulai   !== -1 ? String(row[iEstMulai]   || '').trim() : '',
          est_selesai: iEstSelesai !== -1 ? String(row[iEstSelesai] || '').trim() : '',
          // ⭐ FIELD UNTUK PRINT
          so_ref     : soRef,
          source_loc : srcLoc,
          target_loc : tgtLoc,
          operator   : opName,
          mc_no      : mcNo,
          cut_p      : parseCut(noteVal, 'Cut_P'),
          cut_l      : parseCut(noteVal, 'Cut_L'),
          note       : noteVal
        });
      } else {
        var pObj = {
          Date       : dateStr,
          SPK_No     : spkNo,
          Machine    : mcNo || 'SHR-01',
          Cust       : iCust !== -1 && String(row[iCust]).trim() !== '' ? String(row[iCust]).trim() : '--',
          Parent_SPK : parentSpk || '--',
          Spec_Size  : String(row[iSpecIn] || '').trim(),
          Qty_Tgt    : parseFloat(row[iQtyTgt]) || 0,
          Qty_Act    : parseFloat(row[iQtyAct]) || 0,
          Kg_Tgt     : iKgTgt !== -1 ? parseFloat(row[iKgTgt]) || 0 : 0,
          Kg_Act     : iKgAct !== -1 ? parseFloat(row[iKgAct]) || 0 : 0,
          Status     : statusStr,
          Est_Mulai  : iEstMulai   !== -1 ? String(row[iEstMulai]   || '').trim() : '',
          Est_Selesai: iEstSelesai !== -1 ? String(row[iEstSelesai] || '').trim() : '',
          // ⭐ FIELD UNTUK PRINT (INDUK)
          SO_Ref     : soRef,
          Source_Loc : srcLoc,
          Target_Loc : tgtLoc,
          Operator   : opName,
          Batch_ID   : parentSpk || '--',   // batch = parent (coil source / WIP)
          No_Coil    : coilMap[parentSpk] || getNoCoilByBatch(parentSpk) || '',
          Leader     : leader,
          Note       : noteVal,
          children   : []
        };
        parentList.push(spkNo);
        parentMap[spkNo] = pObj;
      }
    }

    var finalResult = [];
    for (var j = parentList.length - 1; j >= 0; j--) {
      var pKey = parentList[j];
      var parentJob = parentMap[pKey];
      parentJob.children = childrenMap[pKey] || [];
      if (parentJob.children.length > 0 && parentJob.Status !== "CANCELLED") {
        var allDone = true; var anyRunning = false;
        parentJob.children.forEach(function(c) {
          if (c.status !== 'DONE') allDone = false;
          if (c.status === 'RUNNING') anyRunning = true;
        });
        if (allDone) parentJob.Status = "DONE";
        else if (anyRunning) parentJob.Status = "RUNNING";
      }
      finalResult.push(parentJob);
      if (finalResult.length >= 150) break;
    }
    return finalResult;
  } catch (e) {
    return [{ SPK_No: "ERROR", Machine: e.message, children: [] }];
  }
}


function getNextSequence(prefix) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SPK");
  if (!sheet) return 1;
  const lastRow = sheet.getLastRow();
  const yearPrefix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yy");
  if (lastRow < 2) return 1;
  const data = sheet.getRange("A2:A" + lastRow).getValues();
  let maxSeq = 0;
  const regex = new RegExp("^" + prefix + "-" + yearPrefix + "(\\d{4})");
  for (let i = 0; i < data.length; i++) {
    const spk = String(data[i][0]).trim().toUpperCase();
    const match = spk.match(regex);
    if (match) {
      const seq = parseInt(match[1], 10);
      if (seq > maxSeq) { maxSeq = seq; }
    }
  }
  return maxSeq + 1;
}

function generateSpkNoCTL() {
  const seq = getNextSequence("CL");
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yy");
  return "CL-" + year + ("0000" + seq).slice(-4);
}

function generateSpkNoSHR() {
  const seq = getNextSequence("SHR");
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yy");
  return "SHR-" + year + ("0000" + seq).slice(-4);
}
/* =========================================================================
 * getNextSpkNo — Generate nomor SPK berikutnya (scan sheet, max+1)
 * prefix  : 'CTL' | 'SHR' | 'SLT' | 'ALO'
 * allData : spkSheet.getDataRange().getValues()
 * headers : allData[0] yang sudah di-trim
 * Hasil   : PREFIX-YYNNNN  contoh: CTL-260001
 * ========================================================================= */
function getNextSpkNo(prefix, allData, headers) {
  var iSpk = headers.indexOf('SPK_No');
  var tz   = Session.getScriptTimeZone();
  var yy   = Utilities.formatDate(new Date(), tz, 'yy');
  var re   = new RegExp('^' + prefix + '-' + yy + '(\\d{4})$');
  var maxSeq = 0;
  for (var i = 1; i < allData.length; i++) {
    var spk = String(allData[i][iSpk] || '').trim();
    var m   = spk.match(re);
    if (m) {
      var s = parseInt(m[1], 10);
      if (s > maxSeq) maxSeq = s;
    }
  }
  return prefix + '-' + yy + String(maxSeq + 1).padStart(4, '0');
}

/* =========================================================================
 * isNewSpkFormat — Deteksi format baru (CTL/SHR/SLT/ALO-)
 * Dipakai untuk backward compat SPK lama (CL-) saat edit
 * ========================================================================= */
function isNewSpkFormat(spkNo) {
  return /^(CTL|SHR|SLT|ALO)-\d{6}$/.test(String(spkNo || ''));
}

function getMasterItems() { return getSheetData("M_ITEM"); }
function updateSPKStatus(data) { return { success: true }; }
function getEditableSpkList() {
  const sheet = getSheet("SPK");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const iSpk = headers.indexOf("SPK_No");
  const iType = headers.indexOf("SPK_Type");
  const iStatus = headers.indexOf("Status");
  const iParent = headers.indexOf("Parent_SPK");
  const iSpec = headers.indexOf("Input_Spec");

  let result = [];

  for (let i = 1; i < data.length; i++) {
    const type = data[i][iType];
    const status = String(data[i][iStatus] || '').toUpperCase();

    // 🔴 ONLY HEADER + STATUS OPEN/QUEUE
    if (
      type === 'CTL-HEADER' &&
      (status === 'ANTRIAN' || status === 'OPEN')
    ) {
      result.push({
        spk_no: data[i][iSpk],
        parent: data[i][iParent],
        spec  : data[i][iSpec],
        status: status
      });
    }
  }

  return result;
}

// =========================================================================
// TAMBAHAN SpkService.gs — Paste di bawah fungsi getEditableSpkList()
// =========================================================================

/* =========================================================================
 * GET SPK LIST CTL
 * Untuk view-list halaman SPK CTL.
 * Return: array CTL-HEADER sorted newest first (reverse row order)
 * ========================================================================= */
function getSpkListCTL() {
  var sheet = getSheet('SPK');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h){ return String(h).trim(); });

  var iSpk    = hdr.indexOf('SPK_No');
  var iType   = hdr.indexOf('SPK_Type');
  var iTgl    = hdr.indexOf('Tgl_Buat');
  var iItem   = hdr.indexOf('Item_Code');
  var iSpec   = hdr.indexOf('Input_Spec');
  var iKgT    = hdr.indexOf('KG_Target');
  var iStatus = hdr.indexOf('Status');
  var iMC     = hdr.indexOf('MC_No');
  var iCust   = hdr.indexOf('Cust');
  var iOwner  = hdr.indexOf('Owner');
  var iParent = hdr.indexOf('Parent_SPK');
  var iPrio   = hdr.indexOf('Priority');

  var result = [];
  var tz = Session.getScriptTimeZone();

  for (var i = 1; i < data.length; i++) {
    var type = String(data[i][iType] || '').trim();
    if (type !== 'CTL-HEADER') continue;

    var tgl    = data[i][iTgl];
    var tglStr = _formatTglID(tgl, false);

    result.push({
      spk_no    : String(data[i][iSpk]    || ''),
      tgl_buat  : tglStr,
      item_code : String(data[i][iItem]   || ''),
      input_spec: String(data[i][iSpec]   || ''),
      kg_target : parseFloat(data[i][iKgT] || 0),
      status    : String(data[i][iStatus] || '').toUpperCase(),
      mc_no     : String(data[i][iMC]     || ''),
      cust      : String(data[i][iCust]   || ''),
      owner     : String(data[i][iOwner]  || ''),
      batch_id  : String(data[i][iParent] || ''),
      priority  : String(data[i][iPrio]   || '')
    });
  }

  result.reverse(); // Newest first
  return result;
}

/* =========================================================================
 * GET SPK LIST SHR
 * Untuk view-list halaman SPK SHR.
 * Return: array SHR-HEADER (standalone + child) sorted newest first.
 * is_child = true jika Parent_SPK berisi CTL/CL prefix
 * ========================================================================= */
function getSpkListSHR() {
  var sheet = getSheet('SPK');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h){ return String(h).trim(); });

  var iSpk    = hdr.indexOf('SPK_No');
  var iType   = hdr.indexOf('SPK_Type');
  var iTgl    = hdr.indexOf('Tgl_Buat');
  var iItem   = hdr.indexOf('Item_Code');
  var iSpec   = hdr.indexOf('Input_Spec');
  var iQtyT   = hdr.indexOf('Qty_Target');
  var iKgT    = hdr.indexOf('KG_Target');
  var iStatus = hdr.indexOf('Status');
  var iMC     = hdr.indexOf('MC_No');
  var iCust   = hdr.indexOf('Cust');
  var iOwner  = hdr.indexOf('Owner');
  var iParent = hdr.indexOf('Parent_SPK');
  var iPrio   = hdr.indexOf('Priority');
  var iSrcLoc = hdr.indexOf('Source_Loc');

  var result = [];
  var tz = Session.getScriptTimeZone();

  for (var i = 1; i < data.length; i++) {
    var type = String(data[i][iType] || '').trim();
    if (type !== 'SHR-HEADER') continue;

    var tgl    = data[i][iTgl];
    var tglStr = _formatTglID(tgl, false);

    var parentSpk = String(data[i][iParent] || '').trim();
    // is_child = true jika parent adalah SPK CTL (prefix CTL- atau lama CL-)
    var isChild = parentSpk !== '' &&
      (parentSpk.indexOf('CTL-') === 0 ||
       parentSpk.indexOf('CL-')  === 0 ||
       parentSpk.indexOf('CTL-') > -1);

    result.push({
      spk_no    : String(data[i][iSpk]    || ''),
      tgl_buat  : tglStr,
      item_code : String(data[i][iItem]   || ''),
      input_spec: String(data[i][iSpec]   || ''),
      qty_target: parseFloat(data[i][iQtyT] || 0),
      kg_target : parseFloat(data[i][iKgT]  || 0),
      status    : String(data[i][iStatus] || '').toUpperCase(),
      mc_no     : String(data[i][iMC]     || ''),
      cust      : String(data[i][iCust]   || ''),
      owner     : String(data[i][iOwner]  || ''),
      parent_spk: parentSpk,
      is_child  : isChild,
      source_loc: String(data[i][iSrcLoc] || ''),
      priority  : String(data[i][iPrio]   || '')
    });
  }

  result.reverse(); // Newest first
  return result;
}

/* =========================================================================
 * GET FULL SPK DATA — Untuk Edit Mode (exclude rows CANCELLED)
 *   - Include status & tgl_buat di response (BUGFIX)
 *   - Format tgl_buat pakai _formatTglID (DD MMM YYYY HH:mm)
 * ========================================================================= */
function getSpkFullData(spkNo) {
  const sheet = getSheet("SPK");
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h){ return String(h).trim(); });

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach(function(h, idx){
      let val = data[i][idx];
      if (val instanceof Date) val = val.toISOString();
      obj[h] = val;
    });
    rows.push(obj);
  }

  const header = rows.find(function(r){
    return String(r.SPK_No).trim() === String(spkNo).trim() && r.SPK_Type === 'CTL-HEADER';
  });
  if (!header) throw new Error("SPK Header tidak ditemukan: " + spkNo);

  function notCancelled(r){ return String(r.Status || '').toUpperCase() !== 'CANCELLED'; }

  const ctlOuts = rows.filter(function(r){
    return String(r.Parent_SPK).trim() === String(spkNo).trim()
        && r.SPK_Type === 'CTL-OUT' && notCancelled(r);
  });
  ctlOuts.sort(function(a, b){
    return String(a.SPK_No).localeCompare(String(b.SPK_No));
  });

  const out_ctls = ctlOuts.map(function(out){
    const shrHeader = rows.find(function(r){
      return String(r.Parent_SPK).trim() === String(out.SPK_No).trim()
          && r.SPK_Type === 'SHR-HEADER' && notCancelled(r);
    });

    let req_shr = [];
    let shr_header_no = null;
    if (shrHeader) {
      shr_header_no = shrHeader.SPK_No;
      const shrOuts = rows.filter(function(r){
        return String(r.Parent_SPK).trim() === String(shrHeader.SPK_No).trim()
            && r.SPK_Type === 'SHR-OUT' && notCancelled(r);
      });
      shrOuts.sort(function(a, b){
        const aN = parseInt(String(a.SPK_No).split('-').pop(), 10) || 0;
        const bN = parseInt(String(b.SPK_No).split('-').pop(), 10) || 0;
        return aN - bN;
      });
      req_shr = shrOuts.map(function(shr){
        let cutP = 1, cutL = 1, procType = 'Single';
        if (shr.NOTE) {
          const mP = String(shr.NOTE).match(/Cut_P:\s*(\d+)/i);
          const mL = String(shr.NOTE).match(/Cut_L:\s*(\d+)/i);
          if (mP) cutP = parseInt(mP[1], 10);
          if (mL) cutL = parseInt(mL[1], 10);
          if (String(shr.NOTE).toUpperCase().indexOf('MULTY') !== -1) procType = 'Multy';
        }
        return {
          spk_no     : shr.SPK_No,
          item_code  : shr.Item_Code,
          description: shr.Input_Spec,
          target_loc : shr.Target_Loc,
          so_no      : shr.SO_Ref,
          cust       : shr.Cust,
          owner      : shr.Owner_Used || shr.Owner || 'FC',
          cut_p      : cutP, cut_l: cutL,
          qty_plan   : shr.Qty_Target,
          qty_plan_kg: shr.KG_Target,
          proc_type  : procType
        };
      });
    }

    // 🟢 Step 3.B — Detect Stage 2 chain (SHR-HEADER yang Parent = SHR-OUT Stage 1 first)
    let req_shr_stage2 = null;
    if (shrHeader && req_shr.length > 0) {
      const stage1Out01SpkNo = shrHeader.SPK_No + '-01';
      const stage2Header = rows.find(function(r){
        return String(r.Parent_SPK).trim() === stage1Out01SpkNo
            && r.SPK_Type === 'SHR-HEADER' && notCancelled(r);
      });
      if (stage2Header) {
        const stage2Out = rows.find(function(r){
          return String(r.Parent_SPK).trim() === String(stage2Header.SPK_No).trim()
              && r.SPK_Type === 'SHR-OUT' && notCancelled(r);
        });
        if (stage2Out) {
          let cutP2 = 1, cutL2 = 1;
          if (stage2Out.NOTE) {
            const mP = String(stage2Out.NOTE).match(/Cut_P:\s*(\d+)/i);
            const mL = String(stage2Out.NOTE).match(/Cut_L:\s*(\d+)/i);
            if (mP) cutP2 = parseInt(mP[1], 10);
            if (mL) cutL2 = parseInt(mL[1], 10);
          }
          req_shr_stage2 = {
            spk_header_no: stage2Header.SPK_No,
            spk_no       : stage2Out.SPK_No,
            item_code    : stage2Out.Item_Code,
            description  : stage2Out.Input_Spec,
            target_loc   : stage2Out.Target_Loc,
            so_no        : stage2Out.SO_Ref,
            cust         : stage2Out.Cust,
            owner        : stage2Out.Owner_Used || stage2Out.Owner || 'FC',
            machine      : stage2Out.MC_No || 'SHR-03',
            cut_p        : cutP2, cut_l: cutL2,
            qty_plan     : stage2Out.Qty_Target,
            qty_plan_kg  : stage2Out.KG_Target,
            proc_type    : 'Stage2'
          };
        }
      }
    }

    return {
      spk_no     : out.SPK_No,
      item_code  : out.Item_Code,
      description: out.Input_Spec,
      target_loc : out.Target_Loc,
      qty_plan   : out.Qty_Target,
      qty_plan_kg: out.KG_Target,
      qty_actual : parseFloat(out.Qty_Actual) || 0,
      kg_actual  : parseFloat(out.KG_Actual)  || 0,
      so_ref     : out.SO_Ref,
      cust       : out.Cust,
      owner      : out.Owner || 'FC',
      owner_used : out.Owner_Used || 'FC',
      shr_header_no: shr_header_no,
      req_shr    : req_shr,
      req_shr_stage2: req_shr_stage2   // 🟢 Step 3.B
    };
  });
  // ====== CTL-NG rows (NG Cutting Grade 2) ======
  const ctlNgs = rows.filter(function(r){
    return String(r.Parent_SPK).trim() === String(spkNo).trim()
        && r.SPK_Type === 'CTL-NG' && notCancelled(r);
  });
  ctlNgs.sort(function(a, b){
    return String(a.SPK_No).localeCompare(String(b.SPK_No));
  });
  const out_ngs = ctlNgs.map(function(ng){
    return {
      spk_no     : ng.SPK_No,
      t          : ng.T          || 0,
      p          : ng.P          || 0,
      l          : ng.L          || 0,
      input_spec : ng.Input_Spec || '',
      qty        : ng.Qty_Target || 0,
      kg         : ng.KG_Target  || 0,
      owner      : ng.Owner      || 'FC',
      owner_used : ng.Owner_Used || 'FC',
      mc_no      : ng.MC_No      || '',
      status     : ng.Status     || '',
      target_loc : ng.Target_Loc || 'Stok_NG',
      selesai_dt : ng.Selesai_DT ? _formatTglID(ng.Selesai_DT, true) : '—'
    };
  });
  // ✅ Lookup info coil dari Stok_Coil untuk print template (thick/width/no_coil selalu dari Stok_Coil)
  //    Untuk kg_avail: pakai SNAPSHOT dari kolom Coil_Avail_Snapshot di SPK header (data historis saat SPK terbit)
  //    Fallback ke Stok_Coil.KG_Avail hanya kalau snapshot kosong (SPK lama sebelum patch)
  var coilInfo = { thick: 0, width: 0, no_coil: '', kg_avail: 0 };
  var snapshotKgAvail = parseFloat(header.Coil_Avail_Snapshot) || 0;
  try {
    var coilSheet = getSheet(SHEET_NAMES.STOK_COIL);
    if (coilSheet) {
      var coilData = coilSheet.getDataRange().getValues();
      var cH = coilData[0].map(function(h){return String(h).trim();});
      var iB = cH.indexOf('Batch_ID');
      var iT = cH.indexOf('T');
      var iP = cH.indexOf('P');
      var iNC = cH.indexOf('No_Coil');
      var iKA = cH.indexOf('KG_Avail');
      for (var ci = 1; ci < coilData.length; ci++) {
        if (String(coilData[ci][iB]).trim() === String(header.Parent_SPK).trim()) {
          coilInfo.thick    = iT  >= 0 ? (coilData[ci][iT] ||0) : 0;
          coilInfo.width    = iP  >= 0 ? (coilData[ci][iP] ||0) : 0;
          coilInfo.no_coil  = iNC >= 0 ? (coilData[ci][iNC]||'') : '';
          coilInfo.kg_avail = iKA >= 0 ? (parseFloat(coilData[ci][iKA])||0) : 0;
          break;
        }
      }
    }
  } catch(e) { /* fallback to empty */ }
  // Override kg_avail dengan snapshot kalau ada (data saat SPK terbit)
  if (snapshotKgAvail > 0) {
    coilInfo.kg_avail = snapshotKgAvail;
  }

  return {
    header: {
      spk_no     : header.SPK_No,
      status     : String(header.Status || '').toUpperCase(),
      tgl_buat   : _formatTglID(header.Tgl_Buat, true),
      batch_id   : header.Parent_SPK,
      item_code  : header.Item_Code,
      input_spec : header.Input_Spec,
      qty_target : header.Qty_Target,
      kg_target  : header.KG_Target,
      qty_actual : parseFloat(header.Qty_Actual) || 0,
      kg_actual  : parseFloat(header.KG_Actual)  || 0,
      mc_no      : header.MC_No,
      source_loc : header.Source_Loc,
      priority   : header.Priority,
      op         : header.OP,
      owner      : header.Owner || 'FC',
      is_habis   : header.Is_Habis === true,
      thick      : coilInfo.thick,
      width      : coilInfo.width,
      no_coil    : coilInfo.no_coil,
      coil_avail_kg: coilInfo.kg_avail,
      note       : header.NOTE,
      cust       : header.Cust || ''
    },
    out_ctls: out_ctls,
    out_ngs : out_ngs
  };
}

/* =========================================================================
 * UPDATE + APPEND — Inti edit mode
 *   - existing_spk_no diisi → UPDATE row in-place
 *   - existing_spk_no kosong → APPEND huruf berikutnya
 *   - SHR: UPDATE existing per index, APPEND ekstra, CANCEL yang berlebih
 * ========================================================================= */
/* =========================================================================
 * UPDATE + APPEND CTL — Edit mode handler
 * ========================================================================= */
function updateAndAppendSpkCTL(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // 🚩 Sesi 2: Promote DRAFT → ANTRIAN kalau Save final (bukan Save Draft)
    // Cara: cari existing header dgn SPK_No, kalau statusnya DRAFT dan payload
    // TIDAK is_draft, promote status semua rows ke ANTRIAN.
    var promoteFromDraft = false;
    if (data.spk_no && data.is_draft !== true) {
      var _s = getSheet('SPK');
      var _v = _s.getDataRange().getValues();
      var _h = _v[0].map(function(x){ return String(x).trim(); });
      var _iSpk = _h.indexOf('SPK_No'), _iSt = _h.indexOf('Status');
      if (_iSpk >= 0 && _iSt >= 0) {
        for (var _r = 1; _r < _v.length; _r++) {
          if (String(_v[_r][_iSpk]||'').trim() === String(data.spk_no).trim()) {
            if (String(_v[_r][_iSt]||'').trim().toUpperCase() === 'DRAFT') promoteFromDraft = true;
            break;
          }
        }
      }
    }
    // ✅ Validasi stok coil saat edit (Bug 6) — exclude SPK ini sendiri
    if (typeof validateCoilAvailability === 'function' && data.batch_id && data.weight_kg) {
      validateCoilAvailability(data.batch_id, data.weight_kg, data.spk_no);
    }
    const spkSheet = getSheet("SPK");
    const allData  = spkSheet.getDataRange().getValues();
    const headers  = allData[0].map(function(h){ return String(h).trim(); });
    const timestamp = new Date();
    const spkNo    = data.spk_no;
    const generatedSpks = [], updatedSpks = [], cancelledSpks = [];

    function buildRow(rd){ return headers.map(function(h){ return rd[h] !== undefined ? rd[h] : ''; }); }
    function col(name){ var i = headers.indexOf(name); return i >= 0 ? i + 1 : -1; }

    const iSpk    = headers.indexOf("SPK_No");
    const iType   = headers.indexOf("SPK_Type");
    const iParent = headers.indexOf("Parent_SPK");
    const iStatus = headers.indexOf("Status");
    const statusCol = iStatus + 1;

    // Bangun rowMap: SPK_No → { rowNum }
    const rowMap = {};
    for (var i = 1; i < allData.length; i++) {
      var sn = String(allData[i][iSpk] || '').trim();
      if (sn) rowMap[sn] = { rowNum: i + 1 };
    }
    function updateRow(rowNum, rd) {
      Object.keys(rd).forEach(function(k){
        var c = col(k);
        if (c > 0) spkSheet.getRange(rowNum, c).setValue(rd[k]);
      });
    }

    // ── LANGKAH 1: Pre-kalkulasi CT semua OUT ──

    // CT tiap CTL-OUT
    var outCtlCtList = data.out_ctl.map(function(out) {
      return hitungRencanaDurasi(out.item_code, out.qty_plan, out.qty_plan_kg);
    });

    // CT tiap SHR-OUT per CTL-OUT
    var shrCtMatrix = data.out_ctl.map(function(out) {
      if (!out.req_shr || out.req_shr.length === 0) return [];
      return out.req_shr.map(function(shr) {
        return hitungRencanaDurasi(shr.item_code, shr.qty_plan, shr.qty_plan_kg);
      });
    });

    // ── LANGKAH 2: Hitung agregat & UPDATE CTL-HEADER ──
    var coilCt         = hitungRencanaDurasi(data.coil_item_code, data.qty_input, data.weight_kg);
    var hdrPlanSetup   = coilCt.planSetup;
    var hdrPlanRun     = 0;
    var hdrTotalDurasi = hdrPlanSetup;
    outCtlCtList.forEach(function(ct) {
      hdrPlanRun     += ct.planRun;
      hdrTotalDurasi += ct.planSetup + ct.planRun;
    });

    if (rowMap[spkNo]) {
      updateRow(rowMap[spkNo].rowNum, {
        'Priority'           : data.priority,
        'Source_Loc'         : data.source_loc,
        'KG_Target'          : data.weight_kg,
        'OP'                 : data.op,
        'Plan_Setup_Menit'   : hdrPlanSetup,
        'Plan_Run_Menit'     : hdrPlanRun,
        'Total_Durasi_Menit' : hdrTotalDurasi,
        'Is_Habis'           : data.is_habis === true,
        'T'                  : parseFloat(data.thick) || 0
      });
      updatedSpks.push(spkNo);
    }

    // Deteksi format: CTL- = baru (angka), CL- = lama (huruf)
    var useNewFmt  = isNewSpkFormat(spkNo);
    var maxCtlSufx = 0;
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][iType] === 'CTL-OUT' && String(allData[i][iParent]).trim() === spkNo.trim()) {
        var cs = String(allData[i][iSpk]).trim();
        if (useNewFmt) {
          var parts = cs.split('-');
          var n = parseInt(parts[parts.length - 1], 10) || 0;
          if (n > maxCtlSufx) maxCtlSufx = n;
        } else {
          if (cs.length > spkNo.length) {
            var code = cs.charCodeAt(spkNo.length);
            if (code > (maxCtlSufx + 64)) maxCtlSufx = code - 64;
          }
        }
      }
    }

    function findMaxShrOutSuffix(shrHdr) {
      var maxN = 0;
      for (var i = 1; i < allData.length; i++) {
        if (allData[i][iType] === 'SHR-OUT' && String(allData[i][iParent]).trim() === String(shrHdr).trim()) {
          var sf = parseInt(String(allData[i][iSpk]).trim().split('-').pop(), 10) || 0;
          if (sf > maxN) maxN = sf;
        }
      }
      return maxN;
    }
    function getActiveShrOuts(shrHdr) {
      var list = [];
      for (var i = 1; i < allData.length; i++) {
        if (allData[i][iType] === 'SHR-OUT'
            && String(allData[i][iParent]).trim() === String(shrHdr).trim()
            && String(allData[i][iStatus] || '').toUpperCase() !== 'CANCELLED') {
          list.push({ spkNo: String(allData[i][iSpk]).trim(), rowNum: i + 1 });
        }
      }
      list.sort(function(a, b){
        var aN = parseInt(a.spkNo.split('-').pop(), 10) || 0;
        var bN = parseInt(b.spkNo.split('-').pop(), 10) || 0;
        return aN - bN;
      });
      return list;
    }

    var yyNow2     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yy');
    var iSpkColUpd = headers.indexOf('SPK_No');
    var reShrUpd   = new RegExp('^SHR-' + yyNow2 + '(\\d{4})$');
    var maxShrUpd  = 0;
    for (var ru = 1; ru < allData.length; ru++) {
      var mShrU = String(allData[ru][iSpkColUpd] || '').trim().match(reShrUpd);
      if (mShrU) { var su = parseInt(mShrU[1], 10); if (su > maxShrUpd) maxShrUpd = su; }
    }
    var currentShrSeq = maxShrUpd + 1;

    // ── LANGKAH 3: Loop tiap CTL-OUT ──
    data.out_ctl.forEach(function(out, idx) {
      var isExisting = !!(out.existing_spk_no);
      var ctlOutSpkNo;
      var ctOut     = outCtlCtList[idx];
      var shrCtList = shrCtMatrix[idx];

      var ctlOutPayload = {
        'SO_Ref'             : out.so_ref      || '',
        'Cust'               : out.cust         || '',
        'Priority'           : data.priority,
        'Item_Code'          : out.item_code   || '',
        'Input_Spec'         : out.l            || '',
        'Target_Loc'         : out.target_loc  || '',
        'Qty_Target'         : out.qty_plan,
        'KG_Target'          : out.qty_plan_kg,
        'OP'                 : data.op,
        'Owner'              : out.owner      || 'FC',
        'Owner_Used'         : out.owner_used || 'FC',
        'Plan_Setup_Menit'   : ctOut.planSetup,
        'Plan_Run_Menit'     : ctOut.planRun,
        'Total_Durasi_Menit' : 0   // CTL-OUT tidak masuk engine jadwal
      };

      if (isExisting) {
        ctlOutSpkNo = out.existing_spk_no;
        var meta = rowMap[ctlOutSpkNo];
        if (!meta) throw new Error("CTL-OUT existing tidak ditemukan: " + ctlOutSpkNo);
        updateRow(meta.rowNum, ctlOutPayload);
        updatedSpks.push(ctlOutSpkNo);
      } else {
        maxCtlSufx++;
        ctlOutSpkNo = useNewFmt
          ? spkNo + '-' + String(maxCtlSufx).padStart(2, '0')
          : spkNo + String.fromCharCode(64 + maxCtlSufx);
        _appendRowSafe(spkSheet,buildRow(Object.assign({
          'SPK_No'     : ctlOutSpkNo,
          'SPK_Type'   : 'CTL-OUT',
          'Parent_SPK' : spkNo,
          'Tgl_Buat'   : timestamp,
          'Source_Loc' : data.machine,
          'MC_No'      : data.machine,
          'BQ'         : 1,
          'Status'     : 'ANTRIAN',
          'Created_By' : data.created_by
        }, ctlOutPayload)));
        generatedSpks.push(ctlOutSpkNo);
        rowMap[ctlOutSpkNo] = { rowNum: spkSheet.getLastRow() };
      }

      // ── SHR dari CTL ──
      var hasShr       = out.req_shr && out.req_shr.length > 0;
      var existShrHdr  = out.existing_shr_header_no || null;

      // Hitung agregat SHR-HEADER
      var shrHdrPlanRun     = 0;
      var shrHdrTotalDurasi = 0;
      if (hasShr) {
        shrCtList.forEach(function(ct) {
          shrHdrPlanRun     += ct.planRun;
          shrHdrTotalDurasi += ct.planSetup + ct.planRun;
        });
      }

      if (hasShr) {
        var shrSpkNo;

        if (existShrHdr && rowMap[existShrHdr]) {
          // UPDATE SHR-HEADER
          shrSpkNo = existShrHdr;
          updateRow(rowMap[existShrHdr].rowNum, {
            'Priority'           : data.priority,
            'Source_Loc'         : out.target_loc  || '',
            'Item_Code'          : out.item_code   || '',
            'Input_Spec'         : out.l            || '',
            'Qty_Target'         : out.qty_plan,
            'KG_Target'          : out.qty_plan_kg,
            'Owner'              : data.coil_owner || 'FC',
            'Owner_Used'         : data.coil_owner || 'FC',
            'Plan_Setup_Menit'   : 0,
            'Plan_Run_Menit'     : shrHdrPlanRun,
            'Total_Durasi_Menit' : shrHdrTotalDurasi
          });
          updatedSpks.push(shrSpkNo);
        } else {
          // APPEND SHR-HEADER baru
          shrSpkNo = 'SHR-' + yyNow2 + String(currentShrSeq).padStart(4, '0');
          currentShrSeq++;
          _appendRowSafe(spkSheet,buildRow({
            'SPK_No'             : shrSpkNo,
            'SPK_Type'           : 'SHR-HEADER',
            'Parent_SPK'         : ctlOutSpkNo,
            'Tgl_Buat'           : timestamp,
            'Priority'           : data.priority,
            'Source_Loc'         : out.target_loc  || '',
            'Item_Code'          : out.item_code   || '',
            'Input_Spec'         : out.l            || '',
            'Qty_Target'         : out.qty_plan,
            'KG_Target'          : out.qty_plan_kg,
            'MC_No'              : 'SHR-01',
            'Status'             : 'ANTRIAN',
            'Created_By'         : data.created_by,
            'Owner'              : data.coil_owner || 'FC',
            'Owner_Used'         : data.coil_owner || 'FC',
            'Plan_Setup_Menit'   : 0,
            'Plan_Run_Menit'     : shrHdrPlanRun,
            'Total_Durasi_Menit' : shrHdrTotalDurasi,
            'T'                  : parseFloat(data.thick) || 0
          }));
          generatedSpks.push(shrSpkNo);
        }

        // SHR-OUT diff: update existing, append baru, cancel kelebihan
        var existShrOuts = existShrHdr ? getActiveShrOuts(existShrHdr) : [];
        var nextSuffix   = findMaxShrOutSuffix(shrSpkNo);

        out.req_shr.forEach(function(shr, sIdx) {
          var ctSO = shrCtList[sIdx];

          var shrOutPayload = {
            'SO_Ref'             : shr.so_no      || '',
            'Cust'               : shr.cust        || '',
            'Priority'           : data.priority,
            'Item_Code'          : shr.item_code  || '',
            'Input_Spec'         : shr.l           || '',
            'Target_Loc'         : shr.target_loc || '',
            'BQ'                 : shr.cut_p * shr.cut_l,
            'Qty_Target'         : shr.qty_plan,
            'KG_Target'          : shr.qty_plan_kg,
            'NOTE'               : 'Cut_P:' + shr.cut_p + ' Cut_L:' + shr.cut_l + ' ' + (shr.proc_type || ''),
            'Owner'              : data.coil_owner || 'FC',
            'Owner_Used'         : shr.owner       || 'FC',
            'Plan_Setup_Menit'   : ctSO.planSetup,
            'Plan_Run_Menit'     : ctSO.planRun,
            'Total_Durasi_Menit' : 0,  // SHR-OUT tidak masuk engine jadwal
            'T'                  : parseFloat(data.thick) || 0
          };

          if (sIdx < existShrOuts.length) {
            updateRow(existShrOuts[sIdx].rowNum, shrOutPayload);
            updatedSpks.push(existShrOuts[sIdx].spkNo);
          } else {
            nextSuffix++;
            var newNo = shrSpkNo + '-' + String(nextSuffix).padStart(2, '0');
            _appendRowSafe(spkSheet,buildRow(Object.assign({
              'SPK_No'     : newNo,
              'SPK_Type'   : 'SHR-OUT',
              'Parent_SPK' : shrSpkNo,
              'Tgl_Buat'   : timestamp,
              'Source_Loc' : 'SHR-01',
              'MC_No'      : 'SHR-01',
              'Status'     : 'ANTRIAN',
              'Created_By' : data.created_by
            }, shrOutPayload)));
            generatedSpks.push(newNo);
          }
        });

        // Cancel SHR-OUT yang kelebihan
        if (out.req_shr.length < existShrOuts.length) {
          for (var k = out.req_shr.length; k < existShrOuts.length; k++) {
            spkSheet.getRange(existShrOuts[k].rowNum, statusCol).setValue("CANCELLED");
            cancelledSpks.push(existShrOuts[k].spkNo);
          }
        }

        // 🟢 Step 3.B — STAGE 2 chain handling (create / update / cancel)
        var stage1Out01SpkNo = shrSpkNo + '-01';

        // Cari existing Stage 2 SHR-HEADER (parent = stage1Out01SpkNo)
        var existStage2Hdr = null;
        for (var i2 = 1; i2 < allData.length; i2++) {
          if (allData[i2][iType] === 'SHR-HEADER'
              && String(allData[i2][iParent]).trim() === stage1Out01SpkNo
              && String(allData[i2][iStatus] || '').toUpperCase() !== 'CANCELLED') {
            existStage2Hdr = { spkNo: String(allData[i2][iSpk]).trim(), rowNum: i2 + 1 };
            break;
          }
        }

        if (out.req_shr_stage2) {
          var s2 = out.req_shr_stage2;
          var stage1Out01 = out.req_shr[0];

          if (existStage2Hdr) {
            // UPDATE Stage 2 HEADER
            updateRow(existStage2Hdr.rowNum, {
              'Priority'   : data.priority,
              'Source_Loc' : stage1Out01.target_loc || '',
              'Item_Code'  : stage1Out01.item_code  || '',
              'Input_Spec' : stage1Out01.l          || '',
              'Qty_Target' : stage1Out01.qty_plan,
              'KG_Target'  : stage1Out01.qty_plan_kg,
              'MC_No'      : s2.machine || 'SHR-03',
              'Owner_Used' : s2.owner || data.coil_owner || 'FC'
            });
            updatedSpks.push(existStage2Hdr.spkNo);

            // UPDATE Stage 2 OUT
            var existStage2OutNo = existStage2Hdr.spkNo + '-01';
            for (var j2 = 1; j2 < allData.length; j2++) {
              if (String(allData[j2][iSpk]).trim() === existStage2OutNo) {
                updateRow(j2 + 1, {
                  'SO_Ref'     : s2.so_no || stage1Out01.so_no || '',
                  'Cust'       : s2.cust  || stage1Out01.cust  || '',
                  'Priority'   : data.priority,
                  'Item_Code'  : s2.item_code  || '',
                  'Input_Spec' : s2.l           || '',
                  'Target_Loc' : s2.target_loc || '',
                  'MC_No'      : s2.machine || 'SHR-03',
                  'Source_Loc' : s2.machine || 'SHR-03',
                  'BQ'         : (s2.cut_p || 1) * (s2.cut_l || 1),
                  'Qty_Target' : s2.qty_plan,
                  'KG_Target'  : s2.qty_plan_kg,
                  'NOTE'       : 'Cut_P:' + s2.cut_p + ' Cut_L:' + s2.cut_l + ' Stage2',
                  'Owner_Used' : s2.owner || 'FC'
                });
                updatedSpks.push(existStage2OutNo);
                break;
              }
            }
          } else {
            // APPEND Stage 2 chain BARU
            var shrSpkNo2 = 'SHR-' + yyNow2 + String(currentShrSeq).padStart(4, '0');
            currentShrSeq++;

            _appendRowSafe(spkSheet, buildRow({
              'SPK_No'             : shrSpkNo2,
              'SPK_Type'           : 'SHR-HEADER',
              'Parent_SPK'         : stage1Out01SpkNo,
              'Tgl_Buat'           : timestamp,
              'Priority'           : data.priority,
              'Source_Loc'         : stage1Out01.target_loc || '',
              'Item_Code'          : stage1Out01.item_code  || '',
              'Input_Spec'         : stage1Out01.l          || '',
              'Qty_Target'         : stage1Out01.qty_plan,
              'KG_Target'          : stage1Out01.qty_plan_kg,
              'MC_No'              : s2.machine || 'SHR-03',
              'Status'             : 'ANTRIAN',
              'Created_By'         : data.created_by,
              'Owner'              : data.coil_owner || 'FC',
              'Owner_Used'         : s2.owner || data.coil_owner || 'FC',
              'Plan_Setup_Menit'   : 0,
              'Plan_Run_Menit'     : 0,
              'Total_Durasi_Menit' : 0,
              'T'                  : parseFloat(data.thick) || 0,
              'NOTE'               : 'STAGE2-PARENT'
            }));
            generatedSpks.push(shrSpkNo2);

            var shr2OutNo = shrSpkNo2 + '-01';
            _appendRowSafe(spkSheet, buildRow({
              'SPK_No'             : shr2OutNo,
              'SPK_Type'           : 'SHR-OUT',
              'SO_Ref'             : s2.so_no || stage1Out01.so_no || '',
              'Cust'               : s2.cust  || stage1Out01.cust  || '',
              'Parent_SPK'         : shrSpkNo2,
              'Tgl_Buat'           : timestamp,
              'Priority'           : data.priority,
              'Source_Loc'         : s2.machine || 'SHR-03',
              'Item_Code'          : s2.item_code  || '',
              'Input_Spec'         : s2.l           || '',
              'Target_Loc'         : s2.target_loc || '',
              'MC_No'              : s2.machine || 'SHR-03',
              'BQ'                 : (s2.cut_p || 1) * (s2.cut_l || 1),
              'Qty_Target'         : s2.qty_plan,
              'KG_Target'          : s2.qty_plan_kg,
              'Status'             : 'ANTRIAN',
              'Created_By'         : data.created_by,
              'NOTE'               : 'Cut_P:' + s2.cut_p + ' Cut_L:' + s2.cut_l + ' Stage2',
              'Owner'              : data.coil_owner || 'FC',
              'Owner_Used'         : s2.owner || 'FC',
              'Plan_Setup_Menit'   : 0,
              'Plan_Run_Menit'     : 0,
              'Total_Durasi_Menit' : 0,
              'T'                  : parseFloat(data.thick) || 0
            }));
            generatedSpks.push(shr2OutNo);
          }
        } else if (existStage2Hdr) {
          // Payload TIDAK ada Stage 2 tapi existing ADA → CANCEL chain
          spkSheet.getRange(existStage2Hdr.rowNum, statusCol).setValue("CANCELLED");
          cancelledSpks.push(existStage2Hdr.spkNo);
          var s2OutNoExist = existStage2Hdr.spkNo + '-01';
          for (var k2 = 1; k2 < allData.length; k2++) {
            if (String(allData[k2][iSpk]).trim() === s2OutNoExist) {
              spkSheet.getRange(k2 + 1, statusCol).setValue("CANCELLED");
              cancelledSpks.push(s2OutNoExist);
              break;
            }
          }
        }

      } else if (existShrHdr && rowMap[existShrHdr]) {
        // Form tidak ada SHR tapi existing SHR-HEADER ada → cancel semua
        spkSheet.getRange(rowMap[existShrHdr].rowNum, statusCol).setValue("CANCELLED");
        cancelledSpks.push(existShrHdr);
        getActiveShrOuts(existShrHdr).forEach(function(so) {
          spkSheet.getRange(so.rowNum, statusCol).setValue("CANCELLED");
          cancelledSpks.push(so.spkNo);
        });
      }
    });

    SpreadsheetApp.flush();

    // 🚩 Sesi 2: Promote DRAFT rows → ANTRIAN kalau editing DRAFT dgn Save final
    if (promoteFromDraft) {
      var _s2 = getSheet('SPK');
      var _v2 = _s2.getDataRange().getValues();
      var _h2 = _v2[0].map(function(x){ return String(x).trim(); });
      var _iSpk2 = _h2.indexOf('SPK_No'), _iSt2 = _h2.indexOf('Status');
      var _iParent = _h2.indexOf('Parent_SPK');
      if (_iSpk2 >= 0 && _iSt2 >= 0) {
        for (var _r2 = 1; _r2 < _v2.length; _r2++) {
          var _sn = String(_v2[_r2][_iSpk2]||'').trim();
          var _pr = String(_v2[_r2][_iParent]||'').trim();
          var _st = String(_v2[_r2][_iSt2]||'').trim().toUpperCase();
          // Promote kalau SPK_No cocok atau parent match spkNo, dan status masih DRAFT
          if (_st === 'DRAFT' && (_sn === spkNo || _pr === spkNo || _sn.indexOf(spkNo) === 0)) {
            _s2.getRange(_r2+1, _iSt2+1).setValue('ANTRIAN');
          }
        }
      }
      SpreadsheetApp.flush();
    }

    // 🚩 Sesi 2: Kalau Save Draft di edit mode, flip rows baru ke DRAFT
    if (data.is_draft === true && generatedSpks.length) {
      _updateStatusToDraft(spkSheet, generatedSpks);
      SpreadsheetApp.flush();
    }

    if (data.is_draft !== true && typeof kalkulasiEstimasiWaktu === 'function') kalkulasiEstimasiWaktu();

    return {
      spk_no          : spkNo,
      generated       : generatedSpks,
      updated         : updatedSpks,
      cancelled       : cancelledSpks,
      count_ctl       : data.out_ctl.length,
      count_new       : generatedSpks.length,
      count_updated   : updatedSpks.length,
      count_cancelled : cancelledSpks.length,
      is_draft        : data.is_draft === true,
      promoted        : promoteFromDraft,
      mode            : 'UPDATE_APPEND'
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * CASCADE CANCEL — Cancel SPK + semua anak/turunan
 * ========================================================================= */
function cascadeCancelSpk(spkNo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet   = getSheet("SPK");
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(function(h){ return String(h).trim(); });
    const iSpk    = headers.indexOf("SPK_No");
    const iParent = headers.indexOf("Parent_SPK");
    const iStatus = headers.indexOf("Status");
 
    if (iSpk === -1 || iStatus === -1) {
      return { success: false, message: "Kolom tidak valid." };
    }
 
    const statusCol = iStatus + 1;
    const cancelled = [];
    const skippedDone = []; // 🟢 BUG-FIX #10a — track DONE rows yang sengaja dilewati
    const queue     = [String(spkNo).trim()];
    const visited   = {};
 
    while (queue.length > 0) {
      const cur = queue.shift();
      if (visited[cur]) continue;
      visited[cur] = true;
 
      for (let i = 1; i < data.length; i++) {
        const rowSpk    = String(data[i][iSpk]    || '').trim();
        const rowParent = String(data[i][iParent] || '').trim();
 
        if (rowSpk === cur) {
          const st = String(data[i][iStatus] || '').toUpperCase();
          // 🟢 BUG-FIX #10a — skip DONE (data integrity: stok sudah tertulis)
          // dan skip CANCELLED (sudah final).
          if (st === 'DONE') {
            skippedDone.push(cur);
          } else if (st !== 'CANCELLED') {
            sheet.getRange(i + 1, statusCol).setValue("CANCELLED");
            cancelled.push(cur);
          }
        }
        // Tetap traverse children meskipun parent DONE/CANCELLED,
        // supaya child ANTRIAN/RUNNING bisa di-cancel.
        if (rowParent === cur && rowSpk && !visited[rowSpk]) {
          queue.push(rowSpk);
        }
      }
    }
 
    SpreadsheetApp.flush();
 
    // Recalculate antrian setelah cascade cancel agar slot kosong langsung terisi
    if (typeof kalkulasiEstimasiWaktu === 'function') {
      kalkulasiEstimasiWaktu();
    }
 
    return {
      success: true,
      cancelled: cancelled,
      skippedDone: skippedDone   // 🟢 frontend bisa kasih tau user kalau ada yang dilewati
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * CASCADE CANCEL WITH PASSWORD AUTH — Wrapper untuk Cancel SPK dari View Detail
 *   - Verifikasi password dulu sebelum eksekusi cascadeCancelSpk
 *   - Password disimpan di constant CANCEL_SPK_PASSWORD (ganti di sini kalau perlu)
 *   - Reuse fungsi cascadeCancelSpk yang sudah ada (tidak diubah)
 *   - Return: { success, cancelled[], message } — format konsisten
 * ========================================================================= */
function cascadeCancelSpkWithAuth(spkNo, password) {
  var CANCEL_SPK_PASSWORD = '321654';

  // 1. Verifikasi password
  if (!password || String(password).trim() !== CANCEL_SPK_PASSWORD) {
    return {
      success: false,
      message: 'Password salah. Aksi cancel dibatalkan.'
    };
  }

  // 2. Validasi spkNo
  if (!spkNo || String(spkNo).trim() === '') {
    return {
      success: false,
      message: 'SPK No kosong. Aksi cancel dibatalkan.'
    };
  }

  // 3. 🟢 BUG-FIX #10b — Hanya SPK status ANTRIAN yang boleh di-cancel
  // total via view detail. SPK RUNNING/DONE/CANCELLED ditolak.
  // Untuk RUNNING: PPIC harus pakai cancel per-OUT di Board Mesin.
  try {
    var sheet   = getSheet("SPK");
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });
    var iSpk    = headers.indexOf("SPK_No");
    var iStatus = headers.indexOf("Status");
    var iType   = headers.indexOf("SPK_Type");
    var target  = String(spkNo).trim();

    var foundStatus = '';
    var foundType   = '';
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iSpk] || '').trim() === target) {
        foundStatus = String(data[i][iStatus] || '').toUpperCase();
        foundType   = String(data[i][iType]   || '').trim();
        break;
      }
    }

    if (!foundStatus) {
      return { success: false, message: 'SPK tidak ditemukan: ' + target };
    }

    if (foundStatus !== 'ANTRIAN') {
      var msg = 'SPK status ' + foundStatus + ' tidak bisa di-cancel total via view detail.';
      if (foundStatus === 'RUNNING') {
        msg += '\n\nUntuk SPK yang sedang berjalan, gunakan cancel per-OUT di Board Mesin (icon X dengan password PPIC).';
      } else if (foundStatus === 'DONE') {
        msg += '\n\nSPK yang sudah selesai tidak dapat dibatalkan.';
      } else if (foundStatus === 'CANCELLED') {
        msg += '\n\nSPK sudah cancelled sebelumnya.';
      }
      return { success: false, message: msg };
    }
  } catch (e) {
    return { success: false, message: 'Error validasi status: ' + e.message };
  }

  // 4. Forward ke cascadeCancelSpk yang sudah ada
  return cascadeCancelSpk(spkNo);
}

/* =========================================================================
 * STEP 5: cancelCtlOutWithAuth (Bug-Fix #10c)
 * Cancel satu CTL-OUT (cascade ke SHR child kalau ada) dengan password gate.
 * Opsional: kalau ngItems disertakan (cancel OUT terakhir dengan NG popup),
 * generate CTL-NG row(s) untuk sisa material coil, lalu update header
 * KG_Actual, Qty_NG, KG_NG (akumulasi).
 *
 * @param {string} spkOutNo - SPK_No CTL-OUT yang mau di-cancel
 * @param {string} password - Harus '321654'
 * @param {Array}  ngItems  - (Opsional) array NG items: [{t,p,l,qty}, ...]
 * @returns {Object} { success, message?, cancelled?, ng_created?, header_kg_actual? }
 * ========================================================================= */
function cancelCtlOutWithAuth(spkOutNo, password, ngItems) {
  var CANCEL_PASSWORD = '321654';

  // 1. Verifikasi password
  if (!password || String(password).trim() !== CANCEL_PASSWORD) {
    return { success: false, message: 'Password salah. Aksi cancel dibatalkan.' };
  }

  // 2. Validasi spkOutNo
  if (!spkOutNo || String(spkOutNo).trim() === '') {
    return { success: false, message: 'SPK OUT No kosong.' };
  }

  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName("SPK");
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });

    var iSpk    = headers.indexOf("SPK_No");
    var iType   = headers.indexOf("SPK_Type");
    var iParent = headers.indexOf("Parent_SPK");
    var iStatus = headers.indexOf("Status");
    var iQtyAct = headers.indexOf("Qty_Actual");
    var iKgAct  = headers.indexOf("KG_Actual");
    var iQtyNg  = headers.indexOf("Qty_NG");
    var iKgNg   = headers.indexOf("KG_NG");
    var iItem   = headers.indexOf("Item_Code");
    var iMC     = headers.indexOf("MC_No");
    var iOwner  = headers.indexOf("Owner");
    var iKgTgt  = headers.indexOf("KG_Target");

    var target = String(spkOutNo).trim();

    // 3. Cari row CTL-OUT
    var outRowIdx  = -1;
    var outStatus  = '';
    var outType    = '';
    var parentSpk  = '';
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iSpk] || '').trim() === target) {
        outRowIdx = i;
        outStatus = String(data[i][iStatus] || '').toUpperCase();
        outType   = String(data[i][iType]   || '').trim();
        parentSpk = String(data[i][iParent] || '').trim();
        break;
      }
    }

    if (outRowIdx === -1) {
      return { success: false, message: 'SPK OUT tidak ditemukan: ' + target };
    }
    if (outType !== 'CTL-OUT') {
      return { success: false, message: 'Fitur ini hanya untuk CTL-OUT. Type: ' + outType };
    }
    if (outStatus === 'DONE') {
      return { success: false, message: 'OUT ' + target + ' sudah DONE - tidak bisa di-cancel.' };
    }
    if (outStatus === 'CANCELLED') {
      return { success: false, message: 'OUT ' + target + ' sudah CANCELLED.' };
    }

    // 4. Cari parent CTL-HEADER info
    var parentRowIdx = -1;
    var parentOwner  = '';
    var parentItem   = '';
    var parentMcNo   = '';
    var parentKgTgt  = 0;
    var parentSpec   = '';
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][iSpk] || '').trim() === parentSpk &&
          String(data[j][iType] || '').trim() === 'CTL-HEADER') {
        parentRowIdx = j;
        parentOwner  = String(data[j][iOwner] || '').trim();
        parentItem   = String(data[j][iItem]  || '').trim();
        parentMcNo   = String(data[j][iMC]    || '').trim();
        parentKgTgt  = parseFloat(data[j][iKgTgt]) || 0;
        break;
      }
    }
    if (parentRowIdx === -1) {
      return { success: false, message: 'Parent CTL-HEADER tidak ditemukan: ' + parentSpk };
    }

    // 5. Cascade cancel (acquires its own lock internally)
    var cancelResult = cascadeCancelSpk(target);
    if (!cancelResult.success) {
      return { success: false, message: 'Cascade cancel gagal: ' + (cancelResult.message || 'unknown') };
    }

    // 6. NG generation + header rollup (fresh lock for atomicity)
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);

    var createdNgList = [];
    var totalNgQty    = 0;
    var totalNgKg     = 0;
    var finalKgAct    = 0;

    try {
      var timestamp = new Date();
      var freshData = sheet.getDataRange().getValues();

      // 6a. Generate NG rows kalau ngItems disertakan
      if (ngItems && Array.isArray(ngItems) && ngItems.length > 0) {
        if (parentItem) {
          var miSheet = ss.getSheetByName('M_ITEM');
          if (miSheet) {
            var miData = miSheet.getDataRange().getValues();
            var miHdr  = miData[0].map(function(h){ return String(h).trim(); });
            var miIc   = miHdr.indexOf('Item_Code');
            var miSp   = miHdr.indexOf('Spec');
            for (var m = 1; m < miData.length; m++) {
              if (String(miData[m][miIc] || '').trim() === parentItem) {
                parentSpec = String(miData[m][miSp] || '').trim();
                break;
              }
            }
          }
        }

        var maxNgSuffix = 0;
        for (var k = 1; k < freshData.length; k++) {
          if (freshData[k][iType] === 'CTL-NG' &&
              String(freshData[k][iParent]).trim() === parentSpk) {
            var mm = String(freshData[k][iSpk]).match(/-NG(\d+)$/);
            if (mm) {
              var nn = parseInt(mm[1], 10) || 0;
              if (nn > maxNgSuffix) maxNgSuffix = nn;
            }
          }
        }

        ngItems.forEach(function(ng, ngIdx) {
          try {
            var ngT = parseFloat(ng.t)     || 0;
            var ngP = parseFloat(ng.p)     || 0;
            var ngL = parseFloat(ng.l)     || 0;
            var ngQ = parseInt(ng.qty, 10) || 0;
            if (ngT <= 0 || ngP <= 0 || ngL <= 0 || ngQ <= 0) {
              Logger.log('cancelCtlOut NG iter ' + ngIdx + ' SKIPPED - invalid dims');
              return;
            }
            maxNgSuffix++;
            var ngSpkNo = parentSpk + '-NG' + String(maxNgSuffix).padStart(2, '0');
            var ngKg    = Math.round(ngT * ngP * ngL * ngQ * 7.85 / 1000000);

            var ngData = {
              'SPK_No'     : ngSpkNo,
              'SPK_Type'   : 'CTL-NG',
              'Parent_SPK' : parentSpk,
              'Tgl_Buat'   : timestamp,
              'Item_Code'  : parentItem,
              'Input_Spec' : ngT + ' x ' + ngP + ' x ' + ngL,
              'Qty_Target' : ngQ,
              'KG_Target'  : ngKg,
              'Qty_Actual' : ngQ,
              'KG_Actual'  : ngKg,
              'MC_No'      : parentMcNo,
              'Owner'      : parentOwner,
              'Owner_Used' : parentOwner,
              'Status'     : 'DONE',
              'Target_Loc' : 'Stok_NG',
              'Selesai_DT' : timestamp,
              'Created_By' : 'PPIC (cancel-out)',
              'T'          : ngT,
              'P'          : ngP,
              'L'          : ngL
            };
            var ngRow = headers.map(function(h){ return ngData[h] !== undefined ? ngData[h] : ''; });
            _appendRowSafe(sheet, ngRow);
            SpreadsheetApp.flush();

            createdNgList.push({
              spk_no: ngSpkNo, item_code: parentItem, spec: parentSpec,
              t: ngT, p: ngP, l: ngL, qty: ngQ, kg: ngKg,
              owner: parentOwner, mc_no: parentMcNo
            });
            totalNgQty += ngQ;
            totalNgKg  += ngKg;
          } catch (ngErr) {
            Logger.log('cancelCtlOut NG iter ' + ngIdx + ' FAILED: ' + ngErr.toString());
          }
        });
      }

      // 6b. Re-read setelah NG appends → recompute header KG_Actual
      var finalData = sheet.getDataRange().getValues();

      var totalKgActOut = 0;
      for (var p = 1; p < finalData.length; p++) {
        if (String(finalData[p][iParent] || '').trim() === parentSpk &&
            String(finalData[p][iType]   || '').trim() === 'CTL-OUT') {
          var st = String(finalData[p][iStatus] || '').toUpperCase();
          if (st !== 'CANCELLED') {
            totalKgActOut += parseFloat(finalData[p][iKgAct]) || 0;
          }
        }
      }

      var curHeaderQtyNg = parseFloat(finalData[parentRowIdx][iQtyNg]) || 0;
      var curHeaderKgNg  = parseFloat(finalData[parentRowIdx][iKgNg])  || 0;

      finalKgAct = Math.round(totalKgActOut + curHeaderKgNg + totalNgKg);
      sheet.getRange(parentRowIdx + 1, iKgAct + 1).setValue(finalKgAct);

      if (totalNgQty > 0 || totalNgKg > 0) {
        sheet.getRange(parentRowIdx + 1, iQtyNg + 1).setValue(curHeaderQtyNg + totalNgQty);
        sheet.getRange(parentRowIdx + 1, iKgNg + 1).setValue(Math.round(curHeaderKgNg + totalNgKg));
      }

      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    // 7. Check header completion → set DONE kalau semua child sudah final
    if (typeof checkHeaderFinishedStatus === 'function') {
      checkHeaderFinishedStatus(parentSpk);
    }

    return {
      success: true,
      cancelled: cancelResult.cancelled || [],
      skippedDone: cancelResult.skippedDone || [],
      ng_created: createdNgList,
      header_kg_actual: finalKgAct
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}
/* =========================================================================
 * BATCH 3: cancelShrOutWithAuth
 * Cancel satu SHR-OUT dengan password gate PPIC. Sama persis pattern
 * dengan cancelCtlOutWithAuth, tapi target type SHR-OUT, SHR-HEADER,
 * SHR-NG. Opsional ngItems untuk cancel OUT terakhir dengan NG popup.
 * ========================================================================= */
function cancelShrOutWithAuth(spkOutNo, password, ngItems) {
  var CANCEL_PASSWORD = '321654';

  // 1. Verifikasi password
  if (!password || String(password).trim() !== CANCEL_PASSWORD) {
    return { success: false, message: 'Password salah. Aksi cancel dibatalkan.' };
  }

  // 2. Validasi spkOutNo
  if (!spkOutNo || String(spkOutNo).trim() === '') {
    return { success: false, message: 'SPK OUT No kosong.' };
  }

  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName("SPK");
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });

    var iSpk    = headers.indexOf("SPK_No");
    var iType   = headers.indexOf("SPK_Type");
    var iParent = headers.indexOf("Parent_SPK");
    var iStatus = headers.indexOf("Status");
    var iQtyAct = headers.indexOf("Qty_Actual");
    var iKgAct  = headers.indexOf("KG_Actual");
    var iQtyNg  = headers.indexOf("Qty_NG");
    var iKgNg   = headers.indexOf("KG_NG");
    var iItem   = headers.indexOf("Item_Code");
    var iMC     = headers.indexOf("MC_No");
    var iOwner  = headers.indexOf("Owner");
    var iKgTgt  = headers.indexOf("KG_Target");

    var target = String(spkOutNo).trim();

    // 3. Cari row SHR-OUT
    var outRowIdx  = -1;
    var outStatus  = '';
    var outType    = '';
    var parentSpk  = '';
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iSpk] || '').trim() === target) {
        outRowIdx = i;
        outStatus = String(data[i][iStatus] || '').toUpperCase();
        outType   = String(data[i][iType]   || '').trim();
        parentSpk = String(data[i][iParent] || '').trim();
        break;
      }
    }

    if (outRowIdx === -1) {
      return { success: false, message: 'SPK OUT tidak ditemukan: ' + target };
    }
    if (outType !== 'SHR-OUT') {
      return { success: false, message: 'Fitur ini hanya untuk SHR-OUT. Type: ' + outType };
    }
    if (outStatus === 'DONE') {
      return { success: false, message: 'OUT ' + target + ' sudah DONE - tidak bisa di-cancel.' };
    }
    if (outStatus === 'CANCELLED') {
      return { success: false, message: 'OUT ' + target + ' sudah CANCELLED.' };
    }

    // 4. Cari parent SHR-HEADER info
    var parentRowIdx = -1;
    var parentOwner  = '';
    var parentItem   = '';
    var parentMcNo   = '';
    var parentKgTgt  = 0;
    var parentSpec   = '';
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][iSpk] || '').trim() === parentSpk &&
          String(data[j][iType] || '').trim() === 'SHR-HEADER') {
        parentRowIdx = j;
        parentOwner  = String(data[j][iOwner] || '').trim();
        parentItem   = String(data[j][iItem]  || '').trim();
        parentMcNo   = String(data[j][iMC]    || '').trim();
        parentKgTgt  = parseFloat(data[j][iKgTgt]) || 0;
        break;
      }
    }
    if (parentRowIdx === -1) {
      return { success: false, message: 'Parent SHR-HEADER tidak ditemukan: ' + parentSpk };
    }

    // 5. Cascade cancel (acquires lock internally)
    var cancelResult = cascadeCancelSpk(target);
    if (!cancelResult.success) {
      return { success: false, message: 'Cascade cancel gagal: ' + (cancelResult.message || 'unknown') };
    }

    // 6. NG generation + header rollup (fresh lock for atomicity)
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);

    var createdNgList = [];
    var totalNgQty    = 0;
    var totalNgKg     = 0;
    var finalKgAct    = 0;

    try {
      var timestamp = new Date();
      var freshData = sheet.getDataRange().getValues();

      // 6a. Generate SHR-NG rows kalau ngItems disertakan
      if (ngItems && Array.isArray(ngItems) && ngItems.length > 0) {
        // Lookup Spec dari M_ITEM
        if (parentItem) {
          var miSheet = ss.getSheetByName('M_ITEM');
          if (miSheet) {
            var miData = miSheet.getDataRange().getValues();
            var miHdr  = miData[0].map(function(h){ return String(h).trim(); });
            var miIc   = miHdr.indexOf('Item_Code');
            var miSp   = miHdr.indexOf('Spec');
            for (var m = 1; m < miData.length; m++) {
              if (String(miData[m][miIc] || '').trim() === parentItem) {
                parentSpec = String(miData[m][miSp] || '').trim();
                break;
              }
            }
          }
        }

        var maxNgSuffix = 0;
        for (var k = 1; k < freshData.length; k++) {
          if (freshData[k][iType] === 'SHR-NG' &&
              String(freshData[k][iParent]).trim() === parentSpk) {
            var mm = String(freshData[k][iSpk]).match(/-NG(\d+)$/);
            if (mm) {
              var nn = parseInt(mm[1], 10) || 0;
              if (nn > maxNgSuffix) maxNgSuffix = nn;
            }
          }
        }

        ngItems.forEach(function(ng, ngIdx) {
          try {
            var ngT = parseFloat(ng.t)     || 0;
            var ngP = parseFloat(ng.p)     || 0;
            var ngL = parseFloat(ng.l)     || 0;
            var ngQ = parseInt(ng.qty, 10) || 0;
            if (ngT <= 0 || ngP <= 0 || ngL <= 0 || ngQ <= 0) {
              Logger.log('cancelShrOut NG iter ' + ngIdx + ' SKIPPED - invalid dims');
              return;
            }
            maxNgSuffix++;
            var ngSpkNo = parentSpk + '-NG' + String(maxNgSuffix).padStart(2, '0');
            var ngKg    = Math.round(ngT * ngP * ngL * ngQ * 7.85 / 1000000);

            var ngData = {
              'SPK_No'     : ngSpkNo,
              'SPK_Type'   : 'SHR-NG',
              'Parent_SPK' : parentSpk,
              'Tgl_Buat'   : timestamp,
              'Item_Code'  : parentItem,
              'Input_Spec' : ngT + ' x ' + ngP + ' x ' + ngL,
              'Qty_Target' : ngQ,
              'KG_Target'  : ngKg,
              'Qty_Actual' : ngQ,
              'KG_Actual'  : ngKg,
              'MC_No'      : parentMcNo,
              'Owner'      : parentOwner,
              'Owner_Used' : parentOwner,
              'Status'     : 'DONE',
              'Target_Loc' : 'Stok_NG',
              'Selesai_DT' : timestamp,
              'Created_By' : 'PPIC (cancel-out)',
              'T'          : ngT,
              'P'          : ngP,
              'L'          : ngL
            };
            var ngRow = headers.map(function(h){ return ngData[h] !== undefined ? ngData[h] : ''; });
            _appendRowSafe(sheet, ngRow);
            SpreadsheetApp.flush();

            createdNgList.push({
              spk_no: ngSpkNo, item_code: parentItem, spec: parentSpec,
              t: ngT, p: ngP, l: ngL, qty: ngQ, kg: ngKg,
              owner: parentOwner, mc_no: parentMcNo
            });
            totalNgQty += ngQ;
            totalNgKg  += ngKg;
          } catch (ngErr) {
            Logger.log('cancelShrOut NG iter ' + ngIdx + ' FAILED: ' + ngErr.toString());
          }
        });
      }

      // 6b. Re-read setelah NG appends → recompute header KG_Actual
      var finalData = sheet.getDataRange().getValues();

      var totalKgActOut = 0;
      for (var p = 1; p < finalData.length; p++) {
        if (String(finalData[p][iParent] || '').trim() === parentSpk &&
            String(finalData[p][iType]   || '').trim() === 'SHR-OUT') {
          var st = String(finalData[p][iStatus] || '').toUpperCase();
          if (st !== 'CANCELLED') {
            totalKgActOut += parseFloat(finalData[p][iKgAct]) || 0;
          }
        }
      }

      var curHeaderQtyNg = parseFloat(finalData[parentRowIdx][iQtyNg]) || 0;
      var curHeaderKgNg  = parseFloat(finalData[parentRowIdx][iKgNg])  || 0;

      finalKgAct = Math.round(totalKgActOut + curHeaderKgNg + totalNgKg);
      sheet.getRange(parentRowIdx + 1, iKgAct + 1).setValue(finalKgAct);

      if (totalNgQty > 0 || totalNgKg > 0) {
        sheet.getRange(parentRowIdx + 1, iQtyNg + 1).setValue(curHeaderQtyNg + totalNgQty);
        sheet.getRange(parentRowIdx + 1, iKgNg + 1).setValue(Math.round(curHeaderKgNg + totalNgKg));
      }

      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    // 7. Check header completion → set DONE kalau semua child sudah final
    if (typeof checkHeaderFinishedStatus === 'function') {
      checkHeaderFinishedStatus(parentSpk);
    }

    return {
      success: true,
      cancelled: cancelResult.cancelled || [],
      skippedDone: cancelResult.skippedDone || [],
      ng_created: createdNgList,
      header_kg_actual: finalKgAct
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}
/* =========================================================================
 * HELPER: Format tanggal ke format Indonesia "DD MMM YYYY [HH:mm]"
 *   - val: Date object atau string ISO/datetime
 *   - withTime: true → tambahkan jam-menit; false → tanggal saja
 *   - Return: "02 Jun 2026" atau "02 Jun 2026 20:04"
 * ========================================================================= */
function _formatTglID(val, withTime) {
  if (!val) return '';
  var d = val instanceof Date ? val : new Date(val);
  if (isNaN(d.getTime())) return String(val);

  var tz = Session.getScriptTimeZone();
  var bulan = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agt','Sep','Okt','Nov','Des'];

  var dd   = Utilities.formatDate(d, tz, 'dd');
  var mIdx = parseInt(Utilities.formatDate(d, tz, 'MM'), 10) - 1;
  var yy   = Utilities.formatDate(d, tz, 'yyyy');
  var bln  = (mIdx >= 0 && mIdx < 12) ? bulan[mIdx] : '';

  if (withTime) {
    var hhmm = Utilities.formatDate(d, tz, 'HH:mm');
    return dd + ' ' + bln + ' ' + yy + ' ' + hhmm;
  }
  return dd + ' ' + bln + ' ' + yy;
}

/* =========================================================================
 * SHR EDIT MODE — Daftar SHR-HEADER yang bisa diedit (status ANTRIAN/OPEN)
 *                 & EXCLUDE SHR yang dibuat dari flow CTL (cuma standalone)
 * ========================================================================= */
function getEditableSpkListSHR() {
  const sheet = getSheet("SPK");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const iSpk    = headers.indexOf("SPK_No");
  const iType   = headers.indexOf("SPK_Type");
  const iStatus = headers.indexOf("Status");
  const iParent = headers.indexOf("Parent_SPK");
  const iSpec   = headers.indexOf("Input_Spec");

  // Set CTL-OUT untuk filter SHR yang berasal dari flow CTL
  const ctlOutSet = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][iType] === 'CTL-OUT') ctlOutSet[String(data[i][iSpk]).trim()] = true;
  }

  let result = [];
  for (let i = 1; i < data.length; i++) {
    const type   = data[i][iType];
    const status = String(data[i][iStatus] || '').toUpperCase();
    const parent = String(data[i][iParent] || '').trim();

    if (type === 'SHR-HEADER'
        && (status === 'ANTRIAN' || status === 'OPEN')
        && !ctlOutSet[parent]) {
      result.push({
        spk_no: data[i][iSpk],
        parent: data[i][iParent],
        spec  : data[i][iSpec],
        status: status
      });
    }
  }
  return result;
}

/* =========================================================================
 * GET FULL SHR DATA — Untuk Edit Mode (exclude rows CANCELLED)
 *   - Include status & tgl_buat di response (BUGFIX)
 *   - Format tgl_buat pakai _formatTglID (DD MMM YYYY HH:mm)
 * ========================================================================= */
function getSpkFullDataSHR(spkNo) {
  const sheet = getSheet("SPK");
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h){ return String(h).trim(); });

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach(function(h, idx){
      let val = data[i][idx];
      if (val instanceof Date) val = val.toISOString();
      obj[h] = val;
    });
    rows.push(obj);
  }

  const header = rows.find(function(r){
    return String(r.SPK_No).trim() === String(spkNo).trim() && r.SPK_Type === 'SHR-HEADER';
  });
  if (!header) throw new Error("SHR-HEADER tidak ditemukan: " + spkNo);

  function notCancelled(r){ return String(r.Status || '').toUpperCase() !== 'CANCELLED'; }

  const shrOuts = rows.filter(function(r){
    return String(r.Parent_SPK).trim() === String(spkNo).trim()
        && r.SPK_Type === 'SHR-OUT' && notCancelled(r);
  });
  shrOuts.sort(function(a, b){
    const aN = parseInt(String(a.SPK_No).split('-').pop(), 10) || 0;
    const bN = parseInt(String(b.SPK_No).split('-').pop(), 10) || 0;
    return aN - bN;
  });

  const out_shr = shrOuts.map(function(shr){
    let cutP = 1, cutL = 1;
    if (shr.NOTE) {
      const mP = String(shr.NOTE).match(/Cut_P:\s*(\d+)/i);
      const mL = String(shr.NOTE).match(/Cut_L:\s*(\d+)/i);
      if (mP) cutP = parseInt(mP[1], 10);
      if (mL) cutL = parseInt(mL[1], 10);
    }
    return {
      spk_no     : shr.SPK_No,
      item_code  : shr.Item_Code,
      description: shr.Input_Spec,
      target_loc : shr.Target_Loc,
      so_no      : shr.SO_Ref,
      cust       : shr.Cust,
      owner      : shr.Owner      || 'FC',
      owner_used : shr.Owner_Used || shr.Owner || 'FC',
      cut_p      : cutP, cut_l: cutL,
      qty_plan   : shr.Qty_Target,
      kg_plan    : shr.KG_Target
    };
  });

  return {
    header: {
      spk_no     : header.SPK_No,
      status     : String(header.Status || '').toUpperCase(),
      tgl_buat   : _formatTglID(header.Tgl_Buat, true),
      batch_id   : header.Parent_SPK,
      item_code  : header.Item_Code,
      input_spec : header.Input_Spec,
        qty_target : header.Qty_Target,
        kg_target  : header.KG_Target,
        mc_no      : header.MC_No,
      source_loc : header.Source_Loc,
      priority   : header.Priority,
      op         : header.OP,
      owner      : header.Owner      || 'FC',
      owner_used : header.Owner_Used || header.Owner || 'FC',
      cust       : header.Cust || ''
    },
    out_shr: out_shr
  };
}
/* =========================================================================
 * updateAndAppendSpkSHR — OWNER FIX + PENDING_CANCEL HANDLER (B1 FIX)
 * Perubahan B1:
 *   - Baca flag shr.pending_cancel dari payload
 *   - Kalau flag ada + row existing → set Status = CANCELLED, skip update/append
 *   - Return value tambahan: cancelled[], count_cancelled
 * ========================================================================= */
function updateAndAppendSpkSHR(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // ✅ Validasi stok sheet/WIP (Bug 3) — exclude SPK ini sendiri
    //    Teruskan source_loc supaya validator tahu cek Stok_Sheet atau Stok_WIP.
    if (typeof validateSheetAvailability === 'function') {
      validateSheetAvailability(data.batch_id, data.qty_input, data.kg_input, data.spk_no, data.source_loc);
    }
    const spkSheet = getSheet("SPK");
    const allData  = spkSheet.getDataRange().getValues();
    const headers  = allData[0].map(function(h){ return String(h).trim(); });
    const timestamp = new Date();
    const spkNo    = data.spk_no;
    const generatedSpks = [], updatedSpks = [], cancelledSpks = [];

    // ✅ Owner material dari batch
    const batchOwner = String(data.batch_owner || data.owner || 'FC').trim().toUpperCase();

    function buildRow(rd){ return headers.map(function(h){ return rd[h] !== undefined ? rd[h] : ''; }); }
    function col(name){ const i = headers.indexOf(name); return i >= 0 ? i + 1 : -1; }

    const iSpk    = headers.indexOf("SPK_No");
    const iType   = headers.indexOf("SPK_Type");
    const iParent = headers.indexOf("Parent_SPK");
    const iStatus = headers.indexOf("Status");
    const statusCol = iStatus + 1;

    const rowMap = {};
    for (let i = 1; i < allData.length; i++) {
      const sn = String(allData[i][iSpk] || '').trim();
      if (sn) rowMap[sn] = { rowNum: i + 1 };
    }
    function updateRow(rowNum, rd) {
      Object.keys(rd).forEach(function(k){
        const c = col(k);
        if (c > 0) spkSheet.getRange(rowNum, c).setValue(rd[k]);
      });
    }

    // Pre-kalkulasi CT
    var outCtList = data.out_shr.map(function(shr) {
      return hitungRencanaDurasi(shr.item_code, shr.qty_plan, shr.kg_plan);
    });

    // Agregat HEADER
    var hdrPlanRun     = 0;
    var hdrTotalDurasi = 0;
    outCtList.forEach(function(ct) {
      hdrPlanRun     += ct.planRun;
      hdrTotalDurasi += ct.planSetup + ct.planRun;
    });

    // UPDATE SHR-HEADER
    if (!rowMap[spkNo]) throw new Error("SHR-HEADER tidak ditemukan: " + spkNo);
    updateRow(rowMap[spkNo].rowNum, {
      'Priority'           : data.priority,
      'Source_Loc'         : data.source_loc,
      'Qty_Target'         : data.qty_input,
      'KG_Target'          : data.kg_input,
      'MC_No'              : data.machine,
      'OP'                 : data.op,
      'Owner'              : batchOwner,
      'Owner_Used'         : batchOwner,
      'Plan_Setup_Menit'   : 0,
      'Plan_Run_Menit'     : hdrPlanRun,
      'Total_Durasi_Menit' : hdrTotalDurasi,
      'T'                  : parseFloat(data.thick) || 0
    });
    updatedSpks.push(spkNo);

    // Cari suffix tertinggi
    let nextSuffix = 0;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][iType] === 'SHR-OUT' && String(allData[i][iParent]).trim() === spkNo.trim()) {
        const sf = parseInt(String(allData[i][iSpk]).trim().split('-').pop(), 10) || 0;
        if (sf > nextSuffix) nextSuffix = sf;
      }
    }

    // UPDATE / APPEND / CANCEL SHR-OUT
    data.out_shr.forEach(function(shr, idx) {

      // ── B1 FIX: PENDING CANCEL HANDLER ──
      // Kalau frontend tandai baris ini untuk cancel + sudah ada di sheet → set CANCELLED, skip sisanya
      if (shr.pending_cancel && shr.existing_spk_no && rowMap[shr.existing_spk_no]) {
        if (statusCol > 0) {
          const curStatus = String(allData[rowMap[shr.existing_spk_no].rowNum - 1][iStatus] || '').toUpperCase();
          // Hanya cancel kalau belum DONE/CANCELLED (proteksi data jadi)
          if (curStatus !== 'DONE' && curStatus !== 'CANCELLED') {
            spkSheet.getRange(rowMap[shr.existing_spk_no].rowNum, statusCol).setValue("CANCELLED");
            cancelledSpks.push(shr.existing_spk_no);
          }
        }
        return; // skip update/append untuk row ini
      }

      var ctO = outCtList[idx];
      var ownerUsed = String(shr.owner_used || batchOwner).trim().toUpperCase();

      var payload = {
        'SO_Ref'             : shr.so_no       || '',
        'Cust'               : shr.cust         || '',
        'Priority'           : data.priority,
        'Item_Code'          : shr.item_code    || '',
        'Input_Spec'         : shr.description  || '',
        'Target_Loc'         : shr.target_loc   || '',
        'BQ'                 : shr.cut_p * shr.cut_l,
        'Qty_Target'         : shr.qty_plan,
        'KG_Target'          : shr.kg_plan,
        'OP'                 : data.op,
        'Owner'              : batchOwner,
        'Owner_Used'         : ownerUsed,
        'NOTE'               : 'Cut_P:' + shr.cut_p + ' Cut_L:' + shr.cut_l,
        'Plan_Setup_Menit'   : ctO.planSetup,
        'Plan_Run_Menit'     : ctO.planRun,
        'Total_Durasi_Menit' : 0,
        'T'                  : parseFloat(data.thick) || 0
      };

      if (shr.existing_spk_no && rowMap[shr.existing_spk_no]) {
        updateRow(rowMap[shr.existing_spk_no].rowNum, payload);
        updatedSpks.push(shr.existing_spk_no);
      } else {
        nextSuffix++;
        const newNo = spkNo + '-' + String(nextSuffix).padStart(2, '0');
        _appendRowSafe(spkSheet,buildRow(Object.assign({
          'SPK_No'     : newNo,
          'SPK_Type'   : 'SHR-OUT',
          'Parent_SPK' : spkNo,
          'Tgl_Buat'   : timestamp,
          'Source_Loc' : data.machine,
          'MC_No'      : data.machine,
          'Status'     : 'ANTRIAN',
          'Created_By' : 'Admin'
        }, payload)));
        generatedSpks.push(newNo);
      }
    });

    SpreadsheetApp.flush();
    if (typeof kalkulasiEstimasiWaktu === 'function') kalkulasiEstimasiWaktu();

    return {
      spk_no          : spkNo,
      generated       : generatedSpks,
      updated         : updatedSpks,
      cancelled       : cancelledSpks,
      count_new       : generatedSpks.length,
      count_updated   : updatedSpks.length,
      count_cancelled : cancelledSpks.length,
      mode            : 'UPDATE_APPEND'
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * 🟢 T1.5b — saveAndRefreshBoard REVISI
 *
 * Perubahan vs T1c:
 *  - HAPUS call writeStokSheet (CTL-OUT DONE & SHR-OUT→Stok_Sheet)
 *  - HAPUS call writeStokWIP   (SHR-OUT→WIP_*)
 *  - HAPUS call writeStokFG    (SHR-OUT→FG_*)
 *  - TETAP writeTraceLog (sumber utama untuk formula Stok_*)
 *  - TETAP writeRekapICT (untuk cross-billing FC↔DRC)
 *  - TETAP _propagateShtBatchToShrChildren
 *  - TETAP T1 NG Trace_Log writing
 *
 * Setelah patch ini, GAS write hanya: SPK, Trace_Log, Rekap_ICT.
 * Stok_Coil/Sheet/WIP/FG semuanya formula-based.
 * ========================================================================= */
function saveAndRefreshBoard(payload, machineNo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const spkSheet  = ss.getSheetByName("SPK");
    if (!spkSheet) return { success: false, message: 'Sheet SPK tidak ditemukan' };
    const timestamp = new Date();
    const tz        = Session.getScriptTimeZone();

    // ===== 1. READ SPK 1× =====
    let rows = spkSheet.getDataRange().getValues();
    const headers = rows[0].map(function(h){ return String(h).trim(); });

    const I = {
      spk     : headers.indexOf("SPK_No"),
      type    : headers.indexOf("SPK_Type"),
      parent  : headers.indexOf("Parent_SPK"),
      status  : headers.indexOf("Status"),
      selesai : headers.indexOf("Selesai_DT"),
      qtyAct  : headers.indexOf("Qty_Actual"),
      kgAct   : headers.indexOf("KG_Actual"),
      qtyNg   : headers.indexOf("Qty_NG"),
      kgNg    : headers.indexOf("KG_NG"),
      item    : headers.indexOf("Item_Code"),
      qtyTgt  : headers.indexOf("Qty_Target"),
      kgTgt   : headers.indexOf("KG_Target"),
      mc      : headers.indexOf("MC_No"),
      op      : headers.indexOf("OP"),
      owner   : headers.indexOf("Owner"),
      ownerU  : headers.indexOf("Owner_Used"),
      tgtLoc  : headers.indexOf("Target_Loc"),
      soRef   : headers.indexOf("SO_Ref"),
      cust    : headers.indexOf("Cust"),
      batch   : headers.indexOf("Batch_ID"),
      spec    : headers.indexOf("Input_Spec"),
      isHabis : headers.indexOf("Is_Habis"),
      prio    : headers.indexOf("Priority"),
      planSeq : headers.indexOf("Plan_Seq")
    };

    // ===== 2. FIND CHILD OUT =====
    const targetSpk = String(payload.spk_no || '').trim();
    if (!targetSpk) return { success: false, message: 'spk_no kosong' };

    let childRowIdx = -1, childArrIdx = -1;
    let parentSpkNo = '', childRowData = null;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][I.spk]).trim() === targetSpk
          && (rows[i][I.type] === 'CTL-OUT' || rows[i][I.type] === 'SHR-OUT')) {
        childRowIdx  = i + 1;
        childArrIdx  = i;
        parentSpkNo  = String(rows[i][I.parent]).trim();
        childRowData = rows[i];
        break;
      }
    }
    if (childRowIdx === -1) return { success: false, message: 'SPK OUT tidak ditemukan: ' + targetSpk };

    const qtyAct = Math.round(parseFloat(payload.qty_actual) || 0);
    const kgAct  = Math.round(parseFloat(payload.kg_actual)  || 0);
    if (qtyAct <= 0) return { success: false, message: 'Qty Actual harus > 0' };

    // ===== 2.X POKA-YOKE: SHR-HEADER Qty_Used wajib di OUT terakhir =====
    // 🟢 BUG-FIX #2 — Validasi backend: kalau ini SHR-OUT terakhir (yang akan
    // trigger header DONE), payload.qty_used_header WAJIB ada. Throw sebelum
    // write supaya tidak ada partial state.
    if (parentSpkNo && parentSpkNo.indexOf('SHR-') === 0) {
      let remainingOut = 0;
      for (let i = 1; i < rows.length; i++) {
        if (i === childArrIdx) continue;
        const rt = String(rows[i][I.type] || '');
        const rp = String(rows[i][I.parent] || '').trim();
        const rs = String(rows[i][I.status] || '').toUpperCase();
        if (rp === parentSpkNo && rt === 'SHR-OUT' && rs !== 'DONE' && rs !== 'CANCELLED') {
          remainingOut++;
        }
      }
      if (remainingOut === 0) {
        const qu = parseFloat(payload.qty_used_header);
        if (!qu || qu <= 0) {
          throw new Error('🚫 Qty_Used SHR-HEADER wajib diisi pada OUT terakhir. ' +
                          'Lengkapi section "Finalize Header" di modal sebelum submit.');
        }
      }
    }

    // ===== 3. UPDATE CHILD ROW =====
    spkSheet.getRange(childRowIdx, I.qtyAct  + 1).setValue(qtyAct);
    spkSheet.getRange(childRowIdx, I.kgAct   + 1).setValue(kgAct);
    spkSheet.getRange(childRowIdx, I.status  + 1).setValue('DONE');
    spkSheet.getRange(childRowIdx, I.selesai + 1).setValue(timestamp);
    rows[childArrIdx][I.qtyAct]  = qtyAct;
    rows[childArrIdx][I.kgAct]   = kgAct;
    rows[childArrIdx][I.status]  = 'DONE';
    rows[childArrIdx][I.selesai] = timestamp;

    // ===== 4. FIND PARENT HEADER =====
    let parentRowIdx = -1, parentArrIdx = -1;
    let parentQtyTgt = 0, parentKgTgt = 0;
    let parentMcNo = '', parentOwner = '', parentItemCode = '';
    let parentSourceBatch = '';
    let headerIsHabis = false;
    let totalKgActOut = 0;

    if (parentSpkNo) {
      for (let i = 1; i < rows.length; i++) {
        const rowSpk    = String(rows[i][I.spk] || '').trim();
        const rowParent = String(rows[i][I.parent] || '').trim();
        const rowType   = String(rows[i][I.type] || '');
        const rowStatus = String(rows[i][I.status] || '').toUpperCase();

        if (rowSpk === parentSpkNo) {
          parentRowIdx       = i + 1;
          parentArrIdx       = i;
          parentQtyTgt       = Number(rows[i][I.qtyTgt]) || 0;
          parentKgTgt        = Number(rows[i][I.kgTgt])  || 0;
          parentMcNo         = String(rows[i][I.mc]      || '').trim();
          parentOwner        = String(rows[i][I.owner]   || '').trim();
          parentItemCode     = String(rows[i][I.item]    || '').trim();
          parentSourceBatch  = String(rows[i][I.parent]  || '').trim();
          headerIsHabis      = (rows[i][I.isHabis] === true || String(rows[i][I.isHabis]).toUpperCase() === 'TRUE');
        }
        if (rowParent === parentSpkNo
            && (rowType === 'SHR-OUT' || rowType === 'CTL-OUT')
            && rowStatus !== 'CANCELLED') {
          totalKgActOut += Number(rows[i][I.kgAct]) || 0;
        }
      }
    }

    // ===== 5. NG ITEMS — BATCH APPEND =====
    const isCtlHeader = parentSpkNo && parentSpkNo.indexOf('CTL-') === 0;
    const isShrHeader = parentSpkNo && parentSpkNo.indexOf('SHR-') === 0;
    const isAnyHeader = isCtlHeader || isShrHeader;
    const ngType      = isCtlHeader ? 'CTL-NG' : (isShrHeader ? 'SHR-NG' : '');

    let createdNgList = [];
    let totalNgQty = 0, totalNgKg = 0;
    let parentSpec = '';

    const miSheet = ss.getSheetByName('M_ITEM');
    let miMap = {};
    if (miSheet) {
      const miData = miSheet.getDataRange().getValues();
      const miHdr  = miData[0].map(function(h){ return String(h).trim(); });
      const miIc   = miHdr.indexOf('Item_Code');
      const miDesc = miHdr.indexOf('Description');
      const miSpec = miHdr.indexOf('Spec');
      const miT    = miHdr.indexOf('T');
      const miP    = miHdr.indexOf('P');
      const miL    = miHdr.indexOf('L');
      const miUom  = miHdr.indexOf('UoM_MC');
      for (let m = 1; m < miData.length; m++) {
        const k = String(miData[m][miIc] || '').trim();
        if (!k) continue;
        miMap[k] = {
          desc: String(miData[m][miDesc] || '').trim(),
          spec: String(miData[m][miSpec] || '').trim(),
          t   : String(miData[m][miT]    || '').trim(),
          p   : String(miData[m][miP]    || '').trim(),
          l   : String(miData[m][miL]    || '').trim(),
          uom : String(miData[m][miUom]  || 'Sht').trim()
        };
      }
    }
    if (parentItemCode && miMap[parentItemCode]) parentSpec = miMap[parentItemCode].spec;

    if (isAnyHeader && payload.ng_items && Array.isArray(payload.ng_items) && payload.ng_items.length > 0) {
      let maxNgSuffix = 0;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][I.type] === ngType && String(rows[i][I.parent]).trim() === parentSpkNo) {
          const m = String(rows[i][I.spk]).match(/-NG(\d+)$/);
          if (m) { const n = parseInt(m[1], 10) || 0; if (n > maxNgSuffix) maxNgSuffix = n; }
        }
      }

      const ngRowsToAppend = [];
      let baseRowCount = 1;
      for (let r = rows.length - 1; r >= 1; r--) {
        if (rows[r][I.spk] && String(rows[r][I.spk]).trim() !== '') {
          baseRowCount = r + 1; break;
        }
      }

      payload.ng_items.forEach(function(ng) {
        const ngT = parseFloat(ng.t)     || 0;
        const ngP = parseFloat(ng.p)     || 0;
        const ngL = parseFloat(ng.l)     || 0;
        const ngQ = parseInt(ng.qty, 10) || 0;
        if (ngT <= 0 || ngP <= 0 || ngL <= 0 || ngQ <= 0) return;

        maxNgSuffix++;
        const ngSpkNo = parentSpkNo + '-NG' + String(maxNgSuffix).padStart(2, '0');
        const ngKg    = Math.round(ngT * ngP * ngL * ngQ * 7.85 / 1000000);

        const ngData = {
          'SPK_No'     : ngSpkNo,
          'SPK_Type'   : ngType,
          'Parent_SPK' : parentSpkNo,
          'Batch_ID'   : ngSpkNo,            // 🟢 BUG-FIX #3 — NG self-batch (konvensi sama dgn Trace_Log)
          'Tgl_Buat'   : timestamp,
          'Item_Code'  : parentItemCode,
          'Input_Spec' : ngT + ' x ' + ngP + ' x ' + ngL,
          'Qty_Target' : ngQ,
          'KG_Target'  : ngKg,
          'Qty_Actual' : ngQ,
          'KG_Actual'  : ngKg,
          'MC_No'      : parentMcNo,
          'Owner'      : parentOwner,
          'Owner_Used' : parentOwner,
          'Status'     : 'DONE',
          'Target_Loc' : 'Stok_NG',
          'Selesai_DT' : timestamp,
          'Created_By' : payload.created_by || 'operator',
          'T'          : ngT, 'P': ngP, 'L': ngL
        };
        const ngRow = headers.map(function(h){ return ngData[h] !== undefined ? ngData[h] : ''; });

        ngRowsToAppend.push(ngRow);
        rows.push(ngRow);

        createdNgList.push({
          spk_no    : ngSpkNo, item_code : parentItemCode, spec : parentSpec,
          t : ngT, p : ngP, l : ngL, qty : ngQ, kg : ngKg,
          owner : parentOwner, mc_no : parentMcNo
        });

        totalNgQty += ngQ;
        totalNgKg  += ngKg;
      });

      if (ngRowsToAppend.length > 0) {
        const startRow = baseRowCount + 1;
        spkSheet.getRange(startRow, 1, ngRowsToAppend.length, headers.length).setValues(ngRowsToAppend);
      }
    }

    // ===== 6. UPDATE HEADER ROLLUP =====
    if (parentRowIdx !== -1) {
      // 🟢 BUG-FIX #2 — Resolve qty_used_header untuk SHR (kalau dikirim)
      const qtyUsedHeader = isShrHeader ? (parseFloat(payload.qty_used_header) || 0) : 0;
      const bqHeader = (isShrHeader && parentQtyTgt > 0) ? (parentKgTgt / parentQtyTgt) : 0;

      let finalKgAct;
      if (isCtlHeader) {
        finalKgAct = headerIsHabis ? Math.round(parentKgTgt) : Math.round(totalKgActOut + totalNgKg);
      } else if (isShrHeader) {
        // 🟢 BUG-FIX #2 — SHR-HEADER KG_Actual = Qty_Used × BQ_input
        //   Bukan sum OUT+NG, karena SHR ada scrap/serbuk loss yg tidak ter-track.
        //   Fallback ke parentKgTgt kalau qty_used belum ada (OUT non-terakhir).
        finalKgAct = qtyUsedHeader > 0
          ? Math.round(qtyUsedHeader * bqHeader)
          : Math.round(parentKgTgt);
      } else {
        finalKgAct = Math.round(totalKgActOut);
      }

      // 🟢 Qty_Actual logic per type:
      //   CTL-HEADER : = KG_Actual (coil unit = KG)
      //   SHR-HEADER : = Qty_Used (lembar input yang diproses operator)
      //   else       : = Qty_Target
      let qtyActFinal;
      if (isCtlHeader)      qtyActFinal = finalKgAct;
      else if (isShrHeader) qtyActFinal = qtyUsedHeader > 0 ? qtyUsedHeader : parentQtyTgt;
      else                  qtyActFinal = parentQtyTgt;

      spkSheet.getRange(parentRowIdx, I.qtyAct + 1).setValue(qtyActFinal);
      spkSheet.getRange(parentRowIdx, I.kgAct  + 1).setValue(finalKgAct);
      rows[parentArrIdx][I.qtyAct] = qtyActFinal;
      rows[parentArrIdx][I.kgAct]  = finalKgAct;

      if (isAnyHeader && (totalNgQty > 0 || totalNgKg > 0)) {
        const curHeaderQtyNg = parseFloat(rows[parentArrIdx][I.qtyNg]) || 0;
        const curHeaderKgNg  = parseFloat(rows[parentArrIdx][I.kgNg])  || 0;
        const newQtyNg = curHeaderQtyNg + totalNgQty;
        const newKgNg  = Math.round(curHeaderKgNg + totalNgKg);
        spkSheet.getRange(parentRowIdx, I.qtyNg + 1).setValue(newQtyNg);
        spkSheet.getRange(parentRowIdx, I.kgNg  + 1).setValue(newKgNg);
        rows[parentArrIdx][I.qtyNg] = newQtyNg;
        rows[parentArrIdx][I.kgNg]  = newKgNg;
      }
    }

    // ===== 7. CHECK HEADER COMPLETION =====
    let headerDone = false;
    const targetOutType = isCtlHeader ? 'CTL-OUT' : 'SHR-OUT';
    if (parentSpkNo && parentRowIdx !== -1) {
      let anyChildPending = false;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][I.type] === targetOutType && String(rows[i][I.parent]).trim() === parentSpkNo) {
          const st = String(rows[i][I.status] || '').toUpperCase();
          if (st !== 'DONE' && st !== 'CANCELLED') { anyChildPending = true; break; }
        }
      }
      if (!anyChildPending) {
        spkSheet.getRange(parentRowIdx, I.status  + 1).setValue('DONE');
        spkSheet.getRange(parentRowIdx, I.selesai + 1).setValue(timestamp);
        rows[parentArrIdx][I.status]  = 'DONE';
        rows[parentArrIdx][I.selesai] = timestamp;
        headerDone = true;
      }
    }

    // ===== 8. SINGLE FLUSH =====
    SpreadsheetApp.flush();

    // ===== 9. WRITE TRACE_LOG (untuk OUT row yang DONE) =====
    // 🟢 T1.5: writeStokSheet/WIP/FG dihapus. Stok_* derive dari Trace_Log + SPK
    if (childRowData && qtyAct > 0) {
      const spkType   = String(childRowData[I.type]   || '');
      const targetLoc = String(childRowData[I.tgtLoc] || '').trim();
      const owner     = String(childRowData[I.owner]  || '').trim();
      const ownerUsed = String(childRowData[I.ownerU] || '').trim();
      const itemCode  = String(childRowData[I.item]   || '').trim();
      const soRef     = String(childRowData[I.soRef]  || '').trim();
      const cust      = String(childRowData[I.cust]   || '').trim();
      const mcNo      = String(childRowData[I.mc]     || '').trim();
      const op        = String(childRowData[I.op]     || '').trim();
      const srcBatch  = I.batch > -1 ? String(childRowData[I.batch] || '').trim() : '';

      const mi = miMap[itemCode] || { desc:'', spec:'', t:'', p:'', l:'', uom:'Sht' };
      const rootBatch    = getRootBatch(srcBatch) || srcBatch;
      const supplierInfo = getCoilSupplierInfo(rootBatch);

      if (spkType === 'CTL-OUT') {
        // 🟢 FIX: CTL-OUT respect target_loc + prefix match untuk future-proof
        // Recognize: WIP_* (Cust, Stamping, Transit_Shearing, dll) + FG_* (Cust, RM_Stamping, dll) + Stok_Sheet
        // Skip: Stok_Coil (sisa coil), Scrap_Area (scrap) — bukan output produksi
        const ctlIsWIP = (targetLoc && targetLoc.indexOf('WIP_') === 0);
        const ctlIsFG  = (targetLoc && targetLoc.indexOf('FG_')  === 0);
        const ctlIsSht = (targetLoc === 'Stok_Sheet');

        let ctlNewBatch = '';
        let ctlTraceType = '';

        if (ctlIsWIP) {
          ctlNewBatch  = generateBatchId('WIP');
          ctlTraceType = 'WIP';
        } else if (ctlIsFG) {
          const fgType = getFgBatchType(srcBatch, spkType); // CTL-OUT → FSH (FC-SH)
          ctlNewBatch  = generateBatchId(fgType);
          ctlTraceType = (targetLoc === 'FG_Cust') ? 'FGC' : 'FGS';
        } else if (ctlIsSht) {
          ctlNewBatch  = generateBatchId('SHT');
          ctlTraceType = 'SHEET';
        }

        if (ctlNewBatch) {
          writeTraceLog({
            batch_id : ctlNewBatch, tgl_buat : timestamp, level : 1, type : ctlTraceType,
            source_batch : srcBatch || rootBatch, root_batch : rootBatch,
            spk_ref : payload.spk_no, gr_ref : rootBatch,
            item_code : itemCode, description : mi.desc, spec : mi.spec,
            t : mi.t, p : mi.p, l_dim : mi.l, qty : qtyAct, kg : kgAct,
            operator : op, mc_no : mcNo, tgl_prod : timestamp,
            supplier : supplierInfo.supplier, no_po : supplierInfo.no_po, no_do : supplierInfo.no_do,
            owner : owner, owner_used : ownerUsed
          });

          // Propagate batch (SHT/WIP/FG) ke SHR child kalau ada
          _propagateShtBatchToShrChildren(spkSheet, rows, headers, I, payload.spk_no, ctlNewBatch);
        }

        // Cross-billing kalau owner ≠ owner_used (FC↔DRC) DAN target_loc = FG
        // 🔒 GUARD: hanya trigger kalau hasilnya sudah jadi FG (FG_Cust / FG_RM_Stamping)
        if (owner && ownerUsed && owner !== ownerUsed && ctlIsFG) {
          writeRekapICT({
            tgl : timestamp, spk_no : payload.spk_no, item_code : itemCode,
            description : mi.desc, dari_owner : owner, ke_owner : ownerUsed,
            qty : qtyAct, kg : kgAct
          });
        }
      }

      if (spkType === 'SHR-OUT') {
        // 🟢 FIX: prefix match untuk future-proof
        // Recognize: WIP_* (Cust, Stamping, Transit_Shearing, dll) + FG_* + Stok_Sheet
        // Skip: Stok_Coil, Scrap_Area
        const isWIP = (targetLoc && targetLoc.indexOf('WIP_') === 0);
        const isFG  = (targetLoc && targetLoc.indexOf('FG_')  === 0);
        const isSht = (targetLoc === 'Stok_Sheet');
        let newBatch = '';
        let traceType = '';

        if (isWIP) {
          newBatch  = generateBatchId('WIP');
          traceType = 'WIP';
        } else if (isFG) {
          const fgType = getFgBatchType(srcBatch, spkType);
          newBatch  = generateBatchId(fgType);
          traceType = (targetLoc === 'FG_Cust') ? 'FGC' : 'FGS';
        } else if (isSht) {
          newBatch  = generateBatchId('SHT');
          traceType = 'SHEET';
        }

        if (newBatch) {
          writeTraceLog({
            batch_id : newBatch, tgl_buat : timestamp, level : 2, type : traceType,
            source_batch : srcBatch, root_batch : rootBatch, spk_ref : payload.spk_no, gr_ref : rootBatch,
            item_code : itemCode, description : mi.desc, spec : mi.spec,
            t : mi.t, p : mi.p, l_dim : mi.l, qty : qtyAct, kg : kgAct,
            operator : op, mc_no : mcNo, tgl_prod : timestamp,
            supplier : supplierInfo.supplier, no_po : supplierInfo.no_po, no_do : supplierInfo.no_do,
            owner : owner, owner_used : ownerUsed
          });
        }
        // 🔒 GUARD FG-only: hanya trigger cross-billing kalau hasilnya sudah jadi FG
        if (owner && ownerUsed && owner !== ownerUsed && isFG) {
          writeRekapICT({
            tgl : timestamp, spk_no : payload.spk_no, item_code : itemCode,
            description : mi.desc, dari_owner : owner, ke_owner : ownerUsed,
            qty : qtyAct, kg : kgAct
          });
        }
      }
    }

    // ===== 9b. TRACE_LOG UNTUK NG =====
    if (createdNgList.length > 0) {
      try {
        const ngRootBatch    = getRootBatch(parentSourceBatch) || parentSourceBatch;
        const ngSupplierInfo = getCoilSupplierInfo(ngRootBatch);
        const ngLevel        = isCtlHeader ? 1 : 2;

        createdNgList.forEach(function(ng) {
          try {
            writeTraceLog({
              batch_id : ng.spk_no, tgl_buat : timestamp, level : ngLevel, type : 'NG',
              source_batch : parentSourceBatch, root_batch : ngRootBatch,
              spk_ref : ng.spk_no, gr_ref : ngRootBatch,
              item_code : ng.item_code,
              description : (miMap[ng.item_code] && miMap[ng.item_code].desc) || '',
              spec : ng.spec || '',
              t : ng.t, p : ng.p, l_dim : ng.l, qty : ng.qty, kg : ng.kg,
              operator : '', mc_no : ng.mc_no, tgl_prod : timestamp,
              supplier : ngSupplierInfo.supplier, no_po : ngSupplierInfo.no_po, no_do : ngSupplierInfo.no_do,
              owner : ng.owner, owner_used : ng.owner
            });
          } catch (eOne) { Logger.log('NG trace skip ' + ng.spk_no + ': ' + eOne.toString()); }
        });
      } catch (eAll) { Logger.log('NG trace batch error: ' + eAll.toString()); }
    }

    // ===== 10. BUILD RESPONSE =====
    const refresh = _buildBoardRefreshFromCache(rows, headers, I, machineNo, parentSpkNo, headerDone, ss, tz);

    return {
      success : true,
      ng_created : createdNgList,
      queue : refresh.queue,
      doneToday : refresh.doneToday,
      activeChildren : refresh.activeChildren,
      headerDone : headerDone,
      activeHeaderSpkNo : refresh.activeHeaderSpkNo
    };

  } catch (e) {
    Logger.log('saveAndRefreshBoard error: ' + e.toString() + ' stack: ' + (e.stack || ''));
    return { success: false, message: e.message || String(e) };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * 🟢 T1 HELPER — _propagateShtBatchToShrChildren
 *
 * Saat CTL-OUT DONE dan menghasilkan batch SHT baru, helper ini mengisi
 * kolom Batch_ID di semua SHR-HEADER & SHR-OUT yang chain dari CTL-OUT
 * tersebut (Parent_SPK = ctlOutSpkNo untuk SHR-HEADER, lalu cascade ke
 * SHR-OUT child-nya).
 *
 * Update dilakukan di:
 *  - Sheet SPK (via setValue)
 *  - Cache `rows` (supaya response refresh konsisten)
 *
 * Tidak panik kalau Batch_ID column tidak ada / tidak ada SHR child —
 * silent return.
 * ========================================================================= */
function _propagateShtBatchToShrChildren(spkSheet, rows, headers, I, ctlOutSpkNo, newBatchSHT) {
  if (I.batch < 0) return; // Batch_ID column tidak ada di SPK
  if (!ctlOutSpkNo || !newBatchSHT) return;

  const targetParent = String(ctlOutSpkNo).trim();

  // Step 1: Find SHR-HEADER yang Parent_SPK = CTL-OUT ini
  const shrHdrSpks = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][I.type] === 'SHR-HEADER' && String(rows[i][I.parent] || '').trim() === targetParent) {
      const spk = String(rows[i][I.spk] || '').trim();
      if (spk) {
        shrHdrSpks[spk] = true;
        spkSheet.getRange(i + 1, I.batch + 1).setValue(newBatchSHT);
        rows[i][I.batch] = newBatchSHT;
      }
    }
  }

  if (Object.keys(shrHdrSpks).length === 0) return; // no SHR child

  // Step 2: Cascade ke SHR-OUT yang Parent_SPK = salah satu SHR-HEADER di atas
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][I.type] === 'SHR-OUT' && shrHdrSpks[String(rows[i][I.parent] || '').trim()]) {
      spkSheet.getRange(i + 1, I.batch + 1).setValue(newBatchSHT);
      rows[i][I.batch] = newBatchSHT;
    }
  }
}

/* =========================================================================
 * 🟢 T1 PATCH — cancelOutAndRefresh (NG Trace_Log)
 *
 * Perubahan vs versi T1 awal (patch 2b):
 *  - Setelah cancel CTL-OUT/SHR-OUT sukses dan ada NG ter-generate,
 *    tulis Trace_Log untuk setiap NG row supaya chain ke batch coil
 *    asal tetap terhubung.
 *
 * cancelCtlOutWithAuth / cancelShrOutWithAuth (function lama) TIDAK
 * DIUBAH — wrapper ini hanya menambah Trace_Log step di luar.
 * ========================================================================= */
function cancelOutAndRefresh(spkOutNo, password, ngItems, machineNo, parentHeaderSpkNo) {
  try {
    const isShr = String(machineNo || '').toUpperCase().indexOf('SHR') === 0;

    const cancelResult = isShr
      ? cancelShrOutWithAuth(spkOutNo, password, ngItems)
      : cancelCtlOutWithAuth(spkOutNo, password, ngItems);

    if (!cancelResult || !cancelResult.success) return cancelResult;

    // 🟢 T1: Write Trace_Log untuk setiap NG yang ter-generate dari cancel
    if (cancelResult.ng_created && cancelResult.ng_created.length > 0 && parentHeaderSpkNo) {
      try {
        _writeTraceLogForCancelNg(cancelResult.ng_created, isShr, parentHeaderSpkNo);
      } catch (eTrace) {
        Logger.log('Trace_Log NG (cancel) error: ' + eTrace.toString());
      }
    }

    // Build refresh dari single SPK read
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spkSheet = ss.getSheetByName('SPK');
    const rows = spkSheet.getDataRange().getValues();
    const headers = rows[0].map(function(h){ return String(h).trim(); });
    const I = {
      spk     : headers.indexOf("SPK_No"),
      type    : headers.indexOf("SPK_Type"),
      parent  : headers.indexOf("Parent_SPK"),
      status  : headers.indexOf("Status"),
      selesai : headers.indexOf("Selesai_DT"),
      mc      : headers.indexOf("MC_No"),
      prio    : headers.indexOf("Priority"),
      planSeq : headers.indexOf("Plan_Seq")
    };

    let headerDone = false;
    if (parentHeaderSpkNo) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][I.spk] || '').trim() === String(parentHeaderSpkNo).trim()) {
          const st = String(rows[i][I.status] || '').toUpperCase();
          if (st === 'DONE' || st === 'CANCELLED') headerDone = true;
          break;
        }
      }
    }

    const tz = Session.getScriptTimeZone();
    const refresh = _buildBoardRefreshFromCache(rows, headers, I, machineNo, parentHeaderSpkNo, headerDone, ss, tz);

    return Object.assign({}, cancelResult, {
      queue : refresh.queue,
      doneToday : refresh.doneToday,
      activeChildren : refresh.activeChildren,
      headerDone : headerDone,
      activeHeaderSpkNo : refresh.activeHeaderSpkNo
    });
  } catch (e) {
    Logger.log('cancelOutAndRefresh error: ' + e.toString());
    return { success: false, message: e.message || String(e) };
  }
}

/* =========================================================================
 * 🟢 T1 HELPER — _writeTraceLogForCancelNg
 *
 * Cari source batch dari parent header (CTL-HEADER/SHR-HEADER) lewat 1×
 * SPK read, lalu loop tulis Trace_Log untuk setiap NG yang ter-generate
 * saat cancel-out.
 *
 * Tidak crash kalau ada error per item — log lalu lanjut.
 * ========================================================================= */
function _writeTraceLogForCancelNg(ngList, isShr, parentHeaderSpkNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spkSheet = ss.getSheetByName('SPK');
  if (!spkSheet) return;

  const rows = spkSheet.getDataRange().getValues();
  const headers = rows[0].map(function(h){ return String(h).trim(); });
  const iSpk    = headers.indexOf('SPK_No');
  const iParent = headers.indexOf('Parent_SPK');
  const iItem   = headers.indexOf('Item_Code');

  // Lookup parent header's source batch (Parent_SPK dari parent header)
  let parentSourceBatch = '';
  let parentDesc = '';
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iSpk] || '').trim() === String(parentHeaderSpkNo).trim()) {
      parentSourceBatch = String(rows[i][iParent] || '').trim();
      break;
    }
  }

  // Lookup deskripsi item dari M_ITEM (sekali, untuk semua NG)
  const miSheet = ss.getSheetByName('M_ITEM');
  let miMap = {};
  if (miSheet && ngList.length > 0) {
    const miData = miSheet.getDataRange().getValues();
    const miHdr  = miData[0].map(function(h){ return String(h).trim(); });
    const miIc   = miHdr.indexOf('Item_Code');
    const miDesc = miHdr.indexOf('Description');
    for (let m = 1; m < miData.length; m++) {
      const k = String(miData[m][miIc] || '').trim();
      if (k) miMap[k] = String(miData[m][miDesc] || '').trim();
    }
  }

  const ngRootBatch    = (typeof getRootBatch === 'function' ? getRootBatch(parentSourceBatch) : '') || parentSourceBatch;
  const ngSupplierInfo = (typeof getCoilSupplierInfo === 'function')
                          ? getCoilSupplierInfo(ngRootBatch)
                          : { supplier:'', no_po:'', no_do:'' };
  const ngLevel  = isShr ? 2 : 1;
  const timestamp = new Date();

  ngList.forEach(function(ng) {
    try {
      writeTraceLog({
        batch_id    : ng.spk_no,
        tgl_buat    : timestamp,
        level       : ngLevel,
        type        : 'NG',
        source_batch: parentSourceBatch,
        root_batch  : ngRootBatch,
        spk_ref     : ng.spk_no,
        gr_ref      : ngRootBatch,
        item_code   : ng.item_code,
        description : miMap[ng.item_code] || '',
        spec        : ng.spec || '',
        t           : ng.t, p : ng.p, l_dim : ng.l,
        qty         : ng.qty, kg : ng.kg,
        operator    : '',
        mc_no       : ng.mc_no,
        tgl_prod    : timestamp,
        supplier    : ngSupplierInfo.supplier,
        no_po       : ngSupplierInfo.no_po,
        no_do       : ngSupplierInfo.no_do,
        owner       : ng.owner,
        owner_used  : ng.owner
      });
    } catch (e) {
      Logger.log('Trace_Log cancel-NG write skip ' + ng.spk_no + ': ' + e.toString());
    }
  });
}

/* =========================================================================
 * INTERNAL HELPER — Build queue + doneToday + activeChildren dari cache rows
 * Tidak melakukan re-read SPK sheet. Dipakai oleh saveAndRefreshBoard
 * dan cancelOutAndRefresh.
 * ========================================================================= */
function _buildBoardRefreshFromCache(rows, headers, I, machineNo, parentSpkNo, headerDone, ss, tz) {
  const mcUpper       = String(machineNo || '').toUpperCase();
  const isShrMachine  = mcUpper.indexOf('SHR') === 0;
  const headerType    = isShrMachine ? 'SHR-HEADER' : 'CTL-HEADER';
  const outType       = isShrMachine ? 'SHR-OUT'    : 'CTL-OUT';

  const now = new Date();
  const todayStart    = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrowStart = new Date(todayStart.getTime() + 24*60*60*1000);

  // Build Stok_Coil map hanya untuk CTL (untuk KG_Fisik & Mat_Spec)
  let stokCoilMap = {};
  if (!isShrMachine) {
    const stkSheet = ss.getSheetByName('Stok_Coil');
    if (stkSheet) {
      const stkData = stkSheet.getDataRange().getValues();
      if (stkData.length > 1) {
        const stkHdr = stkData[0].map(function(h){ return String(h).trim(); });
        const sB  = stkHdr.indexOf('Batch_ID');
        const sKi = stkHdr.indexOf('KG_In');
        const sKd = stkHdr.indexOf('KG_Done');
        const sSp = stkHdr.indexOf('Spec');
        if (sB > -1) {
          for (let i = 1; i < stkData.length; i++) {
            const kgIn   = parseFloat(stkData[i][sKi]) || 0;
            const kgDone = parseFloat(stkData[i][sKd]) || 0;
            stokCoilMap[String(stkData[i][sB]).trim()] = {
              kg_fisik: kgIn - kgDone,
              spec    : sSp > -1 ? String(stkData[i][sSp]).trim() : ''
            };
          }
        }
      }
    }
  }

  function rowToObj(row) {
    const obj = {};
    headers.forEach(function(h, idx) {
      let val = row[idx];
      if (val instanceof Date) val = Utilities.formatDate(val, tz, 'dd-MM-yyyy HH:mm:ss');
      obj[h] = val;
    });
    return obj;
  }

  let queue = [];
  let doneToday = [];
  let activeChildren = [];
  let activeHeaderSpkNo = null;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const t  = String(row[I.type] || '').trim();
    const mc = String(row[I.mc]   || '').trim().toUpperCase();
    const st = String(row[I.status] || '').trim();

    // Header rows untuk mesin ini
    if (t === headerType && mc === mcUpper) {
      if (st === 'ANTRIAN' || st === 'RUNNING') {
        const obj = rowToObj(row);
        if (!isShrMachine) {
          const coilData = stokCoilMap[String(row[I.parent]).trim()] || { kg_fisik: 0, spec: '' };
          obj['KG_Fisik'] = coilData.kg_fisik;
          obj['Mat_Spec'] = coilData.spec;
        }
        queue.push(obj);
        if (st === 'RUNNING') activeHeaderSpkNo = String(row[I.spk]).trim();
      } else if (st === 'DONE') {
        const sel = row[I.selesai];
        if (sel instanceof Date && sel >= todayStart && sel < tomorrowStart) {
          doneToday.push(rowToObj(row));
        }
      }
    }

    // Active children (kalau parent masih running)
    // 🟢 BUG-FIX #4 — Override BQ supaya konsisten dgn fetchBoardChildData
    //    BQ di sheet untuk CTL-OUT = 1 (placeholder), board butuh kg/lbr asli
    //    untuk auto-calc KG = qty × BQ di modal.
    if (!headerDone && parentSpkNo && t === outType
        && String(row[I.parent] || '').trim() === parentSpkNo) {
      const childObj = rowToObj(row);
      const tQty = parseFloat(childObj['Qty_Target']) || 1;
      const tKg  = parseFloat(childObj['KG_Target']) || 0;
      childObj['BQ'] = tKg / tQty;
      activeChildren.push(childObj);
    }
  }

  // Sort queue by priority + Plan_Seq
  const prioWeight = { 'Urgent': 3, 'High': 2, 'Normal': 1 };
  queue.sort(function(a, b) {
    const d = (prioWeight[b.Priority] || 1) - (prioWeight[a.Priority] || 1);
    if (d !== 0) return d;
    return (parseFloat(a.Plan_Seq) || 999) - (parseFloat(b.Plan_Seq) || 999);
  });

  // Sort doneToday by Selesai_DT desc
  doneToday.sort(function(a, b) {
    return String(b.Selesai_DT || '').localeCompare(String(a.Selesai_DT || ''));
  });

  return {
    queue : queue,
    doneToday : doneToday,
    activeChildren : activeChildren,
    activeHeaderSpkNo : activeHeaderSpkNo
  };
}

/* =========================================================================
 * 🟢 BUG-FIX #2 — finalizeShrBatch (SHR All-in-One Batch Finalize)
 *
 * Function baru terpisah dari saveAndRefreshBoard (yang masih dipakai CTL).
 * Handle 1-shot submit: semua SHR-OUT + SHR-HEADER selesai sekaligus dari
 * 1 modal komprehensif di board SHR.
 *
 * Payload structure:
 *   {
 *     header_spk : 'SHR-260001',     // SPK_No SHR-HEADER
 *     qty_used   : 10,               // lembar input diproses (operator input)
 *     outs       : [
 *       { spk_no: 'SHR-260001-01', qty_actual: 58 },
 *       { spk_no: 'SHR-260001-02', qty_actual: 18 }
 *     ],
 *     ng_items   : [...],            // optional, same structure as saveAndRefreshBoard
 *     created_by : 'operator'
 *   }
 *
 * Behavior:
 *   - All OUTs updated to DONE in single transaction
 *   - HEADER finalized: Qty_Actual=qty_used, KG_Actual=qty_used*BQ_HEADER
 *   - kg_actual per OUT auto-calc backend = qty_actual × (KG_Target/Qty_Target)
 *   - NG rows appended dengan Batch_ID = ngSpkNo (Bug 3 consistent)
 *   - Returns same response shape as saveAndRefreshBoard (queue, doneToday, dst)
 * ========================================================================= */
function finalizeShrBatch(payload, machineNo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // ===== 1. VALIDASI PAYLOAD =====
    if (!payload || !payload.header_spk) {
      return { success: false, message: 'header_spk kosong' };
    }
    const headerSpk = String(payload.header_spk).trim();
    const qtyUsed   = parseFloat(payload.qty_used) || 0;
    if (qtyUsed <= 0) {
      return { success: false, message: 'Qty_Used wajib > 0 (poka-yoke)' };
    }
    if (!Array.isArray(payload.outs) || payload.outs.length === 0) {
      return { success: false, message: 'Tidak ada OUT yang di-finalize' };
    }

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const spkSheet  = ss.getSheetByName("SPK");
    if (!spkSheet) return { success: false, message: 'Sheet SPK tidak ditemukan' };
    const timestamp = new Date();
    const tz        = Session.getScriptTimeZone();

    // ===== 2. READ SPK 1× + COLUMN MAP =====
    let rows = spkSheet.getDataRange().getValues();
    const headers = rows[0].map(function(h){ return String(h).trim(); });
    const I = {
      spk     : headers.indexOf("SPK_No"),
      type    : headers.indexOf("SPK_Type"),
      parent  : headers.indexOf("Parent_SPK"),
      status  : headers.indexOf("Status"),
      mulai   : headers.indexOf("Mulai_DT"),
      selesai : headers.indexOf("Selesai_DT"),
      qtyAct  : headers.indexOf("Qty_Actual"),
      kgAct   : headers.indexOf("KG_Actual"),
      qtyNg   : headers.indexOf("Qty_NG"),
      kgNg    : headers.indexOf("KG_NG"),
      item    : headers.indexOf("Item_Code"),
      qtyTgt  : headers.indexOf("Qty_Target"),
      kgTgt   : headers.indexOf("KG_Target"),
      mc      : headers.indexOf("MC_No"),
      owner   : headers.indexOf("Owner"),
      tgtLoc  : headers.indexOf("Target_Loc"),
      // 🟢 Fix Trace_Log — extra fields untuk build payload writeTraceLog
      batch   : headers.indexOf("Batch_ID"),
      ownerU  : headers.indexOf("Owner_Used"),
      soRef   : headers.indexOf("SO_Ref"),
      cust    : headers.indexOf("Cust"),
      op      : headers.indexOf("OP")
    };

    // ===== 3. FIND HEADER ROW =====
    let headerRowIdx = -1, headerArrIdx = -1;
    let parentQtyTgt = 0, parentKgTgt = 0;
    let parentMcNo = '', parentOwner = '', parentItemCode = '';
    let parentMulai = null;
    // 🟢 Fix Trace_Log — parent context untuk writeTraceLog
    let parentBatchId = '';

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][I.spk]).trim() === headerSpk
          && String(rows[i][I.type]).trim() === 'SHR-HEADER') {
        headerRowIdx    = i + 1;
        headerArrIdx    = i;
        parentQtyTgt    = Number(rows[i][I.qtyTgt]) || 0;
        parentKgTgt     = Number(rows[i][I.kgTgt])  || 0;
        parentMcNo      = String(rows[i][I.mc]      || '').trim();
        parentOwner     = String(rows[i][I.owner]   || '').trim();
        parentItemCode  = String(rows[i][I.item]    || '').trim();
        parentMulai     = rows[i][I.mulai] instanceof Date ? rows[i][I.mulai] : null;
        parentBatchId   = I.batch > -1 ? String(rows[i][I.batch] || '').trim() : '';
        const hdrStatus = String(rows[i][I.status] || '').toUpperCase();
        if (hdrStatus === 'DONE') {
          return { success: false, message: 'SHR-HEADER sudah DONE: ' + headerSpk };
        }
        if (hdrStatus === 'CANCELLED') {
          return { success: false, message: 'SHR-HEADER sudah CANCELLED: ' + headerSpk };
        }
        break;
      }
    }
    if (headerRowIdx === -1) {
      return { success: false, message: 'SHR-HEADER tidak ditemukan: ' + headerSpk };
    }

    // ===== 4. FIND ALL SHR-OUT (parent = headerSpk) =====
    const outMap = {}; // spk_no → {rowIdx, arrIdx, qtyTgt, kgTgt, status}
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][I.parent]).trim() === headerSpk
          && String(rows[i][I.type]).trim() === 'SHR-OUT') {
        outMap[String(rows[i][I.spk]).trim()] = {
          rowIdx : i + 1,
          arrIdx : i,
          qtyTgt : Number(rows[i][I.qtyTgt]) || 0,
          kgTgt  : Number(rows[i][I.kgTgt])  || 0,
          status : String(rows[i][I.status] || '').toUpperCase()
        };
      }
    }

    // ===== 5. VALIDASI PAYLOAD vs SHEET =====
    // a. Semua OUT di payload harus ada di sheet, dan belum DONE/CANCELLED
    for (let p = 0; p < payload.outs.length; p++) {
      const oSpk = String(payload.outs[p].spk_no || '').trim();
      const oQty = parseFloat(payload.outs[p].qty_actual) || 0;
      if (!oSpk) return { success: false, message: 'spk_no OUT kosong di index ' + p };
      if (oQty <= 0) return { success: false, message: 'qty_actual OUT ' + oSpk + ' harus > 0' };
      if (!outMap[oSpk]) {
        return { success: false, message: 'SHR-OUT tidak ditemukan: ' + oSpk };
      }
      if (outMap[oSpk].status === 'DONE' || outMap[oSpk].status === 'CANCELLED') {
        return { success: false, message: 'SHR-OUT ' + oSpk + ' sudah ' + outMap[oSpk].status };
      }
    }
    // b. Semua OUT non-DONE/CANCELLED di sheet harus ada di payload (poka-yoke)
    const payloadSpkSet = {};
    payload.outs.forEach(function(o){ payloadSpkSet[String(o.spk_no).trim()] = true; });
    const missingOuts = [];
    Object.keys(outMap).forEach(function(spk){
      if (outMap[spk].status !== 'DONE' && outMap[spk].status !== 'CANCELLED'
          && !payloadSpkSet[spk]) {
        missingOuts.push(spk);
      }
    });
    if (missingOuts.length > 0) {
      return { success: false, message: 'OUT belum di-finalize: ' + missingOuts.join(', ') };
    }

    // ===== 6. UPDATE EACH OUT ROW (DONE) =====
    payload.outs.forEach(function(o) {
      const oSpk = String(o.spk_no).trim();
      const oQty = parseFloat(o.qty_actual) || 0;
      const info = outMap[oSpk];
      const bqOut = info.qtyTgt > 0 ? (info.kgTgt / info.qtyTgt) : 0;
      const oKg   = Math.round(oQty * bqOut);

      spkSheet.getRange(info.rowIdx, I.qtyAct  + 1).setValue(oQty);
      spkSheet.getRange(info.rowIdx, I.kgAct   + 1).setValue(oKg);
      spkSheet.getRange(info.rowIdx, I.status  + 1).setValue('DONE');
      spkSheet.getRange(info.rowIdx, I.selesai + 1).setValue(timestamp);
      if (!(rows[info.arrIdx][I.mulai] instanceof Date)) {
        spkSheet.getRange(info.rowIdx, I.mulai + 1).setValue(parentMulai || timestamp);
      }

      rows[info.arrIdx][I.qtyAct]  = oQty;
      rows[info.arrIdx][I.kgAct]   = oKg;
      rows[info.arrIdx][I.status]  = 'DONE';
      rows[info.arrIdx][I.selesai] = timestamp;
    });

    // ===== 7. APPEND NG ROWS (kalau ada) =====
    let totalNgQty = 0, totalNgKg = 0;
    const createdNgList = [];

    if (Array.isArray(payload.ng_items) && payload.ng_items.length > 0) {
      // M_ITEM map (untuk derive spec NG)
      const miSheet = ss.getSheetByName('M_ITEM');
      let parentSpec = '';
      if (miSheet) {
        const miData = miSheet.getDataRange().getValues();
        const miHdr  = miData[0].map(function(h){ return String(h).trim(); });
        const miIc   = miHdr.indexOf('Item_Code');
        const miSpec = miHdr.indexOf('Spec');
        for (let m = 1; m < miData.length; m++) {
          if (String(miData[m][miIc] || '').trim() === parentItemCode) {
            parentSpec = String(miData[m][miSpec] || '').trim();
            break;
          }
        }
      }

      // Cari max NG suffix existing
      let maxNgSuffix = 0;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][I.type] === 'SHR-NG' && String(rows[i][I.parent]).trim() === headerSpk) {
          const m = String(rows[i][I.spk]).match(/-NG(\d+)$/);
          if (m) { const n = parseInt(m[1], 10) || 0; if (n > maxNgSuffix) maxNgSuffix = n; }
        }
      }

      // Cari base row count
      let baseRowCount = 1;
      for (let r = rows.length - 1; r >= 1; r--) {
        if (rows[r][I.spk] && String(rows[r][I.spk]).trim() !== '') {
          baseRowCount = r + 1; break;
        }
      }

      const ngRowsToAppend = [];
      payload.ng_items.forEach(function(ng) {
        const ngT = parseFloat(ng.t)     || 0;
        const ngP = parseFloat(ng.p)     || 0;
        const ngL = parseFloat(ng.l)     || 0;
        const ngQ = parseInt(ng.qty, 10) || 0;
        if (ngT <= 0 || ngP <= 0 || ngL <= 0 || ngQ <= 0) return;

        maxNgSuffix++;
        const ngSpkNo = headerSpk + '-NG' + String(maxNgSuffix).padStart(2, '0');
        const ngKg    = Math.round(ngT * ngP * ngL * ngQ * 7.85 / 1000000);

        const ngData = {
          'SPK_No'     : ngSpkNo,
          'SPK_Type'   : 'SHR-NG',
          'Parent_SPK' : headerSpk,
          'Batch_ID'   : ngSpkNo,            // 🟢 Bug 3 consistent (NG self-batch)
          'Tgl_Buat'   : timestamp,
          'Item_Code'  : parentItemCode,
          'Input_Spec' : ngT + ' x ' + ngP + ' x ' + ngL,
          'Qty_Target' : ngQ,
          'KG_Target'  : ngKg,
          'Qty_Actual' : ngQ,
          'KG_Actual'  : ngKg,
          'MC_No'      : parentMcNo,
          'Owner'      : parentOwner,
          'Owner_Used' : parentOwner,
          'Status'     : 'DONE',
          'Target_Loc' : 'Stok_NG',
          'Mulai_DT'   : timestamp,
          'Selesai_DT' : timestamp,
          'Created_By' : payload.created_by || 'operator',
          'T'          : ngT, 'P': ngP, 'L': ngL
        };

        const ngRow = headers.map(function(h){ return ngData[h] !== undefined ? ngData[h] : ''; });
        ngRowsToAppend.push(ngRow);
        rows.push(ngRow.slice());

        totalNgQty += ngQ;
        totalNgKg  += ngKg;
        createdNgList.push({
          spk_no : ngSpkNo,
          spec   : parentSpec || (ngT + 'x' + ngP + 'x' + ngL),
          t      : ngT, p: ngP, l: ngL,
          qty    : ngQ, kg: ngKg,
          owner  : parentOwner, mc_no: parentMcNo
        });
      });

      if (ngRowsToAppend.length > 0) {
        const startRow = baseRowCount + 1;
        spkSheet.getRange(startRow, 1, ngRowsToAppend.length, headers.length).setValues(ngRowsToAppend);
      }
    }

    // ===== 8. FINALIZE HEADER ROW =====
    const bqHeader = parentQtyTgt > 0 ? (parentKgTgt / parentQtyTgt) : 0;
    const finalKgAct = Math.round(qtyUsed * bqHeader);

    spkSheet.getRange(headerRowIdx, I.qtyAct  + 1).setValue(qtyUsed);
    spkSheet.getRange(headerRowIdx, I.kgAct   + 1).setValue(finalKgAct);
    spkSheet.getRange(headerRowIdx, I.status  + 1).setValue('DONE');
    spkSheet.getRange(headerRowIdx, I.selesai + 1).setValue(timestamp);
    if (!(rows[headerArrIdx][I.mulai] instanceof Date)) {
      spkSheet.getRange(headerRowIdx, I.mulai + 1).setValue(parentMulai || timestamp);
    }
    if (totalNgQty > 0 || totalNgKg > 0) {
      const curHeaderQtyNg = parseFloat(rows[headerArrIdx][I.qtyNg]) || 0;
      const curHeaderKgNg  = parseFloat(rows[headerArrIdx][I.kgNg])  || 0;
      spkSheet.getRange(headerRowIdx, I.qtyNg + 1).setValue(curHeaderQtyNg + totalNgQty);
      spkSheet.getRange(headerRowIdx, I.kgNg  + 1).setValue(Math.round(curHeaderKgNg + totalNgKg));
    }

    rows[headerArrIdx][I.qtyAct]  = qtyUsed;
    rows[headerArrIdx][I.kgAct]   = finalKgAct;
    rows[headerArrIdx][I.status]  = 'DONE';
    rows[headerArrIdx][I.selesai] = timestamp;

    // ===== 8.5. WRITE TRACE_LOG (untuk setiap SHR-OUT DONE + NG) =====
    // 🟢 Bug fix — sebelumnya finalizeShrBatch skip writeTraceLog
    // Prefix match untuk WIP_/FG_ (future-proof)
    try {
      // Ambil desc/spec/t/p/l dari M_ITEM (sekali baca)
      const miMap = {};
      const miSheet = ss.getSheetByName('M_ITEM');
      if (miSheet) {
        const miData = miSheet.getDataRange().getValues();
        const miHdr  = miData[0].map(function(h){ return String(h).trim(); });
        const miIc   = miHdr.indexOf('Item_Code');
        const miDesc = miHdr.indexOf('Description');
        const miSpec = miHdr.indexOf('Spec');
        const miT    = miHdr.indexOf('T');
        const miP    = miHdr.indexOf('P');
        const miL    = miHdr.indexOf('L');
        for (let m = 1; m < miData.length; m++) {
          const ic = String(miData[m][miIc] || '').trim();
          if (ic) {
            miMap[ic] = {
              desc: String(miData[m][miDesc] || '').trim(),
              spec: String(miData[m][miSpec] || '').trim(),
              t   : miData[m][miT] || '',
              p   : miData[m][miP] || '',
              l   : miData[m][miL] || ''
            };
          }
        }
      }

      const rootBatch    = (typeof getRootBatch === 'function') ? (getRootBatch(parentBatchId) || parentBatchId) : parentBatchId;
      const supplierInfo = (typeof getCoilSupplierInfo === 'function') ? getCoilSupplierInfo(rootBatch) : { supplier:'', no_po:'', no_do:'' };

      // Loop each SHR-OUT DONE
      payload.outs.forEach(function(o) {
        try {
          const oSpk = String(o.spk_no).trim();
          const info = outMap[oSpk];
          if (!info) return;

          const outRow    = rows[info.arrIdx];
          const targetLoc = I.tgtLoc > -1 ? String(outRow[I.tgtLoc] || '').trim() : '';
          const itemCode  = I.item   > -1 ? String(outRow[I.item]   || '').trim() : '';
          const owner     = I.owner  > -1 ? String(outRow[I.owner]  || '').trim() : parentOwner;
          const ownerUsed = I.ownerU > -1 ? String(outRow[I.ownerU] || '').trim() : owner;
          const mcNo      = I.mc     > -1 ? String(outRow[I.mc]     || '').trim() : parentMcNo;
          const op        = I.op     > -1 ? String(outRow[I.op]     || '').trim() : '';
          const qtyAct    = parseFloat(o.qty_actual) || 0;
          const kgAct     = parseFloat(outRow[I.kgAct]) || 0;

          // Prefix match target_loc → batch type
          const isWIP = (targetLoc && targetLoc.indexOf('WIP_') === 0);
          const isFG  = (targetLoc && targetLoc.indexOf('FG_')  === 0);
          const isSht = (targetLoc === 'Stok_Sheet');

          if (!isWIP && !isFG && !isSht) return;  // skip Scrap/Stok_Coil/lainnya

          let newBatch = '';
          let traceType = '';
          if (isWIP) {
            newBatch  = generateBatchId('WIP');
            traceType = 'WIP';
          } else if (isFG) {
            const fgType = (typeof getFgBatchType === 'function') ? getFgBatchType(parentBatchId, 'SHR-OUT') : 'FGC';
            newBatch  = generateBatchId(fgType);
            traceType = (targetLoc === 'FG_Cust') ? 'FGC' : 'FGS';
          } else if (isSht) {
            newBatch  = generateBatchId('SHT');
            traceType = 'SHEET';
          }
          if (!newBatch) return;

          const mi = miMap[itemCode] || { desc:'', spec:'', t:'', p:'', l:'' };

          writeTraceLog({
            batch_id     : newBatch,
            tgl_buat     : timestamp,
            level        : 2,
            type         : traceType,
            source_batch : parentBatchId,
            root_batch   : rootBatch,
            spk_ref      : oSpk,
            gr_ref       : rootBatch,
            item_code    : itemCode,
            description  : mi.desc,
            spec         : mi.spec,
            t            : mi.t, p: mi.p, l_dim: mi.l,
            qty          : qtyAct, kg: kgAct,
            operator     : op, mc_no: mcNo, tgl_prod: timestamp,
            supplier     : supplierInfo.supplier,
            no_po        : supplierInfo.no_po,
            no_do        : supplierInfo.no_do,
            owner        : owner, owner_used: ownerUsed
          });

          // Cross-billing Rekap_ICT (Owner ≠ Owner_Used) — 🔒 GUARD FG-only
          // hanya trigger kalau hasilnya sudah jadi FG (FG_Cust / FG_RM_Stamping)
          if (owner && ownerUsed && owner !== ownerUsed && isFG && typeof writeRekapICT === 'function') {
            writeRekapICT({
              tgl: timestamp, spk_no: oSpk, item_code: itemCode,
              description: mi.desc, dari_owner: owner, ke_owner: ownerUsed,
              qty: qtyAct, kg: kgAct
            });
          }

          // 🟢 Propagate batch baru ke SHR-HEADER Stage 2 (kalau ada)
          if (typeof _propagateShtBatchToShrChildren === 'function') {
            try {
              _propagateShtBatchToShrChildren(spkSheet, rows, headers, I, oSpk, newBatch);
            } catch (eProp) {
              Logger.log('Stage2 propagate skip ' + oSpk + ': ' + eProp);
            }
          }
        } catch (eOut) {
          Logger.log('Trace_Log OUT skip ' + o.spk_no + ': ' + eOut);
        }
      });

      // Loop NG rows (kalau ada)
      if (createdNgList.length > 0) {
        createdNgList.forEach(function(ng) {
          try {
            const miNg = miMap[parentItemCode] || { desc:'', spec:'', t:'', p:'', l:'' };
            writeTraceLog({
              batch_id     : ng.spk_no, tgl_buat: timestamp, level: 2, type: 'NG',
              source_batch : parentBatchId, root_batch: rootBatch,
              spk_ref      : ng.spk_no, gr_ref: rootBatch,
              item_code    : parentItemCode,
              description  : miNg.desc, spec: ng.spec || miNg.spec,
              t            : ng.t, p: ng.p, l_dim: ng.l,
              qty          : ng.qty, kg: ng.kg,
              operator     : '', mc_no: ng.mc_no, tgl_prod: timestamp,
              supplier     : supplierInfo.supplier,
              no_po        : supplierInfo.no_po,
              no_do        : supplierInfo.no_do,
              owner        : ng.owner, owner_used: ng.owner
            });
          } catch (eNg) {
            Logger.log('Trace_Log NG skip ' + ng.spk_no + ': ' + eNg);
          }
        });
      }
    } catch (eTrace) {
      Logger.log('Trace_Log section error: ' + eTrace + ' | stack: ' + (eTrace.stack || ''));
    }

    // ===== 9. BUILD REFRESH RESPONSE =====
    const refresh = _buildBoardRefreshFromCache(rows, headers, I, machineNo, headerSpk, true, ss, tz);

    return {
      success           : true,
      message           : 'SHR-HEADER ' + headerSpk + ' selesai (semua OUT DONE)',
      headerDone        : true,
      queue             : refresh.queue,
      doneToday         : refresh.doneToday,
      activeChildren    : refresh.activeChildren,
      ng_created        : createdNgList,
      kg_actual_header  : finalKgAct,
      qty_actual_header : qtyUsed
    };

  } catch (e) {
    Logger.log('finalizeShrBatch error: ' + e.toString() + ' stack: ' + (e.stack || ''));
    return { success: false, message: e.message || String(e) };
  } finally {
    lock.releaseLock();
  }
}
/* =========================================================================
 * fetchBoardBundleSHR — TAMBAHAN 2026-07-08
 * ------------------------------------------------------------------------
 * Bulk-load data untuk Board SHR dalam 1 backend call:
 *   - queue     : SHR-HEADER (ANTRIAN + RUNNING) untuk machineNo
 *   - done      : SHR-HEADER (DONE hari ini) untuk machineNo
 *   - Setiap HEADER punya property _outs = list SHR-OUT children (preloaded)
 *
 * Menggantikan pola frontend lama yang panggil 2 backend call terpisah
 * (fetchBoardQueueDataSHR + fetchTodayDoneJobs) lalu N call fetchBoardChildData.
 *
 * READ-ONLY, tidak modify sheet. Aman untuk deploy tanpa risiko regresi.
 * Fungsi existing (fetchBoardQueueDataSHR, fetchTodayDoneJobs,
 * fetchBoardChildData) TIDAK dimodifikasi — masih dipakai board CTL dan
 * flow finalize existing.
 * ========================================================================= */
function fetchBoardBundleSHR(machineNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SPK");
  if (!sheet) throw new Error("Sheet SPK not found");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(function(h) { return String(h).trim(); });

  const iType    = headers.indexOf("SPK_Type");
  const iMc      = headers.indexOf("MC_No");
  const iStatus  = headers.indexOf("Status");
  const iParent  = headers.indexOf("Parent_SPK");
  const iSpkNo   = headers.indexOf("SPK_No");
  const iSelesai = headers.indexOf("Selesai_DT");

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "dd-MM-yyyy");
  const mcUp = String(machineNo).trim().toUpperCase();

  const queue = [];
  const done = [];
  const headerMap = {};   // spkNo → objek HEADER (reference)
  const outCache = [];    // temp: {parent, rowIdx}

  function toObj(rowIdx) {
    const obj = {};
    headers.forEach(function(h, idx){
      let val = rows[rowIdx][idx];
      if (val instanceof Date) val = Utilities.formatDate(val, tz, "dd-MM-yyyy HH:mm:ss");
      obj[h] = val;
    });
    return obj;
  }

  // Pass 1: kumpulkan HEADER match mesin ini + cache SHR-OUT
  for (let i = 1; i < rows.length; i++) {
    const spkType = rows[i][iType];

    if (spkType === 'SHR-HEADER' &&
        String(rows[i][iMc]).trim().toUpperCase() === mcUp) {
      const status = rows[i][iStatus];
      const spkNo = String(rows[i][iSpkNo]).trim();

      if (status === 'ANTRIAN' || status === 'RUNNING') {
        const obj = toObj(i);
        obj._outs = [];
        queue.push(obj);
        headerMap[spkNo] = obj;
      } else if (status === 'DONE') {
        const selesaiDt = rows[i][iSelesai];
        const selesaiDate = selesaiDt instanceof Date
          ? Utilities.formatDate(selesaiDt, tz, "dd-MM-yyyy") : '';
        if (selesaiDate === today) {
          const obj = toObj(i);
          obj._outs = [];
          done.push(obj);
          headerMap[spkNo] = obj;
        }
      }
    } else if (spkType === 'SHR-OUT') {
      outCache.push({ parent: String(rows[i][iParent]).trim(), rowIdx: i });
    }
  }

  // Pass 2: attach SHR-OUT ke HEADER
  outCache.forEach(function(oc){
    const hdrObj = headerMap[oc.parent];
    if (!hdrObj) return;
    const outObj = toObj(oc.rowIdx);
    const tQty = parseFloat(outObj['Qty_Target']) || 1;
    const tKg  = parseFloat(outObj['KG_Target'])  || 0;
    outObj['BQ'] = tKg / tQty;
    hdrObj._outs.push(outObj);
  });

  // Sort queue by priority (Urgent → High → Normal), lalu Plan_Seq
  const prioWeight = { 'Urgent': 3, 'High': 2, 'Normal': 1 };
  queue.sort(function(a, b) {
    const d = (prioWeight[b.Priority] || 1) - (prioWeight[a.Priority] || 1);
    if (d !== 0) return d;
    return (parseFloat(a.Plan_Seq) || 999) - (parseFloat(b.Plan_Seq) || 999);
  });

  return { queue: queue, done: done };
}

/* ============================================================================
 * REPRINT LABEL — v1.2 (fix: batch & lot pull from Trace_Log, not SPK)
 * ========================================================================= */
function getReprintLabelData(parentSpkNo) {
  try {
    if (!parentSpkNo) return { ok: false, msg: 'SPK No kosong' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('SPK');
    if (!sh) return { ok: false, msg: 'Sheet SPK tidak ditemukan' };

    var data = sh.getDataRange().getValues();
    var header = data[0];
    var col = {};
    header.forEach(function(h, i){ col[h] = i; });

    // ── Build M_ITEM lookup ────────────────────────────────────────────
    var mItemMap = {};
    var mSh = ss.getSheetByName('M_ITEM');
    if (mSh) {
      var mData = mSh.getDataRange().getValues();
      var mCol = {};
      mData[0].forEach(function(h, i){ mCol[h] = i; });
      for (var mi = 1; mi < mData.length; mi++) {
        var code = String(mData[mi][mCol.Item_Code] || '').trim();
        if (!code) continue;
        mItemMap[code] = {
          spec: String(mData[mi][mCol.Spec] || '').trim(),
          t:    Number(mData[mi][mCol.T]) || 0,
          p:    mData[mi][mCol.P] || '',
          l:    mData[mi][mCol.L] || ''
        };
      }
    }

    // ── Build Trace_Log lookup by SPK_Ref ──────────────────────────────
    // Map: SPK_Ref → { batch_id, source_batch, root_batch, tgl_prod }
    var traceMap = {};
    var tSh = ss.getSheetByName('Trace_Log');
    if (tSh && tSh.getLastRow() > 1) {
      var tData = tSh.getDataRange().getValues();
      var tCol = {};
      tData[0].forEach(function(h, i){ tCol[String(h).trim()] = i; });
      for (var ti = 1; ti < tData.length; ti++) {
        var spkRef = String(tData[ti][tCol.SPK_Ref] || '').trim();
        if (!spkRef) continue;
        // Kalau ada multiple entry per SPK_Ref, ambil yang pertama (biasanya 1 SPK OUT = 1 entry)
        if (traceMap[spkRef]) continue;
        traceMap[spkRef] = {
          batch_id:     String(tData[ti][tCol.Batch_ID] || '').trim(),
          source_batch: String(tData[ti][tCol.Source_Batch] || '').trim(),
          root_batch:   String(tData[ti][tCol.Root_Batch] || '').trim(),
          tgl_prod:     tData[ti][tCol.Tgl_Prod] || tData[ti][tCol.Tgl_Buat] || null
        };
      }
    }

    // ── Cari HEADER row ────────────────────────────────────────────────
    var headerRow = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col.SPK_No]) === String(parentSpkNo)) {
        headerRow = data[i]; break;
      }
    }
    if (!headerRow) return { ok: false, msg: 'SPK ' + parentSpkNo + ' tidak ditemukan' };

    var headerSpkType = String(headerRow[col.SPK_Type] || '');
    if (headerSpkType.indexOf('HEADER') < 0) {
      return { ok: false, msg: 'SPK ' + parentSpkNo + ' bukan HEADER (type: ' + headerSpkType + ')' };
    }

    // ── Cari semua OUT ─────────────────────────────────────────────────
    var outRows = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col.Parent_SPK]) === String(parentSpkNo)
          && String(data[i][col.SPK_Type] || '').indexOf('OUT') >= 0) {
        outRows.push(data[i]);
      }
    }

    // ── Mapping WIP: CTL-OUT → SHR-HEADER → SHR-OUT ────────────────────
    var shrHeaderByParent = {};
    var shrOutByParent    = {};
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var t = String(r[col.SPK_Type] || '');
      if (t === 'SHR-HEADER') {
        var p = String(r[col.Parent_SPK] || '');
        if (!shrHeaderByParent[p]) shrHeaderByParent[p] = [];
        shrHeaderByParent[p].push(String(r[col.SPK_No]));
      } else if (t === 'SHR-OUT') {
        var pp = String(r[col.Parent_SPK] || '');
        if (!shrOutByParent[pp]) shrOutByParent[pp] = [];
        shrOutByParent[pp].push(r);
      }
    }

    // ── Helpers ────────────────────────────────────────────────────────
    function _labelTypeFromTargetLoc(tgt) {
      var s = String(tgt || '').trim();
      if (s === 'FG_Cust')         return 'FG';
      if (s === 'FG_RM_Stamping')  return 'FG_STP';
      if (s === 'WIP_Cust')        return 'WIP';
      if (s === 'WIP_Stamping')    return 'WIP_STP';
      if (s === 'Stok_Sheet')      return 'STOK_SHT';
      return 'FG';
    }

    function _parseSpecPure(rawSpec) {
      if (!rawSpec) return '';
      var s = String(rawSpec).trim();
      var m = s.match(/^([A-Za-z][A-Za-z0-9\-\/\.]*(?:\s[A-Za-z][A-Za-z0-9\-\/\.]*)?)/);
      if (m) return m[1].trim();
      var sp = s.indexOf(' ');
      return sp > 0 ? s.substring(0, sp) : s;
    }

    function _resolveFields(row) {
      var itemCode = String(row[col.Item_Code] || '').trim();
      var mi = mItemMap[itemCode] || {};
      var rawSpec = String(row[col.Input_Spec] || '').trim();
      var spec = mi.spec || _parseSpecPure(rawSpec);
      var t = mi.t || Number(row[col.T]) || 0;
      var p = mi.p || row[col.P] || '';
      var l = mi.l || row[col.L] || '';
      return { itemCode: itemCode, spec: spec, t: t, p: p, l: l };
    }

    function _getShrChildrenForCtlOut(ctlOutSpkNo) {
      var result = [];
      var shrHeaders = shrHeaderByParent[ctlOutSpkNo] || [];
      shrHeaders.forEach(function(shrHdrSpkNo){
        var childOuts = shrOutByParent[shrHdrSpkNo] || [];
        childOuts.forEach(function(r){
          var st = String(r[col.Status] || '').toUpperCase();
          if (st === 'CANCELLED') return;
          var f = _resolveFields(r);
          result.push({
            target_loc: String(r[col.Target_Loc] || ''),
            p:          f.p,
            l:          f.l,
            qty:        Number(r[col.Qty_Target] || 0),
            cust:       String(r[col.Cust] || '')
          });
        });
      });
      return result;
    }

    // ── Build OUT array ────────────────────────────────────────────────
    var outs = outRows.map(function(r){
      var spkNoOut  = String(r[col.SPK_No] || '');
      var targetLoc = String(r[col.Target_Loc] || '');
      var labelType = _labelTypeFromTargetLoc(targetLoc);
      var f = _resolveFields(r);
      var status = String(r[col.Status] || '').toUpperCase();

      var qtyActual = Number(r[col.Qty_Actual] || 0);
      var qtyTarget = Number(r[col.Qty_Target] || 0);
      var qty       = qtyActual > 0 ? qtyActual : qtyTarget;

      var qtyNG = Number(r[col.Qty_NG] || 0);
      var kgNG  = Number(r[col.KG_NG] || 0);
      var hasNG = qtyNG > 0;

      // ── Batch child & Lot dari Trace_Log ────────────────────────────
      var trace = traceMap[spkNoOut] || null;
      var batchAvailable = trace !== null && trace.batch_id !== '';
      var batchChild = batchAvailable ? trace.batch_id : '';
      var lotProd    = batchAvailable ? String(trace.batch_id ? spkNoOut : spkNoOut) : spkNoOut;
      // Note: SPK_Ref di Trace_Log = SPK OUT No → sama dengan spkNoOut, dipakai sebagai Lot Prod

      var tglProd;
      if (trace && trace.tgl_prod) {
        tglProd = trace.tgl_prod instanceof Date ? trace.tgl_prod.toISOString() : String(trace.tgl_prod);
      } else {
        var selesai = r[col.Selesai_DT];
        var buat    = r[col.Tgl_Buat];
        var t2 = selesai || buat || new Date();
        tglProd = t2 instanceof Date ? t2.toISOString() : String(t2);
      }

      // ── NG batch (fallback kalau Trace_Log gak ada entry NG) ────────
      var ngSpkRef = spkNoOut + '-NG';
      var ngTrace  = traceMap[ngSpkRef] || null;
      var ngBatch  = ngTrace && ngTrace.batch_id ? ngTrace.batch_id : ngSpkRef;

      return {
        spk_no:           spkNoOut,
        target_loc:       targetLoc,
        label_type:       labelType,
        status:           status,
        cust:             String(r[col.Cust] || ''),
        item_code:        f.itemCode,
        spec:             f.spec,
        t:                f.t,
        p:                f.p,
        l:                f.l,
        qty:              qty,
        batch_id:         batchChild,     // dari Trace_Log
        lot:              spkNoOut,       // SPK_Ref di Trace_Log = SPK OUT No
        batch_available:  batchAvailable, // false = blokir print (belum DONE)
        tgl_prod:         tglProd,
        has_ng:           hasNG,
        ng: hasNG ? {
          qty:        qtyNG,
          kg:         kgNG,
          mc:         String(r[col.MC_No] || ''),
          owner:      String(r[col.Owner_Used] || r[col.Owner] || 'FC'),
          source_spk: spkNoOut,
          batch_id:   ngBatch,
          lot:        ngSpkRef
        } : null,
        shr_children: (labelType === 'WIP' || labelType === 'WIP_STP')
          ? _getShrChildrenForCtlOut(spkNoOut)
          : []
      };
    });

    var hTgl = headerRow[col.Tgl_Buat];
    return {
      ok: true,
      header: {
        spk_no:    parentSpkNo,
        spk_type:  headerSpkType,
        status:    String(headerRow[col.Status] || ''),
        tgl_buat:  hTgl instanceof Date ? hTgl.toISOString() : String(hTgl || '')
      },
      outs: outs
    };
  } catch (e) {
    return { ok: false, msg: 'Error: ' + e.message };
  }
}