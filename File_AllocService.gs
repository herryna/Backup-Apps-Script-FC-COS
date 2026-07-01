/* =========================================================================
 * ALLOC SERVICE — Allocated Material without Machine Process
 * File: File_AllocService.gs (tambahkan sebagai file baru)
 * ========================================================================= */

/**
 * Ambil data inisialisasi Form Allocated Material:
 *   1. Batch Stok_Sheet yang masih available (Qty_Avail > 0)
 *   2. Daftar SO aktif untuk poka-yoke lookup
 */
function getInitDataAlloc() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── 1. Stok_Sheet ──────────────────────────────────────────────────────
    var stokSht = ss.getSheetByName('Stok_Sheet');
    if (!stokSht) throw new Error("Sheet 'Stok_Sheet' tidak ditemukan");

    var sd  = stokSht.getDataRange().getValues();
    var sh  = sd[0].map(function(h) { return String(h).trim(); });

    // Helper: cari kolom toleran terhadap variasi nama/huruf besar-kecil
    function findCol(headers, names) {
      for (var n = 0; n < names.length; n++) {
        for (var c = 0; c < headers.length; c++) {
          if (headers[c].toLowerCase() === names[n].toLowerCase()) return c;
        }
      }
      return -1;
    }

    var siB  = findCol(sh, ['Batch_ID']);
    var siI  = findCol(sh, ['Item_Code']);
    var siD  = findCol(sh, ['Description']);
    var siSp = findCol(sh, ['Spec']);
    var siT  = findCol(sh, ['T']);
    var siP  = findCol(sh, ['P']);
    var siL  = findCol(sh, ['L_Dim', 'L_dim', 'L']);
    var siQI = findCol(sh, ['Qty_In']);
    var siKI = findCol(sh, ['KG_In']);
    var siQA = findCol(sh, ['Qty_Avail']);
    var siKA = findCol(sh, ['Kg_Avail', 'KG_Avail']);
    var siOw = findCol(sh, ['Owner']);

    var batches = [];
    for (var i = 1; i < sd.length; i++) {
      var r     = sd[i];
      var bid   = String(r[siB] || '').trim();
      if (!bid) continue;
      var qAvl  = parseFloat(r[siQA]) || 0;
      if (qAvl <= 0) continue;
      var qIn   = parseFloat(r[siQI]) || 1;
      var kIn   = parseFloat(r[siKI]) || 0;
      batches.push({
        batch_id  : bid,
        item_code : String(r[siI]  || '').trim(),
        desc      : String(r[siD]  || '').trim(),
        spec      : String(r[siSp] || '').trim(),
        T         : String(r[siT]  || '').trim(),
        P         : String(r[siP]  || '').trim(),
        L         : String(r[siL]  || '').trim(),
        qty_in    : qIn,
        kg_in     : kIn,
        qty_avail : qAvl,
        kg_avail  : parseFloat(r[siKA]) || 0,
        kg_per_sht: qIn > 0 ? Math.round(kIn / qIn * 10000) / 10000 : 0,
        owner     : String(r[siOw] || 'FC').trim().toUpperCase()
      });
    }

    // ── 2. SO aktif (untuk poka-yoke) ──────────────────────────────────────
    var soSht = ss.getSheetByName('SO');
    if (!soSht) throw new Error("Sheet 'SO' tidak ditemukan");

    var od  = soSht.getDataRange().getValues();
    var oh  = od[0].map(function(h) { return String(h).trim(); });

    var oiSoNo = findCol(oh, ['SO_No']);
    var oiCust = findCol(oh, ['Cust']);
    var oiCode = findCol(oh, ['Item Code', 'Item_Code']);
    var oiDesc = findCol(oh, ['Description']);
    var oiStat = findCol(oh, ['STATUS', 'Status']);
    var oiBlQ  = findCol(oh, ['BL_Q']);
    var oiT    = findCol(oh, ['T']);
    var oiP    = findCol(oh, ['P']);
    var oiL    = findCol(oh, ['L']);

    var soList = [];
    for (var i = 1; i < od.length; i++) {
      var r    = od[i];
      var soNo = String(r[oiSoNo] || '').trim();
      if (!soNo) continue;
      var stat = String(r[oiStat] || '').toUpperCase().trim();
      if (stat === 'DONE' || stat === 'CLOSED' || stat === 'CANCELLED') continue;

      soList.push({
        so_no    : soNo,
        cust     : String(r[oiCust] || '').trim(),
        item_code: String(r[oiCode] || '').trim(),
        desc     : String(r[oiDesc] || '').trim(),
        T        : oiT  !== -1 ? String(r[oiT]  || '').trim() : '',
        P        : oiP  !== -1 ? String(r[oiP]  || '').trim() : '',
        L        : oiL  !== -1 ? String(r[oiL]  || '').trim() : '',
        bl_q     : parseFloat(r[oiBlQ]) || 0
      });
    }

    return { success: true, batches: batches, so_list: soList };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


/**
 * Simpan Alokasi Material
 * Flow: ALLOC (header) + N ALLOC-OUT → SPK
 *       N baris → Stok_FG
 *       Baris Owner ≠ Owner_Used → Rekap_ICT
 *
 * Payload:
 * {
 *   batch_id, batch_item_code, batch_spec, batch_T, batch_P, batch_L,
 *   batch_owner, kg_per_sht, total_qty, total_kg, created_by, note,
 *   out_alloc: [{ so_ref, cust, item_code, description, qty, kg, owner_used }, ...]
 * }
 */
function saveSPK_ALLOC(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var spkSht   = ss.getSheetByName('SPK');
    var fgSht    = ss.getSheetByName('Stok_FG');
    var ictSht   = ss.getSheetByName('Rekap_ICT');
    var tz       = Session.getScriptTimeZone();

    if (!spkSht) throw new Error("Sheet 'SPK' tidak ditemukan");

    var spkData  = spkSht.getDataRange().getValues();
    var spkHdr   = spkData[0].map(function(h) { return String(h).trim(); });
    var ts       = new Date();

    // ── Generate nomor ALO-YYXXXX ──
    var yr     = Utilities.formatDate(ts, tz, 'yy');
    var prefix = 'ALO-' + yr;
    var iSpkNo = spkHdr.indexOf('SPK_No');
    var maxSeq = 0;
    for (var i = 1; i < spkData.length; i++) {
      var sn = String(spkData[i][iSpkNo] || '').trim();
      if (sn.indexOf(prefix) === 0) {
        var q = parseInt(sn.replace(prefix, ''), 10);
        if (!isNaN(q) && q > maxSeq) maxSeq = q;
      }
    }
    var allocNo = prefix + ('0000' + (maxSeq + 1)).slice(-4);

    function buildRow(rd) {
      return spkHdr.map(function(h) { return rd[h] !== undefined ? rd[h] : ''; });
    }

    var bOwner   = String(data.batch_owner || 'FC').trim().toUpperCase();
    var totalQty = parseFloat(data.total_qty) || 0;
    var totalKg  = parseFloat(data.total_kg)  || 0;

    // ── ALLOC HEADER ──────────────────────────────────────────────────────
    spkSht.appendRow(buildRow({
      'SPK_No'            : allocNo,
      'SPK_Type'          : 'ALLOC-HEADER',
      'Parent_SPK'        : data.batch_id,
      'Tgl_Buat'          : ts,
      'Item_Code'         : data.batch_item_code || '',
      'Input_Spec'        : data.batch_desc || data.batch_spec || '',
      'Qty_Target'        : totalQty,
      'KG_Target'         : totalKg,
      'Qty_Actual'        : totalQty,
      'KG_Actual'         : totalKg,
      'Status'            : 'DONE',
      'Mulai_DT'          : ts,
      'Selesai_DT'        : ts,
      'Source_Loc'        : 'Stok_Sheet',
      'Target_Loc'        : 'FG_Cust',
      'Owner'             : bOwner,
      'Owner_Used'        : bOwner,
      'Created_By'        : data.created_by || 'Admin',
      'NOTE'              : data.note        || '',
      'Plan_Setup_Menit'  : 0,
      'Plan_Run_Menit'    : 0,
      'Total_Durasi_Menit': 0
    }));

    // ── Siapkan header FG & ICT (baca sekali) ──
    var fgHdr  = fgSht  ? fgSht.getDataRange().getValues()[0].map(function(h) { return String(h).trim(); }) : [];
    var ictHdr = ictSht ? ictSht.getDataRange().getValues()[0].map(function(h) { return String(h).trim(); }) : [];
    var hasIct = false;

    // ── ALLOC-OUT per baris ───────────────────────────────────────────────
    data.out_alloc.forEach(function(out, idx) {
      var outNo     = allocNo + '-' + (idx + 1);
      var oUsed     = String(out.owner_used || bOwner).trim().toUpperCase();
      var oQty      = parseFloat(out.qty) || 0;
      var oKg       = Math.round(oQty * parseFloat(data.kg_per_sht || 0) * 100) / 100;

      // SPK ALLOC-OUT
      spkSht.appendRow(buildRow({
        'SPK_No'            : outNo,
        'SPK_Type'          : 'ALLOC-OUT',
        'SO_Ref'            : out.so_ref      || '',
        'Cust'              : out.cust        || '',
        'Parent_SPK'        : allocNo,
        'Tgl_Buat'          : ts,
        'Item_Code'         : out.item_code   || data.batch_item_code || '',
        'Input_Spec'        : out.description || data.batch_spec      || '',
        'Qty_Target'        : oQty,
        'KG_Target'         : oKg,
        'Qty_Actual'        : oQty,
        'KG_Actual'         : oKg,
        'Status'            : 'DONE',
        'Mulai_DT'          : ts,
        'Selesai_DT'        : ts,
        'Source_Loc'        : data.batch_id,  // penting: untuk SUMIFS Stok_Sheet
        'Target_Loc'        : 'FG_Cust',
        'Owner'             : bOwner,
        'Owner_Used'        : oUsed,
        'Created_By'        : data.created_by || 'Admin',
        'Plan_Setup_Menit'  : 0,
        'Plan_Run_Menit'    : 0,
        'Total_Durasi_Menit': 0
      }));

      // Stok_FG — tulis ke baris pertama yang Batch_ID-nya kosong
      // (hindari appendRow yang loncat melewati baris berformula)
      if (fgSht && fgHdr.length > 0) {
        var fgRowData = fgHdr.map(function(h) {
          switch(h) {
            case 'Batch_ID'    : return data.batch_id           || '';
            case 'Tgl_Output'  : return ts;
            case 'SPK_Ref'     : return outNo;
            case 'SO_Ref'      : return out.so_ref              || '';
            case 'Item_Code'   : return out.item_code           || data.batch_item_code || '';
            case 'Description' : return out.description         || data.batch_spec      || '';
            case 'Spec'        : return data.batch_spec         || '';
            case 'T'           : return data.batch_T            || '';
            case 'P'           : return data.batch_P            || '';
            case 'L_dim'       : return data.batch_L            || '';
            case 'UoM'         : return 'Sht';
            case 'Qty_In'      : return oQty;
            case 'KG_In'       : return oKg;
            case 'NOTE'        : return data.note               || '';
            case 'Owner'       : return bOwner;
            case 'Target_Loc'  : return out.target_loc          || 'FG_Cust';
            default            : return '';
          }
        });

        // Cari baris pertama yang Batch_ID-nya kosong (kolom A = index 0)
        var iBatchFG   = fgHdr.indexOf('Batch_ID');
        var fgAllVals  = fgSht.getDataRange().getValues();
        var targetFGRow = fgAllVals.length + 1; // fallback: baris setelah data terakhir

        for (var r = 1; r < fgAllVals.length; r++) {
          if (!fgAllVals[r][iBatchFG] || String(fgAllVals[r][iBatchFG]).trim() === '') {
            targetFGRow = r + 1; // +1 karena index sheet 1-based
            break;
          }
        }

        fgSht.getRange(targetFGRow, 1, 1, fgHdr.length).setValues([fgRowData]);
      }

      // Rekap_ICT — jika cross-billing (Owner ≠ Owner_Used)
      if (ictSht && ictHdr.length > 0 && bOwner !== oUsed) {
        hasIct = true;
        ictSht.appendRow(ictHdr.map(function(h) {
          switch(h) {
            case 'Tgl_Transfer': return ts;
            case 'SPK_No'      : return outNo;
            case 'Item_Code'   : return out.item_code || data.batch_item_code || '';
            case 'Spec'        : return out.description || data.batch_spec || '';
            case 'Dari_Owner'  : return bOwner;
            case 'Ke_Owner'    : return oUsed;
            case 'Qty_Sht'     : return oQty;
            case 'Qty_KG'      : return oKg;
            default            : return '';
          }
        }));
      }
    });

    SpreadsheetApp.flush();

    return {
      success  : true,
      alloc_no : allocNo,
      count_out: data.out_alloc.length,
      has_ict  : hasIct
    };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}