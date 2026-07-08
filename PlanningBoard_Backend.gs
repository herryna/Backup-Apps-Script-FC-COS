/* =========================================================================
 * PLANNING BOARD — Backend Functions
 * Versi: v2 — Dynamic machine list dari M_MC (support SLT-01 + SHR-03)
 * Perubahan v2:
 *   - Baca daftar mesin dari M_MC (filter Is_Active=TRUE, sort Display_Order)
 *   - Return machines[] metadata → frontend render kolom dinamis
 *   - Cascade MC_No generic: SHR-OUT & SLT-OUT (CTL-OUT tidak perlu)
 * ========================================================================= */

/* =========================================================================
 * HELPER: Normalize nilai Estimasi_Jam ke "dd MMM HH:mm"
 * ========================================================================= */
function normalizeEstJam(val, tz) {
  if (!val || val === '') return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, "dd MMM HH:mm");
  }
  return String(val).trim();
}

/* =========================================================================
 * HELPER: Baca daftar mesin aktif dari M_MC
 *   Filter : Is_Active === TRUE
 *   Sort   : Display_Order ASC
 *   Return : [{mc_no, mc_name, type, display_order}]
 *   Fallback jika M_MC bermasalah → list legacy 3 mesin
 * ========================================================================= */
function getMachineList_() {
  var fallback = [
    { mc_no: 'CTL-01', mc_name: 'CTL-01', type: 'CTL', display_order: 2 },
    { mc_no: 'SHR-01', mc_name: 'SHR-01', type: 'SHR', display_order: 3 },
    { mc_no: 'SHR-02', mc_name: 'SHR-02', type: 'SHR', display_order: 4 }
  ];

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('M_MC');
    if (!sheet) return fallback;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return fallback;

    var headers   = data[0].map(function(h) { return String(h).trim(); });
    var iMcNo     = headers.indexOf('MC_No');
    var iMcName   = headers.indexOf('MC_Name');
    var iType     = headers.indexOf('Type');
    var iIsActive = headers.indexOf('Is_Active');
    var iDispOrd  = headers.indexOf('Display_Order');

    if (iMcNo === -1) return fallback;

    var list = [];
    for (var i = 1; i < data.length; i++) {
      var mcNo = String(data[i][iMcNo] || '').trim();
      if (!mcNo) continue;

      // Cek Is_Active kalau kolom-nya ada
      if (iIsActive !== -1) {
        var isAct  = data[i][iIsActive];
        var actStr = String(isAct).toUpperCase().trim();
        if (isAct !== true && actStr !== 'TRUE' && actStr !== 'YA' && actStr !== '1') continue;
      }

      list.push({
        mc_no        : mcNo,
        mc_name      : iMcName  !== -1 ? String(data[i][iMcName] || mcNo).trim() : mcNo,
        type         : iType    !== -1 ? String(data[i][iType]   || '').toUpperCase().trim() : '',
        display_order: iDispOrd !== -1 ? (Number(data[i][iDispOrd]) || 999) : 999
      });
    }

    if (list.length === 0) return fallback;

    list.sort(function(a, b) {
      if (a.display_order !== b.display_order) return a.display_order - b.display_order;
      var typeOrder = { 'SLT': 0, 'CTL': 1, 'SHR': 2 };
      var tA = typeOrder[a.type] !== undefined ? typeOrder[a.type] : 9;
      var tB = typeOrder[b.type] !== undefined ? typeOrder[b.type] : 9;
      if (tA !== tB) return tA - tB;
      return a.mc_no < b.mc_no ? -1 : (a.mc_no > b.mc_no ? 1 : 0);
    });
    return list;
  } catch (e) {
    Logger.log('getMachineList_ error: ' + e.message);
    return fallback;
  }
}

/**
 * Ambil data antrian aktif per mesin untuk Planning Board
 * v2: dynamic mesin dari M_MC + tetap kirim out_details untuk popup
 * Return: { success, data: {mc_no: [...]}, machines: [{mc_no, type, ...}] }
 */
function getPlanningBoardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Get machine list dulu ──
    var machineList = getMachineList_();
    var machineSet  = {};   // untuk quick lookup {mc_no: true}
    var result      = {};   // {mc_no: []}
    machineList.forEach(function(m) {
      machineSet[m.mc_no] = true;
      result[m.mc_no]     = [];
    });

    var sheet   = ss.getSheetByName("SPK");
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var tz      = Session.getScriptTimeZone();

    var iSpk         = headers.indexOf("SPK_No");
    var iType        = headers.indexOf("SPK_Type");
    var iMc          = headers.indexOf("MC_No");
    var iParent      = headers.indexOf("Parent_SPK");
    var iSpec        = headers.indexOf("Input_Spec");
    var iCust        = headers.indexOf("Cust");
    var iSoRef       = headers.indexOf("SO_Ref");
    var iPrio        = headers.indexOf("Priority");
    var iStatus      = headers.indexOf("Status");
    var iPlanSeq     = headers.indexOf("Plan_Seq");
    var iTglDelivery = headers.indexOf("Tgl_Delivery");
    var iEstMulai    = headers.indexOf("Estimasi_Jam_Mulai");
    var iEstSelesai  = headers.indexOf("Estimasi_Jam_Selesai");
    var iDurasi      = headers.indexOf("Total_Durasi_Menit");
    var iQtyTgt      = headers.indexOf("Qty_Target");
    var iKgTgt       = headers.indexOf("KG_Target");
    var iQtyAct      = headers.indexOf("Qty_Actual");
    var iKgAct       = headers.indexOf("KG_Actual");
    var iTgtLoc      = headers.indexOf("Target_Loc");
    var iOwnerUsed   = headers.indexOf("Owner_Used");

    // ── Pass 1: custMap ──
    var custMap = {};
    for (var i = 1; i < data.length; i++) {
      var rowType   = String(data[i][iType]   || '').trim();
      var rowParent = String(data[i][iParent] || '').trim();
      if (rowType.indexOf('OUT') === -1 || !rowParent) continue;

      var cust  = String(data[i][iCust]  || '').trim();
      var soRef = String(data[i][iSoRef] || '').trim();
      if (!custMap[rowParent]) custMap[rowParent] = { custs: [], sos: [] };
      if (cust  && custMap[rowParent].custs.indexOf(cust)  === -1) custMap[rowParent].custs.push(cust);
      if (soRef && custMap[rowParent].sos.indexOf(soRef)   === -1) custMap[rowParent].sos.push(soRef);
    }

    // ── Pass 1.5: outDetailsMap (untuk popup) ──
    var outDetailsMap = {};
    for (var i = 1; i < data.length; i++) {
      var rowType   = String(data[i][iType]   || '').trim();
      var rowParent = String(data[i][iParent] || '').trim();
      var rowStatus = String(data[i][iStatus] || '').toUpperCase();

      if (rowType.indexOf('OUT') === -1 || !rowParent) continue;
      if (rowStatus === 'CANCELLED') continue;

      if (!outDetailsMap[rowParent]) outDetailsMap[rowParent] = { first_spec: '', outs: [] };

      var outSpec = String(data[i][iSpec] || '').trim();

      if (outDetailsMap[rowParent].outs.length === 0) {
        outDetailsMap[rowParent].first_spec = outSpec;
      }

      outDetailsMap[rowParent].outs.push({
        spk_no    : String(data[i][iSpk]    || '').trim(),
        spec      : outSpec,
        so_ref    : String(data[i][iSoRef]  || '').trim(),
        cust      : String(data[i][iCust]   || '').trim(),
        qty_tgt   : iQtyTgt    !== -1 ? (Number(data[i][iQtyTgt])    || 0) : 0,
        kg_tgt    : iKgTgt     !== -1 ? (Number(data[i][iKgTgt])     || 0) : 0,
        qty_act   : iQtyAct    !== -1 ? (Number(data[i][iQtyAct])    || 0) : 0,
        kg_act    : iKgAct     !== -1 ? (Number(data[i][iKgAct])     || 0) : 0,
        target_loc: iTgtLoc    !== -1 ? String(data[i][iTgtLoc]    || '').trim() : '',
        owner_used: iOwnerUsed !== -1 ? String(data[i][iOwnerUsed] || '').trim() : '',
        status    : rowStatus
      });
    }

    // ── Pass 2: Ambil HEADER rows aktif ──
    for (var i = 1; i < data.length; i++) {
      var spkType = String(data[i][iType]   || '').trim();
      var status  = String(data[i][iStatus] || '').toUpperCase();
      var mc      = String(data[i][iMc]     || '').trim();

      if (spkType.indexOf('-HEADER') === -1) continue;
      if (status !== 'ANTRIAN' && status !== 'RUNNING') continue;
      if (!machineSet[mc]) continue;   // skip kalau mesin tidak ada di M_MC

      var spkNo      = String(data[i][iSpk]  || '').trim();
      var headerSpec = String(data[i][iSpec] || '').trim();
      var custInfo   = custMap[spkNo]       || { custs: [], sos: [] };
      var outInfo    = outDetailsMap[spkNo] || { first_spec: '', outs: [] };

      // SHR-HEADER → pakai spec OUT pertama (nama produk).
      // CTL/SLT-HEADER → pakai HEADER spec (coil input).
      var displaySpec = (spkType === 'SHR-HEADER' && outInfo.first_spec)
                        ? outInfo.first_spec
                        : headerSpec;

      var tglDelivery = '';
      if (iTglDelivery !== -1 && data[i][iTglDelivery]) {
        var d = data[i][iTglDelivery];
        tglDelivery = d instanceof Date
          ? Utilities.formatDate(d, tz, "dd/MM/yyyy")
          : String(d).trim();
      }

      result[mc].push({
        spk_no           : spkNo,
        spk_type         : spkType,
        mc_no            : mc,
        parent_spk       : String(data[i][iParent] || '').trim(),
        input_spec       : displaySpec,
        input_spec_header: headerSpec,
        cust_list        : custInfo.custs.join(', ') || '--',
        so_list          : custInfo.sos.join(', ')   || '--',
        priority         : String(data[i][iPrio]   || 'Normal').trim(),
        status           : status,
        plan_seq         : iPlanSeq !== -1 ? (Number(data[i][iPlanSeq]) || 0) : 0,
        tgl_delivery     : tglDelivery,
        est_mulai        : iEstMulai   !== -1 ? normalizeEstJam(data[i][iEstMulai],   tz) : '',
        est_selesai      : iEstSelesai !== -1 ? normalizeEstJam(data[i][iEstSelesai], tz) : '',
        total_durasi     : iDurasi     !== -1 ? (Number(data[i][iDurasi])  || 0) : 0,
        qty_target       : iQtyTgt     !== -1 ? (Number(data[i][iQtyTgt])  || 0) : 0,
        kg_target        : iKgTgt      !== -1 ? (Number(data[i][iKgTgt])   || 0) : 0,
        out_details      : outInfo.outs
      });
    }

    // ── Sort tiap mesin: Running → Plan_Seq → Priority ──
    machineList.forEach(function(m) {
      result[m.mc_no].sort(function(a, b) {
        if (a.status === 'RUNNING' && b.status !== 'RUNNING') return -1;
        if (b.status === 'RUNNING' && a.status !== 'RUNNING') return  1;
        if (a.plan_seq > 0 && b.plan_seq > 0) return a.plan_seq - b.plan_seq;
        if (a.plan_seq > 0 && b.plan_seq === 0) return -1;
        if (a.plan_seq === 0 && b.plan_seq > 0) return  1;
        var wA = a.priority === 'URGENT' ? 2 : (a.priority === 'HIGH' ? 1 : 0);
        var wB = b.priority === 'URGENT' ? 2 : (b.priority === 'HIGH' ? 1 : 0);
        return wB - wA;
      });
    });

    return { success: true, data: result, machines: machineList };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Simpan urutan planning (Plan_Seq) & assignment mesin (MC_No)
 * v2: Cascade generic — SHR-OUT & SLT-OUT
 * Input: { updates: [{ spk_no, plan_seq, mc_no }, ...] }
 */
function savePlanningBoard(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SPK");
    var rows    = sheet.getDataRange().getValues();
    var headers = rows[0].map(function(h) { return String(h).trim(); });

    var iSpk     = headers.indexOf("SPK_No");
    var iPlanSeq = headers.indexOf("Plan_Seq");
    var iMc      = headers.indexOf("MC_No");
    var iType    = headers.indexOf("SPK_Type");
    var iParent  = headers.indexOf("Parent_SPK");
    var iSrcLoc  = headers.indexOf("Source_Loc");

    if (iPlanSeq === -1) {
      return { success: false, message: "Kolom Plan_Seq tidak ditemukan di sheet SPK." };
    }

    // Build rowMap: SPK_No → rowNum (1-based)
    var rowMap = {};
    for (var i = 1; i < rows.length; i++) {
      var sn = String(rows[i][iSpk] || '').trim();
      if (sn) rowMap[sn] = i + 1;
    }

    // v2: Build childMap generic — SHR-OUT & SLT-OUT
    // CTL-OUT tidak perlu (CTL cuma 1 mesin, tidak bisa cross-drag)
    var childMap = {};
    for (var i = 1; i < rows.length; i++) {
      var rType   = String(rows[i][iType]   || '').trim();
      var rParent = String(rows[i][iParent] || '').trim();
      if ((rType === 'SHR-OUT' || rType === 'SLT-OUT') && rParent) {
        if (!childMap[rParent]) childMap[rParent] = [];
        childMap[rParent].push(i + 1);
      }
    }

    // Apply updates
    payload.updates.forEach(function(upd) {
      var rowNum = rowMap[upd.spk_no];
      if (!rowNum) return;

      // Update Plan_Seq di HEADER
      if (iPlanSeq !== -1) {
        sheet.getRange(rowNum, iPlanSeq + 1).setValue(upd.plan_seq);
      }

      // Update MC_No di HEADER + CASCADE ke OUT anak
      if (iMc !== -1 && upd.mc_no) {
        sheet.getRange(rowNum, iMc + 1).setValue(upd.mc_no);

        var children = childMap[upd.spk_no] || [];
        children.forEach(function(childRow) {
          sheet.getRange(childRow, iMc + 1).setValue(upd.mc_no);
          if (iSrcLoc !== -1) {
            sheet.getRange(childRow, iSrcLoc + 1).setValue(upd.mc_no);
          }
        });
      }
    });

    SpreadsheetApp.flush();

    // Recalculate jadwal setelah urutan/mesin berubah
    if (typeof kalkulasiEstimasiWaktu === 'function') {
      kalkulasiEstimasiWaktu();
    }

    return { success: true, count: payload.updates.length };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}