/* =========================================================================
 * PLANNING BOARD — Backend Functions
 * Versi: Fix cascade MC_No + OUT details untuk popup
 * ========================================================================= */
/* =========================================================================
 * HELPER: Normalize nilai Estimasi_Jam ke "dd MMM HH:mm"
 *   - Jika Sheets sudah terlanjur ubah jadi Date object → format ulang
 *   - Jika sudah string → biarkan apa adanya
 * ========================================================================= */
function normalizeEstJam(val, tz) {
  if (!val || val === '') return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, "dd MMM HH:mm");
  }
  return String(val).trim();
}
/**
 * Ambil data antrian aktif per mesin untuk Planning Board
 * Return: { 'CTL-01': [...], 'SHR-01': [...], 'SHR-02': [...] }
 * Perubahan:
 *   - Tambah out_details[] per HEADER → dipakai popup di frontend (0 extra GAS call)
 *   - input_spec SHR-HEADER diambil dari OUT pertama (bukan HEADER)
 *   - input_spec_header tetap disimpan → ditampilkan di popup sebagai "Material Input"
 */
function getPlanningBoardData() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
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

    // ── Pass 1: custMap — ringkasan customer per HEADER ──
    var custMap = {}; // parent_spk → { custs: [], sos: [] }
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

    // ── Pass 1.5: outDetailsMap — data OUT lengkap per HEADER (untuk popup) ──
    // Juga ambil spec OUT pertama untuk ditampilkan di kartu (SHR saja)
    var outDetailsMap = {}; // parent_spk → { first_spec: '', outs: [] }
    for (var i = 1; i < data.length; i++) {
      var rowType   = String(data[i][iType]   || '').trim();
      var rowParent = String(data[i][iParent] || '').trim();
      var rowStatus = String(data[i][iStatus] || '').toUpperCase();

      if (rowType.indexOf('OUT') === -1 || !rowParent) continue;
      if (rowStatus === 'CANCELLED') continue;

      if (!outDetailsMap[rowParent]) outDetailsMap[rowParent] = { first_spec: '', outs: [] };

      var outSpec = String(data[i][iSpec] || '').trim();

      // first_spec = spec OUT pertama yang ditemukan
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
    var machines = ['CTL-01', 'SHR-01', 'SHR-02'];
    var result   = { 'CTL-01': [], 'SHR-01': [], 'SHR-02': [] };

    for (var i = 1; i < data.length; i++) {
      var spkType = String(data[i][iType]   || '').trim();
      var status  = String(data[i][iStatus] || '').toUpperCase();
      var mc      = String(data[i][iMc]     || '').trim();

      if (spkType.indexOf('-HEADER') === -1) continue;
      if (status !== 'ANTRIAN' && status !== 'RUNNING') continue;
      if (!result[mc]) continue;

      var spkNo       = String(data[i][iSpk]    || '').trim();
      var headerSpec  = String(data[i][iSpec]   || '').trim(); // spec raw material (untuk popup)
      var custInfo    = custMap[spkNo]       || { custs: [], sos: [] };
      var outInfo     = outDetailsMap[spkNo] || { first_spec: '', outs: [] };

      // Kartu SHR → pakai spec OUT pertama (nama produk hasil potong)
      // Kartu CTL → tetap pakai spec HEADER (nama coil input)
      var displaySpec = (spkType === 'SHR-HEADER' && outInfo.first_spec)
                        ? outInfo.first_spec
                        : headerSpec;

      // Format Tgl_Delivery
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
        input_spec       : displaySpec,          // tampil di kartu
        input_spec_header: headerSpec,           // raw material spec → popup
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
        out_details      : outInfo.outs   // array OUT → popup, 0 extra GAS call
      });
    }

    // ── Sort tiap mesin by Plan_Seq → Priority ──
    machines.forEach(function(mc) {
      result[mc].sort(function(a, b) {
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

    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Simpan urutan planning (Plan_Seq) dan assignment mesin (MC_No)
 * FIX: cascade MC_No + Source_Loc ke semua baris SHR-OUT anak
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

    // Build rowMap: SPK_No → rowNum (sheet row, 1-based)
    var rowMap = {};
    for (var i = 1; i < rows.length; i++) {
      var sn = String(rows[i][iSpk] || '').trim();
      if (sn) rowMap[sn] = i + 1;
    }

    // ✅ FIX: Build childMap → SHR-HEADER spk_no → [rowNums SHR-OUT anak]
    // Dipakai untuk cascade update MC_No & Source_Loc saat mesin diganti
    var childMap = {};
    for (var i = 1; i < rows.length; i++) {
      var rType   = String(rows[i][iType]   || '').trim();
      var rParent = String(rows[i][iParent] || '').trim();
      if (rType === 'SHR-OUT' && rParent) {
        if (!childMap[rParent]) childMap[rParent] = [];
        childMap[rParent].push(i + 1); // simpan row number
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

      // Update MC_No di HEADER + CASCADE ke SHR-OUT anak
      if (iMc !== -1 && upd.mc_no) {
        sheet.getRange(rowNum, iMc + 1).setValue(upd.mc_no);

        // ✅ CASCADE: update MC_No & Source_Loc di semua SHR-OUT anak
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
