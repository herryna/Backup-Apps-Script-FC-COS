/**
 * ═══════════════════════════════════════════════════════════════════
 * SLT_Service.gs — Backend Modul SPK Slitting
 * ═══════════════════════════════════════════════════════════════════
 *
 * SCOPE:
 * Handle semua operasi SPK SLT (Slitting):
 * - Read helpers (batch coil info, master data)
 * - Validation (cut plan, jalur, trim)
 * - Save draft (SLT-HEADER + SLT-OUT)
 * - Save DONE (generate batch turunan ke Trace_Log)
 *
 * KONVENSI:
 * - Function public prefix: (tidak ada prefix, akses dari HTML)
 * - Function private prefix: _slt_xxx (helper internal, jangan panggil dari HTML)
 * - LockService di function save/update, TIDAK nested di helper
 * - Return object: { success: bool, message: string, data: any }
 *
 * DEPENDENCIES:
 * - Helper_CT.gs → getMachineConfig, kalkulasiEstimasiWaktu_SLT,
 *                  hitungLengthCoil, hitungRencanaDurasi
 * - Utils.gs    → getSheet, _appendRowSafe (asumsi ada)
 * - SpkService.gs → getNextSpkNo (reuse pattern)
 *
 * KONSTANTA (constant tidak dari M_Config, ubah di sini kalau perlu):
 * - SLT_DENSITY = 7.85 g/cm³ (density baja, hardcode di Helper_CT)
 * - SLT_TRIM_DEFAULT_MM = 5 mm (default trim, editable di form)
 * - SLT_MAX_JALUR_PER_SETTING = 13 (physical knife slot limit)
 * - SLT_MAX_CUT_PER_JALUR = 10 (sanity cap)
 * ═══════════════════════════════════════════════════════════════════
 */

var SLT_CONSTANTS = {
  TRIM_DEFAULT_MM       : 5,
  MAX_JALUR_PER_SETTING : 13,
  MAX_CUT_PER_JALUR     : 10,
  MIN_TRIM_MM           : 3,   // knife minimum practical
  MAX_TRIM_MM_WARN      : 20   // trim > 20mm → warning yield loss
};

// ═══════════════════════════════════════════════════════════════════
// SECTION: READ HELPERS
// ═══════════════════════════════════════════════════════════════════

/**
 * Lookup detail batch coil dari Stok_Coil.
 * Dipakai saat validasi save dan generate SPK.
 *
 * NOTE: Untuk Type=Coil, kolom P di Stok_Coil = Lebar Coil (bukan panjang).
 *       Length coil dihitung runtime dari KG, T, Lebar via hitungLengthCoil().
 *
 * @param {string} batchId - Batch_ID coil (contoh: "FC-CL-2600100")
 * @return {object} Info coil. Field found=false kalau tidak ketemu.
 *
 * Struktur return:
 * {
 *   batch_id      : "FC-CL-2600100",
 *   item_code     : "3-MMP-CL",
 *   description   : "SPCC 2.90 x 1219 x C",
 *   spec          : "SPCC 2.90x1219xC",
 *   t             : 2.90,        // thickness mm
 *   lebar_coil    : 1219,        // dari kolom P (lebar untuk coil)
 *   qty_avail     : 1,           // biasanya 1 untuk coil
 *   kg_avail      : 9755,        // KG tersisa (KG_In - Keep - Prod - Done)
 *   length_m      : 351.5,       // computed via hitungLengthCoil
 *   owner         : "FC",
 *   source        : "GR",        // atau "Trace_Log" kalau sisa SLT
 *   found         : true
 * }
 */
function _slt_getInputCoilInfo(batchId) {
  var hasil = {
    batch_id    : batchId,
    item_code   : '',
    description : '',
    spec        : '',
    t           : 0,
    lebar_coil  : 0,
    qty_avail   : 0,
    kg_avail    : 0,
    length_m    : 0,
    owner       : '',
    source      : '',
    found       : false
  };

  try {
    if (!batchId) return hasil;
    var target = String(batchId).trim();
    if (!target) return hasil;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Stok_Coil');
    if (!sh) {
      Logger.log('_slt_getInputCoilInfo: sheet Stok_Coil tidak ada');
      return hasil;
    }

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return hasil;

    var hdr = data[0].map(function(h) { return String(h).trim(); });
    var iBatch  = hdr.indexOf('Batch_ID');
    var iItem   = hdr.indexOf('Item_Code');
    var iDesc   = hdr.indexOf('Description');
    var iSpec   = hdr.indexOf('Spec');
    var iT      = hdr.indexOf('T');
    var iP      = hdr.indexOf('P');         // untuk Type=Coil, P = Lebar
    var iQtyAv  = hdr.indexOf('Qty_Avail');
    var iKgAv   = hdr.indexOf('KG_Avail');
    var iOwner  = hdr.indexOf('Owner');
    var iSrcBt  = hdr.indexOf('Source_Batch'); // opsional (kalau ada)

    if (iBatch === -1) {
      Logger.log('_slt_getInputCoilInfo: kolom Batch_ID tidak ada');
      return hasil;
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iBatch] || '').trim() === target) {
        hasil.batch_id    = target;
        hasil.item_code   = iItem  !== -1 ? String(data[i][iItem]  || '') : '';
        hasil.description = iDesc  !== -1 ? String(data[i][iDesc]  || '') : '';
        hasil.spec        = iSpec  !== -1 ? String(data[i][iSpec]  || '') : '';
        hasil.t           = iT     !== -1 ? (Number(data[i][iT])     || 0) : 0;
        hasil.lebar_coil  = iP     !== -1 ? (Number(data[i][iP])     || 0) : 0;
        hasil.qty_avail   = iQtyAv !== -1 ? (Number(data[i][iQtyAv]) || 0) : 0;
        hasil.kg_avail    = iKgAv  !== -1 ? (Number(data[i][iKgAv])  || 0) : 0;
        hasil.owner       = iOwner !== -1 ? String(data[i][iOwner] || '') : '';

        // Identifikasi source: kalau Source_Batch terisi = dari Trace_Log (sisa SLT), else = GR
        if (iSrcBt !== -1 && String(data[i][iSrcBt] || '').trim() !== '') {
          hasil.source = 'Trace_Log';
        } else {
          hasil.source = 'GR';
        }

        // Compute length_m
        if (typeof hitungLengthCoil === 'function') {
          hasil.length_m = hitungLengthCoil(hasil.kg_avail, hasil.t, hasil.lebar_coil);
        }

        hasil.found = true;
        break;
      }
    }

    return hasil;

  } catch (e) {
    Logger.log('Error _slt_getInputCoilInfo(' + batchId + '): ' + e.toString());
    return hasil;
  }
}
/**
 * Lookup Description item dari M_ITEM.
 * Konsisten dengan pola CTL/SHR untuk Input_Spec.
 *
 * @param {string} itemCode
 * @return {string} Description dari M_ITEM, atau '' kalau tidak ketemu
 */
function _slt_getItemDescription(itemCode) {
  try {
    if (!itemCode) return '';
    var target = String(itemCode).trim();
    if (!target) return '';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('M_ITEM');
    if (!sh) return '';

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return '';

    var hdr = data[0].map(function(h) { return String(h).trim(); });
    var iCode = hdr.indexOf('Item_Code');
    var iDesc = hdr.indexOf('Description');
    if (iCode === -1 || iDesc === -1) return '';

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iCode] || '').trim() === target) {
        return String(data[i][iDesc] || '');
      }
    }
    return '';

  } catch (e) {
    Logger.log('Error _slt_getItemDescription: ' + e.toString());
    return '';
  }
}

// ═══════════════════════════════════════════════════════════════════
// SECTION: VALIDATION
// ═══════════════════════════════════════════════════════════════════

/**
 * Cek Item_Code valid di M_ITEM.
 * @param {string} itemCode
 * @return {boolean}
 */
function _slt_isValidItemCode(itemCode) {
  try {
    if (!itemCode) return false;
    var target = String(itemCode).trim();
    if (!target) return false;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('M_ITEM');
    if (!sh) return false;

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return false;
    var iCode = data[0].map(function(h) { return String(h).trim(); }).indexOf('Item_Code');
    if (iCode === -1) return false;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iCode] || '').trim() === target) return true;
    }
    return false;
  } catch (e) {
    Logger.log('Error _slt_isValidItemCode: ' + e.toString());
    return false;
  }
}

/**
 * Validate 1 setting (jalur, trim, lebar budget).
 * Helper internal, dipanggil dari validateSlitingPlan.
 *
 * @param {object} setting     - { trim_kiri, trim_kanan, outs: [...] }
 * @param {number} settingNo   - 1 atau 2 (untuk label error)
 * @param {number} lebarBudget - Lebar coil (setting-1) atau sisa (setting-2)
 * @param {boolean} allowUnder - true kalau boleh < lebarBudget (setting-1 mode split)
 * @return {object} { errors:[], warnings:[], totalJalur:num, lebarTerpakai:num, sumLebarJalur:num }
 */
function _slt_validateSetting(setting, settingNo, lebarBudget, allowUnder) {
  var res = { errors: [], warnings: [], totalJalur: 0, lebarTerpakai: 0, sumLebarJalur: 0 };
  var lbl = 'Setting-' + settingNo;

  var trimKiri  = Number(setting.trim_kiri)  || 0;
  var trimKanan = Number(setting.trim_kanan) || 0;
  var outs      = Array.isArray(setting.outs) ? setting.outs : [];

  // Sum jalur × lebar
  var sumLebarJalur = 0, totalJalur = 0;
  outs.forEach(function(o) {
    var lbr = Number(o.lebar)     || 0;
    var qty = Number(o.qty_jalur) || 0;
    sumLebarJalur += lbr * qty;
    totalJalur    += qty;
  });

  res.sumLebarJalur = sumLebarJalur;
  res.totalJalur    = totalJalur;
  res.lebarTerpakai = trimKiri + sumLebarJalur + trimKanan;

  // V6: Max jalur per setting
  if (totalJalur > SLT_CONSTANTS.MAX_JALUR_PER_SETTING) {
    res.errors.push(lbl + ': Total jalur ' + totalJalur + ' melebihi batas ' +
                    SLT_CONSTANTS.MAX_JALUR_PER_SETTING + ' jalur (physical limit pisau)');
  }
  if (totalJalur === 0) {
    res.errors.push(lbl + ': Belum ada jalur output');
  }

  // V4/V5: Lebar check
  if (lebarBudget > 0) {
    if (allowUnder) {
      // Setting-1 mode split: total lebar boleh <= budget (sisa lanjut ke setting-2)
      if (res.lebarTerpakai > lebarBudget) {
        res.errors.push(lbl + ': Total lebar terpakai ' + res.lebarTerpakai +
                        'mm melebihi lebar coil ' + lebarBudget + 'mm');
      }
    } else {
      // Non-split, atau setting-2: harus EXACT
      if (res.lebarTerpakai !== lebarBudget) {
        var selisih = lebarBudget - res.lebarTerpakai;
        res.errors.push(lbl + ': Total lebar terpakai ' + res.lebarTerpakai +
                        'mm tidak sama dengan lebar tersedia ' + lebarBudget +
                        'mm (selisih ' + selisih + 'mm)');
      }
    }
  }

  // V7: Trim range check
  if (trimKiri > 0 && trimKiri < SLT_CONSTANTS.MIN_TRIM_MM) {
    res.errors.push(lbl + ': Trim Kiri ' + trimKiri + 'mm terlalu kecil (minimum ' +
                    SLT_CONSTANTS.MIN_TRIM_MM + 'mm)');
  }
  if (trimKanan > 0 && trimKanan < SLT_CONSTANTS.MIN_TRIM_MM) {
    res.errors.push(lbl + ': Trim Kanan ' + trimKanan + 'mm terlalu kecil (minimum ' +
                    SLT_CONSTANTS.MIN_TRIM_MM + 'mm)');
  }
  if (trimKiri > SLT_CONSTANTS.MAX_TRIM_MM_WARN) {
    res.warnings.push(lbl + ': Trim Kiri ' + trimKiri + 'mm besar, yield loss meningkat');
  }
  if (trimKanan > SLT_CONSTANTS.MAX_TRIM_MM_WARN) {
    res.warnings.push(lbl + ': Trim Kanan ' + trimKanan + 'mm besar, yield loss meningkat');
  }

  // Setting-1 dalam split: trim_kanan sebaiknya 0
  if (allowUnder && settingNo === 1 && trimKanan > 0) {
    res.warnings.push(lbl + ' (mode split): Trim Kanan sebaiknya 0mm karena sisa lebar akan diproses di Setting-2. Input: ' + trimKanan + 'mm');
  }

  return res;
}

/**
 * Validate full SPK SLT plan.
 * Dipanggil dari frontend sebelum saveSPK_SLT, atau dari backend sebelum write.
 *
 * @param {object} payload - Struktur payload SPK SLT (lihat design D1)
 * @return {object} {
 *   ok       : bool,
 *   errors   : [string],   // errors → block save
 *   warnings : [string],   // warnings → tampil tapi save boleh (butuh konfirmasi frontend)
 *   summary  : object      // ringkasan untuk display
 * }
 */
function validateSlitingPlan(payload) {
  var result = {
    ok       : true,
    errors   : [],
    warnings : [],
    summary  : {
      lebar_coil       : 0,
      lebar_terpakai   : 0,
      sisa_lebar       : 0,
      total_jalur      : 0,
      total_cut        : 0,
      total_batch      : 0,
      length_total_m   : 0,
      length_terpakai_m: 0,
      total_kg_output  : 0,
      yield_pct        : 0
    }
  };

  try {
    if (!payload || typeof payload !== 'object') {
      result.errors.push('Payload kosong atau invalid');
      result.ok = false;
      return result;
    }

    // ═══ V1: BATCH COIL + KG_PROSES ═══
    var coil = null;
    var kgProses = 0;
    if (!payload.batch_id) {
      result.errors.push('Batch coil belum dipilih');
    } else {
      coil = _slt_getInputCoilInfo(payload.batch_id);
      if (!coil.found) {
        result.errors.push('Batch coil ' + payload.batch_id + ' tidak ditemukan di Stok_Coil');
      } else {
        if (coil.kg_avail <= 0) {
          result.errors.push('Batch coil ' + payload.batch_id + ' sudah habis (KG_Avail: ' + coil.kg_avail + ')');
        }
        // kg_proses (default = kg_avail kalau tidak diinput)
        kgProses = Number(payload.kg_proses) || coil.kg_avail;
        if (kgProses > coil.kg_avail) {
          result.errors.push('Total Proses (' + kgProses + ' kg) melebihi KG Avail coil (' + coil.kg_avail + ' kg)');
          kgProses = coil.kg_avail;
        }
        if (kgProses <= 0) {
          result.errors.push('Total Proses harus > 0 kg');
        }
        // Length total = derived dari kg_proses (bukan kg_avail)
        var lengthProses = (coil.lebar_coil > 0 && coil.t > 0)
          ? Math.round(kgProses * 1000 / (coil.t * coil.lebar_coil * 7.85) * 100) / 100
          : 0;
        result.summary.lebar_coil     = coil.lebar_coil;
        result.summary.length_total_m = lengthProses;
        result.summary.kg_proses      = kgProses;
        result.summary.kg_recoil      = Math.max(0, coil.kg_avail - kgProses);
      }
    }

    // ═══ V2 & V3: CUT PLAN ═══
    var cutPlan = Array.isArray(payload.cut_plan) ? payload.cut_plan : [];
    var totalCut = cutPlan.length;
    if (totalCut < 1) {
      result.errors.push('Cut plan minimal 1 cut');
    } else if (totalCut > SLT_CONSTANTS.MAX_CUT_PER_JALUR) {
      result.errors.push('Cut plan maksimal ' + SLT_CONSTANTS.MAX_CUT_PER_JALUR + ' cut (input: ' + totalCut + ')');
    }
    for (var i = 0; i < cutPlan.length; i++) {
      var cLen = Number(cutPlan[i]) || 0;
      if (cLen <= 0) result.errors.push('Cut-' + (i+1) + ' harus > 0 meter (input: ' + cutPlan[i] + ')');
    }
    var lengthTerpakai = cutPlan.reduce(function(s, x) { return s + (Number(x) || 0); }, 0);
    result.summary.total_cut         = totalCut;
    result.summary.length_terpakai_m = Math.round(lengthTerpakai * 100) / 100;

    if (result.summary.length_total_m > 0 && lengthTerpakai > result.summary.length_total_m) {
      result.errors.push('Total cut plan (' + lengthTerpakai + 'm) melebihi length total coil (' +
                         result.summary.length_total_m + 'm)');
    }

    // ═══ V4-V7: SETTINGS ═══
    var s1 = payload.setting_1 || null;
    var s2 = payload.setting_2 || null;
    var isSplit = !!payload.is_split;

    if (!s1) {
      result.errors.push('Setting-1 tidak ada');
    } else {
      var r1 = _slt_validateSetting(s1, 1, result.summary.lebar_coil, isSplit);
      r1.errors.forEach(function(e) { result.errors.push(e); });
      r1.warnings.forEach(function(w) { result.warnings.push(w); });
      result.summary.total_jalur    += r1.totalJalur;
      result.summary.lebar_terpakai += r1.lebarTerpakai;
    }

    if (isSplit) {
      if (!s2 || !Array.isArray(s2.outs) || s2.outs.length === 0) {
        result.errors.push('Split diaktifkan tapi Setting-2 kosong');
      } else {
        var sisaLebar = result.summary.lebar_coil - (s1 ? (Number(s1.trim_kiri) || 0) + _slt_validateSetting(s1, 1, 0, true).sumLebarJalur + (Number(s1.trim_kanan) || 0) : 0);
        var r2 = _slt_validateSetting(s2, 2, sisaLebar, false);
        r2.errors.forEach(function(e) { result.errors.push(e); });
        r2.warnings.forEach(function(w) { result.warnings.push(w); });
        result.summary.total_jalur    += r2.totalJalur;
        result.summary.lebar_terpakai += r2.lebarTerpakai;
      }
    }
    result.summary.sisa_lebar = result.summary.lebar_coil - result.summary.lebar_terpakai;

    // ═══ V9-V10: OUT ROW FIELDS ═══
    var allOuts = [];
    if (s1 && Array.isArray(s1.outs)) {
      s1.outs.forEach(function(o, idx) { allOuts.push({ row: o, setting: 1, idx: idx + 1 }); });
    }
    if (isSplit && s2 && Array.isArray(s2.outs)) {
      s2.outs.forEach(function(o, idx) { allOuts.push({ row: o, setting: 2, idx: idx + 1 }); });
    }

    allOuts.forEach(function(x) {
      var lbl = 'Setting-' + x.setting + ' Jalur #' + x.idx;
      if (!x.row.item_code)   result.errors.push(lbl + ': Item_Code kosong');
      else if (!_slt_isValidItemCode(x.row.item_code)) {
        result.errors.push(lbl + ': Item_Code "' + x.row.item_code + '" tidak ada di M_ITEM');
      }
      if (!x.row.target_loc)  result.errors.push(lbl + ': Target_Loc kosong');
      if (!x.row.so_ref)      result.errors.push(lbl + ': SO_Ref kosong');
      if (!x.row.cust)        result.errors.push(lbl + ': Cust kosong');
      if (!x.row.owner_used)  result.errors.push(lbl + ': Owner_Used kosong');
      if (!Number(x.row.lebar) || Number(x.row.lebar) <= 0) {
        result.errors.push(lbl + ': Lebar harus > 0');
      }
      if (!Number(x.row.qty_jalur) || Number(x.row.qty_jalur) <= 0) {
        result.errors.push(lbl + ': Qty Jalur harus > 0');
      }
    });

    // ═══ V8: KG per roll ═══
    var tVal = coil ? coil.t : (Number(payload.t) || 0);
    if (tVal > 0 && cutPlan.length > 0) {
      var maxCutLen = Math.max.apply(null, cutPlan.map(function(x) { return Number(x) || 0; }));
      allOuts.forEach(function(x) {
        var lebar = Number(x.row.lebar) || 0;
        if (lebar > 0 && maxCutLen > 0) {
          // KG per roll = length(m) × lebar(mm) × t(mm) × 7.85 / 1000
          var kgPerRoll = maxCutLen * lebar * tVal * 7.85 / 1000;
          if (kgPerRoll > 500) {
            result.warnings.push('Setting-' + x.setting + ' Jalur #' + x.idx +
                                 ' (lebar ' + lebar + 'mm × cut ' + maxCutLen + 'm): ~' +
                                 Math.round(kgPerRoll) + ' kg/roll (limit customer umumnya 400-500kg)');
          }
        }
      });
    }

    // ═══ TOTAL BATCH & YIELD ═══
    result.summary.total_batch = result.summary.total_jalur * result.summary.total_cut;

    var totalKgOutput = 0;
    if (tVal > 0 && lengthTerpakai > 0) {
      allOuts.forEach(function(x) {
        var lebar = Number(x.row.lebar) || 0;
        var qty   = Number(x.row.qty_jalur) || 0;
        if (lebar > 0 && qty > 0) {
          totalKgOutput += qty * lengthTerpakai * lebar * tVal * 7.85 / 1000;
        }
      });
    }
    result.summary.total_kg_output = Math.round(totalKgOutput);

    var kgInput = coil ? coil.kg_avail : (Number(payload.kg_input) || 0);
    if (kgInput > 0) {
      result.summary.yield_pct = Math.round((totalKgOutput / kgInput) * 1000) / 10;
    }

    // Final: kalau ada errors → ok=false
    if (result.errors.length > 0) result.ok = false;

    return result;

  } catch (e) {
    Logger.log('Error validateSlitingPlan: ' + e.toString());
    result.ok = false;
    result.errors.push('Error internal validasi: ' + e.toString());
    return result;
  }
}

// ═══════════════════════════════════════════════════════════════════
// SECTION: SAVE DRAFT
// ═══════════════════════════════════════════════════════════════════

/**
 * Helper safe append row. Fallback ke sheet.appendRow kalau _appendRowSafe belum ada di Utils.
 * @param {Sheet} sheet
 * @param {Array} row
 */
function _slt_appendRow(sheet, row) {
  if (typeof _appendRowSafe === 'function') {
    _appendRowSafe(sheet, row);
  } else {
    sheet.appendRow(row);
  }
}

/**
 * Generate SPK_No berikutnya untuk SLT-HEADER.
 * Baca & increment LAST_SLT di M_Config.
 * Format: SLT-YYNNNN (4 digit padded)
 * Contoh: SLT-260001 (SPK ke-1 tahun 2026)
 *
 * @return {string} SPK_No baru
 * @throws Error kalau sheet/key tidak ditemukan
 */
function generateSpkSltNo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('M_Config');
  if (!sh) throw new Error('Sheet M_Config tidak ditemukan');

  var data = sh.getDataRange().getValues();
  var yy = String(new Date().getFullYear()).slice(-2); // "26"
  var keyName = 'LAST_SLT';

  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === keyName) {
      rowIdx = i;
      break;
    }
  }

  if (rowIdx === -1) throw new Error('Key ' + keyName + ' tidak ada di M_Config');

  var currentVal = Number(data[rowIdx][1]) || 0;
  var newVal = currentVal + 1;

  sh.getRange(rowIdx + 1, 2).setValue(newVal);
  SpreadsheetApp.flush();

  var spkNo = 'SLT-' + yy + String(newVal).padStart(4, '0');
  return spkNo;
}

/**
 * Encode NOTE untuk SLT-HEADER row.
 * Format eksplisit (bukan singkatan) supaya operator paham langsung.
 *
 * Contoh output:
 * "Cut Plan: 3 cuts (Cut-1: 120m, Cut-2: 120m, Cut-3: 100m) | Split: Ya |
 *  Setting-1 Trim: Kiri 5mm Kanan 0mm | Setting-2 Trim: Kiri 5mm Kanan 5mm |
 *  Lebar Coil: 1219mm | Length Total: 351.5m | Total Jalur: 14 | Total Roll: 42"
 */
function _slt_encodeHeaderNote(payload, coilInfo, totalJalur, totalRoll, kgProses) {
  var parts = [];

  // Cut Plan
  var cutPlan = payload.cut_plan || [];
  var cutStr = cutPlan.map(function(m, i) {
    return 'Cut-' + (i + 1) + ': ' + m + 'm';
  }).join(', ');
  parts.push('Cut Plan: ' + cutPlan.length + ' cuts (' + cutStr + ')');

  // Split flag
  parts.push('Split: ' + (payload.is_split ? 'Ya' : 'Tidak'));

  // Setting-1 Trim
  var s1 = payload.setting_1 || {};
  parts.push('Setting-1 Trim: Kiri ' + (Number(s1.trim_kiri) || 0) + 'mm Kanan ' + (Number(s1.trim_kanan) || 0) + 'mm');

  // Setting-2 Trim (kalau split)
  if (payload.is_split) {
    var s2 = payload.setting_2 || {};
    parts.push('Setting-2 Trim: Kiri ' + (Number(s2.trim_kiri) || 0) + 'mm Kanan ' + (Number(s2.trim_kanan) || 0) + 'mm');
  }

  // Coil & totals
  parts.push('Lebar Coil: ' + coilInfo.lebar_coil + 'mm');
  parts.push('KG Coil Avail: ' + coilInfo.kg_avail + ' kg');

  // Total Proses & Recoil
  var kgP = Number(kgProses) || coilInfo.kg_avail;
  var kgRecoil = Math.max(0, coilInfo.kg_avail - kgP);
  var lengthProses = (coilInfo.t > 0 && coilInfo.lebar_coil > 0)
    ? Math.round(kgP * 1000 / (coilInfo.t * coilInfo.lebar_coil * 7.85) * 100) / 100
    : 0;
  parts.push('Total Proses: ' + kgP + ' kg (' + lengthProses + 'm)');
  if (kgRecoil > 0) {
    var lengthRecoil = (coilInfo.t > 0 && coilInfo.lebar_coil > 0)
      ? Math.round(kgRecoil * 1000 / (coilInfo.t * coilInfo.lebar_coil * 7.85) * 100) / 100
      : 0;
    parts.push('Recoil: ' + kgRecoil + ' kg (' + lengthRecoil + 'm)');
  }

  parts.push('Total Jalur: ' + totalJalur);
  parts.push('Total Roll: ' + totalRoll);

  return parts.join(' | ');
}

/**
 * Encode NOTE untuk SLT-OUT row (per roll).
 * Contoh: "Jalur 4 dari 14 | Cut 2 dari 3 | Setting-1 | Lebar 85mm | Length 120m"
 */
function _slt_encodeOutNote(jalurGlobal, totalJalur, cutIdx, totalCut, settingNo, lebar, lengthM) {
  var parts = [];
  parts.push('Jalur ' + jalurGlobal + ' dari ' + totalJalur);
  parts.push('Cut ' + cutIdx + ' dari ' + totalCut);
  parts.push('Setting-' + settingNo);
  parts.push('Lebar ' + lebar + 'mm');
  parts.push('Length ' + lengthM + 'm');
  return parts.join(' | ');
}

/**
 * Auto-lock Owner_Used berdasarkan aturan bisnis SLT.
 * Konsisten dengan pola SHR/CTL:
 * - Cust=DRC → auto DRC
 * - Target_Loc=Scrap_Area → auto FC
 * - Target_Loc=Stok_Coil → coil_owner (mother coil)
 * - Else: pakai input user (fallback: coil_owner)
 */
function _slt_resolveOwnerUsed(row, coilOwner) {
  var cust      = String(row.cust        || '').trim().toUpperCase();
  var tgtLoc    = String(row.target_loc  || '').trim();
  var userInput = String(row.owner_used  || '').trim().toUpperCase();
  var fallback  = String(coilOwner       || 'FC').trim().toUpperCase();

  if (cust === 'DRC') return 'DRC';
  if (tgtLoc === 'Scrap_Area') return 'FC';
  if (tgtLoc === 'Stok_Coil') return fallback;
  return userInput || fallback;
}

/**
 * Main save function untuk SPK SLT (draft ANTRIAN).
 *
 * Flow:
 *   1. Validate payload via validateSlitingPlan()
 *   2. Lookup coil info from Stok_Coil
 *   3. Generate SPK_No (SLT-YYNNNN)
 *   4. Compute CT via kalkulasiEstimasiWaktu_SLT()
 *   5. Explode payload → build 1 HEADER + N SLT-OUT rows (per jalur × per cut)
 *   6. Write to sheet SPK
 *   7. Trigger scheduling engine (kalkulasiEstimasiWaktu global)
 *
 * @param {object} payload - Struktur payload SPK SLT (lihat design D1)
 * @return {object} { success, spk_no, count_out, total_jalur, total_roll,
 *                    kg_target, ct, message, warnings }
 */
function saveSPK_SLT(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // ═══ 1. VALIDATE ═══
    var validation = validateSlitingPlan(payload);
    if (!validation.ok) {
      return {
        success: false,
        message: 'Validasi gagal:\n• ' + validation.errors.join('\n• '),
        errors: validation.errors
      };
    }

    // ═══ 2. GET COIL INFO ═══
    var coil = _slt_getInputCoilInfo(payload.batch_id);
    if (!coil.found) {
      return { success: false, message: 'Batch coil ' + payload.batch_id + ' tidak ditemukan' };
    }

    // ═══ 3. PREP SHEET ═══
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var spkSheet = ss.getSheetByName('SPK');
    if (!spkSheet) throw new Error('Sheet SPK tidak ditemukan');

    var headers = spkSheet.getRange(1, 1, 1, spkSheet.getLastColumn()).getValues()[0]
                  .map(function(h) { return String(h).trim(); });

    function buildRow(rowData) {
      return headers.map(function(h) {
        return rowData[h] !== undefined ? rowData[h] : '';
      });
    }

    var timestamp    = new Date();
    var cutPlan      = payload.cut_plan || [];
    var totalCut     = cutPlan.length;
    var totalLengthM = cutPlan.reduce(function(s, x) { return s + (Number(x) || 0); }, 0);
    var isSplit      = !!payload.is_split;
    var mcNo         = payload.mc_no || 'SLT-01';
    var coilOwner    = String(coil.owner || 'FC').trim().toUpperCase();

    // ═══ 4. HITUNG TOTAL JALUR ═══
    var totalJalur = 0;
    var settingList = [payload.setting_1];
    if (isSplit) settingList.push(payload.setting_2);
    settingList.forEach(function(s) {
      if (s && Array.isArray(s.outs)) {
        s.outs.forEach(function(o) { totalJalur += (Number(o.qty_jalur) || 0); });
      }
    });
    var totalRoll = totalJalur * totalCut;

    // ═══ 5. GENERATE SPK_NO & CT ═══
    var spkNo = generateSpkSltNo();
    var ct = kalkulasiEstimasiWaktu_SLT(mcNo, totalJalur, totalLengthM, totalCut);

    // ═══ 6. BUILD SLT-OUT rows (explode jalur × cut) ═══
    var outRows = [];
    var kgTotalHeader = 0;
    var jalurGlobal = 0;

    settingList.forEach(function(settingObj, settingIdx) {
      if (!settingObj || !Array.isArray(settingObj.outs)) return;
      var settingNo = settingIdx + 1;

      settingObj.outs.forEach(function(specRow) {
        var qtyJalur  = Number(specRow.qty_jalur) || 0;
        var lebar     = Number(specRow.lebar) || 0;
        var itemCode  = String(specRow.item_code || '');
        var targetLoc = String(specRow.target_loc || 'Stok_FG');
        var soRef     = String(specRow.so_ref || '');
        var cust      = String(specRow.cust || '');
        var ownerUsed = _slt_resolveOwnerUsed(specRow, coilOwner);

        // Loop per jalur individu
        for (var j = 0; j < qtyJalur; j++) {
          jalurGlobal++;
          var jalurStr = String(jalurGlobal).padStart(2, '0');

          // Loop per cut
          for (var c = 0; c < totalCut; c++) {
            var cutLen = Number(cutPlan[c]) || 0;
            var cutIdx = c + 1;

            // KG per roll = length × lebar × T × 7.85 / 1000
            var kgRoll = Math.round(cutLen * lebar * coil.t * 7.85 / 1000);

            var outSpkNo = spkNo + '-J' + jalurStr + '-C' + cutIdx;
            var outNote  = _slt_encodeOutNote(jalurGlobal, totalJalur, cutIdx, totalCut, settingNo, lebar, cutLen);

            var outRow = {
              'SPK_No'             : outSpkNo,
              'SPK_Type'           : 'SLT-OUT',
              'Parent_SPK'         : spkNo,
              'Batch_ID'           : payload.batch_id,
              'Tgl_Buat'           : timestamp,
              'Priority'           : payload.priority || 'Normal',
              'Source_Loc'         : 'Stok_Coil',
              'Item_Code'          : itemCode,
              'Input_Spec'         : _slt_getItemDescription(itemCode) || (lebar + 'mm x ' + cutLen + 'm'),
              'BQ'                 : 1,
              'Qty_Target'         : kgRoll,   // Qty = KG (per aturan SLT)
              'KG_Target'          : kgRoll,
              'MC_No'              : mcNo,
              'Status'             : 'ANTRIAN',
              'Created_By'         : payload.created_by || 'PPIC',
              'Owner'              : coilOwner,
              'Owner_Used'         : ownerUsed,
              'Target_Loc'         : targetLoc,
              'SO_Ref'             : soRef,
              'Cust'               : cust,
              'Plan_Setup_Menit'   : 0,
              'Plan_Run_Menit'     : 0,
              'Total_Durasi_Menit' : 0,   // OUT tidak masuk scheduling engine
              'T'                  : coil.t,
              'NOTE'               : outNote
            };

            outRows.push(outRow);
            kgTotalHeader += kgRoll;
          }
        }
      });
    });

    // ═══ 7. BUILD SLT-HEADER ═══
    // KG_Target HEADER = kg_proses (yg masuk mesin), bukan sum OUT.
    // Selisih kg_proses - sum(OUT) = trim loss.
    // Selisih kg_avail - kg_proses = recoil (mother coil masih avail).
    var kgProsesHeader = Number(payload.kg_proses) || coil.kg_avail;
    if (kgProsesHeader > coil.kg_avail) kgProsesHeader = coil.kg_avail;

    var headerNote = _slt_encodeHeaderNote(payload, coil, totalJalur, totalRoll, kgProsesHeader);

    var headerRow = {
      'SPK_No'             : spkNo,
      'SPK_Type'           : 'SLT-HEADER',
      'Parent_SPK'         : '',
      'Batch_ID'           : payload.batch_id,
      'Tgl_Buat'           : timestamp,
      'Priority'           : payload.priority || 'Normal',
      'Source_Loc'         : 'Stok_Coil',
      'Item_Code'          : coil.item_code,
      'Input_Spec'         : _slt_getItemDescription(coil.item_code) || coil.spec,
      'BQ'                 : 1,
      'Qty_Target'         : kgProsesHeader,
      'KG_Target'          : kgProsesHeader,
      'MC_No'              : mcNo,
      'Status'             : 'ANTRIAN',
      'Created_By'         : payload.created_by || 'PPIC',
      'Owner'              : coilOwner,
      'Owner_Used'         : coilOwner,
      'Plan_Setup_Menit'   : ct.planSetup,
      'Plan_Run_Menit'     : ct.planRun,
      'Total_Durasi_Menit' : ct.totalDurasi,
      'T'                  : coil.t,
      'NOTE'               : headerNote
    };

    // ═══ 8. WRITE ke sheet SPK ═══
    // HEADER first
    _slt_appendRow(spkSheet, buildRow(headerRow));
    // OUTs
    outRows.forEach(function(r) {
      _slt_appendRow(spkSheet, buildRow(r));
    });

    SpreadsheetApp.flush();

    // ═══ 9. TRIGGER SCHEDULING ═══
    if (typeof kalkulasiEstimasiWaktu === 'function') {
      try {
        kalkulasiEstimasiWaktu();
      } catch (e) {
        Logger.log('Warn: kalkulasiEstimasiWaktu error: ' + e.toString());
      }
    }

    // ═══ 10. RETURN ═══
    return {
      success     : true,
      spk_no      : spkNo,
      count_out   : outRows.length,
      total_jalur : totalJalur,
      total_roll  : totalRoll,
      kg_target   : kgTotalHeader,
      ct          : ct,
      message     : 'SPK ' + spkNo + ' berhasil dibuat: ' + outRows.length +
                    ' row SLT-OUT (' + totalJalur + ' jalur × ' + totalCut + ' cut)',
      warnings    : validation.warnings || []
    };

  } catch (e) {
    Logger.log('Error saveSPK_SLT: ' + e.toString() + '\n' + (e.stack || ''));
    return { success: false, message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════════
// SECTION: DONE HANDLER + BATCH TURUNAN
// ═══════════════════════════════════════════════════════════════════
//
// FLOW 3-ROLE HANDOFF:
//
//   ANTRIAN
//     ↓ operator: start di board
//   RUNNING
//     ↓ operator: klik "Cut-N Selesai" tiap cut fisik keluar dari mesin
//     ↓ input: konfirmasi jumlah roll fisik (default = qty_target, adjust kalau ada NG)
//     ↓ [helper] cek Cut_Progress = total_cut ? → auto-transit
//   MESIN_SELESAI  🚦 mesin bebas untuk SPK berikutnya
//     ↓ packing: input berat aktual per roll (bertahap, bisa lintas hari)
//     ↓ [helper] cek semua SLT-OUT.Qty_Actual > 0 ? → auto-transit
//   MENUNGGU_APPROVAL  🚦 PPIC review antrian
//     ↓ ppic: klik "Tutup SPK" di menu approval
//     ↓ [action] validate 100% roll ditimbang
//     ↓ [action] generate batch turunan → Trace_Log
//     ↓ [action] Rekap_ICT auto (kalau Owner ≠ Owner_Used)
//   DONE
//
// PUBLIC FUNCTIONS:
//   - sltMarkCutSelesai(payload)     → Operator (Board Mesin SLT)
//   - sltInputWeightBatch(payload)   → Packing (Page Packing baru)
//   - sltCloseSpkFinal(payload)      → PPIC (Menu Approval)
//
// PRIVATE HELPERS: _slt_xxx
// ═══════════════════════════════════════════════════════════════════

// ─── STATUS CONSTANTS ─────────────────────────────────────────────
var SLT_STATUS = {
  ANTRIAN           : 'ANTRIAN',
  RUNNING           : 'RUNNING',
  MESIN_SELESAI     : 'MESIN_SELESAI',
  MENUNGGU_APPROVAL : 'MENUNGGU_APPROVAL',
  DONE              : 'DONE',
  CANCELLED         : 'CANCELLED'
};

// ═══════════════════════════════════════════════════════════════════
// PRIVATE HELPERS
// ═══════════════════════════════════════════════════════════════════

/**
 * Get SPK sheet + parsed headers + column index map.
 * Return: { sheet, data, headers, I } — I = column index dict
 */
function _slt_getSpkSheetContext() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SPK');
  if (!sheet) throw new Error('Sheet SPK tidak ditemukan');

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Sheet SPK kosong');

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var I = {
    spk      : headers.indexOf('SPK_No'),
    type     : headers.indexOf('SPK_Type'),
    parent   : headers.indexOf('Parent_SPK'),
    status   : headers.indexOf('Status'),
    qtyTgt   : headers.indexOf('Qty_Target'),
    kgTgt    : headers.indexOf('KG_Target'),
    qtyAct   : headers.indexOf('Qty_Actual'),
    kgAct    : headers.indexOf('KG_Actual'),
    qtyNg    : headers.indexOf('Qty_NG'),
    kgNg     : headers.indexOf('KG_NG'),
    item     : headers.indexOf('Item_Code'),
    mulai    : headers.indexOf('Mulai_DT'),
    selesai  : headers.indexOf('Selesai_DT'),
    ownerU   : headers.indexOf('Owner_Used'),
    owner    : headers.indexOf('Owner'),
    tgtLoc   : headers.indexOf('Target_Loc'),
    soRef    : headers.indexOf('SO_Ref'),
    cust     : headers.indexOf('Cust'),
    batchId  : headers.indexOf('Batch_ID'),
    inputSpec: headers.indexOf('Input_Spec'),
    t        : headers.indexOf('T'),
    note     : headers.indexOf('NOTE'),
    op       : headers.indexOf('OP'),
    leader   : headers.indexOf('Leader'),
    cutProg  : headers.indexOf('Cut_Progress')
  };

  if (I.cutProg === -1) {
    throw new Error('Kolom Cut_Progress belum ditambahkan di sheet SPK. Lihat instruksi step prerequisite.');
  }

  return { sheet: sheet, data: data, headers: headers, I: I };
}

/**
 * Find SPK-HEADER row by SPK_No (SLT-HEADER).
 * Return: { rowIdx, rowData } — rowIdx = 1-based, atau null kalau tidak ketemu
 */
function _slt_findHeader(data, I, spkNo) {
  var target = String(spkNo || '').trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][I.spk] || '').trim() === target &&
        String(data[i][I.type] || '').trim() === 'SLT-HEADER') {
      return { rowIdx: i + 1, rowData: data[i] };
    }
  }
  return null;
}

/**
 * Get semua SLT-OUT row untuk 1 SPK.
 * Return: array of { rowIdx, rowData }
 */
function _slt_getAllOuts(data, I, spkHeaderNo) {
  var target = String(spkHeaderNo || '').trim();
  var outs = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][I.parent] || '').trim() === target &&
        String(data[i][I.type] || '').trim() === 'SLT-OUT') {
      outs.push({ rowIdx: i + 1, rowData: data[i] });
    }
  }
  return outs;
}

/**
 * Parse Cut_Progress "1,2,3" → [1,2,3]
 */
function _slt_parseCutProgress(str) {
  if (!str) return [];
  return String(str).split(',')
    .map(function(x) { return parseInt(x, 10); })
    .filter(function(x) { return !isNaN(x) && x > 0; });
}

/**
 * Extract total_cut dari NOTE HEADER.
 * Format NOTE: "Cut Plan: 3 cuts (Cut-1: 120m, ...) | ..."
 */
function _slt_extractTotalCut(noteStr) {
  var m = String(noteStr || '').match(/Cut Plan:\s*(\d+)\s*cuts/i);
  return m ? parseInt(m[1], 10) : 0;
}
/**
 * Parse cut plan array dari NOTE HEADER.
 * Format NOTE: "Cut Plan: 3 cuts (Cut-1: 120m, Cut-2: 120m, Cut-3: 100m) | ..."
 * Return: [120, 120, 100]
 */
function _slt_parseCutPlan(noteStr) {
  var arr = [];
  if (!noteStr) return arr;
  var re = /Cut-\d+:\s*([\d.]+)\s*m/g;
  var m;
  while ((m = re.exec(noteStr)) !== null) {
    arr.push(parseFloat(m[1]) || 0);
  }
  return arr;
}

/**
 * Parse trim + split info dari NOTE HEADER.
 * Return: { is_split, s1_trim_kiri, s1_trim_kanan, s2_trim_kiri, s2_trim_kanan, lebar_coil }
 */
function _slt_parseTrimInfo(noteStr) {
  var info = {
    is_split: false,
    s1_trim_kiri: 0, s1_trim_kanan: 0,
    s2_trim_kiri: 0, s2_trim_kanan: 0,
    lebar_coil: 0
  };
  if (!noteStr) return info;

  var mSplit = noteStr.match(/Split:\s*(\w+)/i);
  if (mSplit) info.is_split = /ya|yes|true/i.test(mSplit[1]);

  var mS1 = noteStr.match(/Setting-1 Trim:\s*Kiri\s*(\d+)mm\s*Kanan\s*(\d+)mm/i);
  if (mS1) {
    info.s1_trim_kiri = parseInt(mS1[1], 10) || 0;
    info.s1_trim_kanan = parseInt(mS1[2], 10) || 0;
  }

  var mS2 = noteStr.match(/Setting-2 Trim:\s*Kiri\s*(\d+)mm\s*Kanan\s*(\d+)mm/i);
  if (mS2) {
    info.s2_trim_kiri = parseInt(mS2[1], 10) || 0;
    info.s2_trim_kanan = parseInt(mS2[2], 10) || 0;
  }

  var mLC = noteStr.match(/Lebar Coil:\s*(\d+)mm/i);
  if (mLC) info.lebar_coil = parseInt(mLC[1], 10) || 0;

  return info;
}

/** Parse lebar dari NOTE OUT: "Lebar 85mm" → 85 */
function _slt_parseLebarFromNote(outNote) {
  if (!outNote) return 0;
  var m = outNote.match(/Lebar\s*(\d+)\s*mm/i);
  return m ? (parseInt(m[1], 10) || 0) : 0;
}

/** Parse setting number dari NOTE OUT: "Setting-2" → 2 */
function _slt_parseSettingFromNote(outNote) {
  if (!outNote) return 1;
  var m = outNote.match(/Setting-(\d+)/i);
  return m ? (parseInt(m[1], 10) || 1) : 1;
}
/**
 * Parse SPK_No SLT-OUT: "SLT-260001-J05-C2" → { spk_header: 'SLT-260001', jalur: 5, cut: 2 }
 * Return null kalau format tidak match
 */
function _slt_parseOutSpkNo(spkNo) {
  var m = String(spkNo || '').match(/^(SLT-\d+)-J(\d+)-C(\d+)$/);
  if (!m) return null;
  return {
    spk_header : m[1],
    jalur      : parseInt(m[2], 10),
    cut        : parseInt(m[3], 10)
  };
}

/**
 * Lookup master data 1 item dari M_ITEM.
 * Return: { desc, spec, t, p, l } — kosong string kalau tidak ketemu
 */
function _slt_lookupItemMaster(itemCode) {
  var empty = { desc: '', spec: '', t: '', p: '', l: '' };
  try {
    if (!itemCode) return empty;
    var target = String(itemCode).trim();
    if (!target) return empty;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('M_ITEM');
    if (!sh) return empty;

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return empty;

    var hdr = data[0].map(function(h) { return String(h).trim(); });
    var iCode = hdr.indexOf('Item_Code');
    var iDesc = hdr.indexOf('Description');
    var iSpec = hdr.indexOf('Spec');
    var iT    = hdr.indexOf('T');
    var iP    = hdr.indexOf('P');
    var iL    = hdr.indexOf('L');
    if (iCode === -1) return empty;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iCode] || '').trim() === target) {
        return {
          desc: iDesc !== -1 ? String(data[i][iDesc] || '').trim() : '',
          spec: iSpec !== -1 ? String(data[i][iSpec] || '').trim() : '',
          t   : iT    !== -1 ? (data[i][iT] || '') : '',
          p   : iP    !== -1 ? (data[i][iP] || '') : '',
          l   : iL    !== -1 ? (data[i][iL] || '') : ''
        };
      }
    }
    return empty;
  } catch (e) {
    Logger.log('Error _slt_lookupItemMaster: ' + e.toString());
    return empty;
  }
}

// ═══════════════════════════════════════════════════════════════════
// PUBLIC: sltMarkCutSelesai — dipanggil OPERATOR dari Board Mesin
// ═══════════════════════════════════════════════════════════════════

/**
 * Operator input: cut ke-N sudah selesai dari mesin.
 *
 * @param {object} payload {
 *   spk_no          : 'SLT-260001',        // SLT-HEADER SPK_No
 *   cut_no          : 1,                    // cut yang selesai (1-based)
 *   roll_actual     : 14,                   // jumlah roll fisik yg keluar (default = qty jalur)
 *   ng_from_machine : 0,                    // roll NG dari mesin (crack visual dll)
 *   operator        : 'Operator-A',
 *   note            : 'roll #3 crack sedikit'  // optional
 * }
 * @return {object} { success, message, new_status, cut_progress, is_mesin_selesai }
 */
function sltMarkCutSelesai(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // ═══ 1. VALIDATE INPUT ═══
    if (!payload || !payload.spk_no) {
      return { success: false, message: 'spk_no kosong' };
    }
    var cutNo = parseInt(payload.cut_no, 10);
    if (!cutNo || cutNo < 1) {
      return { success: false, message: 'cut_no harus > 0' };
    }
    var rollActual = parseInt(payload.roll_actual, 10) || 0;
    if (rollActual < 0) {
      return { success: false, message: 'roll_actual tidak boleh negatif' };
    }
    var ngMachine = parseInt(payload.ng_from_machine, 10) || 0;

    // ═══ 2. GET SPK CONTEXT ═══
    var ctx = _slt_getSpkSheetContext();
    var header = _slt_findHeader(ctx.data, ctx.I, payload.spk_no);
    if (!header) {
      return { success: false, message: 'SLT-HEADER ' + payload.spk_no + ' tidak ditemukan' };
    }

    var currentStatus = String(header.rowData[ctx.I.status] || '').trim();
    if (currentStatus !== SLT_STATUS.RUNNING && currentStatus !== SLT_STATUS.ANTRIAN) {
      return {
        success: false,
        message: 'SPK ' + payload.spk_no + ' status ' + currentStatus +
                 ' (harus ANTRIAN atau RUNNING untuk mark cut selesai)'
      };
    }

    // ═══ 3. VALIDATE CUT PROGRESS ═══
    var noteStr = String(header.rowData[ctx.I.note] || '');
    var totalCut = _slt_extractTotalCut(noteStr);
    if (totalCut === 0) {
      return { success: false, message: 'Tidak bisa parse total_cut dari NOTE HEADER' };
    }
    if (cutNo > totalCut) {
      return { success: false, message: 'cut_no ' + cutNo + ' > total_cut ' + totalCut };
    }

    var progressStr = String(header.rowData[ctx.I.cutProg] || '');
    var progressArr = _slt_parseCutProgress(progressStr);

    if (progressArr.indexOf(cutNo) !== -1) {
      return { success: false, message: 'Cut ' + cutNo + ' sudah pernah ditandai selesai' };
    }

    // Append cut_no ke progress
    progressArr.push(cutNo);
    progressArr.sort(function(a, b) { return a - b; });
    var newProgressStr = progressArr.join(',');

    // ═══ 4. UPDATE HEADER ═══
    var timestamp = new Date();
    ctx.sheet.getRange(header.rowIdx, ctx.I.cutProg + 1).setValue(newProgressStr);

    // Set Mulai_DT kalau cut pertama
    var currentMulai = header.rowData[ctx.I.mulai];
    if (!currentMulai || currentMulai === '') {
      ctx.sheet.getRange(header.rowIdx, ctx.I.mulai + 1).setValue(timestamp);
    }

    // Update Status ke RUNNING kalau masih ANTRIAN
    if (currentStatus === SLT_STATUS.ANTRIAN) {
      ctx.sheet.getRange(header.rowIdx, ctx.I.status + 1).setValue(SLT_STATUS.RUNNING);
    }

    // Update OP kalau ada input operator
    if (payload.operator && ctx.I.op !== -1) {
      ctx.sheet.getRange(header.rowIdx, ctx.I.op + 1).setValue(String(payload.operator));
    }

    // Append NG dari mesin ke Qty_NG header
    if (ngMachine > 0 && ctx.I.qtyNg !== -1) {
      var currentNg = Number(header.rowData[ctx.I.qtyNg]) || 0;
      ctx.sheet.getRange(header.rowIdx, ctx.I.qtyNg + 1).setValue(currentNg + ngMachine);
    }

    // ═══ 5. CEK AUTO-TRANSIT ke MESIN_SELESAI ═══
    var isMesinSelesai = false;
    var newStatus = SLT_STATUS.RUNNING;

    if (progressArr.length >= totalCut) {
      ctx.sheet.getRange(header.rowIdx, ctx.I.status + 1).setValue(SLT_STATUS.MESIN_SELESAI);
      ctx.sheet.getRange(header.rowIdx, ctx.I.selesai + 1).setValue(timestamp);
      isMesinSelesai = true;
      newStatus = SLT_STATUS.MESIN_SELESAI;
    }

    SpreadsheetApp.flush();

    // ═══ 6. TRIGGER SCHEDULING ═══
    if (isMesinSelesai && typeof kalkulasiEstimasiWaktu === 'function') {
      try { kalkulasiEstimasiWaktu(); }
      catch (e) { Logger.log('Warn kalkulasiEstimasiWaktu: ' + e); }
    }

    // ═══ 7. RETURN ═══
    return {
      success          : true,
      message          : 'Cut ' + cutNo + ' ditandai selesai' +
                         (isMesinSelesai ? ' | Mesin bebas untuk SPK berikutnya' : ''),
      new_status       : newStatus,
      cut_progress     : newProgressStr,
      total_cut        : totalCut,
      cuts_done        : progressArr.length,
      is_mesin_selesai : isMesinSelesai
    };

  } catch (e) {
    Logger.log('Error sltMarkCutSelesai: ' + e.toString() + '\n' + (e.stack || ''));
    return { success: false, message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════════
// PUBLIC: sltInputWeightBatch — dipanggil PACKING dari Page Packing
// ═══════════════════════════════════════════════════════════════════

/**
 * Packing input berat aktual per roll (bertahap OK, bisa partial).
 *
 * @param {object} payload {
 *   spk_no  : 'SLT-260001',                   // SLT-HEADER SPK_No
 *   weights : [                                // partial OK (misal 20 dari 42)
 *     { out_spk: 'SLT-260001-J01-C1', kg_actual: 232, qty_ng: 0,  kg_ng: 0,  ng_reason: '' },
 *     { out_spk: 'SLT-260001-J01-C2', kg_actual: 230, qty_ng: 0,  kg_ng: 0,  ng_reason: '' },
 *     { out_spk: 'SLT-260001-J05-C2', kg_actual: 200, qty_ng: 25, kg_ng: 25, ng_reason: 'crack' },
 *     { out_spk: 'SLT-260001-J10-C1', kg_actual: 0,   qty_ng: 232,kg_ng: 232,ng_reason: 'Hilang' }
 *   ],
 *   packing_by : 'Packing-A'
 * }
 * @return {object} { success, message, updated_count, remaining_count, is_menunggu_approval }
 */
function sltInputWeightBatch(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // ═══ 1. VALIDATE ═══
    if (!payload || !payload.spk_no) {
      return { success: false, message: 'spk_no kosong' };
    }
    if (!Array.isArray(payload.weights) || payload.weights.length === 0) {
      return { success: false, message: 'weights array kosong' };
    }

    // ═══ 2. GET CONTEXT ═══
    var ctx = _slt_getSpkSheetContext();
    var header = _slt_findHeader(ctx.data, ctx.I, payload.spk_no);
    if (!header) {
      return { success: false, message: 'SLT-HEADER ' + payload.spk_no + ' tidak ditemukan' };
    }

    var currentStatus = String(header.rowData[ctx.I.status] || '').trim();
    if (currentStatus !== SLT_STATUS.MESIN_SELESAI && currentStatus !== SLT_STATUS.MENUNGGU_APPROVAL) {
      return {
        success: false,
        message: 'SPK ' + payload.spk_no + ' status ' + currentStatus +
                 ' (harus MESIN_SELESAI atau MENUNGGU_APPROVAL untuk input berat)'
      };
    }

    // ═══ 3. LOOP UPDATE PER OUT ═══
    var allOuts = _slt_getAllOuts(ctx.data, ctx.I, payload.spk_no);
    var outMap = {};
    allOuts.forEach(function(o) {
      outMap[String(o.rowData[ctx.I.spk]).trim()] = o;
    });

    var updatedCount = 0;
    var errors = [];

    payload.weights.forEach(function(w) {
      var outSpk = String(w.out_spk || '').trim();
      if (!outSpk) {
        errors.push('out_spk kosong');
        return;
      }
      var out = outMap[outSpk];
      if (!out) {
        errors.push('SLT-OUT ' + outSpk + ' tidak ditemukan');
        return;
      }

      var kgAct = Number(w.kg_actual) || 0;
      var qtyNg = Number(w.qty_ng)    || 0;
      var kgNg  = Number(w.kg_ng)     || 0;

      if (kgAct < 0 || qtyNg < 0 || kgNg < 0) {
        errors.push(outSpk + ': nilai negatif tidak valid');
        return;
      }

      // Write ke sheet (Qty_Actual = KG_Actual untuk SLT)
      ctx.sheet.getRange(out.rowIdx, ctx.I.qtyAct + 1).setValue(kgAct);
      ctx.sheet.getRange(out.rowIdx, ctx.I.kgAct + 1).setValue(kgAct);
      if (qtyNg > 0 || kgNg > 0) {
        ctx.sheet.getRange(out.rowIdx, ctx.I.qtyNg + 1).setValue(qtyNg);
        ctx.sheet.getRange(out.rowIdx, ctx.I.kgNg + 1).setValue(kgNg);
      }

      // Append NG reason ke NOTE OUT (kalau ada)
      if (w.ng_reason && String(w.ng_reason).trim() !== '') {
        var oldNote = String(out.rowData[ctx.I.note] || '');
        var newNote = oldNote + ' | NG: ' + String(w.ng_reason).trim();
        ctx.sheet.getRange(out.rowIdx, ctx.I.note + 1).setValue(newNote);
      }

      updatedCount++;
    });

    if (errors.length > 0) {
      SpreadsheetApp.flush();
      return {
        success: false,
        message: 'Sebagian gagal: ' + errors.join('; '),
        updated_count: updatedCount,
        errors: errors
      };
    }

    // ═══ 4. RE-READ untuk cek all weights done ═══
    SpreadsheetApp.flush();
    ctx = _slt_getSpkSheetContext();  // re-read data
    var allOutsAfter = _slt_getAllOuts(ctx.data, ctx.I, payload.spk_no);

    var remaining = 0;
    allOutsAfter.forEach(function(o) {
      var kgAct = Number(o.rowData[ctx.I.kgAct]) || 0;
      var kgNg  = Number(o.rowData[ctx.I.kgNg])  || 0;
      // "Sudah ditimbang" = kg_actual > 0 OR (kg_ng > 0 dan ada ng_reason)
      if (kgAct <= 0 && kgNg <= 0) remaining++;
    });

    // ═══ 5. AUTO-TRANSIT ke MENUNGGU_APPROVAL ═══
    var isMenungguApproval = false;
    var headerAfter = _slt_findHeader(ctx.data, ctx.I, payload.spk_no);
    if (remaining === 0 && String(headerAfter.rowData[ctx.I.status] || '').trim() === SLT_STATUS.MESIN_SELESAI) {
      ctx.sheet.getRange(headerAfter.rowIdx, ctx.I.status + 1).setValue(SLT_STATUS.MENUNGGU_APPROVAL);
      isMenungguApproval = true;
      SpreadsheetApp.flush();
    }

    return {
      success              : true,
      message              : 'Berhasil update ' + updatedCount + ' roll' +
                             (isMenungguApproval ? ' | SPK siap approval PPIC' : ' | Sisa ' + remaining + ' roll'),
      updated_count        : updatedCount,
      remaining_count      : remaining,
      is_menunggu_approval : isMenungguApproval
    };

  } catch (e) {
    Logger.log('Error sltInputWeightBatch: ' + e.toString() + '\n' + (e.stack || ''));
    return { success: false, message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════════
// PUBLIC: sltCloseSpkFinal — dipanggil PPIC dari Menu Approval
// ═══════════════════════════════════════════════════════════════════

/**
 * PPIC final approval: close SPK + generate batch turunan ke Trace_Log.
 *
 * Menggunakan helper existing (writeTraceLog, writeRekapICT, getRootBatch,
 * getCoilSupplierInfo) supaya konsisten dengan pattern CTL/SHR.
 *
 * @param {object} payload {
 *   spk_no      : 'SLT-260001',
 *   approved_by : 'Herryna'
 * }
 * @return {object} { success, message, batch_generated, ict_generated, total_kg_actual }
 */
function sltCloseSpkFinal(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // ═══ 1. VALIDATE ═══
    if (!payload || !payload.spk_no) {
      return { success: false, message: 'spk_no kosong' };
    }
    var approvedBy = String(payload.approved_by || 'PPIC').trim();

    // ═══ 2. GET CONTEXT ═══
    var ctx = _slt_getSpkSheetContext();
    var header = _slt_findHeader(ctx.data, ctx.I, payload.spk_no);
    if (!header) {
      return { success: false, message: 'SLT-HEADER ' + payload.spk_no + ' tidak ditemukan' };
    }

    var currentStatus = String(header.rowData[ctx.I.status] || '').trim();
    if (currentStatus !== SLT_STATUS.MENUNGGU_APPROVAL) {
      return {
        success: false,
        message: 'SPK ' + payload.spk_no + ' status ' + currentStatus +
                 ' (harus MENUNGGU_APPROVAL untuk close). Pastikan semua roll sudah ditimbang.'
      };
    }

    // ═══ 3. VALIDATE 100% WEIGHTED ═══
    var allOuts = _slt_getAllOuts(ctx.data, ctx.I, payload.spk_no);
    if (allOuts.length === 0) {
      return { success: false, message: 'Tidak ada SLT-OUT untuk SPK ' + payload.spk_no };
    }

    var missing = [];
    allOuts.forEach(function(o) {
      var kgAct = Number(o.rowData[ctx.I.kgAct]) || 0;
      var kgNg  = Number(o.rowData[ctx.I.kgNg])  || 0;
      if (kgAct <= 0 && kgNg <= 0) {
        missing.push(String(o.rowData[ctx.I.spk]));
      }
    });

    if (missing.length > 0) {
      return {
        success: false,
        message: 'Ada ' + missing.length + ' roll belum ditimbang: ' + missing.slice(0, 3).join(', ') +
                 (missing.length > 3 ? ' dan ' + (missing.length - 3) + ' lagi...' : '')
      };
    }

    // ═══ 4. PREP DATA COMMON ═══
    var timestamp    = new Date();
    var batchIdCoil  = String(header.rowData[ctx.I.batchId] || '').trim();
    var mcNoHeader   = String(header.rowData[ctx.I.mc]      || '').trim() ||
                       String(header.rowData[ctx.I.mc !== -1 ? ctx.I.mc : 0] || 'SLT-01').trim();
    var mcNoFinal    = ctx.headers.indexOf('MC_No') !== -1
                       ? String(header.rowData[ctx.headers.indexOf('MC_No')] || 'SLT-01').trim()
                       : 'SLT-01';
    var opHeader     = ctx.I.op !== -1 ? String(header.rowData[ctx.I.op] || '').trim() : '';
    var ownerHeader  = String(header.rowData[ctx.I.owner]  || '').trim();
    var tglProdFinal = header.rowData[ctx.I.selesai] || timestamp;  // Q2a: dari Selesai_DT HEADER

    // Root batch + supplier info (reuse existing helpers)
    var rootBatch = (typeof getRootBatch === 'function')
                    ? (getRootBatch(batchIdCoil) || batchIdCoil)
                    : batchIdCoil;
    var supplierInfo = (typeof getCoilSupplierInfo === 'function')
                       ? getCoilSupplierInfo(rootBatch)
                       : { supplier: '', no_po: '', no_do: '' };

    // ═══ 5. LOOP UPDATE OUT + GENERATE BATCH TURUNAN ═══
    var totalKgActual = 0;
    var totalKgNg     = 0;
    var batchCount    = 0;
    var ictCount      = 0;

    allOuts.forEach(function(o) {
      var outSpk    = String(o.rowData[ctx.I.spk]).trim();
      var itemCode  = String(o.rowData[ctx.I.item] || '').trim();
      var kgAct     = Number(o.rowData[ctx.I.kgAct]) || 0;
      var kgNg      = Number(o.rowData[ctx.I.kgNg])  || 0;
      var tgtLoc    = String(o.rowData[ctx.I.tgtLoc] || '').trim();
      var ownerUsed = String(o.rowData[ctx.I.ownerU] || '').trim();
      var soRef     = String(o.rowData[ctx.I.soRef] || '').trim();
      var cust      = String(o.rowData[ctx.I.cust]  || '').trim();
      var oldNote   = String(o.rowData[ctx.I.note]  || '').trim();

      // Update SLT-OUT Status=DONE, Selesai_DT
      ctx.sheet.getRange(o.rowIdx, ctx.I.status + 1).setValue(SLT_STATUS.DONE);
      ctx.sheet.getRange(o.rowIdx, ctx.I.selesai + 1).setValue(timestamp);

      totalKgActual += kgAct;
      totalKgNg     += kgNg;

      // Skip batch turunan kalau kg_actual = 0 (roll hilang / all NG)
      if (kgAct <= 0) return;

      // Parse jalur & cut dari outSpk
      var parsed = _slt_parseOutSpkNo(outSpk);
      if (!parsed) {
        Logger.log('WARN: cannot parse outSpk ' + outSpk + ', skip batch turunan');
        return;
      }

      // Generate Batch_ID: FC-SL-{YY}{NNNNN}-J{JJ}-C{C}
      var batchPrefix = '';
      if (tgtLoc === 'Stok_Coil') {
        // Sisa coil dari SLT → prefix FCL
        batchPrefix = (typeof generateBatchId === 'function')
                      ? generateBatchId('FCL')
                      : ('FC-CL-' + String(new Date().getFullYear()).slice(-2) + String(Date.now()).slice(-5));
      } else {
        // Strip FG → prefix FSL
        batchPrefix = (typeof generateBatchId === 'function')
                      ? generateBatchId('FSL')
                      : generateFslBatchId();
      }
      var jalurStr = String(parsed.jalur).padStart(2, '0');
      var batchIdFull = batchPrefix + '-J' + jalurStr + '-C' + parsed.cut;

      // Lookup M_ITEM untuk desc/spec/t/p/l (Q2c: L_dim dari M_ITEM, konsisten CTL/SHR)
      var mi = _slt_lookupItemMaster(itemCode);

      // Determine trace type (konsisten pattern CTL/SHR)
      var traceType = 'FGS';  // default
      if (tgtLoc === 'Stok_Coil') {
        traceType = 'COIL';
      } else if (tgtLoc && tgtLoc.indexOf('WIP_') === 0) {
        traceType = 'WIP';
      } else if (tgtLoc && tgtLoc.indexOf('FG_') === 0) {
        traceType = (tgtLoc === 'FG_Cust') ? 'FGC' : 'FGS';
      } else if (tgtLoc === 'Stok_Sheet') {
        traceType = 'SHEET';
      }

      // ═══ WRITE TRACE_LOG (pakai helper existing) ═══
      if (typeof writeTraceLog === 'function') {
        try {
          writeTraceLog({
            batch_id     : batchIdFull,
            tgl_buat     : timestamp,
            level        : 1,                         // konsisten CTL: input dari coil GR → level 1
            type         : traceType,
            source_batch : batchIdCoil,
            root_batch   : rootBatch,
            spk_ref      : payload.spk_no,
            gr_ref       : rootBatch,
            item_code    : itemCode,
            description  : mi.desc,
            spec         : mi.spec,
            t            : mi.t,
            p            : mi.p,                      // dari M_ITEM (lebar strip)
            l_dim        : mi.l,                      // dari M_ITEM (bentuk "C" utk coil-like)
            qty          : kgAct,                     // SLT: Qty = KG
            kg           : kgAct,
            operator     : opHeader,
            mc_no        : mcNoFinal,
            tgl_prod     : tglProdFinal,              // Q2a: dari Selesai_DT HEADER
            supplier     : supplierInfo.supplier,
            no_po        : supplierInfo.no_po,
            no_do        : supplierInfo.no_do,
            owner        : ownerHeader,
            owner_used   : ownerUsed
          });
          batchCount++;
        } catch (eTL) {
          Logger.log('WARN writeTraceLog skip ' + outSpk + ': ' + eTL);
        }
      } else {
        Logger.log('WARN: writeTraceLog function tidak tersedia');
      }

      // ═══ WRITE REKAP_ICT (kalau Owner ≠ Owner_Used) ═══
      // 🔒 GUARD FG-only: valid FG targets = Stok_FG, FG_RM_Stamping, FG_Cust
      // Skip: Stok_Coil (sisa coil), WIP_*, Stok_Sheet — belum jadi FG
      var sltIsFG = (tgtLoc === 'Stok_FG' || (tgtLoc && tgtLoc.indexOf('FG_') === 0));
      if (ownerHeader && ownerUsed &&
          ownerHeader.toUpperCase() !== ownerUsed.toUpperCase() &&
          sltIsFG &&
          typeof writeRekapICT === 'function') {
        try {
          writeRekapICT({
            tgl        : timestamp,
            spk_no     : payload.spk_no,
            item_code  : itemCode,
            description: mi.desc,
            spec       : mi.spec,
            dari_owner : ownerHeader,
            ke_owner   : ownerUsed,
            qty        : kgAct,                     // SLT: Qty = KG
            kg         : kgAct
          });
          ictCount++;
        } catch (eICT) {
          Logger.log('WARN writeRekapICT skip ' + outSpk + ': ' + eICT);
        }
      }
    });

    // ═══ 6. UPDATE SLT-HEADER FINAL ═══
    ctx.sheet.getRange(header.rowIdx, ctx.I.status + 1).setValue(SLT_STATUS.DONE);
    ctx.sheet.getRange(header.rowIdx, ctx.I.qtyAct + 1).setValue(totalKgActual);
    ctx.sheet.getRange(header.rowIdx, ctx.I.kgAct + 1).setValue(totalKgActual);
    if (totalKgNg > 0) {
      ctx.sheet.getRange(header.rowIdx, ctx.I.kgNg + 1).setValue(totalKgNg);
      ctx.sheet.getRange(header.rowIdx, ctx.I.qtyNg + 1).setValue(totalKgNg);
    }

    // Append approval info ke NOTE HEADER
    var oldHeaderNote = String(header.rowData[ctx.I.note] || '');
    var newHeaderNote = oldHeaderNote + ' | Closed by: ' + approvedBy + ' @ ' + timestamp.toISOString();
    ctx.sheet.getRange(header.rowIdx, ctx.I.note + 1).setValue(newHeaderNote);

    // Set Leader kalau kolom ada
    if (ctx.I.leader !== -1) {
      ctx.sheet.getRange(header.rowIdx, ctx.I.leader + 1).setValue(approvedBy);
    }

    SpreadsheetApp.flush();

    // ═══ 7. TRIGGER SCHEDULING ═══
    if (typeof kalkulasiEstimasiWaktu === 'function') {
      try { kalkulasiEstimasiWaktu(); }
      catch (e) { Logger.log('Warn kalkulasiEstimasiWaktu: ' + e); }
    }

    // ═══ 8. RETURN ═══
    return {
      success          : true,
      message          : 'SPK ' + payload.spk_no + ' berhasil ditutup. ' +
                         batchCount + ' batch turunan ke Trace_Log' +
                         (ictCount > 0 ? ', ' + ictCount + ' entry Rekap_ICT' : ''),
      batch_generated  : batchCount,
      ict_generated    : ictCount,
      total_kg_actual  : totalKgActual,
      total_kg_ng      : totalKgNg
    };

  } catch (e) {
    Logger.log('Error sltCloseSpkFinal: ' + e.toString() + '\n' + (e.stack || ''));
    return { success: false, message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}


// ═══════════════════════════════════════════════════════════════════
// PUBLIC: Query helpers untuk UI Board & Page Packing & Approval
// ═══════════════════════════════════════════════════════════════════

/**
 * List semua SPK SLT dengan status tertentu (untuk UI filter).
 * @param {string} status - contoh: 'MESIN_SELESAI', 'MENUNGGU_APPROVAL', atau '' untuk semua
 * @return {array} list SLT-HEADER summary
 *
 * FIX (5A-1): Convert semua Date object ke ISO string sebelum return.
 * google.script.run tidak bisa serialize raw Date object di array → hasil corrupt di frontend.
 */
function sltListByStatus(status) {
  try {
    var ctx = _slt_getSpkSheetContext();
    var filter = String(status || '').trim().toUpperCase();
    var results = [];

    for (var i = 1; i < ctx.data.length; i++) {
      if (String(ctx.data[i][ctx.I.type] || '').trim() !== 'SLT-HEADER') continue;

      var st = String(ctx.data[i][ctx.I.status] || '').trim().toUpperCase();
      if (filter && st !== filter) continue;

      results.push({
        spk_no      : String(ctx.data[i][ctx.I.spk] || ''),
        status      : st,
        item_code   : String(ctx.data[i][ctx.I.item] || ''),
        input_spec  : String(ctx.data[i][ctx.I.inputSpec] || ''),
        batch_id    : String(ctx.data[i][ctx.I.batchId] || ''),
        kg_target   : Number(ctx.data[i][ctx.I.kgTgt]) || 0,
        kg_actual   : Number(ctx.data[i][ctx.I.kgAct]) || 0,
        cut_progress: String(ctx.data[i][ctx.I.cutProg] || ''),
        mulai_dt    : _slt_serializeDate(ctx.data[i][ctx.I.mulai]),
        selesai_dt  : _slt_serializeDate(ctx.data[i][ctx.I.selesai]),
        owner       : String(ctx.data[i][ctx.I.owner] || ''),
        cust        : String(ctx.data[i][ctx.I.cust] || '')
      });
    }
    return results;

  } catch (e) {
    Logger.log('Error sltListByStatus: ' + e.toString());
    return [];
  }
}

/**
 * Helper: Convert Date object ke ISO string (safe untuk google.script.run serialization).
 * Return '' kalau null/undefined/invalid date.
 */
function _slt_serializeDate(val) {
  if (val === null || val === undefined || val === '') return '';
  try {
    if (val instanceof Date) {
      if (isNaN(val.getTime())) return '';
      return val.toISOString();
    }
    // Kalau bukan Date, return as-is (string, number, dll)
    return String(val);
  } catch (e) {
    return '';
  }
}

/**
 * Get detail 1 SPK SLT (HEADER + semua OUT + progress info).
 * Return: object dengan semua Date di-serialize ke ISO string.
 */
function sltGetSpkDetail(spkNo) {
  try {
    var ctx = _slt_getSpkSheetContext();
    var header = _slt_findHeader(ctx.data, ctx.I, spkNo);
    if (!header) return { found: false };

    var allOuts = _slt_getAllOuts(ctx.data, ctx.I, spkNo);
    var noteStr = String(header.rowData[ctx.I.note] || '');
    var totalCut = _slt_extractTotalCut(noteStr);
    var progressArr = _slt_parseCutProgress(String(header.rowData[ctx.I.cutProg] || ''));
    var cutPlan = _slt_parseCutPlan(noteStr);
    var trimInfo = _slt_parseTrimInfo(noteStr);

    var weightedCount = 0;
    var outsData = allOuts.map(function(o) {
      var kgAct = Number(o.rowData[ctx.I.kgAct]) || 0;
      var kgNg  = Number(o.rowData[ctx.I.kgNg])  || 0;
      var isWeighted = (kgAct > 0 || kgNg > 0);
      if (isWeighted) weightedCount++;

      var outSpkNo = String(o.rowData[ctx.I.spk] || '');
      var parsed = _slt_parseOutSpkNo(outSpkNo);
      var outNote = String(o.rowData[ctx.I.note] || '');

      return {
        out_spk    : outSpkNo,
        jalur      : parsed ? parsed.jalur : 0,
        cut        : parsed ? parsed.cut : 0,
        setting_no : _slt_parseSettingFromNote(outNote),
        item_code  : String(o.rowData[ctx.I.item] || ''),
        input_spec : String(o.rowData[ctx.I.inputSpec] || ''),
        lebar      : _slt_parseLebarFromNote(outNote),
        kg_target  : Number(o.rowData[ctx.I.kgTgt]) || 0,
        kg_actual  : kgAct,
        kg_ng      : kgNg,
        target_loc : String(o.rowData[ctx.I.tgtLoc] || ''),
        owner_used : String(o.rowData[ctx.I.ownerU] || ''),
        so_ref     : String(o.rowData[ctx.I.soRef] || ''),
        cust       : String(o.rowData[ctx.I.cust] || ''),
        note       : outNote,
        status     : String(o.rowData[ctx.I.status] || ''),
        is_weighted: isWeighted
      };
    });

    // Sort outs by jalur then cut
    outsData.sort(function(a, b) {
      if (a.jalur !== b.jalur) return a.jalur - b.jalur;
      return a.cut - b.cut;
    });

    // Additional header field indices
    var iMc       = ctx.headers.indexOf('MC_No');
    var iPriority = ctx.headers.indexOf('Priority');
    var iTglBuat  = ctx.headers.indexOf('Tgl_Buat');
    var iPlSetup  = ctx.headers.indexOf('Plan_Setup_Menit');
    var iPlRun    = ctx.headers.indexOf('Plan_Run_Menit');
    var iTotDur   = ctx.headers.indexOf('Total_Durasi_Menit');
    var iCreated  = ctx.headers.indexOf('Created_By');
    var iT        = ctx.headers.indexOf('T');

    return {
      found: true,
      header: {
        spk_no       : String(header.rowData[ctx.I.spk] || ''),
        status       : String(header.rowData[ctx.I.status] || ''),
        item_code    : String(header.rowData[ctx.I.item] || ''),
        input_spec   : String(header.rowData[ctx.I.inputSpec] || ''),
        batch_id     : String(header.rowData[ctx.I.batchId] || ''),
        kg_target    : Number(header.rowData[ctx.I.kgTgt]) || 0,
        kg_actual    : Number(header.rowData[ctx.I.kgAct]) || 0,
        kg_ng        : Number(header.rowData[ctx.I.kgNg])  || 0,
        cut_progress : String(header.rowData[ctx.I.cutProg] || ''),
        cuts_done    : progressArr.length,
        total_cut    : totalCut,
        cut_plan     : cutPlan,
        trim_info    : trimInfo,
        t            : iT !== -1 ? (Number(header.rowData[iT]) || 0) : 0,
        note         : noteStr,
        owner        : String(header.rowData[ctx.I.owner] || ''),
        cust         : String(header.rowData[ctx.I.cust] || ''),
        mc_no        : iMc !== -1 ? String(header.rowData[iMc] || 'SLT-01') : 'SLT-01',
        priority     : iPriority !== -1 ? String(header.rowData[iPriority] || 'Normal') : 'Normal',
        created_by   : iCreated !== -1 ? String(header.rowData[iCreated] || '') : '',
        tgl_buat     : iTglBuat !== -1 ? _slt_serializeDate(header.rowData[iTglBuat]) : '',
        mulai_dt     : _slt_serializeDate(header.rowData[ctx.I.mulai]),
        selesai_dt   : _slt_serializeDate(header.rowData[ctx.I.selesai]),
        plan_setup   : iPlSetup !== -1 ? (Number(header.rowData[iPlSetup]) || 0) : 0,
        plan_run     : iPlRun !== -1 ? (Number(header.rowData[iPlRun]) || 0) : 0,
        total_durasi : iTotDur !== -1 ? (Number(header.rowData[iTotDur]) || 0) : 0
      },
      outs: outsData,
      progress: {
        cuts_done      : progressArr.length,
        total_cut      : totalCut,
        rolls_weighted : weightedCount,
        rolls_total    : outsData.length,
        rolls_pending  : outsData.length - weightedCount
      }
    };

  } catch (e) {
    Logger.log('Error sltGetSpkDetail: ' + e.toString());
    return { found: false, error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// PUBLIC: sltGetFormData — bundle master data untuk form create SPK SLT
// ═══════════════════════════════════════════════════════════════════

/**
 * Return semua master data yang dibutuhkan form create SPK SLT dalam 1 call.
 * Semua Date object di-serialize ke ISO string (safe untuk google.script.run).
 *
 * @return {object} { stok_coil, m_item, so_open, generated_at }
 */
function sltGetFormData() {
  var result = {
    stok_coil    : [],
    m_item       : [],
    so_open      : [],
    generated_at : new Date().toISOString()
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ═══ 1. STOK_COIL (KG_Avail > 0) ═══
    var shCoil = ss.getSheetByName('Stok_Coil');
    if (shCoil) {
      var data = shCoil.getDataRange().getValues();
      if (data.length >= 2) {
        var hdr = data[0].map(function(h){ return String(h).trim(); });
        var iBatch = hdr.indexOf('Batch_ID');
        var iItem  = hdr.indexOf('Item_Code');
        var iDesc  = hdr.indexOf('Description');
        var iSpec  = hdr.indexOf('Spec');
        var iT     = hdr.indexOf('T');
        var iP     = hdr.indexOf('P');
        var iKgAv  = hdr.indexOf('KG_Avail');
        var iQtyAv = hdr.indexOf('Qty_Avail');
        var iOwner = hdr.indexOf('Owner');

        for (var i = 1; i < data.length; i++) {
          var batch = String(data[i][iBatch] || '').trim();
          if (!batch) continue;
          var kgAv = Number(data[i][iKgAv]) || 0;
          if (kgAv <= 0) continue;

          result.stok_coil.push({
            batch_id    : batch,
            item_code   : String(data[i][iItem] || ''),
            description : String(data[i][iDesc] || ''),
            spec        : String(data[i][iSpec] || ''),
            t           : Number(data[i][iT])     || 0,
            lebar_coil  : Number(data[i][iP])     || 0,
            kg_avail    : kgAv,
            qty_avail   : Number(data[i][iQtyAv]) || 0,
            owner       : String(data[i][iOwner] || '')
          });
        }
      }
    }

    // ═══ 2. M_ITEM ═══
    var shItem = ss.getSheetByName('M_ITEM');
    if (shItem) {
      var data = shItem.getDataRange().getValues();
      if (data.length >= 2) {
        var hdr = data[0].map(function(h){ return String(h).trim(); });
        var iCode = hdr.indexOf('Item_Code');
        var iDesc = hdr.indexOf('Description');
        var iSpec = hdr.indexOf('Spec');
        var iT    = hdr.indexOf('T');
        var iP    = hdr.indexOf('P');
        var iL    = hdr.indexOf('L');
        var iEq   = hdr.indexOf('Equivalent');
        var iType = hdr.indexOf('TYPE');

        for (var i = 1; i < data.length; i++) {
          var code = String(data[i][iCode] || '').trim();
          if (!code) continue;

          result.m_item.push({
            item_code   : code,
            description : String(data[i][iDesc] || ''),
            spec        : String(data[i][iSpec] || ''),
            t           : Number(data[i][iT]) || 0,
            p           : Number(data[i][iP]) || 0,
            l           : String(data[i][iL] || ''),
            equivalent  : String(data[i][iEq] || ''),
            type        : String(data[i][iType] || '')
          });
        }
      }
    }

    // ═══ 3. DEMAND OPEN (SO + STP_REQ) — reuse getDemandData ═══
    // Panggil Demand_Service supaya konsisten dengan modul Total Demand.
    // Field mapping: so_no ← ref_no, bl_q ← net_req, +tipe (SO/STP)
    try {
      if (typeof getDemandData === 'function') {
        var demandData = getDemandData();
        (demandData.demands || []).forEach(function(d) {
          // Filter yang butuh SPK: net_req > 0
          if ((Number(d.net_req) || 0) <= 0) return;
          // Skip status non-actionable
          var st = String(d.status || '').toUpperCase();
          if (st === 'CLOSED' || st === 'CANCELLED' || st === 'DONE' || st === 'FULFILLED') return;

          result.so_open.push({
            so_no         : String(d.ref_no || ''),
            cust          : String(d.cust || ''),
            item_code     : String(d.item_code || ''),
            description   : String(d.description || ''),
            bl_q          : Number(d.net_req) || 0,
            schedule_date : String(d.tgl_needed || ''),
            tipe          : String(d.tipe || 'SO'),
            equivalent    : String(d.equivalent || ''),
            t             : Number(d.t) || 0
          });
        });
      } else {
        // Fallback: direct read SO sheet kalau Demand_Service belum ready
        Logger.log('WARN: getDemandData tidak tersedia, fallback ke SO sheet');
        var shSO = ss.getSheetByName('SO');
        if (shSO) {
          var data = shSO.getDataRange().getValues();
          if (data.length >= 2) {
            var hdr = data[0].map(function(h){ return String(h).trim(); });
            var iSoNo   = hdr.indexOf('SO_No');
            var iCust   = hdr.indexOf('Cust');
            var iItem   = hdr.indexOf('Item_Code');
            var iDesc   = hdr.indexOf('Description');
            var iBlQ    = hdr.indexOf('BL_Q');
            var iStatus = hdr.indexOf('STATUS');
            var iSched  = hdr.indexOf('SCHEDULE_DATE');
            for (var i = 1; i < data.length; i++) {
              var soNo = String(data[i][iSoNo] || '').trim();
              if (!soNo) continue;
              var status = String(data[i][iStatus] || '').trim().toUpperCase();
              if (status === 'CLOSED' || status === 'CANCELLED') continue;
              var blQ = Number(data[i][iBlQ]) || 0;
              if (blQ <= 0) continue;
              result.so_open.push({
                so_no         : soNo,
                cust          : String(data[i][iCust] || ''),
                item_code     : String(data[i][iItem] || ''),
                description   : String(data[i][iDesc] || ''),
                bl_q          : blQ,
                schedule_date : _slt_serializeDate(data[i][iSched]),
                tipe          : 'SO'
              });
            }
          }
        }
      }
    } catch (eDem) {
      Logger.log('Error load demand: ' + eDem.toString());
    }

    return result;

  } catch (e) {
    Logger.log('Error sltGetFormData: ' + e.toString());
    result.error = e.toString();
    return result;
  }
}

// ═══════════════════════════════════════════════════════════════════
// PUBLIC: sltStartProcess — Operator klik "Mulai Proses" (ANTRIAN → RUNNING)
// ═══════════════════════════════════════════════════════════════════

/**
 * @param {object} payload { spk_no, operator, note }
 * @return {object} { success, message, new_status }
 */
function sltStartProcess(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    if (!payload || !payload.spk_no) return { success: false, message: 'spk_no kosong' };
    var operator = String(payload.operator || '').trim();
    if (!operator) return { success: false, message: 'Nama operator wajib diisi' };

    var ctx = _slt_getSpkSheetContext();
    var header = _slt_findHeader(ctx.data, ctx.I, payload.spk_no);
    if (!header) return { success: false, message: 'SLT-HEADER ' + payload.spk_no + ' tidak ditemukan' };

    var currentStatus = String(header.rowData[ctx.I.status] || '').trim();
    if (currentStatus !== SLT_STATUS.ANTRIAN) {
      return { success: false, message: 'SPK status ' + currentStatus + ' (harus ANTRIAN)' };
    }

    // Cek mesin idle (tidak ada SPK lain RUNNING di mesin yang sama)
    var mcNo = String(header.rowData[ctx.headers.indexOf('MC_No')] || 'SLT-01').trim();
    for (var i = 1; i < ctx.data.length; i++) {
      if (String(ctx.data[i][ctx.I.type] || '').trim() !== 'SLT-HEADER') continue;
      if (String(ctx.data[i][ctx.headers.indexOf('MC_No')] || '').trim() !== mcNo) continue;
      var st = String(ctx.data[i][ctx.I.status] || '').trim();
      if (st === SLT_STATUS.RUNNING) {
        return {
          success: false,
          message: 'Mesin ' + mcNo + ' masih RUNNING SPK ' + String(ctx.data[i][ctx.I.spk]) +
                   '. Selesaikan cut terakhir dulu baru bisa start SPK berikutnya.'
        };
      }
    }

    var timestamp = new Date();
    ctx.sheet.getRange(header.rowIdx, ctx.I.status + 1).setValue(SLT_STATUS.RUNNING);
    ctx.sheet.getRange(header.rowIdx, ctx.I.mulai + 1).setValue(timestamp);
    if (ctx.I.op !== -1) ctx.sheet.getRange(header.rowIdx, ctx.I.op + 1).setValue(operator);

    // Append note kalau ada
    if (payload.note && String(payload.note).trim()) {
      var oldNote = String(header.rowData[ctx.I.note] || '');
      var newNote = oldNote + ' | Start: ' + operator + ' - ' + String(payload.note).trim();
      ctx.sheet.getRange(header.rowIdx, ctx.I.note + 1).setValue(newNote);
    }

    SpreadsheetApp.flush();

    if (typeof kalkulasiEstimasiWaktu === 'function') {
      try { kalkulasiEstimasiWaktu(); } catch (e) { Logger.log('Warn: ' + e); }
    }

    return { success: true, message: 'SPK ' + payload.spk_no + ' mulai diproses oleh ' + operator, new_status: SLT_STATUS.RUNNING };

  } catch (e) {
    Logger.log('Error sltStartProcess: ' + e.toString());
    return { success: false, message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════════
// PUBLIC: sltGetBoardData — bundle data Board SLT (4 kolom status)
// ═══════════════════════════════════════════════════════════════════

/**
 * Return semua data untuk Board Mesin SLT dalam 1 call.
 * Filter: hanya SPK di mesin SLT-01 dengan status aktif.
 *
 * @return {object} {
 *   machine_status, counters, antrian, running, menunggu_timbang, menunggu_approval, generated_at
 * }
 */
function sltGetBoardData() {
  var result = {
    machine_status    : 'IDLE',
    counters          : { antrian: 0, running: 0, menunggu_timbang: 0, menunggu_approval: 0 },
    antrian           : [],
    running           : [],
    menunggu_timbang  : [],
    menunggu_approval : [],
    generated_at      : new Date().toISOString()
  };

  try {
    var ctx = _slt_getSpkSheetContext();
    var iMc       = ctx.headers.indexOf('MC_No');
    var iPriority = ctx.headers.indexOf('Priority');
    var iTglBuat  = ctx.headers.indexOf('Tgl_Buat');
    var iCreated  = ctx.headers.indexOf('Created_By');

    for (var i = 1; i < ctx.data.length; i++) {
      if (String(ctx.data[i][ctx.I.type] || '').trim() !== 'SLT-HEADER') continue;
      var mcNo = iMc !== -1 ? String(ctx.data[i][iMc] || '').trim() : '';
      if (mcNo !== 'SLT-01') continue;

      var status = String(ctx.data[i][ctx.I.status] || '').trim().toUpperCase();
      if (['ANTRIAN','RUNNING','MESIN_SELESAI','MENUNGGU_APPROVAL'].indexOf(status) === -1) continue;

      var spkNo = String(ctx.data[i][ctx.I.spk] || '');
      var noteStr = String(ctx.data[i][ctx.I.note] || '');
      var totalCut = _slt_extractTotalCut(noteStr);
      var progressArr = _slt_parseCutProgress(String(ctx.data[i][ctx.I.cutProg] || ''));
      var cutPlan = _slt_parseCutPlan(noteStr);
      var trimInfo = _slt_parseTrimInfo(noteStr);

      // Count OUT rows & weighted
      var outCount = 0, weightedCount = 0, jalurCount = 0;
      for (var j = 1; j < ctx.data.length; j++) {
        if (String(ctx.data[j][ctx.I.parent] || '').trim() !== spkNo) continue;
        if (String(ctx.data[j][ctx.I.type] || '').trim() !== 'SLT-OUT') continue;
        outCount++;
        var kgAct = Number(ctx.data[j][ctx.I.kgAct]) || 0;
        var kgNg  = Number(ctx.data[j][ctx.I.kgNg])  || 0;
        if (kgAct > 0 || kgNg > 0) weightedCount++;
      }
      jalurCount = totalCut > 0 ? Math.round(outCount / totalCut) : outCount;

      var item = {
        spk_no       : spkNo,
        status       : status,
        batch_id     : String(ctx.data[i][ctx.I.batchId] || ''),
        item_code    : String(ctx.data[i][ctx.I.item] || ''),
        input_spec   : String(ctx.data[i][ctx.I.inputSpec] || ''),
        kg_target    : Number(ctx.data[i][ctx.I.kgTgt]) || 0,
        kg_actual    : Number(ctx.data[i][ctx.I.kgAct]) || 0,
        priority     : iPriority !== -1 ? String(ctx.data[i][iPriority] || 'Normal') : 'Normal',
        owner        : String(ctx.data[i][ctx.I.owner] || ''),
        cust         : String(ctx.data[i][ctx.I.cust] || ''),
        cut_plan     : cutPlan,
        total_cut    : totalCut,
        cuts_done    : progressArr.length,
        cut_progress : String(ctx.data[i][ctx.I.cutProg] || ''),
        trim_info    : trimInfo,
        jalur_count  : jalurCount,
        rolls_total  : outCount,
        rolls_weighted : weightedCount,
        rolls_pending  : outCount - weightedCount,
        operator     : ctx.I.op !== -1 ? String(ctx.data[i][ctx.I.op] || '') : '',
        tgl_buat     : iTglBuat !== -1 ? _slt_serializeDate(ctx.data[i][iTglBuat]) : '',
        mulai_dt     : _slt_serializeDate(ctx.data[i][ctx.I.mulai]),
        selesai_dt   : _slt_serializeDate(ctx.data[i][ctx.I.selesai]),
        created_by   : iCreated !== -1 ? String(ctx.data[i][iCreated] || '') : ''
      };

      if (status === 'ANTRIAN')                  { result.antrian.push(item);          result.counters.antrian++; }
      else if (status === 'RUNNING')             { result.running.push(item);          result.counters.running++; }
      else if (status === 'MESIN_SELESAI')       { result.menunggu_timbang.push(item); result.counters.menunggu_timbang++; }
      else if (status === 'MENUNGGU_APPROVAL')   { result.menunggu_approval.push(item);result.counters.menunggu_approval++; }
    }

    // Determine machine status
    if (result.counters.running > 0) result.machine_status = 'RUNNING';
    else if (result.counters.menunggu_approval > 0) result.machine_status = 'WAITING_APPROVAL';
    else if (result.counters.menunggu_timbang > 0) result.machine_status = 'MESIN_SELESAI';
    else result.machine_status = 'IDLE';

    // Sort ANTRIAN by priority DESC then tgl_buat ASC
    result.antrian.sort(function(a, b) {
      var priOrder = { 'Rush': 3, 'Urgent': 2, 'Normal': 1 };
      var pa = priOrder[a.priority] || 1;
      var pb = priOrder[b.priority] || 1;
      if (pa !== pb) return pb - pa;
      return String(a.tgl_buat).localeCompare(String(b.tgl_buat));
    });

    return result;

  } catch (e) {
    Logger.log('Error sltGetBoardData: ' + e.toString());
    result.error = e.toString();
    return result;
  }
}