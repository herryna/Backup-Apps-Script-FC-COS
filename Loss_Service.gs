// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - LOSS SERVICE
//
// Modul: Loss Tracking untuk analisa 4M (Man, Machine, Method, Material)
// GAS writes to : Loss_Log
// Reads from    : M_Loss_Reason, M_Shift_Override, M_Config,
//                 SYS_Sequence, M_MC, SPK
//
// Public API (semua prefix "loss*" utk hindari namespace collision):
//   - lossGetOptions()                : dropdown opsi aktif
//   - lossSaveEvent(payload)          : simpan 1 loss event
//   - lossSaveShared(payload, mcList) : Loss Bersama, N mesin sekali call
//   - lossGetBySpk(spkNo)             : list loss utk 1 SPK
//   - lossGetByMcDate(mcNo, tgl)      : list loss per mesin per hari
//   - lossDelete(lossId, user)        : hapus 1 loss event
//   - lossGetPPT(tgl, mcNo)           : resolver PPT (rule + override)
//   - lossGetSummaryByDate(tgl)       : ringkasan per mesin utk report
// =========================================================================


// -------------------------------------------------------------------------
// PRIVATE HELPERS
// -------------------------------------------------------------------------

function _lossReadSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return { rows: [], hdr: [], idx: {}, sheet: null };
  var d = sh.getDataRange().getValues();
  if (d.length < 1) return { rows: [], hdr: [], idx: {}, sheet: sh };
  var hdr = d[0].map(function(h){ return String(h).trim(); });
  var idx = {};
  for (var i = 0; i < hdr.length; i++) idx[hdr[i]] = i;
  return {
    rows  : d.length > 1 ? d.slice(1) : [],
    hdr   : hdr,
    idx   : idx,
    sheet : sh
  };
}

function _lossFormatDate(dt) {
  if (!(dt instanceof Date)) {
    if (typeof dt === 'string' && dt) dt = new Date(dt);
    else return '';
  }
  if (isNaN(dt.getTime())) return '';
  var y = dt.getFullYear();
  var m = String(dt.getMonth() + 1).padStart(2, '0');
  var d = String(dt.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

function _lossFormatTime(dt) {
  if (!(dt instanceof Date)) return '';
  var h = String(dt.getHours()).padStart(2, '0');
  var m = String(dt.getMinutes()).padStart(2, '0');
  return h + ':' + m;
}

function _lossFormatDT(dt) {
  if (!(dt instanceof Date)) dt = new Date();
  var d = _lossFormatDate(dt);
  var t = _lossFormatTime(dt) + ':' + String(dt.getSeconds()).padStart(2, '0');
  return d + ' ' + t;
}

function _lossParseDate(str) {
  if (str instanceof Date) return str;
  if (!str) return null;
  var s = String(str).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function _lossParseTime(str) {
  if (!str) return null;
  if (str instanceof Date) return { h: str.getHours(), m: str.getMinutes() };
  var s = String(str).trim();
  var match = s.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return null;
  var h = parseInt(match[1]);
  var m = parseInt(match[2]);
  if (h < 0 || h > 23 || m < 0 || m > 59) return null;
  return { h: h, m: m };
}

function _lossCalcDurasi(jamMulai, jamSelesai) {
  var t1 = _lossParseTime(jamMulai);
  var t2 = _lossParseTime(jamSelesai);
  if (!t1 || !t2) return 0;
  var s = t1.h * 60 + t1.m;
  var e = t2.h * 60 + t2.m;
  var diff = e - s;
  if (diff < 0) diff += 24 * 60; // cross midnight (safety, jarang kejadian)
  return diff;
}

function _lossYYMM(dt) {
  if (!(dt instanceof Date)) dt = new Date();
  var y = String(dt.getFullYear()).slice(-2);
  var m = String(dt.getMonth() + 1).padStart(2, '0');
  return y + m;
}

/**
 * Next sequence dari SYS_Sequence untuk (prefix, yymm).
 * Auto-append row kalau belum ada.
 * WAJIB dipanggil dalam LockService di caller.
 */
function _lossNextSeq(prefix, yymm) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SYS_Sequence');
  if (!sh) throw new Error('Sheet SYS_Sequence tidak ditemukan');
  var d = sh.getDataRange().getValues();
  var hdr = d[0].map(function(h){ return String(h).trim(); });
  var iTipe = hdr.indexOf('Tipe_Doc');
  var iBt   = hdr.indexOf('Bulan_Tahun');
  var iSeq  = hdr.indexOf('Last_Seq');
  if (iTipe === -1 || iBt === -1 || iSeq === -1) {
    throw new Error('SYS_Sequence header kurang (butuh Tipe_Doc, Bulan_Tahun, Last_Seq)');
  }
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][iTipe]).trim() === prefix && String(d[i][iBt]).trim() === yymm) {
      var next = Number(d[i][iSeq] || 0) + 1;
      sh.getRange(i + 1, iSeq + 1).setValue(next);
      return next;
    }
  }
  sh.appendRow([prefix, yymm, 1]);
  return 1;
}

/**
 * Cek Code loss valid & aktif. Return meta.
 */
function _lossValidateCode(code) {
  var out = { code: code, kategori: '', subKategori: '', stopType: '', valid: false };
  if (!code) return out;
  var t = String(code).trim().toUpperCase();
  var r = _lossReadSheet('M_Loss_Reason');
  if (r.idx.Code === undefined) return out;
  for (var i = 0; i < r.rows.length; i++) {
    var c = String(r.rows[i][r.idx.Code] || '').trim().toUpperCase();
    if (c !== t) continue;
    var isActive = String(r.rows[i][r.idx.Is_Active] || '').toUpperCase();
    if (isActive !== 'TRUE') return out;
    out.code = c;
    out.kategori = String(r.rows[i][r.idx.Kategori_4M] || '').trim();
    out.subKategori = String(r.rows[i][r.idx.Sub_Kategori] || '').trim();
    out.stopType = String(r.rows[i][r.idx.Stop_Type] || 'STOP').trim().toUpperCase();
    out.valid = true;
    return out;
  }
  return out;
}

function _lossValidateMc(mcNo) {
  if (!mcNo) return false;
  var t = String(mcNo).trim();
  var r = _lossReadSheet('M_MC');
  if (r.idx.MC_No === undefined) return false;
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][r.idx.MC_No] || '').trim() === t) return true;
  }
  return false;
}

/**
 * SPK_No boleh kosong (loss ke PPT bukan ke SPK saat mesin idle).
 * Kalau diisi, harus exist di sheet SPK.
 */
function _lossValidateSpk(spkNo) {
  if (!spkNo) return true;
  var t = String(spkNo).trim();
  var r = _lossReadSheet('SPK');
  if (r.idx.SPK_No === undefined) return false;
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][r.idx.SPK_No] || '').trim() === t) return true;
  }
  return false;
}

function _lossRowToObj(row, idx) {
  var tgl = row[idx.Tgl_Loss];
  var tglStr = (tgl instanceof Date) ? _lossFormatDate(tgl) : String(tgl || '').trim();
  var jamMul = row[idx.Jam_Mulai];
  var jamSel = row[idx.Jam_Selesai];
  var loggedAt = row[idx.Logged_At];
  return {
    loss_id         : String(row[idx.Loss_ID] || '').trim(),
    spk_no          : String(row[idx.SPK_No] || '').trim(),
    mc_no           : String(row[idx.MC_No] || '').trim(),
    tgl_loss        : tglStr,
    jam_mulai       : (jamMul instanceof Date) ? _lossFormatTime(jamMul) : String(jamMul || '').trim(),
    jam_selesai     : (jamSel instanceof Date) ? _lossFormatTime(jamSel) : String(jamSel || '').trim(),
    durasi_menit    : Number(row[idx.Durasi_Menit]) || 0,
    kategori_4m     : String(row[idx.Kategori_4M] || '').trim(),
    sub_kategori    : String(row[idx.Sub_Kategori] || '').trim(),
    code            : String(row[idx.Code] || '').trim(),
    detail          : String(row[idx.Detail] || '').trim(),
    is_shared       : String(row[idx.Is_Shared] || 'N').trim().toUpperCase(),
    shared_group_id : String(row[idx.Shared_Group_ID] || '').trim(),
    stop_type       : String(row[idx.Stop_Type] || 'STOP').trim().toUpperCase(),
    logged_by       : String(row[idx.Logged_By] || '').trim(),
    logged_at       : (loggedAt instanceof Date) ? _lossFormatDT(loggedAt) : String(loggedAt || '').trim()
  };
}


// -------------------------------------------------------------------------
// PPT RESOLVER
// -------------------------------------------------------------------------

/**
 * Ambil PPT (Planned Production Time) untuk (tgl, mcNo).
 * Rule:
 *   1. Cek M_Shift_Override (tgl, mcNo)   -> specific mesin
 *   2. Cek M_Shift_Override (tgl, "")     -> global override tanggal
 *   3. Fallback:
 *      - Weekend (Sat/Sun) -> 0
 *      - Weekday           -> M_Config.PPT_Default_Menit (default 480)
 *
 * @param {string|Date} tgl - "YYYY-MM-DD" atau Date object
 * @param {string} mcNo     - "CTL-01" dll (opsional, kosong = ambil global)
 * @return {number} menit PPT
 */
function lossGetPPT(tgl, mcNo) {
  var dt = _lossParseDate(tgl);
  if (!dt) return 0;
  var tglStr = _lossFormatDate(dt);
  var mc = String(mcNo || '').trim();

  var r = _lossReadSheet('M_Shift_Override');
  if (r.idx.Tgl !== undefined && r.idx.MC_No !== undefined && r.idx.PPT_Menit !== undefined) {
    var specific = null;
    var globalOv = null;
    for (var i = 0; i < r.rows.length; i++) {
      var rowDt = r.rows[i][r.idx.Tgl];
      var rowDtStr = (rowDt instanceof Date) ? _lossFormatDate(rowDt) : String(rowDt || '').trim();
      if (rowDtStr !== tglStr) continue;
      var rowMc = String(r.rows[i][r.idx.MC_No] || '').trim();
      var ppt = Number(r.rows[i][r.idx.PPT_Menit] || 0);
      if (mc && rowMc === mc) specific = ppt;
      else if (rowMc === '') globalOv = ppt;
    }
    if (specific !== null) return specific;
    if (globalOv !== null) return globalOv;
  }

  var dow = dt.getDay(); // 0=Sun, 6=Sat
  if (dow === 0 || dow === 6) return 0;

  var cfg = _lossReadSheet('M_Config');
  if (cfg.idx.Key !== undefined && cfg.idx.Value !== undefined) {
    for (var j = 0; j < cfg.rows.length; j++) {
      if (String(cfg.rows[j][cfg.idx.Key]).trim() === 'PPT_Default_Menit') {
        return Number(cfg.rows[j][cfg.idx.Value] || 480) || 480;
      }
    }
  }
  return 480;
}


// -------------------------------------------------------------------------
// PUBLIC: GET DROPDOWN OPTIONS
// -------------------------------------------------------------------------

/**
 * Return list sub-kategori loss aktif utk dropdown UI.
 * Di-sort by Display_Order.
 */
function lossGetOptions() {
  try {
    var r = _lossReadSheet('M_Loss_Reason');
    var opts = [];
    if (!r.rows.length) return { success: true, options: [] };
    for (var i = 0; i < r.rows.length; i++) {
      var isActive = String(r.rows[i][r.idx.Is_Active] || '').toUpperCase();
      if (isActive !== 'TRUE') continue;
      opts.push({
        code         : String(r.rows[i][r.idx.Code] || '').trim(),
        kategori     : String(r.rows[i][r.idx.Kategori_4M] || '').trim(),
        subKategori  : String(r.rows[i][r.idx.Sub_Kategori] || '').trim(),
        stopType     : String(r.rows[i][r.idx.Stop_Type] || 'STOP').trim().toUpperCase(),
        displayOrder : Number(r.rows[i][r.idx.Display_Order] || 999) || 999
      });
    }
    opts.sort(function(a, b){ return a.displayOrder - b.displayOrder; });
    return { success: true, options: opts };
  } catch (e) {
    return { success: false, message: e.message, options: [] };
  }
}


// -------------------------------------------------------------------------
// PUBLIC: SAVE 1 LOSS EVENT
// -------------------------------------------------------------------------

/**
 * Simpan 1 loss event ke Loss_Log.
 * @param {object} payload {
 *   spk_no      : string  (optional, kosong = loss ke PPT bukan SPK)
 *   mc_no       : string  (wajib)
 *   tgl_loss    : "YYYY-MM-DD"
 *   jam_mulai   : "HH:mm"
 *   jam_selesai : "HH:mm"
 *   code        : string  (wajib, mis "MCH-04")
 *   detail      : string  (wajib utk OTH-03 / OTH-04)
 *   logged_by   : string
 * }
 * @return { success, loss_id, durasi_menit, message }
 */
function lossSaveEvent(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) { return { success: false, message: 'Sistem sedang sibuk. Coba lagi.' }; }

  try {
    if (!payload) throw new Error('Payload kosong');

    var mcNo    = String(payload.mc_no || '').trim();
    var tglLoss = String(payload.tgl_loss || '').trim();
    var jamMul  = String(payload.jam_mulai || '').trim();
    var jamSel  = String(payload.jam_selesai || '').trim();
    var code    = String(payload.code || '').trim().toUpperCase();
    var spkNo   = String(payload.spk_no || '').trim();
    var detail  = String(payload.detail || '').trim();
    var user    = String(payload.logged_by || 'system').trim();

    if (!mcNo) throw new Error('MC_No wajib diisi');
    if (!_lossValidateMc(mcNo)) throw new Error('MC_No "' + mcNo + '" tidak ditemukan di M_MC');
    if (!tglLoss) throw new Error('Tgl_Loss wajib diisi');
    if (!jamMul || !jamSel) throw new Error('Jam_Mulai dan Jam_Selesai wajib diisi');
    var durasi = _lossCalcDurasi(jamMul, jamSel);
    if (durasi <= 0) throw new Error('Durasi tidak valid (Jam_Selesai harus > Jam_Mulai)');

    var info = _lossValidateCode(code);
    if (!info.valid) throw new Error('Code "' + code + '" tidak valid atau tidak aktif');
    if ((code === 'OTH-03' || code === 'OTH-04') && !detail) {
      throw new Error('Kode ' + code + ' wajib mengisi field Detail');
    }
    if (spkNo && !_lossValidateSpk(spkNo)) {
      throw new Error('SPK_No "' + spkNo + '" tidak ditemukan');
    }

    var now  = new Date();
    var yymm = _lossYYMM(now);
    var seq  = _lossNextSeq('LSS', yymm);
    var lossId = 'LSS-' + yymm + '-' + String(seq).padStart(4, '0');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Loss_Log');
    if (!sh) throw new Error('Sheet Loss_Log tidak ditemukan');
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                .map(function(h){ return String(h).trim(); });

    var rowMap = {
      Loss_ID         : lossId,
      SPK_No          : spkNo,
      MC_No           : mcNo,
      Tgl_Loss        : tglLoss,
      Jam_Mulai       : jamMul,
      Jam_Selesai     : jamSel,
      Durasi_Menit    : durasi,
      Kategori_4M     : info.kategori,
      Sub_Kategori    : info.subKategori,
      Code            : info.code,
      Detail          : detail,
      Is_Shared       : 'N',
      Shared_Group_ID : '',
      Stop_Type       : info.stopType,
      Logged_By       : user,
      Logged_At       : now
    };
    var row = hdr.map(function(h){ return rowMap.hasOwnProperty(h) ? rowMap[h] : ''; });
    sh.appendRow(row);

    return {
      success      : true,
      loss_id      : lossId,
      durasi_menit : durasi,
      message      : 'Loss ' + lossId + ' tersimpan (' + durasi + ' mnt)'
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


// -------------------------------------------------------------------------
// PUBLIC: LOSS BERSAMA (SHARED LOSS)
// -------------------------------------------------------------------------

/**
 * Loss Bersama: satu event kena banyak mesin.
 * Insert N row dgn Shared_Group_ID sama. Setiap row punya Loss_ID sendiri.
 *
 * @param {object} payload {
 *   tgl_loss, jam_mulai, jam_selesai, code, detail, logged_by
 * }
 * @param {Array} mcList [{ mc_no, spk_no }] — 1 item per mesin. spk_no boleh ""
 * @return { success, group_id, loss_ids:[], count, message }
 */
function lossSaveShared(payload, mcList) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { success: false, message: 'Sistem sedang sibuk. Coba lagi.' }; }

  try {
    if (!payload) throw new Error('Payload kosong');
    if (!Array.isArray(mcList) || mcList.length === 0) throw new Error('Minimal 1 mesin');

    var tglLoss = String(payload.tgl_loss || '').trim();
    var jamMul  = String(payload.jam_mulai || '').trim();
    var jamSel  = String(payload.jam_selesai || '').trim();
    var code    = String(payload.code || '').trim().toUpperCase();
    var detail  = String(payload.detail || '').trim();
    var user    = String(payload.logged_by || 'system').trim();

    if (!tglLoss) throw new Error('Tgl_Loss wajib diisi');
    if (!jamMul || !jamSel) throw new Error('Jam_Mulai dan Jam_Selesai wajib diisi');
    var durasi = _lossCalcDurasi(jamMul, jamSel);
    if (durasi <= 0) throw new Error('Durasi tidak valid');
    var info = _lossValidateCode(code);
    if (!info.valid) throw new Error('Code "' + code + '" tidak valid atau tidak aktif');
    if ((code === 'OTH-03' || code === 'OTH-04') && !detail) {
      throw new Error('Kode ' + code + ' wajib mengisi field Detail');
    }

    for (var i = 0; i < mcList.length; i++) {
      var m = String(mcList[i].mc_no || '').trim();
      if (!m) throw new Error('MC_No kosong di baris ke-' + (i + 1));
      if (!_lossValidateMc(m)) throw new Error('MC_No "' + m + '" tidak valid');
      var sp = String(mcList[i].spk_no || '').trim();
      if (sp && !_lossValidateSpk(sp)) throw new Error('SPK "' + sp + '" tidak ditemukan');
    }

    var now  = new Date();
    var yymm = _lossYYMM(now);
    var seqG = _lossNextSeq('LSG', yymm);
    var groupId = 'LSG-' + yymm + '-' + String(seqG).padStart(4, '0');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Loss_Log');
    if (!sh) throw new Error('Sheet Loss_Log tidak ditemukan');
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                .map(function(h){ return String(h).trim(); });

    var lossIds = [];
    var newRows = [];
    for (var j = 0; j < mcList.length; j++) {
      var seqL = _lossNextSeq('LSS', yymm);
      var lossId = 'LSS-' + yymm + '-' + String(seqL).padStart(4, '0');
      lossIds.push(lossId);
      var rowMap = {
        Loss_ID         : lossId,
        SPK_No          : String(mcList[j].spk_no || '').trim(),
        MC_No           : String(mcList[j].mc_no || '').trim(),
        Tgl_Loss        : tglLoss,
        Jam_Mulai       : jamMul,
        Jam_Selesai     : jamSel,
        Durasi_Menit    : durasi,
        Kategori_4M     : info.kategori,
        Sub_Kategori    : info.subKategori,
        Code            : info.code,
        Detail          : detail,
        Is_Shared       : 'Y',
        Shared_Group_ID : groupId,
        Stop_Type       : info.stopType,
        Logged_By       : user,
        Logged_At       : now
      };
      newRows.push(hdr.map(function(h){ return rowMap.hasOwnProperty(h) ? rowMap[h] : ''; }));
    }

    var startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, newRows.length, hdr.length).setValues(newRows);

    return {
      success  : true,
      group_id : groupId,
      loss_ids : lossIds,
      count    : newRows.length,
      message  : newRows.length + ' loss event tersimpan (group ' + groupId + ', ' + durasi + ' mnt/mesin)'
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


// -------------------------------------------------------------------------
// PUBLIC: QUERY
// -------------------------------------------------------------------------

/**
 * List loss utk 1 SPK.
 */
function lossGetBySpk(spkNo) {
  try {
    if (!spkNo) return { success: false, message: 'SPK_No kosong', items: [] };
    var t = String(spkNo).trim();
    var r = _lossReadSheet('Loss_Log');
    var items = [];
    var total = 0;
    for (var i = 0; i < r.rows.length; i++) {
      if (String(r.rows[i][r.idx.SPK_No] || '').trim() !== t) continue;
      var o = _lossRowToObj(r.rows[i], r.idx);
      total += o.durasi_menit;
      items.push(o);
    }
    return { success: true, count: items.length, total_menit: total, items: items };
  } catch (e) {
    return { success: false, message: e.message, items: [] };
  }
}

/**
 * List loss per mesin per tanggal.
 * mcNo boleh kosong ("") = ambil semua mesin.
 */
function lossGetByMcDate(mcNo, tgl) {
  try {
    var mc = String(mcNo || '').trim();
    var dt = _lossParseDate(tgl);
    if (!dt) return { success: false, message: 'Tanggal invalid', items: [] };
    var tglStr = _lossFormatDate(dt);
    var r = _lossReadSheet('Loss_Log');
    var items = [];
    var total = 0, stopMin = 0, paceMin = 0;
    for (var i = 0; i < r.rows.length; i++) {
      var rowMc = String(r.rows[i][r.idx.MC_No] || '').trim();
      var rowDt = r.rows[i][r.idx.Tgl_Loss];
      var rowDtStr = (rowDt instanceof Date) ? _lossFormatDate(rowDt) : String(rowDt || '').trim();
      if (mc && rowMc !== mc) continue;
      if (rowDtStr !== tglStr) continue;
      var o = _lossRowToObj(r.rows[i], r.idx);
      total += o.durasi_menit;
      if (o.stop_type === 'STOP') stopMin += o.durasi_menit;
      else if (o.stop_type === 'PACE') paceMin += o.durasi_menit;
      items.push(o);
    }
    return {
      success     : true,
      count       : items.length,
      total_menit : total,
      stop_menit  : stopMin,
      pace_menit  : paceMin,
      items       : items
    };
  } catch (e) {
    return { success: false, message: e.message, items: [] };
  }
}

/**
 * Hard delete 1 loss event. Utk audit ringan disimpan di logger.
 */
function lossDelete(lossId, user) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return { success: false, message: 'Sistem sibuk' }; }

  try {
    if (!lossId) throw new Error('Loss_ID kosong');
    var t = String(lossId).trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Loss_Log');
    if (!sh) throw new Error('Sheet Loss_Log tidak ditemukan');
    var d = sh.getDataRange().getValues();
    var hdr = d[0].map(function(h){ return String(h).trim(); });
    var iId = hdr.indexOf('Loss_ID');
    if (iId === -1) throw new Error('Kolom Loss_ID tidak ada');
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][iId]).trim() === t) {
        sh.deleteRow(i + 1);
        Logger.log('Loss ' + t + ' dihapus oleh ' + (user || 'system') + ' @ ' + new Date());
        return { success: true, message: 'Loss ' + t + ' dihapus' };
      }
    }
    throw new Error('Loss ' + t + ' tidak ditemukan');
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


// -------------------------------------------------------------------------
// PUBLIC: SUMMARY UNTUK REPORT
// -------------------------------------------------------------------------

/**
 * Ringkasan PPT + loss per mesin utk 1 tanggal.
 * Dipanggil oleh Report_Produksi_Service (Step 4).
 *
 * @param {string} tgl "YYYY-MM-DD"
 * @return {
 *   success, tgl, per_mesin: {
 *     "CTL-01": { mc_no, ppt_menit, stop_menit, pace_menit, avail_menit, loss_count },
 *     "SHR-01": {...}, ...
 *   }
 * }
 */
function lossGetSummaryByDate(tgl) {
  try {
    var dt = _lossParseDate(tgl);
    if (!dt) return { success: false, message: 'Tanggal invalid' };
    var tglStr = _lossFormatDate(dt);

    var mc = _lossReadSheet('M_MC');
    var mesinList = [];
    for (var i = 0; i < mc.rows.length; i++) {
      var active = String(mc.rows[i][mc.idx.Is_Active] || '').toUpperCase();
      if (active !== 'TRUE') continue;
      mesinList.push(String(mc.rows[i][mc.idx.MC_No] || '').trim());
    }

    var perMesin = {};
    for (var j = 0; j < mesinList.length; j++) {
      var m = mesinList[j];
      var ppt = lossGetPPT(tglStr, m);
      var l = lossGetByMcDate(m, tglStr);
      perMesin[m] = {
        mc_no       : m,
        ppt_menit   : ppt,
        stop_menit  : l.stop_menit || 0,
        pace_menit  : l.pace_menit || 0,
        avail_menit : Math.max(0, ppt - (l.stop_menit || 0)),
        loss_count  : l.count || 0
      };
    }
    return { success: true, tgl: tglStr, per_mesin: perMesin };
  } catch (e) {
    return { success: false, message: e.message };
  }
}


// =========================================================================
// TEST FUNCTIONS — jalankan di GAS editor utk sanity check
// Semua test READ ONLY kecuali _test_lossSaveEvent_INSERT (di-guard uncomment)
// =========================================================================

function _test_lossGetOptions() {
  var r = lossGetOptions();
  Logger.log('Total opsi aktif: ' + r.options.length);
  Logger.log('Sample 3 pertama:');
  for (var i = 0; i < Math.min(3, r.options.length); i++) {
    Logger.log('  ' + JSON.stringify(r.options[i]));
  }
}

function _test_lossValidateCode() {
  Logger.log('MAN-01: ' + JSON.stringify(_lossValidateCode('MAN-01')));
  Logger.log('MCH-04: ' + JSON.stringify(_lossValidateCode('MCH-04')));
  Logger.log('OTH-04: ' + JSON.stringify(_lossValidateCode('OTH-04')));
  Logger.log('XYZ-99 (invalid): ' + JSON.stringify(_lossValidateCode('XYZ-99')));
}

function _test_lossGetPPT() {
  Logger.log('PPT Senin biasa (2026-11-16) CTL-01: ' + lossGetPPT('2026-11-16', 'CTL-01') + ' (expect 480)');
  Logger.log('PPT Sabtu (2026-11-14) CTL-01     : ' + lossGetPPT('2026-11-14', 'CTL-01') + ' (expect 0)');
  Logger.log('PPT Minggu (2026-11-15) SHR-01    : ' + lossGetPPT('2026-11-15', 'SHR-01') + ' (expect 0)');
  var today = _lossFormatDate(new Date());
  Logger.log('PPT hari ini (' + today + ') CTL-01: ' + lossGetPPT(today, 'CTL-01'));
}

function _test_lossCalcDurasi() {
  Logger.log('10:00 -> 10:15 : ' + _lossCalcDurasi('10:00', '10:15') + ' mnt (expect 15)');
  Logger.log('08:30 -> 09:05 : ' + _lossCalcDurasi('08:30', '09:05') + ' mnt (expect 35)');
  Logger.log('14:00 -> 13:00 : ' + _lossCalcDurasi('14:00', '13:00') + ' mnt (expect 1380 - cross day fallback)');
}

/**
 * INSERT ke Loss_Log. Uncomment blok kalau mau test asli.
 * Setelah test, hapus row-nya manual dari sheet.
 */
function _test_lossSaveEvent_INSERT() {
  /*
  var r = lossSaveEvent({
    spk_no      : '',                 // kosong = loss ke PPT, gak nempel SPK
    mc_no       : 'CTL-01',
    tgl_loss    : _lossFormatDate(new Date()),
    jam_mulai   : '10:00',
    jam_selesai : '10:15',
    code        : 'MCH-04',           // Pisau tumpul
    detail      : 'Test dari _test_lossSaveEvent_INSERT',
    logged_by   : 'TEST'
  });
  Logger.log(JSON.stringify(r, null, 2));
  */
  Logger.log('Uncomment blok kode di function ini untuk test insert asli.');
}

/**
 * Test Loss Bersama (INSERT). Uncomment blok kalau mau test asli.
 * Akan bikin N row baru di Loss_Log dgn Shared_Group_ID sama.
 */
function _test_lossSaveShared_INSERT() {
  /*
  var r = lossSaveShared(
    {
      tgl_loss    : _lossFormatDate(new Date()),
      jam_mulai   : '14:00',
      jam_selesai : '14:30',
      code        : 'MCH-07',           // Listrik PLN mati
      detail      : 'Test Loss Bersama',
      logged_by   : 'TEST'
    },
    [
      { mc_no: 'CTL-01', spk_no: '' },
      { mc_no: 'SHR-01', spk_no: '' },
      { mc_no: 'SHR-02', spk_no: '' },
      { mc_no: 'SHR-03', spk_no: '' },
      { mc_no: 'SLT-01', spk_no: '' }
    ]
  );
  Logger.log(JSON.stringify(r, null, 2));
  */
  Logger.log('Uncomment blok kode di function ini untuk test insert asli.');
}

function _test_lossGetSummaryByDate() {
  var today = _lossFormatDate(new Date());
  var r = lossGetSummaryByDate(today);
  Logger.log('Summary tgl ' + today + ':');
  Logger.log(JSON.stringify(r, null, 2));
}