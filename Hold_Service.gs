// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - HOLD SERVICE
//
// Modul: HOLD / RESUME / CANCEL_HOLD untuk switching SPK Urgent
// Dependency: Loss_Service.gs (pakai helper _lossFormatDate, _lossFormatTime,
//             _lossFormatDT, _lossCalcDurasi, _lossYYMM, _lossNextSeq,
//             _lossValidateCode)
// GAS writes to : SPK.Hold_State, Loss_Log (row MTD-07)
// Reads from    : Script Properties (HOLD_PASSWORD), M_Loss_Reason
//
// Public API (prefix "hold*" & "spk*Hold*" utk hindari namespace collision):
//   - holdVerifyPassword(pwd)           : verify + return session token 1 jam
//   - holdCheckToken(tokenStr)          : cek token masih valid
//   - spkHoldRequest(spkNo,mcNo,detail,token,user) : set Hold_State='HOLD'
//   - spkResumeRequest(spkNo,mcNo,token,user)      : clear Hold_State
//   - spkCancelHoldRequest(spkNo,mcNo,token,user)  : ILLEGAL — must RESUME first
//   - spkGetHeldByMc(mcNo)              : list SPK HOLD utk render kolom
//   - spkGetRunningStatusByMc(mcNo)     : cek slot RUNNING kosong
// =========================================================================


// -------------------------------------------------------------------------
// PRIVATE HELPERS - PASSWORD & SESSION TOKEN
// -------------------------------------------------------------------------

function _holdReadSecret() {
  try {
    var p = PropertiesService.getScriptProperties();
    return String(p.getProperty('HOLD_PASSWORD') || '').trim();
  } catch (e) {
    return '';
  }
}

/**
 * Buat token 1 jam. Format: "HT-{now}-{rand}@{expiresMs}"
 * Simple non-signed — cukup untuk konteks internal lapangan.
 */
function _holdMakeToken() {
  var now = Date.now();
  var expires = now + 60 * 60 * 1000; // 1 jam
  var rand = Math.random().toString(36).slice(2, 10);
  var value = 'HT-' + now + '-' + rand + '@' + expires;
  return { value: value, expires_ms: expires, expires_iso: new Date(expires).toISOString() };
}

function _holdValidateToken(tokenStr) {
  if (!tokenStr) return false;
  var s = String(tokenStr).trim();
  var at = s.lastIndexOf('@');
  if (at < 0) return false;
  var exp = Number(s.slice(at + 1)) || 0;
  if (!exp) return false;
  return Date.now() < exp;
}


// -------------------------------------------------------------------------
// PRIVATE HELPERS - SHEET ACCESS
// -------------------------------------------------------------------------

function _holdReadSpk() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SPK');
  if (!sh) throw new Error('Sheet SPK tidak ditemukan');
  var d = sh.getDataRange().getValues();
  var hdr = d[0].map(function(h){ return String(h).trim(); });
  var idx = {
    spk    : hdr.indexOf('SPK_No'),
    type   : hdr.indexOf('SPK_Type'),
    mc     : hdr.indexOf('MC_No'),
    status : hdr.indexOf('Status'),
    hold   : hdr.indexOf('Hold_State')
  };
  if (idx.spk === -1)    throw new Error('Kolom SPK_No tidak ada di sheet SPK');
  if (idx.status === -1) throw new Error('Kolom Status tidak ada di sheet SPK');
  if (idx.mc === -1)     throw new Error('Kolom MC_No tidak ada di sheet SPK');
  if (idx.hold === -1)   throw new Error('Kolom Hold_State belum ditambah di sheet SPK. Tambah kolom baru dgn header "Hold_State" di posisi paling kanan.');
  return { sheet: sh, data: d, hdr: hdr, idx: idx };
}

function _holdFindSpkRow(spkNo) {
  var s = _holdReadSpk();
  var t = String(spkNo).trim();
  for (var i = 1; i < s.data.length; i++) {
    if (String(s.data[i][s.idx.spk]).trim() === t) {
      return { sheet: s.sheet, hdr: s.hdr, idx: s.idx, row: s.data[i], rowNum: i + 1 };
    }
  }
  return null;
}

function _holdCountHoldByMc(mcNo, dataObj) {
  var s = dataObj || _holdReadSpk();
  var mc = String(mcNo).trim().toUpperCase();
  var count = 0;
  for (var i = 1; i < s.data.length; i++) {
    var jMc = String(s.data[i][s.idx.mc]).trim().toUpperCase();
    var jHold = String(s.data[i][s.idx.hold] || '').trim().toUpperCase();
    if (jMc === mc && jHold === 'HOLD') count++;
  }
  return count;
}

/**
 * Cari SPK RUNNING non-hold di mesin tsb.
 * Return { spk_no, spk_type } atau null.
 */
function _holdFindActiveRunning(mcNo, excludeSpkNo, dataObj) {
  var s = dataObj || _holdReadSpk();
  var mc = String(mcNo).trim().toUpperCase();
  var excl = String(excludeSpkNo || '').trim();
  for (var i = 1; i < s.data.length; i++) {
    var jSpk = String(s.data[i][s.idx.spk]).trim();
    if (excl && jSpk === excl) continue;
    var jType = String(s.data[i][s.idx.type]).trim().toUpperCase();
    if (jType.indexOf('OUT') !== -1) continue;  // skip OUT rows
    var jMc = String(s.data[i][s.idx.mc]).trim().toUpperCase();
    var jStatus = String(s.data[i][s.idx.status]).trim().toUpperCase();
    var jHold = String(s.data[i][s.idx.hold] || '').trim().toUpperCase();
    if (jMc === mc && jStatus === 'RUNNING' && jHold !== 'HOLD') {
      return { spk_no: jSpk, spk_type: jType };
    }
  }
  return null;
}


// -------------------------------------------------------------------------
// PRIVATE HELPERS - LOSS_LOG MTD-07 (open-ended)
// -------------------------------------------------------------------------

function _holdInsertOpenLoss(spkNo, mcNo, detail, user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Loss_Log');
  if (!sh) throw new Error('Sheet Loss_Log tidak ditemukan');
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
              .map(function(h){ return String(h).trim(); });

  var info = _lossValidateCode('MTD-07');
  if (!info.valid) throw new Error('Code MTD-07 belum ditambah / aktif di M_Loss_Reason');

  var now = new Date();
  var yymm = _lossYYMM(now);
  var seq = _lossNextSeq('LSS', yymm);
  var lossId = 'LSS-' + yymm + '-' + String(seq).padStart(4, '0');

  var rowMap = {
    Loss_ID         : lossId,
    SPK_No          : spkNo,
    MC_No           : mcNo,
    Tgl_Loss        : _lossFormatDate(now),
    Jam_Mulai       : _lossFormatTime(now),
    Jam_Selesai     : '',       // OPEN — closed saat RESUME
    Durasi_Menit    : 0,        // OPEN — hitung saat RESUME
    Kategori_4M     : info.kategori,
    Sub_Kategori    : info.subKategori,
    Code            : info.code,
    Detail          : String(detail || 'SPK di-HOLD karena prioritas berubah').trim(),
    Is_Shared       : 'N',
    Shared_Group_ID : '',
    Stop_Type       : info.stopType,
    Logged_By       : String(user || 'PPIC').trim(),
    Logged_At       : now
  };
  var row = hdr.map(function(h){ return rowMap.hasOwnProperty(h) ? rowMap[h] : ''; });
  sh.appendRow(row);
  return { loss_id: lossId, jam_mulai: rowMap.Jam_Mulai, tgl_loss: rowMap.Tgl_Loss };
}

/**
 * Cari & close row Loss_Log MTD-07 yg open utk SPK ini (row terakhir yg match).
 * Return { loss_id, durasi } atau null.
 */
function _holdCloseOpenLoss(spkNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Loss_Log');
  if (!sh) return null;
  var d = sh.getDataRange().getValues();
  if (d.length < 2) return null;
  var hdr = d[0].map(function(h){ return String(h).trim(); });
  var iSpk   = hdr.indexOf('SPK_No');
  var iCode  = hdr.indexOf('Code');
  var iJamM  = hdr.indexOf('Jam_Mulai');
  var iJamS  = hdr.indexOf('Jam_Selesai');
  var iDur   = hdr.indexOf('Durasi_Menit');
  var iLoss  = hdr.indexOf('Loss_ID');
  if (iSpk === -1 || iCode === -1 || iJamS === -1 || iDur === -1) return null;

  var t = String(spkNo).trim();
  for (var i = d.length - 1; i >= 1; i--) {
    if (String(d[i][iSpk]).trim() !== t) continue;
    if (String(d[i][iCode]).trim().toUpperCase() !== 'MTD-07') continue;
    var jamSel = d[i][iJamS];
    var jamSelStr = (jamSel instanceof Date) ? _lossFormatTime(jamSel) : String(jamSel || '').trim();
    if (jamSelStr) continue; // sudah closed, skip

    // Ini row yg open — close
    var now = new Date();
    var jamSelNow = _lossFormatTime(now);
    var jamMul = d[i][iJamM];
    var jamMulStr = (jamMul instanceof Date) ? _lossFormatTime(jamMul) : String(jamMul || '').trim();
    var durasi = _lossCalcDurasi(jamMulStr, jamSelNow);

    sh.getRange(i + 1, iJamS + 1).setValue(jamSelNow);
    sh.getRange(i + 1, iDur + 1).setValue(durasi);
    return { loss_id: String(d[i][iLoss]).trim(), durasi: durasi };
  }
  return null;
}

/**
 * Ambil info Loss_Log MTD-07 open utk 1 SPK (utk render "Hold sejak").
 */
function _holdGetOpenLossInfo(spkNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Loss_Log');
  if (!sh) return null;
  var d = sh.getDataRange().getValues();
  if (d.length < 2) return null;
  var hdr = d[0].map(function(h){ return String(h).trim(); });
  var iSpk = hdr.indexOf('SPK_No');
  var iCode = hdr.indexOf('Code');
  var iJamM = hdr.indexOf('Jam_Mulai');
  var iJamS = hdr.indexOf('Jam_Selesai');
  var iTgl = hdr.indexOf('Tgl_Loss');
  var iDet = hdr.indexOf('Detail');
  var iAt = hdr.indexOf('Logged_At');
  var iBy = hdr.indexOf('Logged_By');

  var t = String(spkNo).trim();
  for (var i = d.length - 1; i >= 1; i--) {
    if (String(d[i][iSpk]).trim() !== t) continue;
    if (String(d[i][iCode]).trim().toUpperCase() !== 'MTD-07') continue;
    var jamSel = d[i][iJamS];
    var jamSelStr = (jamSel instanceof Date) ? _lossFormatTime(jamSel) : String(jamSel || '').trim();
    if (jamSelStr) continue;

    var jamMul = d[i][iJamM];
    var tgl = d[i][iTgl];
    var loggedAt = d[i][iAt];
    return {
      tgl        : (tgl instanceof Date) ? _lossFormatDate(tgl) : String(tgl || '').trim(),
      jam        : (jamMul instanceof Date) ? _lossFormatTime(jamMul) : String(jamMul || '').trim(),
      logged_at  : (loggedAt instanceof Date) ? _lossFormatDT(loggedAt) : String(loggedAt || '').trim(),
      logged_by  : String(d[i][iBy] || '').trim(),
      detail     : String(d[i][iDet] || '').trim()
    };
  }
  return null;
}


// -------------------------------------------------------------------------
// PUBLIC: PASSWORD & SESSION
// -------------------------------------------------------------------------

/**
 * Verify password dgn Script Property HOLD_PASSWORD.
 * @return { success, token, expires_iso, message }
 */
function holdVerifyPassword(pwd) {
  try {
    var secret = _holdReadSecret();
    if (!secret) return { success: false, message: 'Password belum di-setup di Script Properties (HOLD_PASSWORD)' };
    if (String(pwd || '').trim() !== secret) return { success: false, message: 'Password salah' };
    var tok = _holdMakeToken();
    return { success: true, token: tok.value, expires_iso: tok.expires_iso };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Cek token masih valid (dipakai frontend utk pre-check sebelum HOLD/RESUME).
 */
function holdCheckToken(tokenStr) {
  return { success: true, valid: _holdValidateToken(tokenStr) };
}


// -------------------------------------------------------------------------
// PUBLIC: HOLD SPK
// -------------------------------------------------------------------------

/**
 * HOLD SPK: set Hold_State='HOLD' + insert Loss_Log MTD-07 open.
 * SPK Status TETAP RUNNING (approach Opsi B).
 * @return { success, loss_id, held_count, message, need_password }
 */
function spkHoldRequest(spkNo, mcNo, detail, token, user) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) { return { success: false, message: 'Sistem sedang sibuk' }; }

  try {
    if (!_holdValidateToken(token)) {
      return { success: false, message: 'Session expired. Silakan input password ulang.', need_password: true };
    }
    if (!spkNo) throw new Error('SPK_No wajib');
    if (!mcNo) throw new Error('MC_No wajib');

    var spkT = String(spkNo).trim();
    var mcT  = String(mcNo).trim();
    var det  = String(detail || '').trim();
    var usr  = String(user || 'PPIC').trim();

    var s = _holdReadSpk();
    var rowIdx = -1;
    for (var i = 1; i < s.data.length; i++) {
      if (String(s.data[i][s.idx.spk]).trim() === spkT) { rowIdx = i; break; }
    }
    if (rowIdx === -1) throw new Error('SPK ' + spkT + ' tidak ditemukan');

    var rowStatus = String(s.data[rowIdx][s.idx.status]).trim().toUpperCase();
    if (rowStatus !== 'RUNNING') throw new Error('SPK harus RUNNING utk di-HOLD (sekarang: ' + rowStatus + ')');

    var rowHold = String(s.data[rowIdx][s.idx.hold] || '').trim().toUpperCase();
    if (rowHold === 'HOLD') throw new Error('SPK sudah dalam status HOLD');

    var rowMc = String(s.data[rowIdx][s.idx.mc]).trim();
    if (rowMc.toUpperCase() !== mcT.toUpperCase()) {
      throw new Error('SPK mesin ' + rowMc + ' tidak match dgn ' + mcT);
    }

    // Cek max 3 HOLD per mesin
    var heldCount = _holdCountHoldByMc(mcT, s);
    if (heldCount >= 3) {
      throw new Error('Sudah ada ' + heldCount + ' SPK di-HOLD di mesin ' + mcT + ' (limit maksimum 3)');
    }

    // Set Hold_State = HOLD
    s.sheet.getRange(rowIdx + 1, s.idx.hold + 1).setValue('HOLD');

    // Insert Loss_Log MTD-07 open
    var lossInfo = _holdInsertOpenLoss(spkT, mcT, det, usr);

    return {
      success    : true,
      loss_id    : lossInfo.loss_id,
      held_count : heldCount + 1,
      message    : 'SPK ' + spkT + ' berhasil di-HOLD (loss ' + lossInfo.loss_id + ' dibuka)'
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


// -------------------------------------------------------------------------
// PUBLIC: RESUME SPK
// -------------------------------------------------------------------------

/**
 * RESUME SPK: clear Hold_State + close Loss_Log MTD-07 open.
 * TOLAK kalau ada RUNNING lain (non-hold) di mesin sama.
 * @return { success, loss_id, durasi_menit, message, need_password }
 */
function spkResumeRequest(spkNo, mcNo, token, user) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) { return { success: false, message: 'Sistem sedang sibuk' }; }

  try {
    if (!_holdValidateToken(token)) {
      return { success: false, message: 'Session expired. Silakan input password ulang.', need_password: true };
    }
    if (!spkNo) throw new Error('SPK_No wajib');
    if (!mcNo) throw new Error('MC_No wajib');

    var spkT = String(spkNo).trim();
    var mcT  = String(mcNo).trim();
    var usr  = String(user || 'PPIC').trim();

    var s = _holdReadSpk();
    var rowIdx = -1;
    for (var i = 1; i < s.data.length; i++) {
      if (String(s.data[i][s.idx.spk]).trim() === spkT) { rowIdx = i; break; }
    }
    if (rowIdx === -1) throw new Error('SPK ' + spkT + ' tidak ditemukan');

    var rowHold = String(s.data[rowIdx][s.idx.hold] || '').trim().toUpperCase();
    if (rowHold !== 'HOLD') throw new Error('SPK ' + spkT + ' tidak dalam status HOLD');

    var rowStatus = String(s.data[rowIdx][s.idx.status]).trim().toUpperCase();
    if (rowStatus !== 'RUNNING') throw new Error('SPK ' + spkT + ' status Status bukan RUNNING (' + rowStatus + ')');

    // Cek slot RUNNING kosong: ada gak SPK lain di mesin sama, Status=RUNNING, Hold_State != HOLD?
    var other = _holdFindActiveRunning(mcT, spkT, s);
    if (other) {
      return {
        success: false,
        message: 'Tidak bisa RESUME: mesin ' + mcT + ' masih RUNNING SPK ' + other.spk_no + '. HOLD / DONE dulu SPK itu.'
      };
    }

    // Close Loss_Log MTD-07 open
    var closed = _holdCloseOpenLoss(spkT);

    // Clear Hold_State
    s.sheet.getRange(rowIdx + 1, s.idx.hold + 1).setValue('');

    return {
      success       : true,
      loss_id       : closed ? closed.loss_id : null,
      durasi_menit  : closed ? closed.durasi : 0,
      message       : 'SPK ' + spkT + ' berhasil RESUME' + (closed ? ' (loss MTD-07 ditutup: ' + closed.durasi + ' mnt)' : '')
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


// -------------------------------------------------------------------------
// PUBLIC: LIST HELD SPK
// -------------------------------------------------------------------------

/**
 * List SPK yg lagi Hold_State='HOLD' di 1 mesin.
 * Include HEADER only (skip OUT rows). Tiap item dilengkapi info Loss_Log open.
 * @return { success, count, items:[] }
 */
function spkGetHeldByMc(mcNo) {
  try {
    if (!mcNo) return { success: false, items: [] };
    var mc = String(mcNo).trim().toUpperCase();
    var s = _holdReadSpk();
    var items = [];
    for (var i = 1; i < s.data.length; i++) {
      var jHold = String(s.data[i][s.idx.hold] || '').trim().toUpperCase();
      if (jHold !== 'HOLD') continue;
      var jMc = String(s.data[i][s.idx.mc]).trim().toUpperCase();
      if (jMc !== mc) continue;
      var jType = String(s.data[i][s.idx.type]).trim().toUpperCase();
      if (jType.indexOf('OUT') !== -1) continue;

      var obj = {};
      s.hdr.forEach(function(h, idx){
        var v = s.data[i][idx];
        if (v instanceof Date) v = _lossFormatDT(v);
        obj[h] = v;
      });
      var lossInfo = _holdGetOpenLossInfo(String(s.data[i][s.idx.spk]).trim());
      obj._hold_tgl       = lossInfo ? lossInfo.tgl : '';
      obj._hold_jam       = lossInfo ? lossInfo.jam : '';
      obj._hold_logged_at = lossInfo ? lossInfo.logged_at : '';
      obj._hold_logged_by = lossInfo ? lossInfo.logged_by : '';
      obj._hold_detail    = lossInfo ? lossInfo.detail : '';
      items.push(obj);
    }
    return { success: true, count: items.length, items: items };
  } catch (e) {
    return { success: false, message: e.message, items: [] };
  }
}

/**
 * Cek status slot RUNNING di mesin. Utility utk frontend poka-yoke.
 * @return { success, active_running: {spk_no, spk_type} | null, held_count }
 */
function spkGetRunningStatusByMc(mcNo) {
  try {
    if (!mcNo) return { success: false, active_running: null, held_count: 0 };
    var s = _holdReadSpk();
    var mcT = String(mcNo).trim();
    var active = _holdFindActiveRunning(mcT, null, s);
    var held = _holdCountHoldByMc(mcT, s);
    return { success: true, active_running: active, held_count: held };
  } catch (e) {
    return { success: false, message: e.message, active_running: null, held_count: 0 };
  }
}


// =========================================================================
// TEST FUNCTIONS
// =========================================================================

function _test_holdReadSecret() {
  var s = _holdReadSecret();
  Logger.log('Secret loaded (len): ' + s.length);
  Logger.log('Kalau 0, artinya HOLD_PASSWORD belum di-set di Script Properties');
}

function _test_holdMakeToken() {
  var t = _holdMakeToken();
  Logger.log('Token: ' + t.value);
  Logger.log('Expires: ' + t.expires_iso);
  Logger.log('Valid now? ' + _holdValidateToken(t.value));
}

function _test_holdVerifyPassword_WRONG() {
  var r = holdVerifyPassword('salahsalah');
  Logger.log(JSON.stringify(r));  // expect success:false
}

function _test_holdReadSpk() {
  try {
    var s = _holdReadSpk();
    Logger.log('Sheet SPK OK. Total row: ' + (s.data.length - 1));
    Logger.log('idx SPK_No:' + s.idx.spk + ' Status:' + s.idx.status + ' MC_No:' + s.idx.mc + ' Hold_State:' + s.idx.hold);
    if (s.idx.hold === -1) Logger.log('❌ Hold_State BELUM ADA — tambahkan kolom Hold_State di sheet SPK');
    else Logger.log('✅ Hold_State kolom di posisi ke-' + (s.idx.hold + 1));
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.message);
  }
}

function _test_spkGetHeldByMc() {
  var r = spkGetHeldByMc('CTL-01');
  Logger.log('Held CTL-01: ' + r.count);
  Logger.log(JSON.stringify(r, null, 2));
}

function _test_spkGetRunningStatusByMc() {
  var r = spkGetRunningStatusByMc('CTL-01');
  Logger.log(JSON.stringify(r, null, 2));
}