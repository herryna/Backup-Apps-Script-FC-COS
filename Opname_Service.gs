// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - OPNAME SERVICE (T6)
//
// Modul: Stock Take / Closing Period Bulanan
// Mode : Hybrid (auto-fill system + manual adjustment fisik)
//
// Sheet utama: Opname_Bulanan
// Schema:
//   Period, Batch_ID, Loc, Item_Code, Owner,
//   Qty_System, KG_System, Qty_Fisik, KG_Fisik,
//   Qty_Selisih, KG_Selisih, Status, Note,
//   Auto_Filled_DT, Locked_DT, Locked_By
//
// Public API:
//   - setupOpnameSheet()          : upgrade schema (one-time, idempotent)
//   - getOpnamePeriods()          : list semua period + summary
//   - getOpnameDetail(period)     : rows untuk 1 period
//   - autoFillOpname(period, by)  : snapshot saldo dari Stok_* (skip Avail=0)
//   - saveOpnameAdjustment(...)   : update Qty_Fisik / KG_Fisik / Note
//   - lockOpnamePeriod(period, by): finalize period (status → FINAL)
//   - deleteOpnameDraft(period)   : reset draft (hanya kalau status=DRAFT)
//   - getOpnameInitData()         : combo untuk page init (periods + active)
// =========================================================================

var OPNAME_SCHEMA_ = [
  'Period', 'Batch_ID', 'Loc', 'Item_Code', 'Owner',
  'Qty_System', 'KG_System',
  'Qty_Fisik', 'KG_Fisik',
  'Qty_Selisih', 'KG_Selisih',
  'Status', 'Note',
  'Auto_Filled_DT', 'Locked_DT', 'Locked_By'
];

// Format period "YYYY-MM" dari Date
function _opnameFormatPeriod(dt) {
  var d = (dt instanceof Date) ? dt : new Date();
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  return y + '-' + m;
}

// Normalize Period value dari sheet — Google Sheets kadang auto-convert "2026-05" jadi Date object.
// Pakai helper ini setiap baca Period column biar dapat string "YYYY-MM" konsisten.
function _opnameNormalizePeriod(rawVal) {
  if (rawVal === null || rawVal === undefined || rawVal === '') return '';
  if (rawVal instanceof Date) return _opnameFormatPeriod(rawVal);
  var s = String(rawVal).trim();
  // Kalau parseable ke Date (mis: "Fri May 01 2026..."), convert
  if (s.length > 10 && /\d{4}/.test(s)) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) return _opnameFormatPeriod(d);
  }
  return s;
}

// Period cutoff = last day of month 23:59:59
function _opnamePeriodCutoff(period) {
  // period = "YYYY-MM"
  var parts = String(period).split('-');
  var y = parseInt(parts[0]) || new Date().getFullYear();
  var m = parseInt(parts[1]) || (new Date().getMonth() + 1);
  // Last day of month: day 0 of next month
  var d = new Date(y, m, 0, 23, 59, 59, 999);
  return d;
}

// Generate list periode default untuk dropdown (12 bulan terakhir + current + 2 ke depan)
function _opnameGenerateDefaultPeriods() {
  var arr = [];
  var now = new Date();
  for (var i = -12; i <= 2; i++) {
    var d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    arr.push(_opnameFormatPeriod(d));
  }
  return arr;
}

function _opnameGetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Opname_Bulanan');
}

function _opnameReadHeader(sheet) {
  if (!sheet) return [];
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
}

function _opnameHdrIdx(hdr) {
  var idx = {};
  for (var i = 0; i < hdr.length; i++) idx[hdr[i]] = i;
  return idx;
}

// =========================================================================
// 1. SETUP / UPGRADE SCHEMA — idempotent
// =========================================================================
function setupOpnameSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Opname_Bulanan');

  if (!sh) {
    sh = ss.insertSheet('Opname_Bulanan');
    sh.getRange(1, 1, 1, OPNAME_SCHEMA_.length).setValues([OPNAME_SCHEMA_]);
    sh.getRange(1, 1, 1, OPNAME_SCHEMA_.length).setFontWeight('bold').setBackground('#f1f5f9');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, OPNAME_SCHEMA_.length);
    var msg1 = 'Sheet Opname_Bulanan dibuat dengan schema baru (16 kolom).';
    Logger.log(msg1);
    try { SpreadsheetApp.getUi().alert(msg1); } catch(e) {}
    return { success: true, action: 'created' };
  }

  var hdr = _opnameReadHeader(sh);
  var hasNewSchema = (hdr.indexOf('Qty_System') > -1 && hdr.indexOf('Qty_Fisik') > -1);

  if (hasNewSchema) {
    var msg2 = 'Schema Opname_Bulanan sudah benar. Tidak ada perubahan.';
    Logger.log(msg2);
    try { SpreadsheetApp.getUi().alert(msg2); } catch(e) {}
    return { success: true, action: 'no-change' };
  }

  // Schema lama (atau partial). Cek apakah ada data.
  var dataRange = sh.getDataRange();
  var lastRow = dataRange.getNumRows();

  if (lastRow > 1) {
    var msg3 = '⚠️ Sheet Opname_Bulanan punya data (' + (lastRow - 1) + ' rows) tapi schema lama.\n\n' +
               'Untuk safety, sheet TIDAK di-overwrite otomatis.\n' +
               'Backup data dulu (Make a Copy), hapus sheet manual, lalu jalankan setupOpnameSheet() lagi.';
    Logger.log(msg3);
    try { SpreadsheetApp.getUi().alert(msg3); } catch(e) {}
    return { success: false, action: 'has-data', message: msg3 };
  }

  // Sheet ada tapi kosong (cuma header lama) → safe to overwrite
  sh.clear();
  sh.getRange(1, 1, 1, OPNAME_SCHEMA_.length).setValues([OPNAME_SCHEMA_]);
  sh.getRange(1, 1, 1, OPNAME_SCHEMA_.length).setFontWeight('bold').setBackground('#f1f5f9');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, OPNAME_SCHEMA_.length);

  var msg4 = 'Sheet Opname_Bulanan di-upgrade ke schema baru (16 kolom).';
  Logger.log(msg4);
  try { SpreadsheetApp.getUi().alert(msg4); } catch(e) {}
  return { success: true, action: 'upgraded' };
}

// =========================================================================
// HELPER: SCAN STOK_* untuk auto-fill
// =========================================================================
function _scanStokForSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = [];

  var stokSheets = [
    { name: 'Stok_Coil',  loc: 'Coil'  },
    { name: 'Stok_Sheet', loc: 'Sheet' },
    { name: 'Stok_WIP',   loc: 'WIP'   },
    { name: 'Stok_FG',    loc: 'FG'    },
    { name: 'Stok_NG',    loc: 'NG'    }
  ];

  function findCol(hdr, names) {
    for (var n = 0; n < names.length; n++) {
      for (var c = 0; c < hdr.length; c++) {
        if (hdr[c].toLowerCase() === names[n].toLowerCase()) return c;
      }
    }
    return -1;
  }

  for (var s = 0; s < stokSheets.length; s++) {
    var sh = ss.getSheetByName(stokSheets[s].name);
    if (!sh) continue;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) continue;

    var hdr = data[0].map(function(h){ return String(h).trim(); });
    var iB    = findCol(hdr, ['Batch_ID']);
    var iI    = findCol(hdr, ['Item_Code']);
    var iOw   = findCol(hdr, ['Owner']);
    var iQA   = findCol(hdr, ['Qty_Avail']);
    var iKA   = findCol(hdr, ['KG_Avail', 'Kg_Avail']);
    if (iB === -1) continue;

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var bid = String(row[iB] || '').trim();
      if (!bid) continue;

      var qty = iQA > -1 ? (parseFloat(row[iQA]) || 0) : 0;
      var kg  = iKA > -1 ? (parseFloat(row[iKA]) || 0) : 0;

      // Skip batch dengan Avail = 0 (D3 decision)
      if (qty <= 0 && kg <= 0) continue;

      result.push({
        batch_id  : bid,
        loc       : stokSheets[s].loc,
        item_code : iI  > -1 ? String(row[iI]  || '').trim() : '',
        owner     : iOw > -1 ? String(row[iOw] || '').trim() : '',
        qty       : qty,
        kg        : kg
      });
    }
  }

  return result;
}

// =========================================================================
// 2. AUTO-FILL OPNAME
// =========================================================================
function autoFillOpname(period, filledBy) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    if (!period || !/^\d{4}-\d{2}$/.test(String(period).trim())) {
      throw new Error('Format period invalid. Pakai YYYY-MM (contoh: 2026-06)');
    }
    period = String(period).trim();

    var sh = _opnameGetSheet();
    if (!sh) throw new Error('Sheet Opname_Bulanan tidak ada. Jalankan setupOpnameSheet() dulu.');

    var hdr = _opnameReadHeader(sh);
    var idx = _opnameHdrIdx(hdr);
    if (idx['Qty_System'] === undefined) {
      throw new Error('Schema sheet Opname_Bulanan lama. Jalankan setupOpnameSheet() dulu.');
    }

    var data = sh.getDataRange().getValues();
    var iPeriod = idx['Period'];
    var iStatus = idx['Status'];

    // Check apakah period sudah ada
    var existingRows = [];
    var hasFinalRow  = false;
    for (var i = 1; i < data.length; i++) {
      if (_opnameNormalizePeriod(data[i][iPeriod]) === period) {
        existingRows.push(i + 1); // 1-indexed row number di sheet
        var st = String(data[i][iStatus] || '').toUpperCase();
        if (st === 'FINAL' || st === 'LOCKED') hasFinalRow = true;
      }
    }

    if (hasFinalRow) {
      throw new Error('Period ' + period + ' sudah FINAL/LOCKED. Tidak bisa auto-fill ulang.');
    }

    // Kalau ada existing DRAFT, hapus dulu (overwrite)
    if (existingRows.length > 0) {
      // Sort descending biar delete dari bawah
      existingRows.sort(function(a, b){ return b - a; });
      for (var d = 0; d < existingRows.length; d++) {
        sh.deleteRow(existingRows[d]);
      }
    }

    // Scan Stok_* sekarang
    var snapshot = _scanStokForSnapshot();
    if (snapshot.length === 0) {
      return {
        success: true,
        period : period,
        count  : 0,
        message: 'Tidak ada batch dengan Qty_Avail > 0 di Stok_*. Sheet Opname_Bulanan tetap kosong untuk period ini.'
      };
    }

    // Build rows
    var now = new Date();
    var newRows = snapshot.map(function(b) {
      var rowObj = {
        'Period'         : period,
        'Batch_ID'       : b.batch_id,
        'Loc'            : b.loc,
        'Item_Code'      : b.item_code,
        'Owner'          : b.owner,
        'Qty_System'     : b.qty,
        'KG_System'      : b.kg,
        'Qty_Fisik'      : '',
        'KG_Fisik'       : '',
        'Qty_Selisih'    : '',
        'KG_Selisih'     : '',
        'Status'         : 'DRAFT',
        'Note'           : '',
        'Auto_Filled_DT' : now,
        'Locked_DT'      : '',
        'Locked_By'      : ''
      };
      return hdr.map(function(h) { return rowObj[h] !== undefined ? rowObj[h] : ''; });
    });

    // Append batch sekaligus
    var lastRow = sh.getLastRow();
    var startRow = lastRow + 1;
    sh.getRange(startRow, 1, newRows.length, hdr.length).setValues(newRows);
    SpreadsheetApp.flush();

    return {
      success: true,
      period : period,
      count  : newRows.length,
      filled_by: filledBy || '',
      message: 'Auto-fill sukses: ' + newRows.length + ' batch dimasukkan ke Opname_Bulanan period ' + period
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// 3. GET PERIODS LIST + SUMMARY
// =========================================================================
function getOpnamePeriods() {
  try {
    var sh = _opnameGetSheet();
    if (!sh) return { success: true, periods: [], default_list: _opnameGenerateDefaultPeriods() };

    var hdr = _opnameReadHeader(sh);
    var idx = _opnameHdrIdx(hdr);
    if (idx['Period'] === undefined) {
      return { success: true, periods: [], default_list: _opnameGenerateDefaultPeriods() };
    }

    var data = sh.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, periods: [], default_list: _opnameGenerateDefaultPeriods() };
    }

    // Group by Period
    var groups = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var p = _opnameNormalizePeriod(row[idx['Period']]);
      if (!p) continue;

      var st     = String(row[idx['Status']] || '').toUpperCase();
      var qSys   = parseFloat(row[idx['Qty_System']]) || 0;
      var kSys   = parseFloat(row[idx['KG_System']])  || 0;
      var qFis   = parseFloat(row[idx['Qty_Fisik']])  || 0;
      var kFis   = parseFloat(row[idx['KG_Fisik']])   || 0;
      var sel_q  = qFis - qSys;
      var sel_k  = kFis - kSys;

      if (!groups[p]) {
        groups[p] = {
          period            : p,
          total_batch       : 0,
          total_qty_system  : 0,
          total_kg_system   : 0,
          total_qty_fisik   : 0,
          total_kg_fisik    : 0,
          total_selisih_qty : 0,
          total_selisih_kg  : 0,
          batch_with_fisik  : 0,
          batch_pending     : 0,
          has_final         : false,
          all_locked        : true,
          locked_dt         : null,
          locked_by         : ''
        };
      }
      groups[p].total_batch++;
      groups[p].total_qty_system += qSys;
      groups[p].total_kg_system  += kSys;
      groups[p].total_qty_fisik  += qFis;
      groups[p].total_kg_fisik   += kFis;
      groups[p].total_selisih_qty += sel_q;
      groups[p].total_selisih_kg  += sel_k;
      if (qFis > 0 || kFis > 0) groups[p].batch_with_fisik++;
      else groups[p].batch_pending++;

      if (st === 'FINAL' || st === 'LOCKED') {
        groups[p].has_final = true;
        var lkdt = row[idx['Locked_DT']];
        if (lkdt instanceof Date && (!groups[p].locked_dt || lkdt > groups[p].locked_dt)) {
          groups[p].locked_dt = lkdt;
          groups[p].locked_by = String(row[idx['Locked_By']] || '');
        }
      } else {
        groups[p].all_locked = false;
      }
    }

    // Convert to array + add status
    var arr = [];
    for (var k in groups) {
      var g = groups[k];
      g.status = (g.has_final && g.all_locked) ? 'LOCKED' : 'DRAFT';
      g.locked_dt_iso = g.locked_dt ? g.locked_dt.toISOString() : '';
      delete g.locked_dt;
      delete g.has_final;
      delete g.all_locked;
      arr.push(g);
    }

    // Sort: terbaru di atas
    arr.sort(function(a, b) { return b.period.localeCompare(a.period); });

    return {
      success: true,
      periods: arr,
      default_list: _opnameGenerateDefaultPeriods()
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// 4. GET DETAIL DATA UNTUK 1 PERIOD
// =========================================================================
function getOpnameDetail(period) {
  try {
    if (!period) throw new Error('Period tidak boleh kosong');
    period = String(period).trim();

    var sh = _opnameGetSheet();
    if (!sh) throw new Error('Sheet Opname_Bulanan tidak ada');

    var hdr = _opnameReadHeader(sh);
    var idx = _opnameHdrIdx(hdr);
    if (idx['Period'] === undefined) throw new Error('Schema sheet Opname_Bulanan tidak valid');

    // Lookup M_ITEM untuk description, spec, T, P, L, Wg/Pce FC
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mItem = ss.getSheetByName('M_ITEM');
    var itemMaster = {};
    if (mItem) {
      var mData = mItem.getDataRange().getValues();
      if (mData.length > 1) {
        var mH = mData[0].map(function(h){ return String(h).trim(); });
        var mIc = mH.indexOf('Item_Code');
        var mDc = mH.indexOf('Description');
        var mSp = mH.indexOf('Spec');
        var mT  = mH.indexOf('T');
        var mP  = mH.indexOf('P');
        var mL  = mH.indexOf('L');
        var mWg = mH.indexOf('Wg/Pce FC');
        if (mWg === -1) mWg = mH.indexOf('Wg/Pce');
        for (var mi = 1; mi < mData.length; mi++) {
          var ic = String(mData[mi][mIc] || '').trim();
          if (!ic) continue;
          itemMaster[ic] = {
            description: mDc >= 0 ? String(mData[mi][mDc] || '') : '',
            spec:        mSp >= 0 ? String(mData[mi][mSp] || '') : '',
            t:           mT  >= 0 ? (parseFloat(mData[mi][mT]) || 0) : 0,
            p:           mP  >= 0 ? (parseFloat(mData[mi][mP]) || 0) : 0,
            l:           mL  >= 0 ? (parseFloat(mData[mi][mL]) || 0) : 0,
            wg_pce_fc:   mWg >= 0 ? (parseFloat(mData[mi][mWg]) || 1) : 1
          };
        }
      }
    }

    // Lookup Stok_NG untuk batch yang ga punya Item_Code (Grade 2/scrap)
    var ngLookup = {};
    var ngSh = ss.getSheetByName('Stok_NG');
    if (ngSh) {
      var ngData = ngSh.getDataRange().getValues();
      if (ngData.length > 1) {
        var nH = ngData[0].map(function(h){ return String(h).trim(); });
        var nB  = nH.indexOf('Batch_ID');
        var nDc = nH.indexOf('Description');
        var nSp = nH.indexOf('Spec');
        var nT  = nH.indexOf('T');
        var nP  = nH.indexOf('P');
        var nL  = nH.indexOf('L');
        if (nL === -1) nL = nH.indexOf('L_dim');
        if (nB >= 0) {
          for (var ni = 1; ni < ngData.length; ni++) {
            var nbid = String(ngData[ni][nB] || '').trim();
            if (!nbid) continue;
            var nT_v = nT >= 0 ? (parseFloat(ngData[ni][nT]) || 0) : 0;
            var nP_v = nP >= 0 ? (parseFloat(ngData[ni][nP]) || 0) : 0;
            var nL_v = nL >= 0 ? (parseFloat(ngData[ni][nL]) || 0) : 0;
            var nDesc = nDc >= 0 ? String(ngData[ni][nDc] || '') : '';
            // Kalau Description kosong, build dari Spec + TxPxL
            if (!nDesc) {
              var nSpec = nSp >= 0 ? String(ngData[ni][nSp] || '') : '';
              var dim = (nT_v && nP_v) ? (nT_v + 'x' + nP_v + (nL_v ? 'x' + nL_v : '') + ' mm') : '';
              nDesc = [nSpec, dim].filter(function(x){return x;}).join(' ');
            }
            ngLookup[nbid] = {
              description: nDesc,
              spec:        nSp >= 0 ? String(ngData[ni][nSp] || '') : '',
              t: nT_v, p: nP_v, l: nL_v,
              wg_pce_fc: 1  // NG default 1 (auto-calc disabled untuk NG)
            };
          }
        }
      }
    }

    var data = sh.getDataRange().getValues();
    var rows = [];
    var allLocked = true;
    var anyLocked = false;

    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (_opnameNormalizePeriod(r[idx['Period']]) !== period) continue;

      var st = String(r[idx['Status']] || '').toUpperCase();
      if (st === 'FINAL' || st === 'LOCKED') anyLocked = true;
      else allLocked = false;

      var qSys = parseFloat(r[idx['Qty_System']]) || 0;
      var kSys = parseFloat(r[idx['KG_System']])  || 0;
      var qFis = parseFloat(r[idx['Qty_Fisik']])  || 0;
      var kFis = parseFloat(r[idx['KG_Fisik']])   || 0;
      var lkdt = r[idx['Locked_DT']];
      var afdt = r[idx['Auto_Filled_DT']];
      var itemCd = String(r[idx['Item_Code']] || '');
      var batchId = String(r[idx['Batch_ID']] || '');
      var loc = String(r[idx['Loc']] || '');
      // Lookup priority: M_ITEM (kalo ada item_code), fallback ke Stok_NG (untuk NG/Grade 2)
      var meta = itemMaster[itemCd];
      if (!meta && (loc === 'NG' || !itemCd)) {
        meta = ngLookup[batchId];
      }
      if (!meta) meta = { description:'', spec:'', t:0, p:0, l:0, wg_pce_fc:1 };

      rows.push({
        row_index    : i + 1, // 1-indexed sheet row (untuk update target)
        period       : period,
        batch_id     : batchId,
        loc          : loc,
        item_code    : itemCd,
        description  : meta.description,
        spec         : meta.spec,
        t            : meta.t,
        p            : meta.p,
        l            : meta.l,
        wg_pce_fc    : meta.wg_pce_fc || 1,
        owner        : String(r[idx['Owner']] || ''),
        qty_system   : qSys,
        kg_system    : kSys,
        qty_fisik    : qFis,
        kg_fisik     : kFis,
        qty_selisih  : qFis - qSys,
        kg_selisih   : kFis - kSys,
        status       : st || 'DRAFT',
        note         : String(r[idx['Note']] || ''),
        auto_filled_dt_iso : afdt instanceof Date ? afdt.toISOString() : '',
        locked_dt_iso      : lkdt instanceof Date ? lkdt.toISOString() : '',
        locked_by    : String(r[idx['Locked_By']] || '')
      });
    }

    // Sort: Loc asc → Description asc (tebal asc) → Batch_ID asc
    // Membantu cek fisik lapangan: per lokasi, urut tebal, lalu batch
    rows.sort(function(a, b) {
      var c = String(a.loc).localeCompare(String(b.loc));
      if (c !== 0) return c;
      // Sort by T (numeric) dulu kalo ada, fallback ke description string
      var ta = parseFloat(a.t) || 0, tb = parseFloat(b.t) || 0;
      if (ta !== tb) return ta - tb;
      c = String(a.description).localeCompare(String(b.description));
      if (c !== 0) return c;
      return String(a.batch_id).localeCompare(String(b.batch_id));
    });

    // Build summary
    var totalBatch  = rows.length;
    var qS = 0, kS = 0, qF = 0, kF = 0, sq = 0, sk = 0, withFisik = 0;
    for (var x = 0; x < rows.length; x++) {
      qS += rows[x].qty_system;  kS += rows[x].kg_system;
      qF += rows[x].qty_fisik;   kF += rows[x].kg_fisik;
      sq += rows[x].qty_selisih; sk += rows[x].kg_selisih;
      if (rows[x].qty_fisik > 0 || rows[x].kg_fisik > 0) withFisik++;
    }

    return JSON.parse(JSON.stringify({
      success: true,
      period: period,
      status: (anyLocked && allLocked) ? 'LOCKED' : 'DRAFT',
      is_locked: (anyLocked && allLocked),
      cutoff_iso: _opnamePeriodCutoff(period).toISOString(),
      rows: rows,
      summary: {
        total_batch: totalBatch,
        batch_with_fisik: withFisik,
        batch_pending: totalBatch - withFisik,
        total_qty_system: qS,
        total_kg_system: kS,
        total_qty_fisik: qF,
        total_kg_fisik: kF,
        total_selisih_qty: sq,
        total_selisih_kg: sk
      }
    }));

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// 5. SAVE ADJUSTMENT (single batch)
// =========================================================================
function saveOpnameAdjustment(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    if (!payload || !payload.period || !payload.batch_id) {
      throw new Error('Payload tidak lengkap (perlu period & batch_id)');
    }

    var sh = _opnameGetSheet();
    if (!sh) throw new Error('Sheet Opname_Bulanan tidak ada');

    var hdr = _opnameReadHeader(sh);
    var idx = _opnameHdrIdx(hdr);

    var data = sh.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (_opnameNormalizePeriod(data[i][idx['Period']]) === payload.period
          && String(data[i][idx['Batch_ID']]).trim() === payload.batch_id) {
        targetRow = i + 1;

        // Check kalau sudah locked → reject
        var st = String(data[i][idx['Status']] || '').toUpperCase();
        if (st === 'FINAL' || st === 'LOCKED') {
          throw new Error('Period ' + payload.period + ' sudah LOCKED. Edit tidak diizinkan.');
        }
        break;
      }
    }
    if (targetRow === -1) throw new Error('Batch ' + payload.batch_id + ' tidak ditemukan di period ' + payload.period);

    var qSys = parseFloat(data[targetRow - 1][idx['Qty_System']]) || 0;
    var kSys = parseFloat(data[targetRow - 1][idx['KG_System']])  || 0;

    var qFis = payload.qty_fisik !== undefined && payload.qty_fisik !== null && payload.qty_fisik !== ''
                ? (parseFloat(payload.qty_fisik) || 0) : '';
    var kFis = payload.kg_fisik !== undefined && payload.kg_fisik !== null && payload.kg_fisik !== ''
                ? (parseFloat(payload.kg_fisik) || 0) : '';
    var note = payload.note !== undefined ? String(payload.note) : '';

    // Update Qty_Fisik
    if (idx['Qty_Fisik'] !== undefined) sh.getRange(targetRow, idx['Qty_Fisik'] + 1).setValue(qFis);
    if (idx['KG_Fisik']  !== undefined) sh.getRange(targetRow, idx['KG_Fisik']  + 1).setValue(kFis);

    // Auto-calc selisih (kalau fisik diisi)
    if (qFis !== '') {
      if (idx['Qty_Selisih'] !== undefined) sh.getRange(targetRow, idx['Qty_Selisih'] + 1).setValue(qFis - qSys);
    } else {
      if (idx['Qty_Selisih'] !== undefined) sh.getRange(targetRow, idx['Qty_Selisih'] + 1).setValue('');
    }
    if (kFis !== '') {
      if (idx['KG_Selisih']  !== undefined) sh.getRange(targetRow, idx['KG_Selisih']  + 1).setValue(kFis - kSys);
    } else {
      if (idx['KG_Selisih']  !== undefined) sh.getRange(targetRow, idx['KG_Selisih']  + 1).setValue('');
    }

    if (idx['Note'] !== undefined) sh.getRange(targetRow, idx['Note'] + 1).setValue(note);

    SpreadsheetApp.flush();

    return {
      success: true,
      batch_id: payload.batch_id,
      qty_selisih: qFis !== '' ? (qFis - qSys) : null,
      kg_selisih:  kFis !== '' ? (kFis - kSys) : null
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// 6. LOCK PERIOD (finalize)
// =========================================================================
function lockOpnamePeriod(period, lockedBy) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    if (!period) throw new Error('Period tidak boleh kosong');
    period = String(period).trim();

    var sh = _opnameGetSheet();
    if (!sh) throw new Error('Sheet Opname_Bulanan tidak ada');

    var hdr = _opnameReadHeader(sh);
    var idx = _opnameHdrIdx(hdr);

    var data = sh.getDataRange().getValues();
    var targetRows = [];
    var alreadyLocked = false;
    for (var i = 1; i < data.length; i++) {
      if (_opnameNormalizePeriod(data[i][idx['Period']]) !== period) continue;
      var st = String(data[i][idx['Status']] || '').toUpperCase();
      if (st === 'FINAL' || st === 'LOCKED') alreadyLocked = true;
      targetRows.push(i + 1);
    }

    if (targetRows.length === 0) throw new Error('Tidak ada data untuk period ' + period);
    if (alreadyLocked) throw new Error('Period ' + period + ' sudah ter-lock sebelumnya.');

    var now = new Date();
    var by = lockedBy || 'system';

    // Batch update via getRangeList kalau bisa, atau loop
    for (var t = 0; t < targetRows.length; t++) {
      var rr = targetRows[t];
      if (idx['Status']    !== undefined) sh.getRange(rr, idx['Status']    + 1).setValue('FINAL');
      if (idx['Locked_DT'] !== undefined) sh.getRange(rr, idx['Locked_DT'] + 1).setValue(now);
      if (idx['Locked_By'] !== undefined) sh.getRange(rr, idx['Locked_By'] + 1).setValue(by);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      period: period,
      locked_count: targetRows.length,
      locked_by: by,
      locked_dt_iso: now.toISOString()
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// 7. DELETE DRAFT PERIOD (reset)
// =========================================================================
function deleteOpnameDraft(period) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    if (!period) throw new Error('Period tidak boleh kosong');
    period = String(period).trim();

    var sh = _opnameGetSheet();
    if (!sh) throw new Error('Sheet Opname_Bulanan tidak ada');

    var hdr = _opnameReadHeader(sh);
    var idx = _opnameHdrIdx(hdr);

    var data = sh.getDataRange().getValues();
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      if (_opnameNormalizePeriod(data[i][idx['Period']]) !== period) continue;
      var st = String(data[i][idx['Status']] || '').toUpperCase();
      if (st === 'FINAL' || st === 'LOCKED') {
        throw new Error('Period ' + period + ' sudah LOCKED. Tidak bisa di-reset.');
      }
      rowsToDelete.push(i + 1);
    }

    if (rowsToDelete.length === 0) {
      return { success: true, period: period, deleted: 0, message: 'Tidak ada data untuk dihapus.' };
    }

    // Delete dari bawah (descending) biar index gak shift
    rowsToDelete.sort(function(a, b){ return b - a; });
    for (var d = 0; d < rowsToDelete.length; d++) {
      sh.deleteRow(rowsToDelete[d]);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      period: period,
      deleted: rowsToDelete.length,
      message: 'Period ' + period + ' di-reset (' + rowsToDelete.length + ' rows dihapus)'
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// 8. GET INIT DATA (combo untuk page open)
// =========================================================================
function getOpnameInitData() {
  try {
    var periods = getOpnamePeriods();
    if (!periods.success) throw new Error(periods.message);

    // Active period suggestion: bulan berjalan
    var now = new Date();
    var currentPeriod = _opnameFormatPeriod(now);

    // Last period in list (untuk auto-select kalau ada)
    var lastPeriod = (periods.periods && periods.periods.length > 0) ? periods.periods[0].period : currentPeriod;

    return {
      success: true,
      periods: periods.periods,
      default_list: periods.default_list,
      current_period: currentPeriod,
      suggested_active: lastPeriod
    };
  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// QUICK TEST HELPERS
// =========================================================================
function _test_setupOpname() {
  var r = setupOpnameSheet();
  Logger.log(JSON.stringify(r, null, 2));
}

function _test_autoFillOpname() {
  // Test auto-fill untuk current month
  var now = new Date();
  var period = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  var r = autoFillOpname(period, 'tester');
  Logger.log(JSON.stringify(r, null, 2));
}

function _test_getOpnamePeriods() {
  var r = getOpnamePeriods();
  Logger.log('Success: ' + r.success);
  Logger.log('Total periods: ' + (r.periods ? r.periods.length : 0));
  if (r.periods) {
    r.periods.forEach(function(p) {
      Logger.log('  ' + p.period + ' | ' + p.status + ' | ' + p.total_batch + ' batch | selisih ' + p.total_selisih_qty + ' pcs / ' + p.total_selisih_kg + ' kg');
    });
  }
}

function _test_getOpnameDetail() {
  var now = new Date();
  var period = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  var r = getOpnameDetail(period);
  Logger.log('Success: ' + r.success);
  if (r.success) {
    Logger.log('Period: ' + r.period + ' | Status: ' + r.status);
    Logger.log('Total batch: ' + r.summary.total_batch);
    Logger.log('First 3 rows:');
    r.rows.slice(0, 3).forEach(function(b) {
      Logger.log('  ' + b.batch_id + ' | ' + b.loc + ' | ' + b.item_code + ' | qty_sys=' + b.qty_system + ' | kg_sys=' + b.kg_system);
    });
  }
}