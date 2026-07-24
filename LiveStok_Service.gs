/* =========================================================================
 * LIVE STOK SERVICE — Backend untuk page_live_stok.html
 * File: LiveStok_Service.gs (BARU — terpisah dari SpkService.gs)
 *
 * Tugas:
 *   1. getLiveStockData()  → return all 5 kategori stok + summary
 *   2. getLiveStockTrace() → delegate ke getBatchTrace() existing
 *
 * Catatan:
 *   - Single round-trip aggregator (1 call dari frontend)
 *   - Aging dihitung server-side (server timezone)
 *   - Toleran terhadap inconsistency naming kolom (Kg_Avail vs KG_Avail,
 *     Tgl_Output vs Tgl_Masuk, L vs L_dim) via findCol() helper
 *   - Filter & search dilakukan di frontend, backend cuma siapkan data lengkap
 * ========================================================================= */

/* =========================================================================
 * KONSTANTA — Aging thresholds per kategori (hari)
 * ========================================================================= */
var LS_AGING_THRESHOLD = {
  coil  : { warn: 45, crit: 60 },
  sheet : { warn: 21, crit: 30 },
  wip   : { warn: 10, crit: 14 },
  fg    : { warn: 10, crit: 14 },
  ng    : { warn: 21, crit: 30 }
};

/* =========================================================================
 * KONSTANTA — Mass Balance (Sprint 1B)
 *   Metric_Name di M_Threshold: MASS_BALANCE_TOLERANCE_KG
 *   Fallback dipakai kalau M_Threshold tidak punya metric ini / inactive.
 * ========================================================================= */
var LS_MASS_BALANCE_METRIC   = 'MASS_BALANCE_TOLERANCE_KG';
var LS_MASS_BALANCE_FALLBACK = { warn: 1, crit: 5 };

/* =========================================================================
 * KONSTANTA — Mapping field per kategori
 *   Disesuaikan dengan inconsistency naming di sheet existing.
 *   Format: { sheetName, dateField (alias list), columns to extract }
 * ========================================================================= */
var LS_CATEGORY_MAP = {
  coil: {
    sheet      : 'Stok_Coil',
    dateAlias  : ['Tgl_Masuk'],
    lengthAlias: [],
    availAlias : ['KG_Avail', 'Kg_Avail'],
    consKind   : 'prod_done',
    hasKeep    : true
  },
  sheet: {
    sheet      : 'Stok_Sheet',
    dateAlias  : ['Tgl_Masuk'],
    lengthAlias: ['L_dim', 'L_Dim', 'L'],
    availAlias : ['KG_Avail', 'Kg_Avail'],
    consKind   : 'prod_done',
    hasKeep    : true
  },
  wip: {
    sheet      : 'Stok_WIP',
    dateAlias  : ['Tgl_Masuk'],
    lengthAlias: ['L_dim', 'L_Dim', 'L'],
    availAlias : ['KG_Avail', 'Kg_Avail'],
    consKind   : 'prod_done',
    hasKeep    : true
  },
  fg: {
    sheet      : 'Stok_FG',
    dateAlias  : ['Tgl_Output', 'Tgl_Masuk'],
    lengthAlias: ['L_dim', 'L_Dim', 'L'],
    availAlias : ['KG_Avail', 'Kg_Avail'],
    consKind   : 'delv',
    hasKeep    : false
  },
  ng: {
    sheet      : 'Stok_NG',
    dateAlias  : ['Tgl_Generated', 'Tgl_Masuk'],
    lengthAlias: ['L', 'L_dim', 'L_Dim'],
    availAlias : ['KG_Avail', 'Kg_Avail'],
    consKind   : 'prod_done',
    hasKeep    : true
  }
};

/* =========================================================================
 * HELPER — Cari index kolom toleran terhadap variasi nama/huruf besar-kecil
 * ========================================================================= */
function _lsFindCol(headers, aliasList) {
  if (!aliasList || aliasList.length === 0) return -1;
  for (var n = 0; n < aliasList.length; n++) {
    var target = String(aliasList[n] || '').toLowerCase();
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c] || '').toLowerCase() === target) return c;
    }
  }
  return -1;
}

/* =========================================================================
 * HELPER — Hitung age dalam hari dari Date object
 * ========================================================================= */
function _lsCalcAgeDays(dateValue, today) {
  if (!dateValue) return null;
  var d = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
  if (isNaN(d.getTime())) return null;
  var msPerDay = 24 * 60 * 60 * 1000;
  // Floor agar konsisten: same-day = 0
  return Math.floor((today.getTime() - d.getTime()) / msPerDay);
}

/* =========================================================================
 * HELPER — Tentukan tier aging: 'fresh' | 'warn' | 'crit' | null
 * ========================================================================= */
function _lsAgingTier(ageDays, category) {
  if (ageDays === null || ageDays < 0) return null;
  var t = LS_AGING_THRESHOLD[category];
  if (!t) return null;
  if (ageDays > t.crit) return 'crit';
  if (ageDays > t.warn) return 'warn';
  return 'fresh';
}

/* =========================================================================
 * HELPER — Tentukan status batch:
 *   'avail' = KG_Avail > 0
 *   'keep'  = KG_Avail = 0 & KG_Keep > 0
 *   'habis' = keduanya 0
 * ========================================================================= */
function _lsStockStatus(kgAvail, kgKeep) {
  var av = parseFloat(kgAvail) || 0;
  var kp = parseFloat(kgKeep)  || 0;
  if (av > 0)  return 'avail';
  if (kp > 0)  return 'keep';
  return 'habis';
}

/* =========================================================================
 * HELPER — Build GR map (Sprint 1B)
 *   Map Batch_ID → {supplier, no_coil}
 *   Dipakai untuk enrich vendor+no_coil di WIP/FG/NG (via Root_Batch)
 *   dan Sheet turunan (via Source_Batch).
 * ========================================================================= */
function _lsBuildGrMap(ss) {
  try {
    var sh = ss.getSheetByName('GR');
    if (!sh) return {};
    var last = sh.getLastRow();
    if (last < 2) return {};
    var data = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
    var hdr  = data[0].map(function(h) { return String(h).trim(); });
    var iB = _lsFindCol(hdr, ['Batch_ID']);
    var iS = _lsFindCol(hdr, ['Supplier']);
    var iN = _lsFindCol(hdr, ['No_Coil']);
    if (iB < 0) return {};
    var map = {};
    for (var r = 1; r < data.length; r++) {
      var b = String(data[r][iB] || '').trim();
      if (!b) continue;
      map[b] = {
        supplier: iS >= 0 ? String(data[r][iS] || '').trim() : '',
        no_coil : iN >= 0 ? String(data[r][iN] || '').trim() : ''
      };
    }
    return map;
  } catch (e) {
    return {};
  }
}

/* =========================================================================
 * HELPER — Ambil threshold Mass Balance dari M_Threshold (Sprint 1B)
 *   Return { warn: <Kg>, crit: <Kg> }
 *   Fallback ke LS_MASS_BALANCE_FALLBACK kalau sheet/metric tidak ada.
 * ========================================================================= */
function _lsGetMassBalanceThreshold(ss) {
  try {
    var sh = ss.getSheetByName('M_Threshold');
    if (!sh) return LS_MASS_BALANCE_FALLBACK;
    var last = sh.getLastRow();
    if (last < 2) return LS_MASS_BALANCE_FALLBACK;
    var data = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
    var hdr  = data[0].map(function(h) { return String(h).trim(); });
    var iN  = _lsFindCol(hdr, ['Metric_Name']);
    var iW  = _lsFindCol(hdr, ['Warning_Threshold']);
    var iA  = _lsFindCol(hdr, ['Alert_Threshold']);
    var iAc = _lsFindCol(hdr, ['Active']);
    if (iN < 0) return LS_MASS_BALANCE_FALLBACK;
    for (var r = 1; r < data.length; r++) {
      var name = String(data[r][iN] || '').trim();
      if (name !== LS_MASS_BALANCE_METRIC) continue;
      var activeVal = iAc >= 0 ? String(data[r][iAc] || 'TRUE').toUpperCase() : 'TRUE';
      if (activeVal === 'FALSE') return LS_MASS_BALANCE_FALLBACK;
      return {
        warn: iW >= 0 ? (parseFloat(data[r][iW]) || LS_MASS_BALANCE_FALLBACK.warn) : LS_MASS_BALANCE_FALLBACK.warn,
        crit: iA >= 0 ? (parseFloat(data[r][iA]) || LS_MASS_BALANCE_FALLBACK.crit) : LS_MASS_BALANCE_FALLBACK.crit
      };
    }
    return LS_MASS_BALANCE_FALLBACK;
  } catch (e) {
    return LS_MASS_BALANCE_FALLBACK;
  }
}

/* =========================================================================
 * HELPER — Tentukan level selisih (Sprint 1B)
 *   Return null | 'warn' | 'crit'
 * ========================================================================= */
function _lsSelisihLevel(selisihKg, threshold) {
  var abs = Math.abs(parseFloat(selisihKg) || 0);
  if (abs > threshold.crit) return 'crit';
  if (abs > threshold.warn) return 'warn';
  return null;
}

/* =========================================================================
 * MAIN — getLiveStockData()
 *   Return all 5 kategori stok dengan aging & status enriched.
 *
 *   Output shape:
 *   {
 *     success: true,
 *     generated_at: '2026-06-25 14:32:00',
 *     summary: { coil: {...}, sheet: {...}, ... , total: {...} },
 *     batches: {
 *       coil  : [ {batch_id, item_code, ...} , ... ],
 *       sheet : [ ... ],
 *       wip   : [ ... ],
 *       fg    : [ ... ],
 *       ng    : [ ... ]
 *     }
 *   }
 * ========================================================================= */
function getLiveStockData() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var tz    = Session.getScriptTimeZone();
    var today = new Date();

    var result = {
      success     : true,
      generated_at: Utilities.formatDate(today, tz, 'yyyy-MM-dd HH:mm:ss'),
      summary     : {},
      batches     : { coil: [], sheet: [], wip: [], fg: [], ng: [] }
    };

    // Total summary akumulator (Sprint 1B: + kg_konsumsi, selisih counts)
    var grandTotal = {
      total_batch      : 0,
      kg_in            : 0,
      kg_avail         : 0,
      kg_keep          : 0,
      kg_konsumsi      : 0,
      warn_count       : 0,
      crit_count       : 0,
      selisih_warn_count: 0,
      selisih_crit_count: 0
    };

    // Sprint 1B: Build enrichment resources sekali di awal
    var grMap    = _lsBuildGrMap(ss);
    var mbThresh = _lsGetMassBalanceThreshold(ss);

    // Loop tiap kategori
    var categories = ['coil', 'sheet', 'wip', 'fg', 'ng'];

    for (var ci = 0; ci < categories.length; ci++) {
      var cat = categories[ci];
      var conf = LS_CATEGORY_MAP[cat];
      var sh   = ss.getSheetByName(conf.sheet);

      // Init category summary (Sprint 1B: + kg_konsumsi, selisih counts)
      result.summary[cat] = {
        total_batch      : 0,
        kg_in            : 0,
        kg_avail         : 0,
        kg_keep          : 0,
        kg_konsumsi      : 0,
        fresh_count      : 0,
        warn_count       : 0,
        crit_count       : 0,
        avail_count      : 0,
        keep_count       : 0,
        habis_count      : 0,
        selisih_warn_count: 0,
        selisih_crit_count: 0,
        sheet_found      : !!sh
      };

      if (!sh) continue;

      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      if (lastRow < 2 || lastCol < 1) continue;

      var data = sh.getRange(1, 1, lastRow, lastCol).getValues();
      var hdr  = data[0].map(function(h) { return String(h).trim(); });

      // Resolve kolom indices (existing)
      var iBatch = _lsFindCol(hdr, ['Batch_ID']);
      var iDate  = _lsFindCol(hdr, conf.dateAlias);
      var iItem  = _lsFindCol(hdr, ['Item_Code']);
      var iDesc  = _lsFindCol(hdr, ['Description']);
      var iSpec  = _lsFindCol(hdr, ['Spec']);
      var iT     = _lsFindCol(hdr, ['T']);
      var iP     = _lsFindCol(hdr, ['P']);
      var iL     = _lsFindCol(hdr, conf.lengthAlias);
      var iQtyIn = _lsFindCol(hdr, ['Qty_In']);
      var iKgIn  = _lsFindCol(hdr, ['KG_In']);
      var iQtyAv = _lsFindCol(hdr, ['Qty_Avail']);
      var iKgAv  = _lsFindCol(hdr, conf.availAlias);
      var iQtyKp = _lsFindCol(hdr, ['Qty_Keep']);
      var iKgKp  = _lsFindCol(hdr, ['KG_Keep']);
      var iStat  = _lsFindCol(hdr, ['Status']);
      var iOwner = _lsFindCol(hdr, ['Owner']);
      var iOwUsd = _lsFindCol(hdr, ['Owner_Used']);
      var iNote  = _lsFindCol(hdr, ['Note', 'NOTE']);
      var iSupp  = _lsFindCol(hdr, ['Supplier']);
      var iNoCoil= _lsFindCol(hdr, ['No_Coil']);
      var iSpkRef= _lsFindCol(hdr, ['SPK_Ref', 'Source_SPK']);
      var iSoRef = _lsFindCol(hdr, ['SO_Ref']);
      var iCust  = _lsFindCol(hdr, ['Cust']);
      var iTgtLoc= _lsFindCol(hdr, ['Target_Loc']);

      // Sprint 1B: kolom baru untuk Konsumsi + Vendor enrichment
      var iQProd = _lsFindCol(hdr, ['Qty_Prod']);
      var iKProd = _lsFindCol(hdr, ['KG_Prod']);
      var iQDone = _lsFindCol(hdr, ['Qty_Done']);
      var iKDone = _lsFindCol(hdr, ['KG_Done']);
      var iQDelv = _lsFindCol(hdr, ['Qty_Delv']);
      var iKDelv = _lsFindCol(hdr, ['KG_Delv']);
      var iRoot  = _lsFindCol(hdr, ['Root_Batch']);
      var iSrc   = _lsFindCol(hdr, ['Source_Batch']);

      // Loop data rows
      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var batchId = iBatch >= 0 ? String(row[iBatch] || '').trim() : '';
        if (!batchId) continue;

        var kgIn   = iKgIn  >= 0 ? (parseFloat(row[iKgIn])  || 0) : 0;
        var kgAv   = iKgAv  >= 0 ? (parseFloat(row[iKgAv])  || 0) : 0;
        var kgKp   = iKgKp  >= 0 ? (parseFloat(row[iKgKp])  || 0) : 0;
        var qtyIn  = iQtyIn >= 0 ? (parseFloat(row[iQtyIn]) || 0) : 0;
        var qtyAv  = iQtyAv >= 0 ? (parseFloat(row[iQtyAv]) || 0) : 0;
        var qtyKp  = iQtyKp >= 0 ? (parseFloat(row[iQtyKp]) || 0) : 0;

        // Sprint 1B: Konsumsi (dari Prod+Done atau Delv untuk FG)
        var qtyProd = iQProd >= 0 ? (parseFloat(row[iQProd]) || 0) : 0;
        var kgProd  = iKProd >= 0 ? (parseFloat(row[iKProd]) || 0) : 0;
        var qtyDone = iQDone >= 0 ? (parseFloat(row[iQDone]) || 0) : 0;
        var kgDone  = iKDone >= 0 ? (parseFloat(row[iKDone]) || 0) : 0;
        var qtyDelv = iQDelv >= 0 ? (parseFloat(row[iQDelv]) || 0) : 0;
        var kgDelv  = iKDelv >= 0 ? (parseFloat(row[iKDelv]) || 0) : 0;

        var qtyKons, kgKons;
        if (conf.consKind === 'delv') {
          qtyKons = qtyDelv;
          kgKons  = kgDelv;
        } else {
          qtyKons = qtyProd + qtyDone;
          kgKons  = kgProd + kgDone;
        }

        // Skip baris yang benar-benar kosong (termasuk konsumsi & delv 0)
        if (kgIn === 0 && kgAv === 0 && kgKp === 0 && qtyIn === 0 && kgKons === 0) continue;

        var dateVal = iDate >= 0 ? row[iDate] : null;
        var ageDays = _lsCalcAgeDays(dateVal, today);
        var tier    = _lsAgingTier(ageDays, cat);
        var status  = _lsStockStatus(kgAv, kgKp);

        // Sprint 1B: Selisih mass balance
        //   Coil/Sheet/WIP/NG: IN = Keep + Konsumsi + Avail
        //   FG (hasKeep=false): IN = Konsumsi(Delv) + Avail
        var kgKeepForBalance = conf.hasKeep ? kgKp : 0;
        var selisihKg = kgIn - kgKeepForBalance - kgKons - kgAv;
        var selisihLevel = _lsSelisihLevel(selisihKg, mbThresh);

        // Sprint 1B: Enrich vendor+no_coil via GR lookup (Root_Batch atau Source_Batch)
        var supplierRaw = iSupp  >= 0 ? String(row[iSupp]  || '').trim() : '';
        var noCoilRaw   = iNoCoil>= 0 ? String(row[iNoCoil]|| '').trim() : '';
        var supplier = supplierRaw;
        var noCoil   = noCoilRaw;
        if ((!supplier || !noCoil) && grMap) {
          var rootB = iRoot >= 0 ? String(row[iRoot] || '').trim() : '';
          var srcB  = iSrc  >= 0 ? String(row[iSrc]  || '').trim() : '';
          var lookupKeys = [];
          if (rootB) lookupKeys.push(rootB);
          if (srcB && srcB !== rootB) lookupKeys.push(srcB);
          for (var lk = 0; lk < lookupKeys.length; lk++) {
            var gr = grMap[lookupKeys[lk]];
            if (gr) {
              if (!supplier) supplier = gr.supplier;
              if (!noCoil)   noCoil   = gr.no_coil;
              if (supplier && noCoil) break;
            }
          }
        }

        // Build batch object
        var spec = iSpec >= 0 ? String(row[iSpec] || '').trim() : '';
        var tDim = iT    >= 0 ? row[iT]   : '';
        var pDim = iP    >= 0 ? row[iP]   : '';
        var lDim = iL    >= 0 ? row[iL]   : '';

        var specComposite = spec;
        if (tDim) specComposite += ' ' + tDim;
        if (pDim) specComposite += 'x' + pDim;
        if (lDim) specComposite += 'x' + lDim;

        var batch = {
          batch_id      : batchId,
          tgl_masuk     : dateVal ? Utilities.formatDate(new Date(dateVal), tz, 'yyyy-MM-dd') : '',
          age_days      : ageDays,
          aging_tier    : tier,
          status        : status,
          item_code     : iItem  >= 0 ? String(row[iItem]  || '').trim() : '',
          description   : iDesc  >= 0 ? String(row[iDesc]  || '').trim() : '',
          spec          : spec,
          spec_composite: specComposite.trim(),
          t             : tDim,
          p             : pDim,
          l             : lDim,
          qty_in        : qtyIn,
          kg_in         : kgIn,
          qty_avail     : qtyAv,
          kg_avail      : kgAv,
          qty_keep      : conf.hasKeep ? qtyKp : 0,
          kg_keep       : conf.hasKeep ? kgKp  : 0,
          qty_konsumsi  : qtyKons,
          kg_konsumsi   : kgKons,
          selisih_kg    : Math.round(selisihKg * 100) / 100,
          selisih_level : selisihLevel,
          stock_status  : iStat  >= 0 ? String(row[iStat]  || '').trim() : '',
          owner         : iOwner >= 0 ? String(row[iOwner] || 'FC').trim().toUpperCase() : 'FC',
          owner_used    : iOwUsd >= 0 ? String(row[iOwUsd] || '').trim().toUpperCase() : '',
          note          : iNote  >= 0 ? String(row[iNote]  || '').trim() : '',
          supplier      : supplier,
          no_coil       : noCoil,
          spk_ref       : iSpkRef>= 0 ? String(row[iSpkRef]|| '').trim() : '',
          so_ref        : iSoRef >= 0 ? String(row[iSoRef] || '').trim() : '',
          cust          : iCust  >= 0 ? String(row[iCust]  || '').trim() : '',
          target_loc    : iTgtLoc>= 0 ? String(row[iTgtLoc]|| '').trim() : ''
        };

        result.batches[cat].push(batch);

        // Update summary
        var s = result.summary[cat];
        s.total_batch++;
        s.kg_in       += kgIn;
        s.kg_avail    += kgAv;
        s.kg_keep     += conf.hasKeep ? kgKp : 0;
        s.kg_konsumsi += kgKons;

        if (tier === 'fresh') s.fresh_count++;
        else if (tier === 'warn') s.warn_count++;
        else if (tier === 'crit') s.crit_count++;

        if (status === 'avail') s.avail_count++;
        else if (status === 'keep') s.keep_count++;
        else if (status === 'habis') s.habis_count++;

        if (selisihLevel === 'warn') s.selisih_warn_count++;
        else if (selisihLevel === 'crit') s.selisih_crit_count++;
      }

      // Akumulasi grand total (Sprint 1B: + kg_konsumsi, selisih counts)
      grandTotal.total_batch       += result.summary[cat].total_batch;
      grandTotal.kg_in             += result.summary[cat].kg_in;
      grandTotal.kg_avail          += result.summary[cat].kg_avail;
      grandTotal.kg_keep           += result.summary[cat].kg_keep;
      grandTotal.kg_konsumsi       += result.summary[cat].kg_konsumsi;
      grandTotal.warn_count        += result.summary[cat].warn_count;
      grandTotal.crit_count        += result.summary[cat].crit_count;
      grandTotal.selisih_warn_count+= result.summary[cat].selisih_warn_count;
      grandTotal.selisih_crit_count+= result.summary[cat].selisih_crit_count;
    }

    result.summary.total = grandTotal;

    // Round numbers untuk display (frontend tetep terima as-is)
    return result;

  } catch (e) {
    return {
      success: false,
      message: e.message || String(e),
      stack  : e.stack || ''
    };
  }
}

/* =========================================================================
 * getLiveStockTrace(batchId)
 *   Delegate ke Trace_Service.getBatchTrace() yang udah ada.
 *   Wrapped untuk konsistensi naming dari frontend page_live_stok.
 * ========================================================================= */
function getLiveStockTrace(batchId) {
  try {
    if (!batchId) throw new Error('Batch ID kosong');
    if (typeof getBatchTrace !== 'function') {
      throw new Error('getBatchTrace() tidak ditemukan di Trace_Service.gs');
    }
    return getBatchTrace(batchId);
  } catch (e) {
    return {
      success: false,
      message: e.message || String(e)
    };
  }
}

/* =========================================================================
 * QUICK TEST — Run dari GAS editor
 * ========================================================================= */
function _test_getLiveStockData() {
  var res = getLiveStockData();
  Logger.log('Success: ' + res.success);
  if (!res.success) {
    Logger.log('Error: ' + res.message);
    Logger.log('Stack: ' + res.stack);
    return;
  }
  Logger.log('Generated at: ' + res.generated_at);
  Logger.log('--- SUMMARY ---');
  ['coil', 'sheet', 'wip', 'fg', 'ng'].forEach(function(cat) {
    var s = res.summary[cat];
    Logger.log(
      cat.toUpperCase() + ': ' +
      'batch=' + s.total_batch + ' | ' +
      'kg_avail=' + Math.round(s.kg_avail) + ' | ' +
      'warn=' + s.warn_count + ' crit=' + s.crit_count + ' | ' +
      'sheet_found=' + s.sheet_found
    );
  });
  Logger.log('--- GRAND TOTAL ---');
  Logger.log(
    'batch=' + res.summary.total.total_batch + ' | ' +
    'kg_avail=' + Math.round(res.summary.total.kg_avail) + ' | ' +
    'warn=' + res.summary.total.warn_count + ' crit=' + res.summary.total.crit_count
  );
}

/* =========================================================================
 * STOCK AUDIT — Sprint 2B
 * ---------------------------------------------------------------------
 * runStockAudit() — Jalankan audit lengkap (heavy). Hasil di-cache 6 jam.
 * getStockAudit() — Baca hasil cached (fast).
 * _clearStockAuditCache() — Utility untuk clear cache.
 * ========================================================================= */
var LS_AUDIT_CACHE_KEY = 'STOCK_AUDIT_RESULT_V2'; // Sprint 2D: struktur baru with loss_analytics
var LS_AUDIT_CACHE_TTL = 21600; // 6 jam
var LS_AUDIT_CACHE_MAX = 100000; // 100KB limit CacheService

function runStockAudit() {
  var startTime = new Date();
  try {
    // ---- Sprint 2B: Mass Balance system-wide ----
    var mbResult = getSystemMassBalance('all');
    if (!mbResult.success) {
      return { success: false, message: 'Mass balance scan failed: ' + mbResult.message };
    }

    // ---- Sprint 2C: SPK Salah Kamar (real backend) ----
    var spkResult = getSpkSalahKamar();
    var spkCheckers;
    if (spkResult && spkResult.success) {
      spkCheckers = {
        checker1: spkResult.checker1,
        checker2: spkResult.checker2,
        checker3: spkResult.checker3,
        checker4: spkResult.checker4,
        checker5: spkResult.checker5
      };
    } else {
      var errMsg = (spkResult && spkResult.message) || 'SPK checker failed';
      spkCheckers = {
        checker1: { title: 'SPK DONE tanpa Trace_Log', level: 'ok', ids: [], error: errMsg },
        checker2: { title: 'Trace_Log tanpa Stok_*',   level: 'ok', ids: [] },
        checker3: { title: 'Owner_Used mismatch',       level: 'ok', ids: [] },
        checker4: { title: 'Target_Loc mismatch',       level: 'ok', ids: [] },
        checker5: { title: 'Orphan batch (no source)',  level: 'ok', ids: [] }
      };
    }

    // ---- Sprint 2C: Reconciliation (real backend) ----
    var reconResult = getReconciliation(100);
    var reconItems    = (reconResult && reconResult.success) ? reconResult.items : [];
    var reconGapItems = (reconResult && reconResult.success) ? reconResult.summary.total_gap_items : 0;
    var reconGapKg    = (reconResult && reconResult.success) ? reconResult.summary.total_gap_kg : 0;
    var totalSpkIssues = (spkResult && spkResult.success) ? spkResult.total_issues : 0;

    // ---- Sprint 2D: Loss Analytics (real backend) ----
    var lossResult = getSpkLossReport({});
    var lossData = (lossResult && lossResult.success) ? lossResult : null;

    var tz = Session.getScriptTimeZone();
    var durationSec = Math.round((new Date() - startTime) / 100) / 10;
    var userEmail = Session.getActiveUser().getEmail() || 'system';
    var userName = userEmail.split('@')[0];

    var auditResult = {
      generated_at : Utilities.formatDate(startTime, tz, 'yyyy-MM-dd HH:mm:ss'),
      duration_sec : durationSec,
      by_user      : userName,
      scope        : 'all',
      summary      : {
        batches_audited: mbResult.summary.total_audited,
        active_count   : mbResult.summary.active,
        historical_count: mbResult.summary.historical,
        mb_warn        : mbResult.summary.warn,
        mb_crit        : mbResult.summary.crit,
        spk_issues     : totalSpkIssues,
        recon_gap_items: reconGapItems,
        recon_gap_kg   : reconGapKg
      },
      mass_balance    : mbResult.exceptions,
      spk_salah_kamar : spkCheckers,
      reconciliation  : reconItems,
      loss_analytics  : lossData
    };

    // ---- Try cache write (silent fail if too big) ----
    try {
      var cache = CacheService.getScriptCache();
      var json = JSON.stringify(auditResult);
      if (json.length < LS_AUDIT_CACHE_MAX) {
        cache.put(LS_AUDIT_CACHE_KEY, json, LS_AUDIT_CACHE_TTL);
        auditResult._cache_status = 'saved (' + json.length + ' bytes)';
      } else {
        auditResult._cache_status = 'skipped (too big: ' + json.length + ' bytes)';
      }
    } catch (cacheErr) {
      auditResult._cache_status = 'error: ' + String(cacheErr);
    }

    return { success: true, result: auditResult };
  } catch (e) {
    return { success: false, message: String(e), stack: e.stack };
  }
}

function getStockAudit() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(LS_AUDIT_CACHE_KEY);
    if (cached) {
      return { success: true, result: JSON.parse(cached), from_cache: true };
    }
    return { success: true, result: null, from_cache: false };
  } catch (e) {
    return { success: false, message: String(e) };
  }
}

function _clearStockAuditCache() {
  try {
    CacheService.getScriptCache().remove(LS_AUDIT_CACHE_KEY);
    return { success: true, message: 'Cache cleared' };
  } catch (e) {
    return { success: false, message: String(e) };
  }
}

/* Test helper */
function _test_runStockAudit() {
  var res = runStockAudit();
  if (!res.success) { Logger.log('ERR: ' + res.message); return; }
  var r = res.result;
  Logger.log('=== AUDIT COMPLETE ===');
  Logger.log('Generated: ' + r.generated_at + ' | Duration: ' + r.duration_sec + 's');
  Logger.log('By: ' + r.by_user + ' | Cache: ' + r._cache_status);
  Logger.log('--- SUMMARY ---');
  Logger.log('audited=' + r.summary.batches_audited +
             ' (active=' + r.summary.active_count +
             ', historical=' + r.summary.historical_count + ')');
  Logger.log('mass_balance: warn=' + r.summary.mb_warn + ' crit=' + r.summary.mb_crit +
             ' | total_exceptions=' + r.mass_balance.length);
  Logger.log('spk_issues=' + r.summary.spk_issues + ' (Sprint 2C)');
  Logger.log('recon_gap=' + r.summary.recon_gap_items + ' items (Sprint 2C)');
}

/* =========================================================================
 * STOCK HEALTH — Sprint 3
 * ---------------------------------------------------------------------
 * getStockHealth(opts) — Coverage matrix + FFS + DoC per material fingerprint
 *   opts.period: 'current' | 'current_next' (default) | 'all_open' | 'custom'
 *   opts.custom_periods: ['2026-07', '2026-08'] (kalau period='custom')
 *
 * Grouping: Equivalent + T + P (fingerprint mother coil)
 * Formula:
 *   Gap = MAX(0, Demand - Fisik_Item)
 *   Additional_Gap = MAX(0, Gap - Mother_Kg_Keep)
 *   Committed = Mother_Kg_Keep + Additional_Gap
 *   FFS = MAX(0, Mother_Kg_Avail - Additional_Gap)
 *   Net_Gap = (Fisik_Item + Mother_Total) - Demand
 *
 * Status:
 *   demand=0 & mother>0     → 'all_ffs'
 *   fisik >= demand         → 'ready'
 *   net_gap >= 0            → 'producible'
 *   net_gap < 0             → 'short'
 *
 * Batch-level: FIFO allocation (oldest first) untuk additional_gap.
 * ========================================================================= */

function getStockHealth(opts) {
  opts = opts || {};
  var period = opts.period || 'current_next';
  var customPeriods = opts.custom_periods || null;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();

    var itemMap = _shBuildItemMap(ss);
    var activePeriods = _shGetActivePeriods(period, customPeriods);
    var currentPeriod = _shCurrentPeriodKey();

    var demandWarnings = [];
    var demandByFp = _shAggregateDemand(ss, activePeriods, currentPeriod, itemMap, demandWarnings);
    var fisikByFp  = _shAggregateFisikItem(ss, itemMap);
    var motherByFp = _shBuildMotherCoilBatches(ss, itemMap);
    var leadTime   = _shGetLeadTime(ss);

    var allFps = {};
    Object.keys(demandByFp).forEach(function(fp) { allFps[fp] = true; });
    Object.keys(fisikByFp).forEach(function(fp) { allFps[fp] = true; });
    Object.keys(motherByFp).forEach(function(fp) { allFps[fp] = true; });

    var materials = [];
    Object.keys(allFps).forEach(function(fp) {
      var dInfo = demandByFp[fp] || { total: 0, current: 0, forecast: 0, items: [] };
      var fInfo = fisikByFp[fp]  || { total: 0, items: [] };
      var mInfo = motherByFp[fp] || { total_kg_avail: 0, total_kg_keep: 0, batches: [] };

      var demand      = dInfo.total;
      var fisikItem   = fInfo.total;
      var motherAvail = mInfo.total_kg_avail;
      var motherKeep  = mInfo.total_kg_keep;

      var gap           = Math.max(0, demand - fisikItem);
      var additionalGap = Math.max(0, gap - motherKeep);
      var committed     = motherKeep + additionalGap;
      var ffs           = Math.max(0, motherAvail - additionalGap);
      var motherTotal   = motherAvail + motherKeep;
      var totalSupply   = fisikItem + motherTotal;
      var netGap        = totalSupply - demand;

      var status;
      if (demand === 0 && motherTotal > 0) status = 'all_ffs';
      else if (fisikItem >= demand)        status = 'ready';
      else if (netGap >= 0)                status = 'producible';
      else                                  status = 'short';

      // FIFO batch allocation
      var batches = mInfo.batches.slice();
      batches.sort(function(a, b) {
        return (a._tgl_ts || 0) - (b._tgl_ts || 0);
      });
      var remaining = additionalGap;
      batches.forEach(function(b) {
        var implicit = Math.min(b.kg_avail, remaining);
        b.committed_kg = Math.round(b.kg_keep + implicit);
        b.ffs_kg       = Math.round(Math.max(0, b.kg_avail - implicit));
        if (b.ffs_kg <= 0) b.batch_status = 'full_committed';
        else if (b.ffs_kg < b.kg_avail) b.batch_status = 'partial_ffs';
        else if (b.kg_keep > 0)      b.batch_status = 'has_keep';
        else                          b.batch_status = 'all_ffs';
        remaining -= implicit;
        delete b._tgl_ts;
      });

      // DoC
      var avgDailyMat = 0;
      fInfo.items.forEach(function(it) { avgDailyMat += (it.avg_daily || 0); });
      dInfo.items.forEach(function(it) {
        // Some items may only appear in demand (not in fisik). Add their avg_daily too.
        var exists = fInfo.items.some(function(f) { return f.item_code === it.item_code; });
        if (!exists) avgDailyMat += (it.avg_daily || 0);
      });
      var docDays = avgDailyMat > 0 ? Math.round((totalSupply / avgDailyMat) * 10) / 10 : null;
      var reorderAlert = docDays !== null && docDays < leadTime;

      var parts = fp.split('|');
      materials.push({
        fingerprint     : fp,
        equivalent      : parts[0],
        t               : parseFloat(parts[1]) || 0,
        display         : parts[0] + ' ' + parts[1] + 'mm',
        demand_total    : Math.round(demand),
        demand_current  : Math.round(dInfo.current),
        demand_forecast : Math.round(dInfo.forecast),
        fisik_item      : Math.round(fisikItem),
        mother_kg_avail : Math.round(motherAvail),
        mother_kg_keep  : Math.round(motherKeep),
        mother_total    : Math.round(motherTotal),
        committed       : Math.round(committed),
        ffs             : Math.round(ffs),
        total_supply    : Math.round(totalSupply),
        net_gap         : Math.round(netGap),
        status          : status,
        doc_days        : docDays,
        avg_daily_kg    : Math.round(avgDailyMat),
        reorder_alert   : reorderAlert,
        lead_time_days  : leadTime,
        batch_count     : batches.length,
        batches         : batches,
        items_demand    : dInfo.items,
        items_fisik     : fInfo.items
      });
    });

    materials.sort(function(a, b) {
      // Sort: short first, then producible, ready, all_ffs
      var order = { short: 1, producible: 2, ready: 3, all_ffs: 4 };
      var oa = order[a.status] || 9, ob = order[b.status] || 9;
      if (oa !== ob) return oa - ob;
      return a.net_gap - b.net_gap;
    });

    var summary = {
      total_materials: materials.length,
      count_ready: 0, count_producible: 0, count_short: 0, count_all_ffs: 0,
      count_reorder: 0,
      total_ffs_kg: 0, total_demand_kg: 0, total_supply_kg: 0,
      total_committed_kg: 0
    };
    materials.forEach(function(m) {
      summary['count_' + m.status]++;
      if (m.reorder_alert) summary.count_reorder++;
      summary.total_ffs_kg       += m.ffs;
      summary.total_demand_kg    += m.demand_total;
      summary.total_supply_kg    += m.total_supply;
      summary.total_committed_kg += m.committed;
    });

    return {
      success  : true,
      materials: materials,
      summary  : summary,
      warnings : demandWarnings,
      meta     : {
        period           : period,
        active_periods   : activePeriods,
        current_period   : currentPeriod,
        lead_time_default: leadTime,
        generated_at     : new Date().toISOString()
      }
    };
  } catch (e) {
    return { success: false, message: String(e), stack: e.stack };
  }
}

/* ===== HELPERS Sprint 3 ===== */

function _shCurrentPeriodKey() {
  var d = new Date();
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
}

function _shPeriodToKey(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2);
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{1,2}$/.test(s)) {
    var parts = s.split('-');
    return parts[0] + '-' + ('0' + parts[1]).slice(-2);
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
  }
  return s.toUpperCase();
}

function _shGetActivePeriods(period, customPeriods) {
  if (period === 'all_open') return null;
  if (period === 'custom' && customPeriods && customPeriods.length > 0) {
    return customPeriods.map(_shPeriodToKey);
  }
  var now = new Date();
  var cur = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
  if (period === 'current') return [cur];
  var next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var nxt = next.getFullYear() + '-' + ('0' + (next.getMonth() + 1)).slice(-2);
  return [cur, nxt];
}

function _shBuildItemMap(ss) {
  var sh = ss.getSheetByName('M_ITEM');
  if (!sh) return {};
  var last = sh.getLastRow();
  if (last < 2) return {};
  var data = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  var hdr = data[0].map(function(h) { return String(h).trim(); });
  var iIc  = _lsFindCol(hdr, ['Item_Code']);
  var iEq  = _lsFindCol(hdr, ['Equivalent']);
  var iT   = _lsFindCol(hdr, ['T']);
  var iP   = _lsFindCol(hdr, ['P']);
  var iDs  = _lsFindCol(hdr, ['Description']);
  var iBQ  = _lsFindCol(hdr, ['BQ']);
  var iWg  = _lsFindCol(hdr, ['Wg/Pce FC', 'Wg_Pce_FC', 'Wg/Pce']);
  var iAvg = _lsFindCol(hdr, ['Avg_Daily_Consumption']);
  if (iIc < 0) return {};
  var map = {};
  for (var r = 1; r < data.length; r++) {
    var ic = String(data[r][iIc] || '').trim();
    if (!ic) continue;
    var eq = iEq >= 0 ? String(data[r][iEq] || '').trim() : '';
    var t  = iT  >= 0 ? parseFloat(data[r][iT])  || 0 : 0;
    var p  = iP  >= 0 ? parseFloat(data[r][iP])  || 0 : 0;
    var fp = (eq || 'UNKNOWN') + '|' + t;
    map[ic] = {
      item_code   : ic,
      equivalent  : eq,
      t: t, p: p,
      fingerprint : fp,
      description : iDs >= 0 ? String(data[r][iDs] || '').trim() : '',
      bq          : iBQ >= 0 ? parseFloat(data[r][iBQ]) || 0 : 0,
      wg_pce      : iWg >= 0 ? parseFloat(data[r][iWg]) || 0 : 0,
      avg_daily   : iAvg >= 0 ? parseFloat(data[r][iAvg]) || 0 : 0
    };
  }
  return map;
}

function _shAggregateDemand(ss, activePeriods, currentPeriod, itemMap, warnings) {
  var result = {};

  var addDemand = function(itemCode, kg, periodKey, sourceSheet) {
    var it = itemMap[itemCode];
    if (!it || !it.fingerprint) {
      if (warnings) {
        warnings.push({
          item_code: itemCode,
          source   : sourceSheet,
          kg       : Math.round(kg),
          reason   : !it ? 'item_not_in_master' : 'no_fingerprint'
        });
      }
      return;
    }
    var fp = it.fingerprint;
    if (!result[fp]) result[fp] = { total: 0, current: 0, forecast: 0, items: [], _itemMap: {} };
    var b = result[fp];
    b.total += kg;
    if (periodKey === currentPeriod) b.current += kg;
    else b.forecast += kg;

    if (!b._itemMap[itemCode]) {
      b._itemMap[itemCode] = { item_code: itemCode, description: it.description, kg_req: 0, current_kg: 0, forecast_kg: 0, avg_daily: it.avg_daily };
      b.items.push(b._itemMap[itemCode]);
    }
    var iRef = b._itemMap[itemCode];
    iRef.kg_req += kg;
    if (periodKey === currentPeriod) iRef.current_kg += kg;
    else iRef.forecast_kg += kg;
  };

  // Skip status: apply to SO + STP_REQ (Opsi A — align dgn Demand_Service)
  var SKIP_STATUS = { 'CANCELLED': 1, 'DONE': 1, 'CLOSED': 1 };

  // SO sheet — kolom KG pakai BL_KG (sisa), fallback SO_KG (total awal)
  var soSh = ss.getSheetByName('SO');
  if (soSh) {
    var last = soSh.getLastRow();
    if (last >= 2) {
      var data = soSh.getRange(1, 1, last, soSh.getLastColumn()).getValues();
      var hdr = data[0].map(function(h) { return String(h).trim(); });
      var iIc = _lsFindCol(hdr, ['Item_Code']);
      var iKg = _lsFindCol(hdr, ['BL_KG', 'SO_KG']);
      var iStat = _lsFindCol(hdr, ['Status']);
      var iPer = _lsFindCol(hdr, ['SO_Period', 'Period']);
      if (iIc >= 0 && iKg >= 0) {
        for (var r = 1; r < data.length; r++) {
          var status = iStat >= 0 ? String(data[r][iStat] || '').toUpperCase().trim() : '';
          if (SKIP_STATUS[status]) continue;
          var pk = iPer >= 0 ? _shPeriodToKey(data[r][iPer]) : currentPeriod;
          if (activePeriods !== null && activePeriods.indexOf(pk) < 0) continue;
          var ic = String(data[r][iIc] || '').trim();
          var kg = parseFloat(data[r][iKg]) || 0;
          if (!ic || kg <= 0) continue;
          addDemand(ic, kg, pk, 'SO');
        }
      }
    }
  }

  // STP_REQ sheet — kolom KG pakai BL_KG (sisa), fallback STP_KG (total awal)
  var stpSh = ss.getSheetByName('STP_REQ');
  if (stpSh) {
    var last = stpSh.getLastRow();
    if (last >= 2) {
      var data = stpSh.getRange(1, 1, last, stpSh.getLastColumn()).getValues();
      var hdr = data[0].map(function(h) { return String(h).trim(); });
      var iIc = _lsFindCol(hdr, ['Item_Code']);
      var iKg = _lsFindCol(hdr, ['BL_KG', 'STP_KG']);
      var iStat = _lsFindCol(hdr, ['Status']);
      var iPer = _lsFindCol(hdr, ['STP_Period']);
      if (iIc >= 0 && iKg >= 0) {
        for (var r = 1; r < data.length; r++) {
          var status = iStat >= 0 ? String(data[r][iStat] || '').toUpperCase().trim() : '';
          if (SKIP_STATUS[status]) continue;
          var pk = iPer >= 0 ? _shPeriodToKey(data[r][iPer]) : currentPeriod;
          if (activePeriods !== null && activePeriods.indexOf(pk) < 0) continue;
          var ic = String(data[r][iIc] || '').trim();
          var kg = parseFloat(data[r][iKg]) || 0;
          if (!ic || kg <= 0) continue;
          addDemand(ic, kg, pk, 'STP_REQ');
        }
      }
    }
  }

  // Clean internal map
  Object.keys(result).forEach(function(fp) {
    delete result[fp]._itemMap;
  });

  return result;
}

function _shAggregateFisikItem(ss, itemMap) {
  var result = {};
  var sheets = ['Stok_Sheet', 'Stok_WIP', 'Stok_FG'];

  var addFisik = function(itemCode, kgAvail, kgKeep, source) {
    var it = itemMap[itemCode];
    if (!it || !it.fingerprint) return;
    var fp = it.fingerprint;
    if (!result[fp]) result[fp] = { total: 0, items: [], _itemMap: {} };
    var b = result[fp];
    var totalKg = kgAvail + kgKeep;
    b.total += totalKg;

    var key = itemCode + '|' + source;
    if (!b._itemMap[key]) {
      b._itemMap[key] = { item_code: itemCode, description: it.description, source: source, kg_avail: 0, kg_keep: 0, avg_daily: it.avg_daily };
      b.items.push(b._itemMap[key]);
    }
    b._itemMap[key].kg_avail += kgAvail;
    b._itemMap[key].kg_keep  += kgKeep;
  };

  sheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 2) return;
    var data = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
    var hdr = data[0].map(function(h) { return String(h).trim(); });
    var iIc = _lsFindCol(hdr, ['Item_Code']);
    var iKa = _lsFindCol(hdr, ['KG_Avail', 'Kg_Avail']);
    var iKk = _lsFindCol(hdr, ['KG_Keep']);
    if (iIc < 0 || iKa < 0) return;
    for (var r = 1; r < data.length; r++) {
      var ic = String(data[r][iIc] || '').trim();
      var kA = parseFloat(data[r][iKa]) || 0;
      var kK = iKk >= 0 ? (parseFloat(data[r][iKk]) || 0) : 0;
      if (!ic || (kA === 0 && kK === 0)) continue;
      addFisik(ic, kA, kK, name.replace('Stok_', ''));
    }
  });

  Object.keys(result).forEach(function(fp) { delete result[fp]._itemMap; });
  return result;
}

function _shBuildMotherCoilBatches(ss, itemMap) {
  var result = {};
  var sh = ss.getSheetByName('Stok_Coil');
  if (!sh) return result;
  var last = sh.getLastRow();
  if (last < 2) return result;
  var data = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  var hdr = data[0].map(function(h) { return String(h).trim(); });
  var iB   = _lsFindCol(hdr, ['Batch_ID']);
  var iTgl = _lsFindCol(hdr, ['Tgl_Masuk']);
  var iIc  = _lsFindCol(hdr, ['Item_Code']);
  var iDs  = _lsFindCol(hdr, ['Description']);
  var iT   = _lsFindCol(hdr, ['T']);
  var iP   = _lsFindCol(hdr, ['P']);
  var iKa  = _lsFindCol(hdr, ['KG_Avail', 'Kg_Avail']);
  var iKk  = _lsFindCol(hdr, ['KG_Keep']);
  var iKi  = _lsFindCol(hdr, ['KG_In']);
  var iOw  = _lsFindCol(hdr, ['Owner']);
  var iSp  = _lsFindCol(hdr, ['Supplier']);
  var iNc  = _lsFindCol(hdr, ['No_Coil']);
  var iSt  = _lsFindCol(hdr, ['Status']);
  if (iB < 0 || iKa < 0) return result;

  for (var r = 1; r < data.length; r++) {
    var bid = String(data[r][iB] || '').trim();
    if (!bid) continue;
    var kA = parseFloat(data[r][iKa]) || 0;
    var kK = iKk >= 0 ? (parseFloat(data[r][iKk]) || 0) : 0;
    if (kA === 0 && kK === 0) continue;

    var ic = iIc >= 0 ? String(data[r][iIc] || '').trim() : '';
    var it = itemMap[ic];
    var eq = it ? it.equivalent : '';
    var t  = iT >= 0 ? parseFloat(data[r][iT]) || 0 : (it ? it.t : 0);
    var p  = iP >= 0 ? parseFloat(data[r][iP]) || 0 : (it ? it.p : 0);
    var fp = (eq || 'UNKNOWN') + '|' + t;

    if (!result[fp]) result[fp] = { total_kg_avail: 0, total_kg_keep: 0, batches: [] };
    result[fp].total_kg_avail += kA;
    result[fp].total_kg_keep  += kK;

    var tglVal = iTgl >= 0 ? data[r][iTgl] : null;
    var tglTs = tglVal instanceof Date ? tglVal.getTime() : (tglVal ? new Date(tglVal).getTime() : 0);
    var ageDays = tglVal ? _lsCalcAgeDays(tglVal, new Date()) : null;

    result[fp].batches.push({
      batch_id   : bid,
      tgl_masuk  : tglVal ? Utilities.formatDate(new Date(tglVal), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      age_days   : ageDays,
      item_code  : ic,
      description: iDs >= 0 ? String(data[r][iDs] || '').trim() : '',
      t: t, p: p,
      kg_in      : iKi >= 0 ? Math.round(parseFloat(data[r][iKi]) || 0) : 0,
      kg_avail   : Math.round(kA),
      kg_keep    : Math.round(kK),
      owner      : iOw >= 0 ? String(data[r][iOw] || 'FC').trim().toUpperCase() : 'FC',
      supplier   : iSp >= 0 ? String(data[r][iSp] || '').trim() : '',
      no_coil    : iNc >= 0 ? String(data[r][iNc] || '').trim() : '',
      status     : iSt >= 0 ? String(data[r][iSt] || '').trim() : '',
      _tgl_ts    : tglTs
    });
  }
  return result;
}

function _shGetLeadTime(ss) {
  var fallback = 14;
  try {
    var sh = ss.getSheetByName('M_Config');
    if (!sh) return fallback;
    var last = sh.getLastRow();
    if (last < 2) return fallback;
    var data = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
    var hdr = data[0].map(function(h) { return String(h).trim(); });
    var iK = _lsFindCol(hdr, ['Key', 'Metric', 'Name']);
    var iV = _lsFindCol(hdr, ['Value']);
    if (iK < 0 || iV < 0) return fallback;
    for (var r = 1; r < data.length; r++) {
      var key = String(data[r][iK] || '').trim().toUpperCase();
      if (key === 'LEAD_TIME_DEFAULT_DAYS') {
        var v = parseFloat(data[r][iV]);
        return (!isNaN(v) && v > 0) ? v : fallback;
      }
    }
    return fallback;
  } catch (e) { return fallback; }
}

/* Test helper */
function _test_getStockHealth() {
  var res = getStockHealth({ period: 'current_next' });
  if (!res.success) { Logger.log('ERR: ' + res.message); return; }
  var s = res.summary;
  Logger.log('=== STOCK HEALTH SUMMARY ===');
  Logger.log('Materials: ' + s.total_materials +
             ' | Ready: ' + s.count_ready +
             ' | Producible: ' + s.count_producible +
             ' | Short: ' + s.count_short +
             ' | All FFS: ' + s.count_all_ffs +
             ' | Reorder: ' + s.count_reorder);
  Logger.log('Total demand: ' + s.total_demand_kg + ' Kg | Supply: ' + s.total_supply_kg + ' Kg | FFS: ' + s.total_ffs_kg + ' Kg');
  Logger.log('Active periods: ' + JSON.stringify(res.meta.active_periods));

  Logger.log('\n=== TOP 5 SHORT / URGENT ===');
  res.materials.slice(0, 5).forEach(function(m) {
    Logger.log(m.display + ' [' + m.status + ']' +
               ' | demand=' + m.demand_total + ' fisik=' + m.fisik_item +
               ' mother=' + m.mother_total + ' committed=' + m.committed +
               ' ffs=' + m.ffs + ' netgap=' + m.net_gap +
               ' DoC=' + m.doc_days + 'd' + (m.reorder_alert ? ' 🚨' : ''));
  });

  Logger.log('\n=== SAMPLE BATCH DRILL-DOWN (1st material) ===');
  if (res.materials.length > 0) {
    res.materials[0].batches.slice(0, 5).forEach(function(b) {
      Logger.log('  ' + b.batch_id + ' [' + b.batch_status + '] ' +
                 (b.no_coil || '-') + '/' + (b.supplier || '-') + '/' + b.owner +
                 ' | avail=' + b.kg_avail + ' keep=' + b.kg_keep +
                 ' committed=' + b.committed_kg + ' ffs=' + b.ffs_kg);
    });
  }
}