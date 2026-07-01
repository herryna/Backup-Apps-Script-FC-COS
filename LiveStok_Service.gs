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
 * KONSTANTA — Mapping field per kategori
 *   Disesuaikan dengan inconsistency naming di sheet existing.
 *   Format: { sheetName, dateField (alias list), columns to extract }
 * ========================================================================= */
var LS_CATEGORY_MAP = {
  coil: {
    sheet      : 'Stok_Coil',
    dateAlias  : ['Tgl_Masuk'],
    lengthAlias: [], // coil gak punya L (continuous)
    availAlias : ['KG_Avail', 'Kg_Avail']
  },
  sheet: {
    sheet      : 'Stok_Sheet',
    dateAlias  : ['Tgl_Masuk'],
    lengthAlias: ['L_dim', 'L_Dim', 'L'],
    availAlias : ['KG_Avail', 'Kg_Avail']
  },
  wip: {
    sheet      : 'Stok_WIP',
    dateAlias  : ['Tgl_Masuk'],
    lengthAlias: ['L_dim', 'L_Dim', 'L'],
    availAlias : ['KG_Avail', 'Kg_Avail']
  },
  fg: {
    sheet      : 'Stok_FG',
    dateAlias  : ['Tgl_Output', 'Tgl_Masuk'],
    lengthAlias: ['L_dim', 'L_Dim', 'L'],
    availAlias : ['KG_Avail', 'Kg_Avail']
  },
  ng: {
    sheet      : 'Stok_NG',
    dateAlias  : ['Tgl_Generated', 'Tgl_Masuk'],
    lengthAlias: ['L', 'L_dim', 'L_Dim'],
    availAlias : ['KG_Avail', 'Kg_Avail']
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

    // Total summary akumulator
    var grandTotal = {
      total_batch: 0,
      kg_in      : 0,
      kg_avail   : 0,
      kg_keep    : 0,
      warn_count : 0,
      crit_count : 0
    };

    // Loop tiap kategori
    var categories = ['coil', 'sheet', 'wip', 'fg', 'ng'];

    for (var ci = 0; ci < categories.length; ci++) {
      var cat = categories[ci];
      var conf = LS_CATEGORY_MAP[cat];
      var sh   = ss.getSheetByName(conf.sheet);

      // Init category summary
      result.summary[cat] = {
        total_batch: 0,
        kg_in      : 0,
        kg_avail   : 0,
        kg_keep    : 0,
        fresh_count: 0,
        warn_count : 0,
        crit_count : 0,
        avail_count: 0,
        keep_count : 0,
        habis_count: 0,
        sheet_found: !!sh
      };

      if (!sh) continue;

      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      if (lastRow < 2 || lastCol < 1) continue;

      var data = sh.getRange(1, 1, lastRow, lastCol).getValues();
      var hdr  = data[0].map(function(h) { return String(h).trim(); });

      // Resolve kolom indices
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

      // Loop data rows
      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var batchId = iBatch >= 0 ? String(row[iBatch] || '').trim() : '';
        if (!batchId) continue; // skip baris kosong (ARRAYFORMULA spill, dll)

        var kgIn   = iKgIn  >= 0 ? (parseFloat(row[iKgIn])  || 0) : 0;
        var kgAv   = iKgAv  >= 0 ? (parseFloat(row[iKgAv])  || 0) : 0;
        var kgKp   = iKgKp  >= 0 ? (parseFloat(row[iKgKp])  || 0) : 0;
        var qtyIn  = iQtyIn >= 0 ? (parseFloat(row[iQtyIn]) || 0) : 0;
        var qtyAv  = iQtyAv >= 0 ? (parseFloat(row[iQtyAv]) || 0) : 0;
        var qtyKp  = iQtyKp >= 0 ? (parseFloat(row[iQtyKp]) || 0) : 0;

        // Skip baris yang fully nol semua (zombie row dari formula)
        if (kgIn === 0 && kgAv === 0 && kgKp === 0 && qtyIn === 0) continue;

        var dateVal = iDate >= 0 ? row[iDate] : null;
        var ageDays = _lsCalcAgeDays(dateVal, today);
        var tier    = _lsAgingTier(ageDays, cat);
        var status  = _lsStockStatus(kgAv, kgKp);

        // Build batch object
        var spec = iSpec >= 0 ? String(row[iSpec] || '').trim() : '';
        var tDim = iT    >= 0 ? row[iT]   : '';
        var pDim = iP    >= 0 ? row[iP]   : '';
        var lDim = iL    >= 0 ? row[iL]   : '';

        // Composite spec untuk display (terutama NG yang gak punya Item_Code)
        var specComposite = spec;
        if (tDim) specComposite += ' ' + tDim;
        if (pDim) specComposite += 'x' + pDim;
        if (lDim) specComposite += 'x' + lDim;

        var batch = {
          batch_id     : batchId,
          tgl_masuk    : dateVal ? Utilities.formatDate(new Date(dateVal), tz, 'yyyy-MM-dd') : '',
          age_days     : ageDays,
          aging_tier   : tier,
          status       : status,
          item_code    : iItem  >= 0 ? String(row[iItem]  || '').trim() : '',
          description  : iDesc  >= 0 ? String(row[iDesc]  || '').trim() : '',
          spec         : spec,
          spec_composite: specComposite.trim(),
          t            : tDim,
          p            : pDim,
          l            : lDim,
          qty_in       : qtyIn,
          kg_in        : kgIn,
          qty_avail    : qtyAv,
          kg_avail     : kgAv,
          qty_keep     : qtyKp,
          kg_keep      : kgKp,
          stock_status : iStat  >= 0 ? String(row[iStat]  || '').trim() : '',
          owner        : iOwner >= 0 ? String(row[iOwner] || 'FC').trim().toUpperCase() : 'FC',
          owner_used   : iOwUsd >= 0 ? String(row[iOwUsd] || '').trim().toUpperCase() : '',
          note         : iNote  >= 0 ? String(row[iNote]  || '').trim() : '',
          supplier     : iSupp  >= 0 ? String(row[iSupp]  || '').trim() : '',
          no_coil      : iNoCoil>= 0 ? String(row[iNoCoil]|| '').trim() : '',
          spk_ref      : iSpkRef>= 0 ? String(row[iSpkRef]|| '').trim() : '',
          so_ref       : iSoRef >= 0 ? String(row[iSoRef] || '').trim() : '',
          cust         : iCust  >= 0 ? String(row[iCust]  || '').trim() : '',
          target_loc   : iTgtLoc>= 0 ? String(row[iTgtLoc]|| '').trim() : ''
        };

        result.batches[cat].push(batch);

        // Update summary
        var s = result.summary[cat];
        s.total_batch++;
        s.kg_in    += kgIn;
        s.kg_avail += kgAv;
        s.kg_keep  += kgKp;

        if (tier === 'fresh') s.fresh_count++;
        else if (tier === 'warn') s.warn_count++;
        else if (tier === 'crit') s.crit_count++;

        if (status === 'avail') s.avail_count++;
        else if (status === 'keep') s.keep_count++;
        else if (status === 'habis') s.habis_count++;
      }

      // Akumulasi grand total
      grandTotal.total_batch += result.summary[cat].total_batch;
      grandTotal.kg_in       += result.summary[cat].kg_in;
      grandTotal.kg_avail    += result.summary[cat].kg_avail;
      grandTotal.kg_keep     += result.summary[cat].kg_keep;
      grandTotal.warn_count  += result.summary[cat].warn_count;
      grandTotal.crit_count  += result.summary[cat].crit_count;
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