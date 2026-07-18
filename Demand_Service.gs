/* =========================================================================
 * DEMAND_SERVICE.GS — Total Demand Aggregator (untuk panduan PPIC planning SPK)
 * File: Demand_Service.gs (BARU — terpisah dari service lain)
 *
 * Tujuan:
 *   Merangkum semua OPEN demand (SO + STP_REQ) sebagai panduan PPIC saat mau
 *   bikin SPK. Setiap baris kasih info:
 *     - Apakah sudah ada SPK-nya (spk_list)
 *     - Berapa qty yang belum tercover (net_req)
 *     - Action rekomendasi (COVER / PARTIAL / ALOKASI_FG / ALOKASI_SHEET /
 *       PERLU_SHR / PERLU_CTL / OVER)
 *   Plus opsi grouping by Equivalent+T untuk opportunity 1 mother coil =
 *   banyak SO/STP sekaligus.
 *
 * Public:
 *   getDemandData()
 *     Return bundled payload (1 round-trip untuk seluruh page):
 *       {
 *         summary:    { total_open_count, total_open_kg,
 *                       need_spk_count, need_spk_kg,
 *                       ctl_group_count, ctl_group_kg },
 *         demands:    [ ... array of demand rows ... ],
 *         ctl_groups: [ ... array of grouped rows by Equivalent+T ... ],
 *         generated_at: ISO timestamp
 *       }
 *
 * Prinsip:
 *   - READ-ONLY. GAS TIDAK menulis ke sheet apa pun di sini.
 *   - Sheet Demand_View lama TIDAK dipakai (bisa dihapus setelah UI verified).
 *   - Toleran naming: STATUS/Status, Kg/KG, dsb.
 *
 * Semantik netReq:
 *   Untuk SO:  netReq = BL_Q         − SPK_pending  (SPK non-cancelled non-DONE)
 *   Untuk STP: netReq = Qty_Sisa     − SPK_pending
 *   FG dan Sheet TIDAK ikut mengurangi netReq — dianggap "opportunity" (bisa
 *   dialokasi via action), bukan commitment otomatis.
 * ========================================================================= */

// =========================================================================
// 1. ENTRY POINT
// =========================================================================
function getDemandData() {
  var soRows    = _demReadSheet('SO');
  var stpRows   = _demReadSheet('STP_REQ');
  var spkRows   = _demReadSheet('SPK');
  var mItemRows = _demReadSheet('M_ITEM');
  var stokFG    = _demReadSheet('Stok_FG');
  var stokSheet = _demReadSheet('Stok_Sheet');

  var itemMap    = _demBuildItemMap(mItemRows);
  var spkIndex   = _demBuildSpkIndex(spkRows);       // { 'REF|ITEM': [spkRow...] }
  var fgIndex    = _demBuildFgIndex(stokFG);          // { 'REF|ITEM': totalQtyAvail }
  var sheetIndex = _demBuildSheetIndex(stokSheet, itemMap);

  var demands = [];

  soRows.forEach(function(row) {
    var d = _demBuildFromSO(row, itemMap, spkIndex, fgIndex, sheetIndex);
    if (d) demands.push(d);
  });

  stpRows.forEach(function(row) {
    var d = _demBuildFromSTP(row, itemMap, spkIndex, fgIndex, sheetIndex);
    if (d) demands.push(d);
  });

  // Sort by tgl_needed ASC (yang paling urgent duluan), fallback ref_no ASC
  demands.sort(function(a, b) {
    var ta = a.tgl_needed_ts || Number.MAX_SAFE_INTEGER;
    var tb = b.tgl_needed_ts || Number.MAX_SAFE_INTEGER;
    if (ta !== tb) return ta - tb;
    return String(a.ref_no).localeCompare(String(b.ref_no));
  });

  var ctlGroups = _demBuildCtlGroups(demands);
  var summary   = _demBuildSummary(demands, ctlGroups);

  return {
    summary:      summary,
    demands:      demands,
    ctl_groups:   ctlGroups,
    generated_at: new Date().toISOString()
  };
}


// =========================================================================
// 2. SHEET READER — preserve raw values (Date object tetap Date)
// =========================================================================
function _demReadSheet(sheetName) {
  var sheet = getSheet(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row.join('').trim() === '') continue;
    var obj = {};
    headers.forEach(function(h, idx){ obj[h] = row[idx]; });
    result.push(obj);
  }
  return result;
}


// =========================================================================
// 3. LOOKUP BUILDERS
// =========================================================================

// UoM di-derive dari kolom TYPE di M_ITEM (bukan dari "Unit of Measure").
// Alasan: untuk customer SUP & HKB, kolom "Unit of Measure" di M_ITEM di-set KG
// (karena SO mereka beli by berat), padahal system COS butuh satuan fisik sesuai
// bentuk material.
//   Part  → PCS
//   Sheet → LBR
//   Coil  → KG
// Fallback '' kalau TYPE kosong/nggak dikenal — biar cepet ketahuan item bermasalah.
function _demUomFromType(type) {
  var t = String(type || '').trim().toLowerCase();
  if (t === 'part')  return 'PCS';
  if (t === 'sheet') return 'LBR';
  if (t === 'coil')  return 'KG';
  return '';
}

function _demBuildItemMap(mItemRows) {
  // { item_code: { description, spec, t, p, l, uom, equivalent, wg_pce } }
  var map = {};
  mItemRows.forEach(function(row) {
    var code = String(row['Item_Code'] || '').trim();
    if (!code) return;
    map[code] = {
      description: String(row['Description'] || '').trim(),
      spec:        String(row['Spec'] || '').trim(),
      t:           _demNum(row['T']),
      p:           _demNum(row['P']),
      l:           _demNum(row['L']),
      type:        String(row['TYPE'] || '').trim(),
      uom:         _demUomFromType(row['TYPE']),
      equivalent:  String(row['Equivalent'] || '').trim(),
      wg_pce:      _demNum(row['Wg/Pce FC'])
    };
  });
  return map;
}

function _demBuildSpkIndex(spkRows) {
  // { 'SO_Ref|Item_Code': [ spkRow, ... ]  }  — hanya non-CANCELLED
  var idx = {};
  spkRows.forEach(function(row) {
    var status = String(row['Status'] || '').toUpperCase();
    if (status === 'CANCELLED') return;
    var soRef = String(row['SO_Ref'] || '').trim();
    var item  = String(row['Item_Code'] || '').trim();
    if (!soRef || !item) return;
    var key = soRef + '|' + item;
    if (!idx[key]) idx[key] = [];
    idx[key].push(row);
  });
  return idx;
}

function _demBuildFgIndex(fgRows) {
  // { 'SO_Ref|Item_Code': totalQtyAvail }
  var idx = {};
  fgRows.forEach(function(row) {
    var soRef = String(row['SO_Ref'] || '').trim();
    var item  = String(row['Item_Code'] || '').trim();
    if (!item) return;
    var qty = _demNum(row['Qty_Avail']);
    if (qty <= 0) return;
    var key = soRef + '|' + item;
    idx[key] = (idx[key] || 0) + qty;
  });
  return idx;
}

function _demBuildSheetIndex(sheetRows, itemMap) {
  // { byItem:   { item_code: totalQtyAvail },
  //   byEquivT: { 'equivalent|t': totalQtyAvail } }
  var byItem = {};
  var byEquivT = {};
  sheetRows.forEach(function(row) {
    var item = String(row['Item_Code'] || '').trim();
    if (!item) return;
    // Toleran naming: Qty_Avail (standard)
    var qty = _demNum(row['Qty_Avail']);
    if (qty <= 0) return;
    byItem[item] = (byItem[item] || 0) + qty;

    var mi = itemMap[item];
    if (mi && mi.equivalent && mi.t) {
      var key = mi.equivalent + '|' + mi.t;
      byEquivT[key] = (byEquivT[key] || 0) + qty;
    }
  });
  return { byItem: byItem, byEquivT: byEquivT };
}


// =========================================================================
// 4. BUILD DEMAND ROW — FROM SO
// =========================================================================
function _demBuildFromSO(soRow, itemMap, spkIndex, fgIndex, sheetIndex) {
  var soNo = String(soRow['SO_No'] || '').trim();
  if (!soNo) return null;

  var itemCode = String(soRow['Item_Code'] || '').trim();
  if (!itemCode) return null;

  var statusRaw = String(soRow['STATUS'] || soRow['Status'] || '').toUpperCase();
  if (statusRaw === 'CANCELLED') return null;
  if (statusRaw === 'DONE')      return null;  // fully delivered — skip
  if (statusRaw === 'CLOSED')    return null;  // manual closed — skip dari planning

  var qtyDemand = _demNum(soRow['BL_Q']);
  var kgDemand  = _demNum(soRow['BL_KG']);

  // Defensive: skip kalau BL_Q = 0 dan bukan OVER
  if (qtyDemand === 0 && statusRaw !== 'OVER') return null;

  var mi = itemMap[itemCode] || {};
  var equivalent = mi.equivalent || '';
  var t = _demNum(soRow['T']) || mi.t || 0;

  var spkKey = soNo + '|' + itemCode;
  var relatedSpk = spkIndex[spkKey] || [];

  var spkListDisplay = _demBuildSpkListForDisplay(relatedSpk);

  // netReq semantik: SPK yang belum DONE dianggap "in progress akan cover"
  var qtySpkPending = _demSumSpk(relatedSpk, function(r) {
    var t2 = String(r['SPK_Type'] || '').toUpperCase();
    var st = String(r['Status']   || '').toUpperCase();
    return (t2 === 'SHR-OUT' || t2 === 'ALLOC-OUT') && st !== 'DONE';
  });

  var netReq = qtyDemand - qtySpkPending;

  var stokFgQty      = fgIndex[spkKey] || 0;
  var stokSheetExact = sheetIndex.byItem[itemCode] || 0;
  var stokSheetStd   = (equivalent && t) ? (sheetIndex.byEquivT[equivalent + '|' + t] || 0) : 0;

  var action = _demComputeAction({
    statusRaw:      statusRaw,
    netReq:         netReq,
    fg:             stokFgQty,
    sheetExact:     stokSheetExact,
    sheetStd:       stokSheetStd,
    qtySpkPending:  qtySpkPending
  });

  var tglNeeded = _demToDate(soRow['SCHEDULE_DATE']);
  var periode   = String(soRow['SO_Period'] || '').trim();
  if (!periode && tglNeeded) periode = _demFormatPeriod(tglNeeded);

  return {
    tipe:             'SO',
    ref_no:           soNo,
    cust:             String(soRow['Cust'] || '').trim(),
    periode:          periode,
    tgl_needed:       _demFormatDate(tglNeeded),
    tgl_needed_ts:    tglNeeded ? tglNeeded.getTime() : 0,
    item_code:        itemCode,
    description:      String(soRow['Description'] || mi.description || '').trim(),
    spec:             String(soRow['Spec'] || mi.spec || '').trim(),
    t:                t,
    p:                _demNum(soRow['P']) || mi.p || 0,
    l:                _demNum(soRow['L']) || mi.l || 0,
    uom:              mi.uom || '',
    qty_demand:       qtyDemand,
    kg_demand:        kgDemand,
    spk_list:         spkListDisplay,
    qty_spk_pending:  qtySpkPending,
    net_req:          netReq,
    stok_fg_qty:      stokFgQty,
    stok_sheet_exact: stokSheetExact,
    stok_sheet_std:   stokSheetStd,
    action:           action,
    owner_used:       String(soRow['Owner_Used'] || 'FC').trim(),
    status:           statusRaw,
    equivalent:       equivalent,
    material_key:     (equivalent && t) ? (equivalent + '|' + t) : ''
  };
}


// =========================================================================
// 5. BUILD DEMAND ROW — FROM STP_REQ
// =========================================================================
function _demBuildFromSTP(stpRow, itemMap, spkIndex, fgIndex, sheetIndex) {
  var stpNo = String(stpRow['STP_No'] || '').trim();
  if (!stpNo) return null;

  var itemCode = String(stpRow['Item_Code'] || '').trim();
  if (!itemCode) return null;

  var statusRaw = String(stpRow['Status'] || '').toUpperCase();
  if (statusRaw !== 'OPEN') return null;

  var qtyReq  = _demNum(stpRow['Qty_Req']);
  var qtyFul  = _demNum(stpRow['Qty_Fulfill']);
  var qtySisa = _demNum(stpRow['Qty_Sisa']);
  if (!qtySisa && qtyReq) qtySisa = qtyReq - qtyFul;
  if (qtySisa <= 0) return null;

  var qtyDemand = qtySisa;

  var mi = itemMap[itemCode] || {};
  var equivalent = mi.equivalent || '';
  var t = _demNum(stpRow['T']) || mi.t || 0;
  var wgPce = mi.wg_pce || 0;

  // KG_Req dari sheet biasanya sudah dihitung; kalau kosong pakai qtyDemand * wg_pce
  var kgDemand = _demNum(stpRow['KG_Req']);
  if (!kgDemand && wgPce) kgDemand = qtyDemand * wgPce;

  var spkKey = stpNo + '|' + itemCode;
  var relatedSpk = spkIndex[spkKey] || [];
  var spkListDisplay = _demBuildSpkListForDisplay(relatedSpk);

  // Untuk STP, Qty_Sisa sudah net dari SPK-DONE (via formula Qty_Fulfill).
  // Jadi netReq = Qty_Sisa − SPK_non_DONE (yang masih antrian/running).
  var qtySpkPending = _demSumSpk(relatedSpk, function(r) {
    var t2 = String(r['SPK_Type'] || '').toUpperCase();
    var st = String(r['Status']   || '').toUpperCase();
    return (t2 === 'SHR-OUT' || t2 === 'ALLOC-OUT') && st !== 'DONE';
  });

  var netReq = qtyDemand - qtySpkPending;

  var stokFgQty      = fgIndex[spkKey] || 0;
  var stokSheetExact = sheetIndex.byItem[itemCode] || 0;
  var stokSheetStd   = (equivalent && t) ? (sheetIndex.byEquivT[equivalent + '|' + t] || 0) : 0;

  var action = _demComputeAction({
    statusRaw:      statusRaw,
    netReq:         netReq,
    fg:             stokFgQty,
    sheetExact:     stokSheetExact,
    sheetStd:       stokSheetStd,
    qtySpkPending:  qtySpkPending
  });

  var tglNeeded = _demToDate(stpRow['Schedule_Date']);
  var periode   = String(stpRow['Periode'] || '').trim();
  if (!periode && tglNeeded) periode = _demFormatPeriod(tglNeeded);

  return {
    tipe:             'STP',
    ref_no:           stpNo,
    cust:             String(stpRow['Cust'] || '').trim(),
    periode:          periode,
    tgl_needed:       _demFormatDate(tglNeeded),
    tgl_needed_ts:    tglNeeded ? tglNeeded.getTime() : 0,
    item_code:        itemCode,
    description:      String(stpRow['Description'] || mi.description || '').trim(),
    spec:             String(stpRow['Spec'] || mi.spec || '').trim(),
    t:                t,
    p:                _demNum(stpRow['P']) || mi.p || 0,
    l:                _demNum(stpRow['L']) || mi.l || 0,
    uom:              mi.uom || '',
    qty_demand:       qtyDemand,
    kg_demand:        kgDemand,
    spk_list:         spkListDisplay,
    qty_spk_pending:  qtySpkPending,
    net_req:          netReq,
    stok_fg_qty:      stokFgQty,
    stok_sheet_exact: stokSheetExact,
    stok_sheet_std:   stokSheetStd,
    action:           action,
    owner_used:       String(stpRow['Owner_Used'] || 'FC').trim(),
    status:           statusRaw,
    equivalent:       equivalent,
    material_key:     (equivalent && t) ? (equivalent + '|' + t) : ''
  };
}


// =========================================================================
// 6. SPK DISPLAY LIST — kolom yang muncul di modal detail
// =========================================================================
function _demBuildSpkListForDisplay(spkArr) {
  if (!spkArr || !spkArr.length) return [];
  return spkArr.map(function(r) {
    return {
      spk_no:           String(r['SPK_No'] || '').trim(),
      spk_type:         String(r['SPK_Type'] || '').trim(),
      parent_spk:       String(r['Parent_SPK'] || '').trim(),
      status:           String(r['Status'] || '').trim(),
      qty_target:       _demNum(r['Qty_Target']),
      kg_target:        _demNum(r['KG_Target']),
      qty_actual:       _demNum(r['Qty_Actual']),
      kg_actual:        _demNum(r['KG_Actual']),
      mc_no:            String(r['MC_No'] || '').trim(),
      priority:         _demNum(r['Priority']),
      tgl_buat:         _demFormatDate(_demToDate(r['Tgl_Buat'])),
      estimasi_mulai:   _demFormatDateTime(_demToDate(r['Estimasi_Jam_Mulai'])),
      estimasi_selesai: _demFormatDateTime(_demToDate(r['Estimasi_Jam_Selesai'])),
      owner:            String(r['Owner'] || '').trim(),
      owner_used:       String(r['Owner_Used'] || '').trim(),
      leader:           String(r['Leader'] || '').trim(),
      note:             String(r['NOTE'] || '').trim()
    };
  }).sort(function(a, b) {
    // Sort by spk_no ASC (chronological karena format YYNNNN)
    return String(a.spk_no).localeCompare(String(b.spk_no));
  });
}


// =========================================================================
// 7. HELPER — sum Qty_Target dari SPK array berdasarkan predicate
// =========================================================================
function _demSumSpk(spkArr, predicate) {
  if (!spkArr || !spkArr.length) return 0;
  var sum = 0;
  for (var i = 0; i < spkArr.length; i++) {
    if (predicate(spkArr[i])) sum += _demNum(spkArr[i]['Qty_Target']);
  }
  return sum;
}


// =========================================================================
// 8. ACTION LOGIC
//   Priority order (top to bottom):
//     1. OVER          → status SO over-delivery
//     2. COVER         → netReq <= 0 (SPK sudah cukup / sudah aman)
//     3. ALOKASI_FG    → ada FG matching yang bisa cover netReq
//     4. ALOKASI_SHEET → ada Sheet exact match yang bisa cover netReq
//     5. PARTIAL       → sudah ada SPK jalan tapi belum cukup, perlu tambah
//     6. PERLU_SHR     → ada Sheet standar (via Equivalent+T), bisa di-cut
//     7. PERLU_CTL     → default, harus proses dari coil
// =========================================================================
function _demComputeAction(ctx) {
  if (ctx.statusRaw === 'OVER') return 'OVER';
  if (ctx.netReq <= 0)          return 'COVER';

  if (ctx.fg > 0 && ctx.fg >= ctx.netReq)                    return 'ALOKASI_FG';
  if (ctx.sheetExact > 0 && ctx.sheetExact >= ctx.netReq)    return 'ALOKASI_SHEET';
  if (ctx.qtySpkPending > 0)                                  return 'PARTIAL';
  if (ctx.sheetStd > 0)                                        return 'PERLU_SHR';
  return 'PERLU_CTL';
}


// =========================================================================
// 9. CTL GROUPING — Equivalent + T
//   Group hanya row yang butuh proses baru (PERLU_CTL / PERLU_SHR / PARTIAL).
//   Row COVER / ALOKASI_* / OVER TIDAK masuk grouping (sudah ada solusi).
//   Owner_Used di-mix (per keputusan Herryna), tapi Owner tetap tampil per row.
// =========================================================================
function _demBuildCtlGroups(demands) {
  var GROUPABLE_ACTIONS = { PERLU_CTL: 1, PERLU_SHR: 1, PARTIAL: 1 };
  var groups = {};

  demands.forEach(function(d) {
    if (d.net_req <= 0) return;
    if (!GROUPABLE_ACTIONS[d.action]) return;
    if (!d.material_key) return;

    if (!groups[d.material_key]) {
      groups[d.material_key] = {
        material_key: d.material_key,
        equivalent:   d.equivalent,
        t:            d.t,
        total_kg:     0,
        total_qty:    0,
        count_so:     0,
        count_stp:    0,
        demand_refs:  []
      };
    }
    var g = groups[d.material_key];
    // Total KG di group hitung berdasar netReq (bukan qty_demand penuh)
    // supaya cerminan yang benar-benar perlu diproses.
    var kgPortion = 0;
    if (d.qty_demand > 0) {
      kgPortion = (d.kg_demand || 0) * (d.net_req / d.qty_demand);
    }
    g.total_kg  += kgPortion;
    g.total_qty += d.net_req;
    if (d.tipe === 'SO')  g.count_so  += 1;
    if (d.tipe === 'STP') g.count_stp += 1;

    // Composite key untuk ref di frontend (satu SO bisa punya banyak item)
    g.demand_refs.push(d.ref_no + '||' + d.item_code);
  });

  var arr = Object.keys(groups).map(function(k) {
    var g = groups[k];
    g.total_kg = Math.round(g.total_kg);
    return g;
  });
  arr.sort(function(a, b) { return b.total_kg - a.total_kg; });
  return arr;
}


// =========================================================================
// 10. SUMMARY CARDS
// =========================================================================
function _demBuildSummary(demands, ctlGroups) {
  var totalOpenCount = 0, totalOpenKg = 0;
  var needSpkCount   = 0, needSpkKg   = 0;
  var ctlKg          = 0;

  var NEED_SPK_ACTIONS = { PERLU_CTL: 1, PERLU_SHR: 1, PARTIAL: 1 };

  demands.forEach(function(d) {
    totalOpenCount += 1;
    totalOpenKg    += (d.kg_demand || 0);

    if (NEED_SPK_ACTIONS[d.action]) {
      needSpkCount += 1;
      // Portion kg untuk qty yang net_req saja
      if (d.qty_demand > 0) {
        needSpkKg += (d.kg_demand || 0) * (d.net_req / d.qty_demand);
      }
    }
  });

  ctlGroups.forEach(function(g) { ctlKg += g.total_kg; });

  return {
    total_open_count: totalOpenCount,
    total_open_kg:    Math.round(totalOpenKg),
    need_spk_count:   needSpkCount,
    need_spk_kg:      Math.round(needSpkKg),
    ctl_group_count:  ctlGroups.length,
    ctl_group_kg:     Math.round(ctlKg)
  };
}


// =========================================================================
// 11. UTIL — number & date helpers
// =========================================================================
function _demNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function _demToDate(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function _demFormatDate(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy');
}

function _demFormatDateTime(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
}

function _demFormatPeriod(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM yyyy');
}


// =========================================================================
// 12. TEST HELPER (opsional — jalankan manual di GAS editor untuk cek payload)
// =========================================================================
function _demTest_getDemandData() {
  var result = getDemandData();
  Logger.log('=== SUMMARY ===');
  Logger.log(JSON.stringify(result.summary, null, 2));
  Logger.log('=== DEMANDS COUNT: ' + result.demands.length + ' ===');
  if (result.demands.length > 0) {
    Logger.log('First row:');
    Logger.log(JSON.stringify(result.demands[0], null, 2));
  }
  Logger.log('=== CTL GROUPS COUNT: ' + result.ctl_groups.length + ' ===');
  if (result.ctl_groups.length > 0) {
    Logger.log('First group:');
    Logger.log(JSON.stringify(result.ctl_groups[0], null, 2));
  }
  return result;
}