/* =========================================================================
 * DEMAND_SERVICE.GS — Demand Planning Aggregator
 * 
 * Sprint 1 Refactor (2026-07-23):
 *   - Semantik Demand STATIC: pakai SO_Q/SO_KG (bukan BL_Q/BL_KG lagi)
 *   - Kolom baru: done_qty/kg (SPK DONE), delv_qty/kg (dari DELV+DELV_STP)
 *   - Status 5-tier: UNPLANNED / PARTIAL / COVERED / PRODUCED / FULFILLED
 *   - Warnings: leak / over_produce / over_delivery
 *   - Info Stok: Coil + Sheet + WIP (+ NG di bawah) by Equal+T fingerprint
 *   - STP_REQ header baru: STP_Q, BL_Q, STP_KG, BL_KG, STP_Period, Schedule_Date
 *   - Backward compat: field lama (action, stok_fg_qty, stok_sheet_*) tetap ada
 *
 * Public:
 *   getDemandData() → { summary, demands, ctl_groups, stok_info, generated_at }
 * ========================================================================= */

// =========================================================================
// 1. ENTRY POINT
// =========================================================================
function getDemandData() {
  var soRows      = _demReadSheet('SO');
  var stpRows     = _demReadSheet('STP_REQ');
  var spkRows     = _demReadSheet('SPK');
  var mItemRows   = _demReadSheet('M_ITEM');
  var stokFG      = _demReadSheet('Stok_FG');
  var stokSheet   = _demReadSheet('Stok_Sheet');
  var stokCoil    = _demReadSheet('Stok_Coil');
  var stokWIP     = _demReadSheet('Stok_WIP');
  var stokNG      = _demReadSheet('Stok_NG');
  var delvRows    = _demReadSheet('DELV');
  var delvStpRows = _demReadSheet('DELV_STP');

  var itemMap    = _demBuildItemMap(mItemRows);
  var spkIndex   = _demBuildSpkIndex(spkRows);
  var fgIndex    = _demBuildFgIndex(stokFG);
  var sheetIndex = _demBuildSheetIndex(stokSheet, itemMap);
  var delvSo     = _demBuildDelvIndex(delvRows, 'SO_No');
  var delvStp    = _demBuildDelvIndex(delvStpRows, 'STP_No');
  var delvBySpk  = _demBuildDelvBySpk(delvRows, delvStpRows);
  var stokInfo   = _demBuildStokInfoIndex(stokCoil, stokSheet, stokWIP, stokNG, itemMap);

  var ctx = {
    itemMap: itemMap, spkIndex: spkIndex,
    fgIndex: fgIndex, sheetIndex: sheetIndex,
    delvSo: delvSo, delvStp: delvStp, delvBySpk: delvBySpk
  };

  var demands = [];
  soRows.forEach(function(row) {
    var d = _demBuildFromSO(row, ctx);
    if (d) demands.push(d);
  });
  stpRows.forEach(function(row) {
    var d = _demBuildFromSTP(row, ctx);
    if (d) demands.push(d);
  });

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
    stok_info:    stokInfo,
    generated_at: new Date().toISOString()
  };
}

// =========================================================================
// 2. SHEET READER
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
function _demUomFromType(type) {
  var t = String(type || '').trim().toLowerCase();
  if (t === 'part')  return 'PCS';
  if (t === 'sheet') return 'LBR';
  if (t === 'coil')  return 'KG';
  return '';
}

function _demBuildItemMap(mItemRows) {
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
  // Include SEMUA SPK kecuali CANCELLED (termasuk DONE, RUNNING, ANTRIAN, HOLD)
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
  var byItem = {};
  var byEquivT = {};
  sheetRows.forEach(function(row) {
    var item = String(row['Item_Code'] || '').trim();
    if (!item) return;
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

// NEW: Delivery index by (Ref_No + Item_Code) — pisah SO dan STP
function _demBuildDelvIndex(delvRows, refCol) {
  // Return: { 'REF|ITEM': { qty: X, kg: Y } }
  var idx = {};
  delvRows.forEach(function(row) {
    var ref  = String(row[refCol] || '').trim();
    var item = String(row['Item_Code'] || '').trim();
    if (!ref || !item) return;
    var q = _demNum(row['Delv_Q']);
    var k = _demNum(row['Delv_KG']);
    var key = ref + '|' + item;
    if (!idx[key]) idx[key] = { qty: 0, kg: 0 };
    idx[key].qty += q;
    idx[key].kg  += k;
  });
  return idx;
}

// NEW: Delivery index by SPK_No — untuk modal detail SPK per-baris
function _demBuildDelvBySpk(delvRows, delvStpRows) {
  var idx = {};
  var addRow = function(row) {
    var spk = String(row['Spk_ref'] || '').trim();
    if (!spk) return;
    var q = _demNum(row['Delv_Q']);
    if (!idx[spk]) idx[spk] = { qty: 0, kg: 0 };
    idx[spk].qty += q;
    idx[spk].kg  += _demNum(row['Delv_KG']);
  };
  delvRows.forEach(addRow);
  delvStpRows.forEach(addRow);
  return idx;
}

// NEW: Stok info by Equal+T fingerprint — untuk modal Info Stok
// Include: Coil, Sheet, WIP, NG (skip FG). NG selalu di posisi bawah.
function _demBuildStokInfoIndex(coilRows, sheetRows, wipRows, ngRows, itemMap) {
  var idx = {}; // { 'EQ|T': [ {posisi, item_code, description, batch_id, owner, avail, uom, tgl_masuk_ts, tgl_masuk_display}, ... ] }

  var posisiOrder = { 'Coil': 1, 'Sheet': 2, 'WIP': 3, 'NG': 4 };

  var addRows = function(rows, posisi) {
    rows.forEach(function(row) {
      var item = String(row['Item_Code'] || '').trim();
      if (!item) return;
      var mi = itemMap[item];
      if (!mi || !mi.equivalent || !mi.t) return;

      // Avail: coba beberapa kolom kandidat
      var avail = _demNum(row['KG_Avail']);
      if (!avail) avail = _demNum(row['Qty_Avail']);
      if (!avail) avail = _demNum(row['Kg_Avail']);
      if (avail <= 0) return;

      var key = mi.equivalent + '|' + mi.t;
      if (!idx[key]) idx[key] = [];

      var tglVal = row['Tgl_Masuk'] || row['Tgl_In'] || row['Tgl_Buat'] || row['Tanggal'];
      var tglDate = _demToDate(tglVal);

      // UoM per posisi: Coil = KG, Sheet = LBR, WIP = PCS, NG = ikut type
      var uom = 'KG';
      if (posisi === 'Sheet') uom = 'LBR';
      else if (posisi === 'WIP') uom = 'PCS';
      else if (posisi === 'NG') uom = mi.uom || 'KG';

      idx[key].push({
        posisi:            posisi,
        posisi_order:      posisiOrder[posisi] || 99,
        item_code:         item,
        description:       String(row['Description'] || mi.description || '').trim(),
        batch_id:          String(row['Batch_ID'] || row['Batch_Id'] || '').trim(),
        owner:             String(row['Owner'] || 'FC').trim(),
        avail:             Math.round(avail),
        uom:               uom,
        tgl_masuk_ts:      tglDate ? tglDate.getTime() : 0,
        tgl_masuk_display: _demFormatDate(tglDate)
      });
    });
  };

  addRows(coilRows,  'Coil');
  addRows(sheetRows, 'Sheet');
  addRows(wipRows,   'WIP');
  addRows(ngRows,    'NG');

  // Sort per fingerprint: by posisi order → by tgl_masuk ASC (FIFO)
  Object.keys(idx).forEach(function(key) {
    idx[key].sort(function(a, b) {
      if (a.posisi_order !== b.posisi_order) return a.posisi_order - b.posisi_order;
      return a.tgl_masuk_ts - b.tgl_masuk_ts;
    });
  });

  return idx;
}

// =========================================================================
// 4. BUILD DEMAND ROW — FROM SO
// =========================================================================
function _demBuildFromSO(soRow, ctx) {
  var soNo = String(soRow['SO_No'] || '').trim();
  if (!soNo) return null;

  var itemCode = String(soRow['Item_Code'] || '').trim();
  if (!itemCode) return null;

  var statusRaw = String(soRow['STATUS'] || soRow['Status'] || '').toUpperCase().trim();
  if (statusRaw === 'CANCELLED') return null;
  if (statusRaw === 'DONE')      return null;
  if (statusRaw === 'CLOSED')    return null;

  // STATIC demand: pakai SO_Q dan SO_KG (bukan BL_Q/BL_KG)
  var qtyDemand = _demNum(soRow['SO_Q']);
  var kgDemand  = _demNum(soRow['SO_KG']);
  if (qtyDemand <= 0) return null;

  var mi = ctx.itemMap[itemCode] || {};
  var equivalent = mi.equivalent || '';
  var t = _demNum(soRow['T']) || mi.t || 0;

  var spkKey = soNo + '|' + itemCode;
  var relatedSpk = ctx.spkIndex[spkKey] || [];

  var spkListDisplay = _demBuildSpkListForDisplay(relatedSpk, ctx.delvBySpk);

  // Aggregate SPK (semua kecuali CANCELLED — filter type SHR-OUT/ALLOC-OUT)
  var spkQty = 0, spkKg = 0, spkCount = 0;
  var doneQty = 0, doneKg = 0;
  relatedSpk.forEach(function(r) {
    var typ = String(r['SPK_Type'] || '').toUpperCase();
    if (typ !== 'SHR-OUT' && typ !== 'ALLOC-OUT' && typ !== 'CTL-OUT' && typ !== 'SLT-OUT') return;
    spkQty += _demNum(r['Qty_Target']);
    spkKg  += _demNum(r['KG_Target']);
    spkCount++;
    var st = String(r['Status'] || '').toUpperCase();
    if (st === 'DONE') {
      doneQty += _demNum(r['Qty_Actual']);
      doneKg  += _demNum(r['KG_Actual']);
    }
  });

  // Delivery — dari DELV
  var delvInfo = ctx.delvSo[spkKey] || { qty: 0, kg: 0 };
  var delvQty = delvInfo.qty;
  var delvKg  = delvInfo.kg;

  // Net Req = Demand - SPK (kalau minus, capped ke 0; warning akan flag)
  var netReq = Math.max(0, qtyDemand - spkQty);

  // Status 5-tier
  var status5 = _demComputeStatus(qtyDemand, spkQty, doneQty, delvQty);

  // Warnings
  var warnings = _demComputeWarnings(qtyDemand, spkQty, doneQty, delvQty);

  // Info Stok key (untuk lookup modal)
  var stokInfoKey = (equivalent && t) ? (equivalent + '|' + t) : '';

  // BACKWARD COMPAT — field lama untuk frontend existing
  var stokFgQty      = ctx.fgIndex[spkKey] || 0;
  var stokSheetExact = ctx.sheetIndex.byItem[itemCode] || 0;
  var stokSheetStd   = (equivalent && t) ? (ctx.sheetIndex.byEquivT[equivalent + '|' + t] || 0) : 0;
  var qtySpkPending  = _demSumSpk(relatedSpk, function(r) {
    var t2 = String(r['SPK_Type'] || '').toUpperCase();
    var st = String(r['Status']   || '').toUpperCase();
    return (t2 === 'SHR-OUT' || t2 === 'ALLOC-OUT' || t2 === 'CTL-OUT' || t2 === 'SLT-OUT') && st !== 'DONE';
  });
  var actionLegacy = _demComputeAction({
    statusRaw:     statusRaw,
    netReq:        qtyDemand - qtySpkPending,
    fg:            stokFgQty,
    sheetExact:    stokSheetExact,
    sheetStd:      stokSheetStd,
    qtySpkPending: qtySpkPending
  });

  var tglNeeded = _demToDate(soRow['SCHEDULE_DATE']);
  var periode   = String(soRow['SO_Period'] || '').trim();
  if (!periode && tglNeeded) periode = _demFormatPeriod(tglNeeded);

  return {
    tipe:             'SO',
    ref_no:           soNo,
    cust:             String(soRow['Cust'] || '').trim(),
    periode:          periode,
    tgl_needed:       _demFormatDateShort(tglNeeded),
    tgl_needed_full:  _demFormatDate(tglNeeded),
    tgl_needed_ts:    tglNeeded ? tglNeeded.getTime() : 0,
    sched_urgent:     _demIsUrgent(tglNeeded),
    item_code:        itemCode,
    description:      String(soRow['Description'] || mi.description || '').trim(),
    spec:             String(soRow['Spec'] || mi.spec || '').trim(),
    t:                t,
    p:                _demNum(soRow['P']) || mi.p || 0,
    l:                _demNum(soRow['L']) || mi.l || 0,
    uom:              mi.uom || '',
    type:             mi.type || '',
    // NEW static demand
    qty_demand:       qtyDemand,
    kg_demand:        kgDemand,
    // NEW SPK aggregate
    spk_qty:          spkQty,
    spk_kg:           spkKg,
    spk_count:        spkCount,
    // NEW Done aggregate
    done_qty:         doneQty,
    done_kg:          doneKg,
    // NEW Delivery
    delv_qty:         delvQty,
    delv_kg:          delvKg,
    // Net Req
    net_req:          netReq,
    // NEW 5-tier status + warnings
    status5:          status5,
    warnings:         warnings,
    // Info Stok reference
    stok_info_key:    stokInfoKey,
    // SPK list (untuk modal)
    spk_list:         spkListDisplay,
    // Owner
    owner_used:       String(soRow['Owner_Used'] || 'FC').trim(),
    // Equivalent
    equivalent:       equivalent,
    material_key:     stokInfoKey,
    // === BACKWARD COMPAT (untuk frontend existing yang belum di-rewrite) ===
    status:           statusRaw,
    qty_spk_pending:  qtySpkPending,
    stok_fg_qty:      stokFgQty,
    stok_sheet_exact: stokSheetExact,
    stok_sheet_std:   stokSheetStd,
    action:           actionLegacy
  };
}

// =========================================================================
// 5. BUILD DEMAND ROW — FROM STP_REQ (header baru)
// =========================================================================
function _demBuildFromSTP(stpRow, ctx) {
  var stpNo = String(stpRow['STP_No'] || '').trim();
  if (!stpNo) return null;

  var itemCode = String(stpRow['Item_Code'] || '').trim();
  if (!itemCode) return null;

  var statusRaw = String(stpRow['STATUS'] || stpRow['Status'] || '').toUpperCase().trim();
  if (statusRaw === 'CANCELLED') return null;
  if (statusRaw === 'DONE')      return null;
  if (statusRaw === 'CLOSED')    return null;

  // STATIC demand: pakai STP_Q dan STP_KG (header baru)
  var qtyDemand = _demNum(stpRow['STP_Q']);
  var kgDemand  = _demNum(stpRow['STP_KG']);
  if (qtyDemand <= 0) return null;

  var mi = ctx.itemMap[itemCode] || {};
  var equivalent = mi.equivalent || '';
  var t = _demNum(stpRow['T']) || mi.t || 0;

  var spkKey = stpNo + '|' + itemCode;
  var relatedSpk = ctx.spkIndex[spkKey] || [];
  var spkListDisplay = _demBuildSpkListForDisplay(relatedSpk, ctx.delvBySpk);

  var spkQty = 0, spkKg = 0, spkCount = 0;
  var doneQty = 0, doneKg = 0;
  relatedSpk.forEach(function(r) {
    var typ = String(r['SPK_Type'] || '').toUpperCase();
    if (typ !== 'SHR-OUT' && typ !== 'ALLOC-OUT' && typ !== 'CTL-OUT' && typ !== 'SLT-OUT') return;
    spkQty += _demNum(r['Qty_Target']);
    spkKg  += _demNum(r['KG_Target']);
    spkCount++;
    var st = String(r['Status'] || '').toUpperCase();
    if (st === 'DONE') {
      doneQty += _demNum(r['Qty_Actual']);
      doneKg  += _demNum(r['KG_Actual']);
    }
  });

  // Delivery dari DELV_STP
  var delvInfo = ctx.delvStp[spkKey] || { qty: 0, kg: 0 };
  var delvQty = delvInfo.qty;
  var delvKg  = delvInfo.kg;

  var netReq = Math.max(0, qtyDemand - spkQty);
  var status5 = _demComputeStatus(qtyDemand, spkQty, doneQty, delvQty);
  var warnings = _demComputeWarnings(qtyDemand, spkQty, doneQty, delvQty);

  var stokInfoKey = (equivalent && t) ? (equivalent + '|' + t) : '';

  // BACKWARD COMPAT
  var stokFgQty      = ctx.fgIndex[spkKey] || 0;
  var stokSheetExact = ctx.sheetIndex.byItem[itemCode] || 0;
  var stokSheetStd   = (equivalent && t) ? (ctx.sheetIndex.byEquivT[equivalent + '|' + t] || 0) : 0;
  var qtySpkPending  = _demSumSpk(relatedSpk, function(r) {
    var t2 = String(r['SPK_Type'] || '').toUpperCase();
    var st = String(r['Status']   || '').toUpperCase();
    return (t2 === 'SHR-OUT' || t2 === 'ALLOC-OUT' || t2 === 'CTL-OUT' || t2 === 'SLT-OUT') && st !== 'DONE';
  });
  var actionLegacy = _demComputeAction({
    statusRaw:     statusRaw,
    netReq:        qtyDemand - qtySpkPending,
    fg:            stokFgQty,
    sheetExact:    stokSheetExact,
    sheetStd:      stokSheetStd,
    qtySpkPending: qtySpkPending
  });

  var tglNeeded = _demToDate(stpRow['Schedule_Date']);
  var periode   = String(stpRow['STP_Period'] || '').trim();
  if (!periode && tglNeeded) periode = _demFormatPeriod(tglNeeded);

  return {
    tipe:             'STP',
    ref_no:           stpNo,
    cust:             String(stpRow['Cust'] || '').trim(),
    periode:          periode,
    tgl_needed:       _demFormatDateShort(tglNeeded),
    tgl_needed_full:  _demFormatDate(tglNeeded),
    tgl_needed_ts:    tglNeeded ? tglNeeded.getTime() : 0,
    sched_urgent:     _demIsUrgent(tglNeeded),
    item_code:        itemCode,
    description:      String(stpRow['Description'] || mi.description || '').trim(),
    spec:             String(stpRow['Spec'] || mi.spec || '').trim(),
    t:                t,
    p:                _demNum(stpRow['P']) || mi.p || 0,
    l:                _demNum(stpRow['L']) || mi.l || 0,
    uom:              mi.uom || '',
    type:             mi.type || '',
    qty_demand:       qtyDemand,
    kg_demand:        kgDemand,
    spk_qty:          spkQty,
    spk_kg:           spkKg,
    spk_count:        spkCount,
    done_qty:         doneQty,
    done_kg:          doneKg,
    delv_qty:         delvQty,
    delv_kg:          delvKg,
    net_req:          netReq,
    status5:          status5,
    warnings:         warnings,
    stok_info_key:    stokInfoKey,
    spk_list:         spkListDisplay,
    owner_used:       String(stpRow['Owner_Used'] || 'FC').trim(),
    equivalent:       equivalent,
    material_key:     stokInfoKey,
    // === BACKWARD COMPAT ===
    status:           statusRaw,
    qty_spk_pending:  qtySpkPending,
    stok_fg_qty:      stokFgQty,
    stok_sheet_exact: stokSheetExact,
    stok_sheet_std:   stokSheetStd,
    action:           actionLegacy
  };
}

// =========================================================================
// 6. SPK DISPLAY LIST — untuk modal detail per row
// =========================================================================
function _demBuildSpkListForDisplay(spkArr, delvBySpk) {
  if (!spkArr || !spkArr.length) return [];
  return spkArr.map(function(r) {
    var spkNo = String(r['SPK_No'] || '').trim();
    var qtyActual = _demNum(r['Qty_Actual']);
    var delvInfo = delvBySpk[spkNo] || { qty: 0, kg: 0 };
    var sisaStok = Math.max(0, qtyActual - delvInfo.qty);
    return {
      spk_no:           spkNo,
      spk_type:         String(r['SPK_Type'] || '').trim(),
      parent_spk:       String(r['Parent_SPK'] || '').trim(),
      status:           String(r['Status'] || '').trim(),
      mc_no:            String(r['MC_No'] || '').trim(),
      qty_target:       _demNum(r['Qty_Target']),
      kg_target:        _demNum(r['KG_Target']),
      qty_actual:       qtyActual,
      kg_actual:        _demNum(r['KG_Actual']),
      delivery_qty:     delvInfo.qty,
      delivery_kg:      delvInfo.kg,
      sisa_stok:        sisaStok,
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
    return String(a.spk_no).localeCompare(String(b.spk_no));
  });
}

// =========================================================================
// 7. STATUS 5-TIER
// =========================================================================
function _demComputeStatus(demand, spk, done, delv) {
  if (spk === 0)              return 'UNPLANNED';
  if (spk < demand)           return 'PARTIAL';
  if (done < demand)          return 'COVERED';
  if (delv < demand)          return 'PRODUCED';
  return 'FULFILLED';
}

// =========================================================================
// 8. WARNINGS
// =========================================================================
function _demComputeWarnings(demand, spk, done, delv) {
  return {
    leak:          delv > done,     // Kirim melebihi produksi
    over_produce:  done > spk,      // Produksi melebihi target SPK
    over_delivery: delv > demand    // Kirim melebihi demand
  };
}

// =========================================================================
// 9. LEGACY — Helper sum & action (backward compat)
// =========================================================================
function _demSumSpk(spkArr, predicate) {
  if (!spkArr || !spkArr.length) return 0;
  var sum = 0;
  for (var i = 0; i < spkArr.length; i++) {
    if (predicate(spkArr[i])) sum += _demNum(spkArr[i]['Qty_Target']);
  }
  return sum;
}

function _demComputeAction(ctx) {
  if (ctx.statusRaw === 'OVER') return 'OVER';
  if (ctx.netReq <= 0)          return 'COVER';
  if (ctx.fg > 0 && ctx.fg >= ctx.netReq)                 return 'ALOKASI_FG';
  if (ctx.sheetExact > 0 && ctx.sheetExact >= ctx.netReq) return 'ALOKASI_SHEET';
  if (ctx.qtySpkPending > 0)                              return 'PARTIAL';
  if (ctx.sheetStd > 0)                                    return 'PERLU_SHR';
  return 'PERLU_CTL';
}

// =========================================================================
// 10. CTL GROUPING (backward compat, still used oleh existing frontend tab CTL)
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
    var kgPortion = 0;
    if (d.qty_demand > 0) kgPortion = (d.kg_demand || 0) * (d.net_req / d.qty_demand);
    g.total_kg  += kgPortion;
    g.total_qty += d.net_req;
    if (d.tipe === 'SO')  g.count_so  += 1;
    if (d.tipe === 'STP') g.count_stp += 1;
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
// 11. SUMMARY
// =========================================================================
function _demBuildSummary(demands, ctlGroups) {
  var totalOpenCount = 0, totalOpenKg = 0;
  var needSpkCount   = 0, needSpkKg   = 0;
  var ctlKg          = 0;

  // New 5-tier counts
  var cntUnplan = 0, cntPartial = 0, cntCovered = 0, cntProduced = 0, cntFulfilled = 0;
  var cntLeak = 0, cntOverProd = 0, cntOverDelv = 0;

  demands.forEach(function(d) {
    totalOpenCount += 1;
    totalOpenKg    += (d.kg_demand || 0);

    // Backward compat need_spk
    if (d.action === 'PERLU_CTL' || d.action === 'PERLU_SHR' || d.action === 'PARTIAL') {
      needSpkCount += 1;
      if (d.qty_demand > 0) needSpkKg += (d.kg_demand || 0) * (d.net_req / d.qty_demand);
    }

    // New status 5-tier counts
    switch (d.status5) {
      case 'UNPLANNED': cntUnplan++;    break;
      case 'PARTIAL':   cntPartial++;   break;
      case 'COVERED':   cntCovered++;   break;
      case 'PRODUCED':  cntProduced++;  break;
      case 'FULFILLED': cntFulfilled++; break;
    }

    if (d.warnings) {
      if (d.warnings.leak)          cntLeak++;
      if (d.warnings.over_produce)  cntOverProd++;
      if (d.warnings.over_delivery) cntOverDelv++;
    }
  });

  ctlGroups.forEach(function(g) { ctlKg += g.total_kg; });

  return {
    // Backward compat
    total_open_count: totalOpenCount,
    total_open_kg:    Math.round(totalOpenKg),
    need_spk_count:   needSpkCount,
    need_spk_kg:      Math.round(needSpkKg),
    ctl_group_count:  ctlGroups.length,
    ctl_group_kg:     Math.round(ctlKg),
    // New status counts
    count_unplanned:  cntUnplan,
    count_partial:    cntPartial,
    count_covered:    cntCovered,
    count_produced:   cntProduced,
    count_fulfilled:  cntFulfilled,
    // Warning counts
    count_warn_leak:          cntLeak,
    count_warn_over_produce:  cntOverProd,
    count_warn_over_delivery: cntOverDelv
  };
}

// =========================================================================
// 12. UTILITIES
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

// NEW: format singkat "02 Jul" untuk Sched column
function _demFormatDateShort(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM');
}

function _demFormatDateTime(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
}

function _demFormatPeriod(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM yyyy');
}

// NEW: cek urgent (≤ 3 hari dari sekarang)
function _demIsUrgent(d) {
  if (!d) return false;
  var now = new Date();
  var diffMs = d.getTime() - now.getTime();
  var diffDays = diffMs / (1000 * 60 * 60 * 24);
  return diffDays <= 3;
}

// =========================================================================
// 13. TEST HELPER
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
  Logger.log('=== STOK INFO KEYS: ' + Object.keys(result.stok_info).length + ' ===');
  var sampleKey = Object.keys(result.stok_info)[0];
  if (sampleKey) {
    Logger.log('Sample stok_info["' + sampleKey + '"]:');
    Logger.log(JSON.stringify(result.stok_info[sampleKey], null, 2));
  }
  return result;
}