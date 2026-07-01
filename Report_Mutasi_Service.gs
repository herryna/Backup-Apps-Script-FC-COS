// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - REPORT SERVICE (T7)
//
// Modul: Laporan Mutasi Bulanan untuk Finance
// Format: Saldo Awal + IN - OUT ± ADJ = Saldo Akhir
//
// Public API:
//   - getMutasiReport(period, filter)
//   - getReportInitData()
//
// Sumber data: GR, Trace_Log, SPK, DELV, Rekap_ICT, Opname_Bulanan, M_ITEM
// =========================================================================

function _rptPrevPeriod(period) {
  var p = String(period).split('-');
  var y = parseInt(p[0]); var m = parseInt(p[1]);
  m -= 1; if (m < 1) { m = 12; y -= 1; }
  return y + '-' + String(m).padStart(2, '0');
}

function _rptPeriodRange(period) {
  var p = String(period).split('-');
  var y = parseInt(p[0]); var m = parseInt(p[1]);
  var start = new Date(y, m - 1, 1, 0, 0, 0, 0);
  var end   = new Date(y, m, 0, 23, 59, 59, 999); // last day of month
  return { start: start, end: end };
}

function _rptInRange(dt, range) {
  if (!(dt instanceof Date)) return false;
  return dt >= range.start && dt <= range.end;
}

function _rptHdrIdx(arr) {
  var m = {};
  for (var i = 0; i < arr.length; i++) m[String(arr[i]).trim()] = i;
  return m;
}

function _rptReadSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return { rows: [], hdr: [], idx: {} };
  var d = sh.getDataRange().getValues();
  if (d.length < 1) return { rows: [], hdr: [], idx: {} };
  var hdr = d[0].map(function(h){ return String(h).trim(); });
  return { rows: d.length > 1 ? d.slice(1) : [], hdr: hdr, idx: _rptHdrIdx(hdr) };
}

// =========================================================================
// HELPER: List periode yang available untuk dropdown
// =========================================================================
function getReportInitData() {
  try {
    // Generate 12 months back + current + 2 ahead
    var arr = [];
    var now = new Date();
    for (var i = -12; i <= 2; i++) {
      var d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      var p = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      arr.push(p);
    }
    // Auto-select bulan lalu (yang biasanya udah siap laporannya)
    var prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    var suggested = prev.getFullYear() + '-' + String(prev.getMonth() + 1).padStart(2, '0');

    // Get list opname periods (untuk indicator saldo awal/akhir availability)
    var op = _rptReadSheet('Opname_Bulanan');
    var lockedPeriods = {};
    if (op.idx['Period'] !== undefined && op.idx['Status'] !== undefined) {
      for (var i = 0; i < op.rows.length; i++) {
        var pp = String(op.rows[i][op.idx['Period']]).trim();
        var st = String(op.rows[i][op.idx['Status']] || '').toUpperCase();
        if (pp && (st === 'FINAL' || st === 'LOCKED')) lockedPeriods[pp] = true;
      }
    }

    return {
      success: true,
      periods: arr,
      suggested: suggested,
      current: now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0'),
      locked_periods: Object.keys(lockedPeriods)
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// =========================================================================
// MAIN: getMutasiReport(period, filter)
// =========================================================================
function getMutasiReport(period, filter) {
  try {
    if (!period || !/^\d{4}-\d{2}$/.test(String(period).trim())) {
      throw new Error('Format period invalid. Pakai YYYY-MM');
    }
    period = String(period).trim();
    filter = filter || {};
    var fOwner = filter.owner ? String(filter.owner).toUpperCase() : '';
    var fLoc   = filter.loc   ? String(filter.loc) : '';

    var range = _rptPeriodRange(period);
    var prevPeriod = _rptPrevPeriod(period);

    // === LOAD ALL DATA ===
    var gr   = _rptReadSheet('GR');
    var tl   = _rptReadSheet('Trace_Log');
    var spk  = _rptReadSheet('SPK');
    var delv = _rptReadSheet('DELV');
    var ict  = _rptReadSheet('Rekap_ICT');
    var op   = _rptReadSheet('Opname_Bulanan');
    var item = _rptReadSheet('M_ITEM');

    // Master item lookup
    var itemMaster = {};
    if (item.idx['Item_Code'] !== undefined) {
      for (var im = 0; im < item.rows.length; im++) {
        var ic = String(item.rows[im][item.idx['Item_Code']] || '').trim();
        if (!ic) continue;
        itemMaster[ic] = {
          description: String(item.rows[im][item.idx['Description']] || ''),
          spec       : String(item.rows[im][item.idx['Spec']] || ''),
          uom        : item.idx['UoM'] !== undefined ? String(item.rows[im][item.idx['UoM']] || '') : ''
        };
      }
    }

    // === SALDO AWAL (dari Opname_Bulanan period sebelumnya) ===
    var saldoAwalByItem  = {}; // { item_code: { qty, kg } }
    var saldoAwalByBatch = {}; // { batch_id: { item_code, loc, owner, qty, kg } }
    var saldoAwalSource  = 'EMPTY'; // 'OPNAME_LOCKED' | 'COMPUTED' | 'EMPTY'

    if (op.idx['Period'] !== undefined) {
      var lockedRowsPrev = [];
      for (var i = 0; i < op.rows.length; i++) {
        var r = op.rows[i];
        if (String(r[op.idx['Period']]).trim() !== prevPeriod) continue;
        var st = String(r[op.idx['Status']] || '').toUpperCase();
        if (st !== 'FINAL' && st !== 'LOCKED') continue;
        lockedRowsPrev.push(r);
      }
      if (lockedRowsPrev.length > 0) {
        saldoAwalSource = 'OPNAME_LOCKED';
        for (var j = 0; j < lockedRowsPrev.length; j++) {
          var rr = lockedRowsPrev[j];
          var bid    = String(rr[op.idx['Batch_ID']] || '').trim();
          var itemCd = String(rr[op.idx['Item_Code']] || '').trim();
          var loc    = String(rr[op.idx['Loc']] || '');
          var owner  = String(rr[op.idx['Owner']] || '');
          // Pakai Fisik kalau ada, fallback System
          var qFis = parseFloat(rr[op.idx['Qty_Fisik']]) || 0;
          var kFis = parseFloat(rr[op.idx['KG_Fisik']])  || 0;
          var qSys = parseFloat(rr[op.idx['Qty_System']]) || 0;
          var kSys = parseFloat(rr[op.idx['KG_System']])  || 0;
          var q = qFis > 0 ? qFis : qSys;
          var k = kFis > 0 ? kFis : kSys;

          // Apply filter
          if (fOwner && owner.toUpperCase() !== fOwner) continue;
          if (fLoc && loc !== fLoc) continue;

          if (!saldoAwalByItem[itemCd]) saldoAwalByItem[itemCd] = { qty: 0, kg: 0 };
          saldoAwalByItem[itemCd].qty += q;
          saldoAwalByItem[itemCd].kg  += k;

          saldoAwalByBatch[bid] = { item_code: itemCd, loc: loc, owner: owner, qty: q, kg: k };
        }
      }
    }

    // === IN — RECEIPT (GR) ===
    var inReceiptByItem  = {};
    var inReceiptByBatch = {};
    if (gr.idx['Tgl_Masuk'] !== undefined) {
      var iGrBat = gr.idx['Batch_ID'];
      var iGrItm = gr.idx['Item_Code'];
      var iGrTgl = gr.idx['Tgl_Masuk'];
      var iGrOwn = gr.idx['Owner'];
      var iGrLoc = gr.idx['Target_Loc'];
      var iGrQ   = gr.idx['QTY_In'] !== undefined ? gr.idx['QTY_In'] : gr.idx['Qty_In'];
      var iGrK   = gr.idx['KG_In'];

      for (var k = 0; k < gr.rows.length; k++) {
        var rg = gr.rows[k];
        var tgl = rg[iGrTgl];
        if (!_rptInRange(tgl, range)) continue;
        var own = iGrOwn !== undefined ? String(rg[iGrOwn] || '').toUpperCase() : '';
        if (fOwner && own !== fOwner) continue;
        var tloc = iGrLoc !== undefined ? String(rg[iGrLoc] || '') : '';
        // Map Target_Loc Stok_Coil/Stok_Sheet → Coil/Sheet
        var locMapped = tloc === 'Stok_Coil' ? 'Coil' : (tloc === 'Stok_Sheet' ? 'Sheet' : tloc);
        if (fLoc && locMapped !== fLoc) continue;

        var bid    = String(rg[iGrBat] || '');
        var itemCd = String(rg[iGrItm] || '');
        var q = parseFloat(rg[iGrQ]) || 0;
        var kg = parseFloat(rg[iGrK]) || 0;

        if (!inReceiptByItem[itemCd]) inReceiptByItem[itemCd] = { qty: 0, kg: 0 };
        inReceiptByItem[itemCd].qty += q;
        inReceiptByItem[itemCd].kg  += kg;

        if (!inReceiptByBatch[bid]) inReceiptByBatch[bid] = { item_code: itemCd, loc: locMapped, owner: own, qty: 0, kg: 0 };
        inReceiptByBatch[bid].qty += q;
        inReceiptByBatch[bid].kg  += kg;
      }
    }

    // === IN — PRODUCTION (Trace_Log Type=SHEET/WIP/FGC/FGS) ===
    var inProdByItem  = {};
    var inProdByBatch = {};
    if (tl.idx['Tgl_Buat'] !== undefined) {
      var iTlBat = tl.idx['Batch_ID'];
      var iTlItm = tl.idx['Item_Code'];
      var iTlTgl = tl.idx['Tgl_Buat'];
      var iTlTyp = tl.idx['Type'];
      var iTlQ   = tl.idx['Qty'];
      var iTlK   = tl.idx['KG'];
      var iTlOwn = tl.idx['Owner'];

      for (var t = 0; t < tl.rows.length; t++) {
        var rt = tl.rows[t];
        var typ = String(rt[iTlTyp] || '').toUpperCase();
        if (typ === 'NG') continue;
        if (typ !== 'SHEET' && typ !== 'WIP' && typ !== 'FGC' && typ !== 'FGS' && typ !== 'FCL') continue;

        var tgl = rt[iTlTgl];
        if (!_rptInRange(tgl, range)) continue;
        var own = iTlOwn !== undefined ? String(rt[iTlOwn] || '').toUpperCase() : '';
        if (fOwner && own !== fOwner) continue;
        // Map type → loc
        var locMap = typ === 'SHEET' ? 'Sheet' : (typ === 'WIP' ? 'WIP' : 'FG');
        if (fLoc && locMap !== fLoc) continue;

        var bid    = String(rt[iTlBat] || '');
        var itemCd = String(rt[iTlItm] || '');
        var q = parseFloat(rt[iTlQ]) || 0;
        var kg = parseFloat(rt[iTlK]) || 0;

        if (!inProdByItem[itemCd]) inProdByItem[itemCd] = { qty: 0, kg: 0 };
        inProdByItem[itemCd].qty += q;
        inProdByItem[itemCd].kg  += kg;

        if (!inProdByBatch[bid]) inProdByBatch[bid] = { item_code: itemCd, loc: locMap, owner: own, qty: 0, kg: 0 };
        inProdByBatch[bid].qty += q;
        inProdByBatch[bid].kg  += kg;
      }
    }

    // === OUT — CONSUMPTION (SPK Header DONE) ===
    // SPK_Type ends with HEADER (CTL-HEADER, SHR-HEADER, SLT-HEADER) = batch jadi input
    var outConsByItem  = {};
    var outConsByBatch = {};
    if (spk.idx['SPK_No'] !== undefined) {
      var iSpkType = spk.idx['SPK_Type'];
      var iSpkBat  = spk.idx['Batch_ID'];
      var iSpkItm  = spk.idx['Item_Code'];
      var iSpkSt   = spk.idx['Status'];
      var iSpkSel  = spk.idx['Selesai_DT'];
      var iSpkQAct = spk.idx['Qty_Actual'];
      var iSpkKAct = spk.idx['KG_Actual'];
      var iSpkOwn  = spk.idx['Owner'];
      var iSpkSrc  = spk.idx['Source_Loc'];

      for (var s = 0; s < spk.rows.length; s++) {
        var rs = spk.rows[s];
        var sTyp = String(rs[iSpkType] || '');
        if (sTyp.indexOf('HEADER') === -1) continue;
        var sStatus = String(rs[iSpkSt] || '').toUpperCase();
        if (sStatus !== 'DONE') continue;
        var sel = rs[iSpkSel];
        if (!_rptInRange(sel, range)) continue;

        var own = iSpkOwn !== undefined ? String(rs[iSpkOwn] || '').toUpperCase() : '';
        if (fOwner && own !== fOwner) continue;
        var srcLoc = iSpkSrc !== undefined ? String(rs[iSpkSrc] || '') : '';
        // Source_Loc bisa 'Stok Coil' / 'Stok_Coil' / 'Stok_Sheet'
        var srcLocMapped = '';
        var slL = srcLoc.toLowerCase().replace(/\s/g, '_');
        if (slL.indexOf('coil') > -1) srcLocMapped = 'Coil';
        else if (slL.indexOf('sheet') > -1) srcLocMapped = 'Sheet';
        else if (slL.indexOf('wip') > -1) srcLocMapped = 'WIP';
        if (fLoc && srcLocMapped !== fLoc) continue;

        var bid    = String(rs[iSpkBat] || '');
        var itemCd = String(rs[iSpkItm] || '');
        var q = parseFloat(rs[iSpkQAct]) || 0;
        var kg = parseFloat(rs[iSpkKAct]) || 0;

        if (!outConsByItem[itemCd]) outConsByItem[itemCd] = { qty: 0, kg: 0 };
        outConsByItem[itemCd].qty += q;
        outConsByItem[itemCd].kg  += kg;

        if (!outConsByBatch[bid]) outConsByBatch[bid] = { item_code: itemCd, loc: srcLocMapped, owner: own, qty: 0, kg: 0 };
        outConsByBatch[bid].qty += q;
        outConsByBatch[bid].kg  += kg;
      }
    }

    // === OUT — DELIVERY ===
    var outDelvByItem  = {};
    var outDelvByBatch = {};
    if (delv.idx['Tanggal'] !== undefined) {
      var iDvTgl = delv.idx['Tanggal'];
      var iDvBat = delv.idx['Batch_ID'];
      var iDvItm = delv.idx['Item_Code'];
      var iDvOwn = delv.idx['Owner'];
      var iDvQ   = delv.idx['Delv_Q'];
      var iDvK   = delv.idx['Delv_KG'];

      for (var d = 0; d < delv.rows.length; d++) {
        var rd = delv.rows[d];
        var tgl = rd[iDvTgl];
        if (!_rptInRange(tgl, range)) continue;
        var own = iDvOwn !== undefined ? String(rd[iDvOwn] || '').toUpperCase() : '';
        if (fOwner && own !== fOwner) continue;
        if (fLoc && fLoc !== 'FG') continue; // Delivery selalu dari FG

        var bid    = String(rd[iDvBat] || '');
        var itemCd = String(rd[iDvItm] || '');
        var q = parseFloat(rd[iDvQ]) || 0;
        var kg = parseFloat(rd[iDvK]) || 0;

        if (!outDelvByItem[itemCd]) outDelvByItem[itemCd] = { qty: 0, kg: 0 };
        outDelvByItem[itemCd].qty += q;
        outDelvByItem[itemCd].kg  += kg;

        if (!outDelvByBatch[bid]) outDelvByBatch[bid] = { item_code: itemCd, loc: 'FG', owner: own, qty: 0, kg: 0 };
        outDelvByBatch[bid].qty += q;
        outDelvByBatch[bid].kg  += kg;
      }
    }

    // === OUT — NG (Trace_Log Type=NG) ===
    var outNgByItem  = {};
    var outNgByBatch = {};
    if (tl.idx['Tgl_Buat'] !== undefined) {
      var iTlBat = tl.idx['Batch_ID'];
      var iTlItm = tl.idx['Item_Code'];
      var iTlTgl = tl.idx['Tgl_Buat'];
      var iTlTyp = tl.idx['Type'];
      var iTlQ   = tl.idx['Qty'];
      var iTlK   = tl.idx['KG'];
      var iTlOwn = tl.idx['Owner'];
      var iTlSrc = tl.idx['Source_Batch'];

      for (var n = 0; n < tl.rows.length; n++) {
        var rn = tl.rows[n];
        var typ = String(rn[iTlTyp] || '').toUpperCase();
        if (typ !== 'NG') continue;
        var tgl = rn[iTlTgl];
        if (!_rptInRange(tgl, range)) continue;
        var own = iTlOwn !== undefined ? String(rn[iTlOwn] || '').toUpperCase() : '';
        if (fOwner && own !== fOwner) continue;
        if (fLoc && fLoc !== 'NG') continue;

        var srcBat = String(rn[iTlSrc] || '');
        var itemCd = String(rn[iTlItm] || '');
        var q = parseFloat(rn[iTlQ]) || 0;
        var kg = parseFloat(rn[iTlK]) || 0;

        if (!outNgByItem[itemCd]) outNgByItem[itemCd] = { qty: 0, kg: 0 };
        outNgByItem[itemCd].qty += q;
        outNgByItem[itemCd].kg  += kg;

        // For per-batch, NG charged ke source batch (parent yang kasih NG)
        if (srcBat) {
          if (!outNgByBatch[srcBat]) outNgByBatch[srcBat] = { item_code: itemCd, loc: 'NG', owner: own, qty: 0, kg: 0 };
          outNgByBatch[srcBat].qty += q;
          outNgByBatch[srcBat].kg  += kg;
        }
      }
    }

    // === TF — Cross Owner (Rekap_ICT) ===
    // Net per perspective: kalau filter Owner=FC, ICT yang masuk FC = +, keluar FC = -
    // Kalau filter All, biarkan 0 (cancel out)
    var tfByItem  = {}; // { item_code: { qty_in, kg_in, qty_out, kg_out, net_qty, net_kg } }
    if (ict.idx['Tgl_Transfer'] !== undefined) {
      var iIcTgl = ict.idx['Tgl_Transfer'];
      var iIcItm = ict.idx['Item_Code'];
      var iIcFr  = ict.idx['Dari_Owner'];
      var iIcTo  = ict.idx['Ke_Owner'];
      var iIcQ   = ict.idx['Qty_Sht'];
      var iIcK   = ict.idx['Qty_KG'];

      for (var c = 0; c < ict.rows.length; c++) {
        var rc = ict.rows[c];
        var tgl = rc[iIcTgl];
        if (!_rptInRange(tgl, range)) continue;

        var itemCd = String(rc[iIcItm] || '');
        var fr = String(rc[iIcFr] || '').toUpperCase();
        var to = String(rc[iIcTo] || '').toUpperCase();
        var q = parseFloat(rc[iIcQ]) || 0;
        var kg = parseFloat(rc[iIcK]) || 0;

        if (!tfByItem[itemCd]) tfByItem[itemCd] = { qty_in: 0, kg_in: 0, qty_out: 0, kg_out: 0 };

        if (fOwner) {
          // Per-entity view: in/out berdasarkan target Owner
          if (to === fOwner) { tfByItem[itemCd].qty_in += q; tfByItem[itemCd].kg_in += kg; }
          if (fr === fOwner) { tfByItem[itemCd].qty_out += q; tfByItem[itemCd].kg_out += kg; }
        } else {
          // Grand Total: mass conservation — setiap transaksi ICT punya 1 sender (out) dan 1 receiver (in)
          // ICT_In total = ICT_Out total = total volume transferred (untuk reconciliation check)
          tfByItem[itemCd].qty_in  += q;
          tfByItem[itemCd].kg_in   += kg;
          tfByItem[itemCd].qty_out += q;
          tfByItem[itemCd].kg_out  += kg;
        }
      }
    }

    // === ADJ — Adjustment (Opname_Bulanan current period LOCKED) ===
    var adjByItem  = {};
    var adjByBatch = {};
    if (op.idx['Period'] !== undefined) {
      for (var a = 0; a < op.rows.length; a++) {
        var ra = op.rows[a];
        if (String(ra[op.idx['Period']]).trim() !== period) continue;
        var st = String(ra[op.idx['Status']] || '').toUpperCase();
        if (st !== 'FINAL' && st !== 'LOCKED') continue;

        var qFis = parseFloat(ra[op.idx['Qty_Fisik']]) || 0;
        var kFis = parseFloat(ra[op.idx['KG_Fisik']])  || 0;
        // Kalau fisik 0, anggap belum diisi (gak ada adjustment)
        if (qFis === 0 && kFis === 0) continue;

        var qSel = parseFloat(ra[op.idx['Qty_Selisih']]) || 0;
        var kSel = parseFloat(ra[op.idx['KG_Selisih']])  || 0;
        var bid    = String(ra[op.idx['Batch_ID']] || '');
        var itemCd = String(ra[op.idx['Item_Code']] || '');
        var own    = String(ra[op.idx['Owner']] || '').toUpperCase();
        var loc    = String(ra[op.idx['Loc']] || '');

        if (fOwner && own !== fOwner) continue;
        if (fLoc && loc !== fLoc) continue;

        if (!adjByItem[itemCd]) adjByItem[itemCd] = { qty: 0, kg: 0 };
        adjByItem[itemCd].qty += qSel;
        adjByItem[itemCd].kg  += kSel;

        adjByBatch[bid] = { item_code: itemCd, loc: loc, owner: own, qty: qSel, kg: kSel };
      }
    }

    // === SALDO AKHIR (dari Opname_Bulanan current period LOCKED) ===
    var saldoAkhirByItem  = {};
    var saldoAkhirByBatch = {};
    var saldoAkhirSource  = 'COMPUTED'; // default fallback

    if (op.idx['Period'] !== undefined) {
      var lockedRowsCurr = [];
      for (var x = 0; x < op.rows.length; x++) {
        if (String(op.rows[x][op.idx['Period']]).trim() !== period) continue;
        var stt = String(op.rows[x][op.idx['Status']] || '').toUpperCase();
        if (stt !== 'FINAL' && stt !== 'LOCKED') continue;
        lockedRowsCurr.push(op.rows[x]);
      }
      if (lockedRowsCurr.length > 0) {
        saldoAkhirSource = 'OPNAME_LOCKED';
        for (var y = 0; y < lockedRowsCurr.length; y++) {
          var rcurr = lockedRowsCurr[y];
          var bid    = String(rcurr[op.idx['Batch_ID']] || '');
          var itemCd = String(rcurr[op.idx['Item_Code']] || '');
          var loc    = String(rcurr[op.idx['Loc']] || '');
          var own    = String(rcurr[op.idx['Owner']] || '');
          var qFis2 = parseFloat(rcurr[op.idx['Qty_Fisik']]) || 0;
          var kFis2 = parseFloat(rcurr[op.idx['KG_Fisik']])  || 0;
          var qSys2 = parseFloat(rcurr[op.idx['Qty_System']]) || 0;
          var kSys2 = parseFloat(rcurr[op.idx['KG_System']])  || 0;
          var q = qFis2 > 0 ? qFis2 : qSys2;
          var k = kFis2 > 0 ? kFis2 : kSys2;

          if (fOwner && own.toUpperCase() !== fOwner) continue;
          if (fLoc && loc !== fLoc) continue;

          if (!saldoAkhirByItem[itemCd]) saldoAkhirByItem[itemCd] = { qty: 0, kg: 0 };
          saldoAkhirByItem[itemCd].qty += q;
          saldoAkhirByItem[itemCd].kg  += k;

          saldoAkhirByBatch[bid] = { item_code: itemCd, loc: loc, owner: own, qty: q, kg: k };
        }
      }
    }

    // === ASSEMBLE BY-ITEM REPORT ===
    var allItems = {};
    function collectItems(obj) { for (var k in obj) allItems[k] = true; }
    collectItems(saldoAwalByItem);
    collectItems(inReceiptByItem);
    collectItems(inProdByItem);
    collectItems(outConsByItem);
    collectItems(outDelvByItem);
    collectItems(outNgByItem);
    collectItems(tfByItem);
    collectItems(adjByItem);
    collectItems(saldoAkhirByItem);

    var byItem = [];
    for (var itc in allItems) {
      var awal = saldoAwalByItem[itc] || { qty: 0, kg: 0 };
      var rcv  = inReceiptByItem[itc] || { qty: 0, kg: 0 };
      var prd  = inProdByItem[itc]    || { qty: 0, kg: 0 };
      var con  = outConsByItem[itc]   || { qty: 0, kg: 0 };
      var dlv  = outDelvByItem[itc]   || { qty: 0, kg: 0 };
      var ng   = outNgByItem[itc]     || { qty: 0, kg: 0 };
      var tf   = tfByItem[itc]        || { qty_in: 0, kg_in: 0, qty_out: 0, kg_out: 0 };
      var aj   = adjByItem[itc]       || { qty: 0, kg: 0 };
      var akh  = saldoAkhirByItem[itc];

      var compQ = awal.qty + rcv.qty + prd.qty - con.qty - dlv.qty - ng.qty + tf.qty_in - tf.qty_out + aj.qty;
      var compK = awal.kg  + rcv.kg  + prd.kg  - con.kg  - dlv.kg  - ng.kg  + tf.kg_in  - tf.kg_out  + aj.kg;
      var actQ  = akh ? akh.qty : compQ;
      var actK  = akh ? akh.kg  : compK;
      var varQ  = actQ - compQ;
      var varK  = actK - compK;

      var meta = itemMaster[itc] || { description: '', spec: '', uom: '' };

      byItem.push({
        item_code: itc,
        description: meta.description,
        spec: meta.spec,
        uom: meta.uom,
        saldo_awal_qty: awal.qty, saldo_awal_kg: awal.kg,
        in_receipt_qty: rcv.qty,  in_receipt_kg: rcv.kg,
        in_prod_qty   : prd.qty,  in_prod_kg   : prd.kg,
        out_cons_qty  : con.qty,  out_cons_kg  : con.kg,
        out_delv_qty  : dlv.qty,  out_delv_kg  : dlv.kg,
        out_ng_qty    : ng.qty,   out_ng_kg    : ng.kg,
        tf_in_qty     : tf.qty_in,  tf_in_kg     : tf.kg_in,
        tf_out_qty    : tf.qty_out, tf_out_kg    : tf.kg_out,
        adj_qty       : aj.qty,   adj_kg       : aj.kg,
        saldo_akhir_qty: actQ,    saldo_akhir_kg: actK,
        computed_qty  : compQ,    computed_kg  : compK,
        variance_qty  : varQ,     variance_kg  : varK,
        has_actual_saldo: !!akh
      });
    }
    byItem.sort(function(a, b) { return String(a.item_code).localeCompare(String(b.item_code)); });

    // === ASSEMBLE BY-BATCH REPORT ===
    var allBatches = {};
    function collectBatches(obj) { for (var k in obj) allBatches[k] = true; }
    collectBatches(saldoAwalByBatch);
    collectBatches(inReceiptByBatch);
    collectBatches(inProdByBatch);
    collectBatches(outConsByBatch);
    collectBatches(outDelvByBatch);
    collectBatches(outNgByBatch);
    collectBatches(adjByBatch);
    collectBatches(saldoAkhirByBatch);

    var byBatch = [];
    for (var bid in allBatches) {
      var aw = saldoAwalByBatch[bid]  || { qty: 0, kg: 0, item_code: '', loc: '', owner: '' };
      var rc = inReceiptByBatch[bid]  || { qty: 0, kg: 0, item_code: '', loc: '', owner: '' };
      var pr = inProdByBatch[bid]     || { qty: 0, kg: 0, item_code: '', loc: '', owner: '' };
      var co = outConsByBatch[bid]    || { qty: 0, kg: 0, item_code: '', loc: '', owner: '' };
      var dl = outDelvByBatch[bid]    || { qty: 0, kg: 0, item_code: '', loc: '', owner: '' };
      var nge = outNgByBatch[bid]     || { qty: 0, kg: 0, item_code: '', loc: '', owner: '' };
      var ad = adjByBatch[bid]        || { qty: 0, kg: 0, item_code: '', loc: '', owner: '' };
      var ah = saldoAkhirByBatch[bid];

      var itc2  = aw.item_code || rc.item_code || pr.item_code || co.item_code || dl.item_code || (ah ? ah.item_code : '') || '';
      var loc2  = aw.loc       || rc.loc       || pr.loc       || co.loc       || dl.loc       || (ah ? ah.loc : '')       || '';
      var own2  = aw.owner     || rc.owner     || pr.owner     || co.owner     || dl.owner     || (ah ? ah.owner : '')     || '';

      var cQ = aw.qty + rc.qty + pr.qty - co.qty - dl.qty - nge.qty + ad.qty;
      var cK = aw.kg  + rc.kg  + pr.kg  - co.kg  - dl.kg  - nge.kg  + ad.kg;
      var aQ = ah ? ah.qty : cQ;
      var aK = ah ? ah.kg  : cK;

      byBatch.push({
        batch_id: bid,
        item_code: itc2,
        loc: loc2,
        owner: own2,
        saldo_awal_qty: aw.qty, saldo_awal_kg: aw.kg,
        in_receipt_qty: rc.qty, in_receipt_kg: rc.kg,
        in_prod_qty   : pr.qty, in_prod_kg   : pr.kg,
        out_cons_qty  : co.qty, out_cons_kg  : co.kg,
        out_delv_qty  : dl.qty, out_delv_kg  : dl.kg,
        out_ng_qty    : nge.qty, out_ng_kg   : nge.kg,
        adj_qty       : ad.qty, adj_kg       : ad.kg,
        saldo_akhir_qty: aQ,    saldo_akhir_kg: aK,
        computed_qty  : cQ,     computed_kg  : cK,
        variance_qty  : aQ - cQ, variance_kg : aK - cK,
        has_actual_saldo: !!ah
      });
    }
    byBatch.sort(function(a, b) { return String(a.batch_id).localeCompare(String(b.batch_id)); });

    // === SUMMARY ===
    var summary = {
      total_items: byItem.length,
      total_batches: byBatch.length,
      total_in_qty: 0, total_in_kg: 0,
      total_out_qty: 0, total_out_kg: 0,
      total_adj_qty: 0, total_adj_kg: 0,
      total_saldo_awal_qty: 0, total_saldo_awal_kg: 0,
      total_saldo_akhir_qty: 0, total_saldo_akhir_kg: 0
    };
    for (var z = 0; z < byItem.length; z++) {
      var r = byItem[z];
      summary.total_in_qty  += r.in_receipt_qty + r.in_prod_qty + r.tf_in_qty;
      summary.total_in_kg   += r.in_receipt_kg  + r.in_prod_kg  + r.tf_in_kg;
      summary.total_out_qty += r.out_cons_qty + r.out_delv_qty + r.out_ng_qty + r.tf_out_qty;
      summary.total_out_kg  += r.out_cons_kg  + r.out_delv_kg  + r.out_ng_kg  + r.tf_out_kg;
      summary.total_adj_qty += r.adj_qty;
      summary.total_adj_kg  += r.adj_kg;
      summary.total_saldo_awal_qty  += r.saldo_awal_qty;
      summary.total_saldo_awal_kg   += r.saldo_awal_kg;
      summary.total_saldo_akhir_qty += r.saldo_akhir_qty;
      summary.total_saldo_akhir_kg  += r.saldo_akhir_kg;
    }

    return {
      success: true,
      period: period,
      prev_period: prevPeriod,
      period_start: range.start.toISOString(),
      period_end: range.end.toISOString(),
      saldo_awal_source: saldoAwalSource,
      saldo_akhir_source: saldoAkhirSource,
      filter: { owner: fOwner, loc: fLoc },
      by_item: byItem,
      by_batch: byBatch,
      summary: summary
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// QUICK TEST
// =========================================================================
function _test_getMutasiReport() {
  var now = new Date();
  var prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var period = prev.getFullYear() + '-' + String(prev.getMonth() + 1).padStart(2, '0');
  var res = getMutasiReport(period, { owner: '', loc: '' });
  Logger.log('Success: ' + res.success);
  if (res.success) {
    Logger.log('Period: ' + res.period + ' | Prev: ' + res.prev_period);
    Logger.log('Saldo Awal Source: ' + res.saldo_awal_source);
    Logger.log('Saldo Akhir Source: ' + res.saldo_akhir_source);
    Logger.log('Total items: ' + res.summary.total_items);
    Logger.log('Total batches: ' + res.summary.total_batches);
    Logger.log('Total IN qty: ' + res.summary.total_in_qty + ' / KG: ' + res.summary.total_in_kg);
    Logger.log('Total OUT qty: ' + res.summary.total_out_qty + ' / KG: ' + res.summary.total_out_kg);
    Logger.log('First 3 items:');
    res.by_item.slice(0, 3).forEach(function(r) {
      Logger.log('  ' + r.item_code + ' | Awal=' + r.saldo_awal_kg + ' | IN=' + (r.in_receipt_kg + r.in_prod_kg) + ' | OUT=' + (r.out_cons_kg + r.out_delv_kg + r.out_ng_kg) + ' | Akhir=' + r.saldo_akhir_kg);
    });
  } else {
    Logger.log('Error: ' + res.message);
  }
}
// =========================================================================
// HELPER: Build details (7 array rincian transaksi untuk export Accounting)
// Dipakai oleh getMutasiReportFull
// =========================================================================
function _rptBuildDetails(period, filter) {
  filter = filter || {};
  var fLoc = filter.loc ? String(filter.loc) : '';
  var range = _rptPeriodRange(period);

  var details = {
    gr:   [],
    delv: [],
    prod: [],
    cons: [],
    ng:   [],
    ict:  [],
    adj:  []
  };

  var gr   = _rptReadSheet('GR');
  var tl   = _rptReadSheet('Trace_Log');
  var spk  = _rptReadSheet('SPK');
  var delv = _rptReadSheet('DELV');
  var ict  = _rptReadSheet('Rekap_ICT');
  var op   = _rptReadSheet('Opname_Bulanan');
  var item = _rptReadSheet('M_ITEM');

  // Master Item lookup (untuk Description di sheet yg ga punya kolom Description)
  var itemMaster = {};
  if (item.idx['Item_Code'] !== undefined) {
    for (var im = 0; im < item.rows.length; im++) {
      var ic = String(item.rows[im][item.idx['Item_Code']] || '').trim();
      if (!ic) continue;
      itemMaster[ic] = {
        description: item.idx['Description'] !== undefined ? String(item.rows[im][item.idx['Description']] || '') : '',
        spec       : item.idx['Spec'] !== undefined ? String(item.rows[im][item.idx['Spec']] || '') : '',
        uom        : item.idx['UoM'] !== undefined ? String(item.rows[im][item.idx['UoM']] || '') : ''
      };
    }
  }

  // ===== GR (Penerimaan) =====
  if (gr.idx['Tgl_Masuk'] !== undefined) {
    for (var i = 0; i < gr.rows.length; i++) {
      var r = gr.rows[i];
      var tgl = r[gr.idx['Tgl_Masuk']];
      if (!_rptInRange(tgl, range)) continue;
      var tloc = gr.idx['Target_Loc'] !== undefined ? String(r[gr.idx['Target_Loc']] || '') : '';
      var locMapped = tloc === 'Stok_Coil' ? 'Coil' : (tloc === 'Stok_Sheet' ? 'Sheet' : tloc);
      if (fLoc && locMapped !== fLoc) continue;

      details.gr.push({
        tgl_masuk:   tgl,
        batch_id:    String(r[gr.idx['Batch_ID']] || ''),
        item_code:   String(r[gr.idx['Item_Code']] || ''),
        description: gr.idx['Description'] !== undefined ? String(r[gr.idx['Description']] || '') : '',
        spec:        gr.idx['Spec'] !== undefined ? String(r[gr.idx['Spec']] || '') : '',
        t:           gr.idx['T'] !== undefined ? r[gr.idx['T']] : '',
        p:           gr.idx['P'] !== undefined ? r[gr.idx['P']] : '',
        l:           gr.idx['L'] !== undefined ? r[gr.idx['L']] : '',
        supplier:    gr.idx['Supplier'] !== undefined ? String(r[gr.idx['Supplier']] || '') : '',
        no_coil:     gr.idx['No_Coil'] !== undefined ? String(r[gr.idx['No_Coil']] || '') : '',
        no_po:       gr.idx['No_PO'] !== undefined ? String(r[gr.idx['No_PO']] || '') : '',
        no_do:       gr.idx['No_DO'] !== undefined ? String(r[gr.idx['No_DO']] || '') : '',
        qty:         parseFloat(r[gr.idx['QTY_In'] !== undefined ? gr.idx['QTY_In'] : gr.idx['Qty_In']]) || 0,
        kg:          parseFloat(r[gr.idx['KG_In']]) || 0,
        owner:       String(r[gr.idx['Owner']] || ''),
        target_loc:  locMapped
      });
    }
  }

  // ===== DELV (Penjualan) =====
  if (delv.idx['Tanggal'] !== undefined) {
    for (var i = 0; i < delv.rows.length; i++) {
      var r = delv.rows[i];
      var tgl = r[delv.idx['Tanggal']];
      if (!_rptInRange(tgl, range)) continue;
      if (fLoc && fLoc !== 'FG') continue;

      details.delv.push({
        tanggal:          tgl,
        sj_no:            delv.idx['SJ_No'] !== undefined ? String(r[delv.idx['SJ_No']] || '') : '',
        so_no:            delv.idx['SO_No'] !== undefined ? String(r[delv.idx['SO_No']] || '') : '',
        batch_id:         String(r[delv.idx['Batch_ID']] || ''),
        item_code:        String(r[delv.idx['Item_Code']] || ''),
        description:      delv.idx['Description'] !== undefined ? String(r[delv.idx['Description']] || '') : '',
        uom:              delv.idx['UoM'] !== undefined ? String(r[delv.idx['UoM']] || '') : '',
        delv_q:           parseFloat(r[delv.idx['Delv_Q']]) || 0,
        delv_kg:          parseFloat(r[delv.idx['Delv_KG']]) || 0,
        cust:             delv.idx['Cust'] !== undefined ? String(r[delv.idx['Cust']] || '') : '',
        owner:            String(r[delv.idx['Owner']] || ''),
        spk_ref:          delv.idx['Spk_ref'] !== undefined ? String(r[delv.idx['Spk_ref']] || '') : (delv.idx['SPK_Ref'] !== undefined ? String(r[delv.idx['SPK_Ref']] || '') : ''),
        no_armada:        delv.idx['No_Armada'] !== undefined ? String(r[delv.idx['No_Armada']] || '') : '',
        driver:           delv.idx['Driver'] !== undefined ? String(r[delv.idx['Driver']] || '') : '',
        diterima_finance: delv.idx['Diterima_Finance'] !== undefined ? String(r[delv.idx['Diterima_Finance']] || '') : ''
      });
    }
  }

  // ===== PRODUKSI (Trace_Log IN types) =====
  if (tl.idx['Tgl_Buat'] !== undefined) {
    for (var i = 0; i < tl.rows.length; i++) {
      var r = tl.rows[i];
      var typ = String(r[tl.idx['Type']] || '').toUpperCase();
      if (typ === 'NG') continue;
      if (typ !== 'SHEET' && typ !== 'WIP' && typ !== 'FGC' && typ !== 'FGS' && typ !== 'FCL') continue;
      var tgl = r[tl.idx['Tgl_Buat']];
      if (!_rptInRange(tgl, range)) continue;
      var locMap = typ === 'SHEET' ? 'Sheet' : (typ === 'WIP' ? 'WIP' : 'FG');
      if (fLoc && locMap !== fLoc) continue;

      details.prod.push({
        tgl_buat:     tgl,
        batch_id:     String(r[tl.idx['Batch_ID']] || ''),
        level:        tl.idx['Level'] !== undefined ? r[tl.idx['Level']] : '',
        type:         typ,
        source_batch: tl.idx['Source_Batch'] !== undefined ? String(r[tl.idx['Source_Batch']] || '') : '',
        spk_ref:      tl.idx['SPK_Ref'] !== undefined ? String(r[tl.idx['SPK_Ref']] || '') : '',
        item_code:    String(r[tl.idx['Item_Code']] || ''),
        description:  tl.idx['Description'] !== undefined ? String(r[tl.idx['Description']] || '') : '',
        spec:         tl.idx['Spec'] !== undefined ? String(r[tl.idx['Spec']] || '') : '',
        t:            tl.idx['T'] !== undefined ? r[tl.idx['T']] : '',
        p:            tl.idx['P'] !== undefined ? r[tl.idx['P']] : '',
        qty:          parseFloat(r[tl.idx['Qty']]) || 0,
        kg:           parseFloat(r[tl.idx['KG']]) || 0,
        operator:     tl.idx['Operator'] !== undefined ? String(r[tl.idx['Operator']] || '') : '',
        mc_no:        tl.idx['MC_No'] !== undefined ? String(r[tl.idx['MC_No']] || '') : '',
        owner:        String(r[tl.idx['Owner']] || '')
      });
    }
  }

  // ===== KONSUMSI (SPK Header DONE) =====
  if (spk.idx['SPK_No'] !== undefined) {
    for (var i = 0; i < spk.rows.length; i++) {
      var r = spk.rows[i];
      var typ = String(r[spk.idx['SPK_Type']] || '');
      if (typ.indexOf('HEADER') === -1) continue;
      var status = String(r[spk.idx['Status']] || '').toUpperCase();
      if (status !== 'DONE') continue;
      var sel = r[spk.idx['Selesai_DT']];
      if (!_rptInRange(sel, range)) continue;
      var srcLoc = spk.idx['Source_Loc'] !== undefined ? String(r[spk.idx['Source_Loc']] || '') : '';
      var srcLocMapped = '';
      var slL = srcLoc.toLowerCase().replace(/\s/g, '_');
      if (slL.indexOf('coil') > -1) srcLocMapped = 'Coil';
      else if (slL.indexOf('sheet') > -1) srcLocMapped = 'Sheet';
      else if (slL.indexOf('wip') > -1) srcLocMapped = 'WIP';
      if (fLoc && srcLocMapped !== fLoc) continue;

      var itemCd = String(r[spk.idx['Item_Code']] || '');
      var meta = itemMaster[itemCd] || { description: '' };

      details.cons.push({
        spk_no:      String(r[spk.idx['SPK_No']] || ''),
        spk_type:    typ,
        tgl_buat:    spk.idx['Tgl_Buat'] !== undefined ? r[spk.idx['Tgl_Buat']] : '',
        selesai_dt:  sel,
        batch_id:    String(r[spk.idx['Batch_ID']] || ''),
        item_code:   itemCd,
        description: meta.description,
        qty_actual:  parseFloat(r[spk.idx['Qty_Actual']]) || 0,
        kg_actual:   parseFloat(r[spk.idx['KG_Actual']]) || 0,
        mc_no:       spk.idx['MC_No'] !== undefined ? String(r[spk.idx['MC_No']] || '') : '',
        op:          spk.idx['OP'] !== undefined ? String(r[spk.idx['OP']] || '') : '',
        owner:       String(r[spk.idx['Owner']] || ''),
        owner_used:  spk.idx['Owner_Used'] !== undefined ? String(r[spk.idx['Owner_Used']] || '') : '',
        source_loc:  srcLocMapped,
        target_loc:  spk.idx['Target_Loc'] !== undefined ? String(r[spk.idx['Target_Loc']] || '') : ''
      });
    }
  }

  // ===== NG (Trace_Log Type=NG) =====
  if (tl.idx['Tgl_Buat'] !== undefined) {
    for (var i = 0; i < tl.rows.length; i++) {
      var r = tl.rows[i];
      var typ = String(r[tl.idx['Type']] || '').toUpperCase();
      if (typ !== 'NG') continue;
      var tgl = r[tl.idx['Tgl_Buat']];
      if (!_rptInRange(tgl, range)) continue;
      if (fLoc && fLoc !== 'NG') continue;

      details.ng.push({
        tgl_buat:     tgl,
        batch_id_ng:  String(r[tl.idx['Batch_ID']] || ''),
        source_batch: tl.idx['Source_Batch'] !== undefined ? String(r[tl.idx['Source_Batch']] || '') : '',
        spk_ref:      tl.idx['SPK_Ref'] !== undefined ? String(r[tl.idx['SPK_Ref']] || '') : '',
        item_code:    String(r[tl.idx['Item_Code']] || ''),
        description:  tl.idx['Description'] !== undefined ? String(r[tl.idx['Description']] || '') : '',
        qty:          parseFloat(r[tl.idx['Qty']]) || 0,
        kg:           parseFloat(r[tl.idx['KG']]) || 0,
        operator:     tl.idx['Operator'] !== undefined ? String(r[tl.idx['Operator']] || '') : '',
        mc_no:        tl.idx['MC_No'] !== undefined ? String(r[tl.idx['MC_No']] || '') : '',
        owner:        String(r[tl.idx['Owner']] || '')
      });
    }
  }

  // ===== ICT (Rekap_ICT) =====
  if (ict.idx['Tgl_Transfer'] !== undefined) {
    for (var i = 0; i < ict.rows.length; i++) {
      var r = ict.rows[i];
      var tgl = r[ict.idx['Tgl_Transfer']];
      if (!_rptInRange(tgl, range)) continue;

      var itemCd = String(r[ict.idx['Item_Code']] || '');
      var meta = itemMaster[itemCd] || { description: '' };

      details.ict.push({
        tgl_transfer: tgl,
        spk_no:       ict.idx['SPK_No'] !== undefined ? String(r[ict.idx['SPK_No']] || '') : '',
        item_code:    itemCd,
        description:  ict.idx['Description'] !== undefined ? String(r[ict.idx['Description']] || '') : meta.description,
        dari_owner:   ict.idx['Dari_Owner'] !== undefined ? String(r[ict.idx['Dari_Owner']] || '') : '',
        ke_owner:     ict.idx['Ke_Owner'] !== undefined ? String(r[ict.idx['Ke_Owner']] || '') : '',
        qty_sht:      parseFloat(r[ict.idx['Qty_Sht']]) || 0,
        qty_kg:       parseFloat(r[ict.idx['Qty_KG']]) || 0
      });
    }
  }

  // ===== ADJUSTMENT (Opname current period LOCKED, selisih ≠ 0) =====
  if (op.idx['Period'] !== undefined) {
    for (var i = 0; i < op.rows.length; i++) {
      var r = op.rows[i];
      if (String(r[op.idx['Period']]).trim() !== period) continue;
      var st = String(r[op.idx['Status']] || '').toUpperCase();
      if (st !== 'FINAL' && st !== 'LOCKED') continue;
      var qSel = parseFloat(r[op.idx['Qty_Selisih']]) || 0;
      var kSel = parseFloat(r[op.idx['KG_Selisih']]) || 0;
      if (qSel === 0 && kSel === 0) continue;
      var loc = String(r[op.idx['Loc']] || '');
      if (fLoc && loc !== fLoc) continue;

      details.adj.push({
        period:       String(r[op.idx['Period']]).trim(),
        batch_id:     String(r[op.idx['Batch_ID']] || ''),
        item_code:    String(r[op.idx['Item_Code']] || ''),
        loc:          loc,
        owner:        String(r[op.idx['Owner']] || ''),
        qty_system:   parseFloat(r[op.idx['Qty_System']]) || 0,
        kg_system:    parseFloat(r[op.idx['KG_System']]) || 0,
        qty_fisik:    parseFloat(r[op.idx['Qty_Fisik']]) || 0,
        kg_fisik:     parseFloat(r[op.idx['KG_Fisik']]) || 0,
        qty_selisih:  qSel,
        kg_selisih:   kSel,
        note:         op.idx['Note'] !== undefined ? String(r[op.idx['Note']] || '') : '',
        locked_by:    op.idx['Locked_By'] !== undefined ? String(r[op.idx['Locked_By']] || '') : ''
      });
    }
  }

  return details;
}

// =========================================================================
// MAIN: getMutasiReportFull(period, filter)
// Return: ringkasan per owner (FC/DRC/Grand) + rincian semua transaksi
// Dipakai oleh page_report.html untuk render dashboard + export Excel
// =========================================================================
function getMutasiReportFull(period, filter) {
  try {
    if (!period || !/^\d{4}-\d{2}$/.test(String(period).trim())) {
      throw new Error('Format period invalid. Pakai YYYY-MM');
    }
    period = String(period).trim();
    filter = filter || {};
    var fLoc = filter.loc ? String(filter.loc) : '';

    // Panggil engine existing 3x untuk 3 perspektif owner
    var grand = getMutasiReport(period, { owner: '', loc: fLoc });
    if (!grand.success) throw new Error('Grand Total gagal: ' + grand.message);

    var fc = getMutasiReport(period, { owner: 'FC', loc: fLoc });
    if (!fc.success) throw new Error('FC report gagal: ' + fc.message);

    var drc = getMutasiReport(period, { owner: 'DRC', loc: fLoc });
    if (!drc.success) throw new Error('DRC report gagal: ' + drc.message);

    // Build rincian transaksi (untuk export Accounting)
    var details = _rptBuildDetails(period, { loc: fLoc });

    var payload = {
      success: true,
      period: period,
      prev_period: grand.prev_period,
      period_start: grand.period_start,
      period_end: grand.period_end,
      saldo_awal_source: grand.saldo_awal_source,
      saldo_akhir_source: grand.saldo_akhir_source,
      filter: { loc: fLoc },
      fc:    { by_item: fc.by_item,    by_batch: fc.by_batch,    summary: fc.summary },
      drc:   { by_item: drc.by_item,   by_batch: drc.by_batch,   summary: drc.summary },
      grand: { by_item: grand.by_item, by_batch: grand.by_batch, summary: grand.summary },
      details: details
    };

    // Force JSON-safe round-trip: konversi Date → ISO string, drop undefined, sanitize NaN/Infinity
    // Wajib untuk google.script.run yang strict soal payload non-primitive
    return JSON.parse(JSON.stringify(payload));

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// QUICK TEST untuk getMutasiReportFull
// =========================================================================
function _test_getMutasiReportFull() {
  var res = getMutasiReportFull('2026-06', {});
  Logger.log('=== getMutasiReportFull TEST ===');
  Logger.log('Success: ' + res.success);
  if (!res.success) { Logger.log('Error: ' + res.message); return; }
  Logger.log('Period: ' + res.period + ' | Prev: ' + res.prev_period);
  Logger.log('Saldo Awal: ' + res.saldo_awal_source + ' | Saldo Akhir: ' + res.saldo_akhir_source);
  Logger.log('--- FC ---');
  Logger.log('  items=' + res.fc.summary.total_items + ' | in_kg=' + res.fc.summary.total_in_kg + ' | out_kg=' + res.fc.summary.total_out_kg);
  Logger.log('--- DRC ---');
  Logger.log('  items=' + res.drc.summary.total_items + ' | in_kg=' + res.drc.summary.total_in_kg + ' | out_kg=' + res.drc.summary.total_out_kg);
  Logger.log('--- GRAND ---');
  Logger.log('  items=' + res.grand.summary.total_items + ' | in_kg=' + res.grand.summary.total_in_kg + ' | out_kg=' + res.grand.summary.total_out_kg);
  Logger.log('--- DETAILS ---');
  Logger.log('  GR=' + res.details.gr.length +
             ' | Delv=' + res.details.delv.length +
             ' | Prod=' + res.details.prod.length +
             ' | Cons=' + res.details.cons.length +
             ' | NG=' + res.details.ng.length +
             ' | ICT=' + res.details.ict.length +
             ' | Adj=' + res.details.adj.length);
}