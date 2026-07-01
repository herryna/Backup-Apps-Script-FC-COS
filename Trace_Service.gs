// =========================================================================
// CORE OPERATIONS SYSTEM (COS) - TRACE SERVICE (T2 v2 - dengan Mass Balance)
//
// 5 fungsi public:
//   - getBatchTrace(batchId)         : genealogy tree (ancestors + descendants) + flat array
//   - getBatchMutasi(batchId)        : timeline event kronologis
//   - searchBatchByFilter(f)         : pencarian batch dengan filter
//   - getBatchMassBalance(batchId)   : MASS BALANCE per batch (BARU)
//   - getThresholds()                : ambil threshold dari M_Threshold (BARU)
//
// Data sources (READ-ONLY):
//   GR, Trace_Log, SPK, DELV, Rekap_ICT,
//   Stok_Coil, Stok_Sheet, Stok_WIP, Stok_FG, Stok_NG, M_Threshold
// =========================================================================

var TRACE_CACHE_ = null; // in-execution cache

function _hdrIdx(hdr) {
  var idx = {};
  for (var i = 0; i < hdr.length; i++) idx[hdr[i]] = i;
  return idx;
}

function _loadTraceData() {
  if (TRACE_CACHE_) return TRACE_CACHE_;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function readSheet(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return { rows: [], hdr: [], idx: {} };
    var d = sh.getDataRange().getValues();
    if (d.length < 2) return { rows: [], hdr: d[0] ? d[0].map(function(h){return String(h).trim();}) : [], idx: {} };
    var hdr = d[0].map(function(h){ return String(h).trim(); });
    return { rows: d.slice(1), hdr: hdr, idx: _hdrIdx(hdr) };
  }

  TRACE_CACHE_ = {
    gr   : readSheet('GR'),
    tl   : readSheet('Trace_Log'),
    spk  : readSheet('SPK'),
    dv   : readSheet('DELV'),
    ict  : readSheet('Rekap_ICT'),
    fg   : readSheet('Stok_FG'),
    // ── BARU untuk Mass Balance ──
    coil : readSheet('Stok_Coil'),
    sht  : readSheet('Stok_Sheet'),
    wip  : readSheet('Stok_WIP'),
    ng   : readSheet('Stok_NG'),
    thr  : readSheet('M_Threshold')
  };

  _buildTraceIndexes(TRACE_CACHE_);
  return TRACE_CACHE_;
}

function _buildTraceIndexes(cache) {
  // ---- GR index ----
  cache.gr.byBatch = {};
  var iGB = cache.gr.idx['Batch_ID'];
  if (iGB !== undefined) {
    for (var i = 0; i < cache.gr.rows.length; i++) {
      var b = String(cache.gr.rows[i][iGB] || '').trim();
      if (b) cache.gr.byBatch[b] = i;
    }
  }

  // ---- Trace_Log index (by batch + by source) ----
  cache.tl.byBatch = {};
  cache.tl.bySrc   = {};
  var iTB = cache.tl.idx['Batch_ID'];
  var iTS = cache.tl.idx['Source_Batch'];
  if (iTB !== undefined) {
    for (var j = 0; j < cache.tl.rows.length; j++) {
      var bb = String(cache.tl.rows[j][iTB] || '').trim();
      if (bb) cache.tl.byBatch[bb] = j;
      var ss2 = iTS !== undefined ? String(cache.tl.rows[j][iTS] || '').trim() : '';
      if (ss2) {
        if (!cache.tl.bySrc[ss2]) cache.tl.bySrc[ss2] = [];
        cache.tl.bySrc[ss2].push(j);
      }
    }
  }

  // ---- DELV index by Batch_ID ----
  cache.dv.byBatch = {};
  var iDB = cache.dv.idx['Batch_ID'];
  if (iDB !== undefined) {
    for (var k = 0; k < cache.dv.rows.length; k++) {
      var b3 = String(cache.dv.rows[k][iDB] || '').trim();
      if (b3) {
        if (!cache.dv.byBatch[b3]) cache.dv.byBatch[b3] = [];
        cache.dv.byBatch[b3].push(k);
      }
    }
  }

  // ---- SPK index by Batch_ID + by SPK_No ----
  cache.spk.byBatch = {};
  cache.spk.bySpkNo = {};
  var iSpkB  = cache.spk.idx['Batch_ID'];
  var iSpkNo = cache.spk.idx['SPK_No'];
  for (var m = 0; m < cache.spk.rows.length; m++) {
    var b4 = iSpkB !== undefined ? String(cache.spk.rows[m][iSpkB] || '').trim() : '';
    if (b4) {
      if (!cache.spk.byBatch[b4]) cache.spk.byBatch[b4] = [];
      cache.spk.byBatch[b4].push(m);
    }
    var sn = iSpkNo !== undefined ? String(cache.spk.rows[m][iSpkNo] || '').trim() : '';
    if (sn) cache.spk.bySpkNo[sn] = m;
  }

  // ---- Rekap_ICT index by SPK_No ----
  cache.ict.bySpkNo = {};
  var iIctSpk = cache.ict.idx['SPK_No'];
  if (iIctSpk !== undefined) {
    for (var n = 0; n < cache.ict.rows.length; n++) {
      var s2 = String(cache.ict.rows[n][iIctSpk] || '').trim();
      if (s2) {
        if (!cache.ict.bySpkNo[s2]) cache.ict.bySpkNo[s2] = [];
        cache.ict.bySpkNo[s2].push(n);
      }
    }
  }

  // ---- Stok_FG index by Batch_ID ----
  cache.fg.byBatch = {};
  var iFgB = cache.fg.idx['Batch_ID'];
  if (iFgB !== undefined) {
    for (var p = 0; p < cache.fg.rows.length; p++) {
      var b5 = String(cache.fg.rows[p][iFgB] || '').trim();
      if (b5) cache.fg.byBatch[b5] = p;
    }
  }

  // ---- Stok_Coil/Sheet/WIP/NG indexes (BARU) ----
  ['coil','sht','wip','ng'].forEach(function(key) {
    cache[key].byBatch = {};
    var iB = cache[key].idx['Batch_ID'];
    if (iB !== undefined) {
      for (var i = 0; i < cache[key].rows.length; i++) {
        var bid = String(cache[key].rows[i][iB] || '').trim();
        if (bid) cache[key].byBatch[bid] = i;
      }
    }
  });

  // ---- M_Threshold index by Metric_Name (BARU) ----
  cache.thr.byMetric = {};
  var iThM = cache.thr.idx['Metric_Name'];
  if (iThM !== undefined) {
    for (var t = 0; t < cache.thr.rows.length; t++) {
      var mn = String(cache.thr.rows[t][iThM] || '').trim();
      if (mn) cache.thr.byMetric[mn] = t;
    }
  }
}

function _resolveBatch(batchId, cache) {
  var b = String(batchId).trim();

  // Coba GR dulu (level 0 / root)
  if (cache.gr.byBatch[b] !== undefined) {
    var ri = cache.gr.byBatch[b];
    var rg = cache.gr.rows[ri];
    var gI = cache.gr.idx;
    var iGrQty = gI['QTY_In'] !== undefined ? gI['QTY_In'] : gI['Qty_In'];
    return {
      found        : true,
      source       : 'GR',
      batch_id     : b,
      level        : 0,
      type         : 'GR',
      tgl_buat     : rg[gI['Tgl_Masuk']],
      item_code    : String(rg[gI['Item_Code']] || ''),
      description  : String(rg[gI['Description']] || ''),
      spec         : String(rg[gI['Spec']] || ''),
      t            : rg[gI['T']] || '',
      p            : rg[gI['P']] || '',
      l            : rg[gI['L']] || '',
      qty          : iGrQty !== undefined ? (parseFloat(rg[iGrQty]) || 0) : 0,
      kg           : parseFloat(rg[gI['KG_In']]) || 0,
      owner        : String(rg[gI['Owner']] || ''),
      owner_used   : String(rg[gI['Owner']] || ''),
      target_loc   : String(rg[gI['Target_Loc']] || ''),
      supplier     : String(rg[gI['Supplier']] || ''),
      no_po        : String(rg[gI['No_PO']] || ''),
      no_do        : String(rg[gI['No_DO']] || ''),
      no_coil      : String(rg[gI['No_Coil']] || ''),
      source_batch : '',
      root_batch   : b,
      spk_ref      : '',
      mc_no        : '',
      operator     : '',
      full_chain   : b
    };
  }

  // Coba Trace_Log
  if (cache.tl.byBatch[b] !== undefined) {
    var ti = cache.tl.byBatch[b];
    var rt = cache.tl.rows[ti];
    var tI = cache.tl.idx;
    return {
      found        : true,
      source       : 'Trace_Log',
      batch_id     : b,
      level        : parseInt(rt[tI['Level']]) || 0,
      type         : String(rt[tI['Type']] || ''),
      tgl_buat     : rt[tI['Tgl_Buat']],
      item_code    : String(rt[tI['Item_Code']] || ''),
      description  : String(rt[tI['Description']] || ''),
      spec         : String(rt[tI['Spec']] || ''),
      t            : rt[tI['T']] || '',
      p            : rt[tI['P']] || '',
      l            : rt[tI['L_dim']] || '',
      qty          : parseFloat(rt[tI['Qty']]) || 0,
      kg           : parseFloat(rt[tI['KG']]) || 0,
      owner        : String(rt[tI['Owner']] || ''),
      owner_used   : String(rt[tI['Owner_Used']] || ''),
      supplier     : String(rt[tI['Supplier']] || ''),
      no_po        : String(rt[tI['No_PO']] || ''),
      no_do        : String(rt[tI['No_DO']] || ''),
      source_batch : String(rt[tI['Source_Batch']] || ''),
      root_batch   : String(rt[tI['Root_Batch']] || ''),
      spk_ref      : String(rt[tI['SPK_Ref']] || ''),
      mc_no        : String(rt[tI['MC_No']] || ''),
      operator     : String(rt[tI['Operator']] || ''),
      full_chain   : String(rt[tI['Full_Chain']] || '')
    };
  }

  return { found: false, batch_id: b };
}

function _sanitizeDates(obj) {
  if (!obj) return obj;
  for (var k in obj) {
    if (obj[k] instanceof Date) obj[k] = obj[k].toISOString();
  }
  return obj;
}

function _sanitizeTree(node) {
  if (!node) return null;
  _sanitizeDates(node);
  if (node.children && node.children.length > 0) {
    var arr = [];
    for (var i = 0; i < node.children.length; i++) {
      arr.push(_sanitizeTree(node.children[i]));
    }
    node.children = arr;
  }
  return node;
}

// =========================================================================
// PUBLIC API 1: getBatchTrace(batchId)
//   Returns tree (nested) + flat array (ancestors..current..descendants)
//   + meta (root, supplier, total descendants) + deliveries list
// =========================================================================
function getBatchTrace(batchId) {
  try {
    if (!batchId) throw new Error('batchId kosong');
    var cache = _loadTraceData();
    var b = String(batchId).trim();

    var current = _resolveBatch(b, cache);
    if (!current.found) {
      return { success: false, message: 'Batch tidak ditemukan: ' + b };
    }

    // ---- ANCESTORS (backward dari current ke root) ----
    var ancestors = [];
    var seenA = {};
    seenA[b] = true;
    var cur = current;
    var safety = 0;
    while (cur.source_batch && safety < 50) {
      if (seenA[cur.source_batch]) break;
      seenA[cur.source_batch] = true;
      var parent = _resolveBatch(cur.source_batch, cache);
      if (!parent.found) break;
      ancestors.unshift(parent);
      cur = parent;
      safety++;
    }

    // ---- DESCENDANTS (forward BFS) ----
    var flatDesc = [];
    var queue = [{ batch: b, level: 0 }];
    var seenD = {};
    seenD[b] = true;
    var safety2 = 0;
    while (queue.length > 0 && safety2 < 5000) {
      var item = queue.shift();
      var childIdxs = cache.tl.bySrc[item.batch] || [];
      for (var k = 0; k < childIdxs.length; k++) {
        var cr = cache.tl.rows[childIdxs[k]];
        var childBatch = String(cr[cache.tl.idx['Batch_ID']] || '').trim();
        if (!childBatch || seenD[childBatch]) continue;
        seenD[childBatch] = true;
        var resolved = _resolveBatch(childBatch, cache);
        if (resolved.found) {
          resolved.depth_from_current = item.level + 1;
          resolved.parent_in_tree = item.batch;
          flatDesc.push(resolved);
          queue.push({ batch: childBatch, level: item.level + 1 });
        }
      }
      safety2++;
    }

    // ---- BUILD TREE STRUCTURE (recursive) ----
    function buildSubTree(batchId, seen, depth) {
      if (depth > 30) return null;
      if (seen[batchId]) return null;
      seen[batchId] = true;
      var node = _resolveBatch(batchId, cache);
      if (!node.found) return null;
      node.children = [];
      var ci = cache.tl.bySrc[batchId] || [];
      for (var k = 0; k < ci.length; k++) {
        var cr2 = cache.tl.rows[ci[k]];
        var cb = String(cr2[cache.tl.idx['Batch_ID']] || '').trim();
        if (!cb) continue;
        var sub = buildSubTree(cb, seen, depth + 1);
        if (sub) node.children.push(sub);
      }
      return node;
    }

    var tree = buildSubTree(b, {}, 0);

    // ---- FLAT ARRAY ----
    var flat = [];
    for (var ai = 0; ai < ancestors.length; ai++) {
      ancestors[ai].depth_from_current = -(ancestors.length - ai);
      flat.push(ancestors[ai]);
    }
    current.depth_from_current = 0;
    flat.push(current);
    for (var di = 0; di < flatDesc.length; di++) flat.push(flatDesc[di]);

    // ---- DELIVERIES ----
    var deliveries = [];
    var allBatchesForDelv = [b];
    for (var fi = 0; fi < flatDesc.length; fi++) allBatchesForDelv.push(flatDesc[fi].batch_id);
    for (var ab = 0; ab < allBatchesForDelv.length; ab++) {
      var dvIdxs = cache.dv.byBatch[allBatchesForDelv[ab]] || [];
      for (var dv = 0; dv < dvIdxs.length; dv++) {
        var rd = cache.dv.rows[dvIdxs[dv]];
        var dI = cache.dv.idx;
        deliveries.push({
          batch_id  : allBatchesForDelv[ab],
          sj_no     : String(rd[dI['SJ_No']] || ''),
          tgl       : rd[dI['Tanggal']],
          so_no     : String(rd[dI['SO_No']] || ''),
          cust      : String(rd[dI['Cust']] || ''),
          item_code : String(rd[dI['Item_Code']] || ''),
          qty       : parseFloat(rd[dI['Delv_Q']]) || 0,
          kg        : parseFloat(rd[dI['Delv_KG']]) || 0,
          driver    : String(rd[dI['Driver']] || ''),
          armada    : String(rd[dI['No_Armada']] || '')
        });
      }
    }

    // ---- META ----
    var rootBatch = '';
    if (ancestors.length > 0) rootBatch = ancestors[0].batch_id;
    else if (current.root_batch) rootBatch = current.root_batch;
    else rootBatch = current.batch_id;

    var supplier = '', noPo = '', noDo = '';
    if (ancestors.length > 0 && ancestors[0].source === 'GR') {
      supplier = ancestors[0].supplier;
      noPo     = ancestors[0].no_po;
      noDo     = ancestors[0].no_do;
    } else if (current.source === 'GR') {
      supplier = current.supplier; noPo = current.no_po; noDo = current.no_do;
    } else {
      supplier = current.supplier; noPo = current.no_po; noDo = current.no_do;
    }

    var totalDelvQty = 0, totalDelvKg = 0;
    for (var de = 0; de < deliveries.length; de++) {
      totalDelvQty += deliveries[de].qty;
      totalDelvKg  += deliveries[de].kg;
    }

    var meta = {
      root_batch        : rootBatch,
      supplier          : supplier,
      no_po             : noPo,
      no_do             : noDo,
      total_ancestors   : ancestors.length,
      total_descendants : flatDesc.length,
      total_deliveries  : deliveries.length,
      total_delv_qty    : totalDelvQty,
      total_delv_kg     : totalDelvKg,
      is_delivered      : deliveries.length > 0
    };

    // ---- SANITIZE DATES ----
    for (var fl = 0; fl < flat.length; fl++) _sanitizeDates(flat[fl]);
    tree = _sanitizeTree(tree);
    for (var dl = 0; dl < deliveries.length; dl++) _sanitizeDates(deliveries[dl]);

    return {
      success    : true,
      batch_id   : b,
      meta       : meta,
      tree       : tree,
      flat       : flat,
      deliveries : deliveries
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// PUBLIC API 2: getBatchMutasi(batchId)
//   Returns timeline kronologis (created, SPK start/done/cancel, ICT, delivery)
// =========================================================================
function getBatchMutasi(batchId) {
  try {
    if (!batchId) throw new Error('batchId kosong');
    var cache = _loadTraceData();
    var b = String(batchId).trim();
    var tz = Session.getScriptTimeZone();

    var events = [];

    // ---- EVENT 1: CREATED (GR atau Trace_Log) ----
    if (cache.gr.byBatch[b] !== undefined) {
      var rg = cache.gr.rows[cache.gr.byBatch[b]];
      var gI = cache.gr.idx;
      var iGrQty2 = gI['QTY_In'] !== undefined ? gI['QTY_In'] : gI['Qty_In'];
      events.push({
        tgl          : rg[gI['Tgl_Masuk']],
        event_type   : 'GR_IN',
        event_label  : 'Goods Receipt (Material Masuk)',
        source_sheet : 'GR',
        ref          : String(rg[gI['No_DO']] || ''),
        qty          : iGrQty2 !== undefined ? (parseFloat(rg[iGrQty2]) || 0) : 0,
        kg           : parseFloat(rg[gI['KG_In']]) || 0,
        note         : 'Supplier: ' + String(rg[gI['Supplier']] || '-') + ' | PO: ' + String(rg[gI['No_PO']] || '-') + ' | Loc: ' + String(rg[gI['Target_Loc']] || '-')
      });
    }
    if (cache.tl.byBatch[b] !== undefined) {
      var rt = cache.tl.rows[cache.tl.byBatch[b]];
      var tI = cache.tl.idx;
      var typ = String(rt[tI['Type']] || '');
      events.push({
        tgl          : rt[tI['Tgl_Buat']],
        event_type   : typ === 'NG' ? 'NG_CREATED' : 'PROD_OUTPUT',
        event_label  : typ === 'NG' ? 'NG Generated' : 'Production Output (' + typ + ')',
        source_sheet : 'Trace_Log',
        ref          : String(rt[tI['SPK_Ref']] || ''),
        qty          : parseFloat(rt[tI['Qty']]) || 0,
        kg           : parseFloat(rt[tI['KG']]) || 0,
        note         : 'Dari: ' + String(rt[tI['Source_Batch']] || '-') + ' | MC: ' + String(rt[tI['MC_No']] || '-') + ' | OP: ' + String(rt[tI['Operator']] || '-')
      });
    }

    // ---- EVENT 2: SPK lifecycle (saat batch ini dipakai sebagai source) ----
    var spkIdxs = cache.spk.byBatch[b] || [];
    var sI = cache.spk.idx;
    for (var si = 0; si < spkIdxs.length; si++) {
      var rs = cache.spk.rows[spkIdxs[si]];
      var spkType = String(rs[sI['SPK_Type']] || '');
      var spkNo   = String(rs[sI['SPK_No']] || '');
      var status  = String(rs[sI['Status']] || '').toUpperCase();
      var mulai   = rs[sI['Mulai_DT']];
      var selesai = rs[sI['Selesai_DT']];
      var qtyAct  = parseFloat(rs[sI['Qty_Actual']]) || 0;
      var kgAct   = parseFloat(rs[sI['KG_Actual']]) || 0;
      var mcNo    = String(rs[sI['MC_No']] || '');
      var op      = String(rs[sI['OP']] || '');
      var tglBuat = rs[sI['Tgl_Buat']];

      if (tglBuat instanceof Date) {
        events.push({
          tgl          : tglBuat,
          event_type   : 'SPK_CREATE',
          event_label  : 'SPK Dibuat (' + spkType + ')',
          source_sheet : 'SPK',
          ref          : spkNo,
          qty          : 0, kg: 0,
          note         : 'Mesin: ' + mcNo + ' | OP: ' + op + ' | Status: ' + status
        });
      }
      if (mulai instanceof Date) {
        events.push({
          tgl          : mulai,
          event_type   : 'SPK_START',
          event_label  : 'SPK Mulai (' + spkType + ')',
          source_sheet : 'SPK',
          ref          : spkNo,
          qty          : 0, kg: 0,
          note         : 'Mesin: ' + mcNo + ' | OP: ' + op
        });
      }
      if (selesai instanceof Date && status === 'DONE') {
        events.push({
          tgl          : selesai,
          event_type   : 'SPK_DONE',
          event_label  : 'SPK Selesai (' + spkType + ')',
          source_sheet : 'SPK',
          ref          : spkNo,
          qty          : qtyAct, kg: kgAct,
          note         : 'Mesin: ' + mcNo + ' | Hasil: ' + qtyAct + ' pcs / ' + kgAct + ' kg'
        });
      }
      if (status === 'CANCELLED') {
        events.push({
          tgl          : selesai instanceof Date ? selesai : tglBuat,
          event_type   : 'SPK_CANCEL',
          event_label  : 'SPK Dibatalkan (' + spkType + ')',
          source_sheet : 'SPK',
          ref          : spkNo,
          qty          : 0, kg: 0,
          note         : 'Mesin: ' + mcNo
        });
      }

      // ICT TRANSFER
      var ictIdxs = cache.ict.bySpkNo[spkNo] || [];
      var icI = cache.ict.idx;
      for (var ic = 0; ic < ictIdxs.length; ic++) {
        var rc = cache.ict.rows[ictIdxs[ic]];
        events.push({
          tgl          : rc[icI['Tgl_Transfer']],
          event_type   : 'ICT_TRANSFER',
          event_label  : 'Cross-Owner Transfer (FC↔DRC)',
          source_sheet : 'Rekap_ICT',
          ref          : spkNo,
          qty          : parseFloat(rc[icI['Qty_Sht']]) || 0,
          kg           : parseFloat(rc[icI['Qty_KG']]) || 0,
          note         : String(rc[icI['Dari_Owner']] || '') + ' → ' + String(rc[icI['Ke_Owner']] || '')
        });
      }
    }

    // ---- EVENT 3: DELIVERY ----
    var dvIdxs = cache.dv.byBatch[b] || [];
    var dI = cache.dv.idx;
    for (var di = 0; di < dvIdxs.length; di++) {
      var rd = cache.dv.rows[dvIdxs[di]];
      events.push({
        tgl          : rd[dI['Tanggal']],
        event_type   : 'DELV_SHIP',
        event_label  : 'Surat Jalan / Delivery',
        source_sheet : 'DELV',
        ref          : String(rd[dI['SJ_No']] || ''),
        qty          : parseFloat(rd[dI['Delv_Q']]) || 0,
        kg           : parseFloat(rd[dI['Delv_KG']]) || 0,
        note         : 'Cust: ' + String(rd[dI['Cust']] || '-') + ' | SO: ' + String(rd[dI['SO_No']] || '-') + ' | Armada: ' + String(rd[dI['No_Armada']] || '-')
      });
    }

    // ---- SORT BY DATE ASCENDING ----
    events.sort(function(a, b) {
      var ta = a.tgl instanceof Date ? a.tgl.getTime() : 0;
      var tb = b.tgl instanceof Date ? b.tgl.getTime() : 0;
      return ta - tb;
    });

    var out = [];
    for (var e = 0; e < events.length; e++) {
      var ev = events[e];
      if (ev.tgl instanceof Date) {
        ev.tgl_iso   = ev.tgl.toISOString();
        ev.tgl_label = Utilities.formatDate(ev.tgl, tz, 'dd MMM yy HH:mm');
      } else {
        ev.tgl_iso = ''; ev.tgl_label = '-';
      }
      delete ev.tgl;
      out.push(ev);
    }

    return {
      success  : true,
      batch_id : b,
      events   : out,
      total    : out.length
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// PUBLIC API 3: searchBatchByFilter(filter)
// =========================================================================
function searchBatchByFilter(filter) {
  try {
    filter = filter || {};
    var cache = _loadTraceData();
    var tz = Session.getScriptTimeZone();

    var fBatch = filter.batch_id ? String(filter.batch_id).trim().toUpperCase() : '';
    var fItem  = filter.item_code ? String(filter.item_code).trim().toUpperCase() : '';
    var fOwner = filter.owner ? String(filter.owner).trim().toUpperCase() : '';
    var fLoc   = filter.loc ? String(filter.loc).trim().toUpperCase() : '';
    var fFrom  = filter.date_from ? new Date(filter.date_from) : null;
    var fTo    = filter.date_to ? new Date(filter.date_to) : null;
    if (fTo) fTo.setHours(23, 59, 59, 999);
    var maxResults = parseInt(filter.limit) || 200;

    function matchLoc(type, targetLoc) {
      if (!fLoc) return true;
      var t  = String(type || '').toUpperCase();
      var tl = String(targetLoc || '').toUpperCase();
      if (fLoc === 'COIL'  && (t === 'GR' && tl === 'STOK_COIL')) return true;
      if (fLoc === 'SHEET' && (t === 'SHEET' || (t === 'GR' && tl === 'STOK_SHEET'))) return true;
      if (fLoc === 'WIP'   && t === 'WIP') return true;
      if (fLoc === 'FG'    && (t === 'FGC' || t === 'FGS' || t === 'FCL')) return true;
      if (fLoc === 'NG'    && t === 'NG') return true;
      return false;
    }

    function matchDate(tgl) {
      if (!fFrom && !fTo) return true;
      if (!(tgl instanceof Date)) return false;
      if (fFrom && tgl < fFrom) return false;
      if (fTo && tgl > fTo) return false;
      return true;
    }

    var results = [];

    // ---- GR batches ----
    var iGB = cache.gr.idx['Batch_ID'];
    var iGrQty3 = cache.gr.idx['QTY_In'] !== undefined ? cache.gr.idx['QTY_In'] : cache.gr.idx['Qty_In'];
    if (iGB !== undefined) {
      for (var i = 0; i < cache.gr.rows.length && results.length < maxResults; i++) {
        var r = cache.gr.rows[i];
        var gI = cache.gr.idx;
        var b = String(r[iGB] || '').trim();
        if (!b) continue;
        var item  = String(r[gI['Item_Code']] || '');
        var owner = String(r[gI['Owner']] || '');
        var tgl   = r[gI['Tgl_Masuk']];
        var tloc  = String(r[gI['Target_Loc']] || '');

        if (fBatch && b.toUpperCase().indexOf(fBatch) === -1) continue;
        if (fItem && item.toUpperCase().indexOf(fItem) === -1) continue;
        if (fOwner && owner.toUpperCase() !== fOwner) continue;
        if (!matchLoc('GR', tloc)) continue;
        if (!matchDate(tgl)) continue;

        results.push({
          batch_id    : b,
          source      : 'GR',
          level       : 0,
          type        : 'GR',
          tgl_iso     : tgl instanceof Date ? tgl.toISOString() : '',
          tgl_label   : tgl instanceof Date ? Utilities.formatDate(tgl, tz, 'dd MMM yy') : '-',
          item_code   : item,
          description : String(r[gI['Description']] || ''),
          spec        : String(r[gI['Spec']] || ''),
          t           : r[gI['T']] || '',
          p           : r[gI['P']] || '',
          l           : r[gI['L']] || '',
          qty         : iGrQty3 !== undefined ? (parseFloat(r[iGrQty3]) || 0) : 0,
          kg          : parseFloat(r[gI['KG_In']]) || 0,
          owner       : owner,
          loc         : tloc,
          supplier    : String(r[gI['Supplier']] || ''),
          no_po       : String(r[gI['No_PO']] || ''),
          no_coil     : String(r[gI['No_Coil']] || '')
        });
      }
    }

    // ---- Trace_Log batches ----
    var iTB = cache.tl.idx['Batch_ID'];
    if (iTB !== undefined) {
      for (var j = 0; j < cache.tl.rows.length && results.length < maxResults; j++) {
        var rt = cache.tl.rows[j];
        var tI = cache.tl.idx;
        var bb = String(rt[iTB] || '').trim();
        if (!bb) continue;
        var item2  = String(rt[tI['Item_Code']] || '');
        var owner2 = String(rt[tI['Owner']] || '');
        var tgl2   = rt[tI['Tgl_Buat']];
        var typ    = String(rt[tI['Type']] || '');

        if (fBatch && bb.toUpperCase().indexOf(fBatch) === -1) continue;
        if (fItem && item2.toUpperCase().indexOf(fItem) === -1) continue;
        if (fOwner && owner2.toUpperCase() !== fOwner) continue;
        if (!matchLoc(typ, '')) continue;
        if (!matchDate(tgl2)) continue;

        results.push({
          batch_id     : bb,
          source       : 'Trace_Log',
          level        : parseInt(rt[tI['Level']]) || 0,
          type         : typ,
          tgl_iso      : tgl2 instanceof Date ? tgl2.toISOString() : '',
          tgl_label    : tgl2 instanceof Date ? Utilities.formatDate(tgl2, tz, 'dd MMM yy') : '-',
          item_code    : item2,
          description  : String(rt[tI['Description']] || ''),
          spec         : String(rt[tI['Spec']] || ''),
          t            : rt[tI['T']] || '',
          p            : rt[tI['P']] || '',
          l            : rt[tI['L_dim']] || '',
          qty          : parseFloat(rt[tI['Qty']]) || 0,
          kg           : parseFloat(rt[tI['KG']]) || 0,
          owner        : owner2,
          source_batch : String(rt[tI['Source_Batch']] || ''),
          root_batch   : String(rt[tI['Root_Batch']] || ''),
          spk_ref      : String(rt[tI['SPK_Ref']] || '')
        });
      }
    }

    results.sort(function(a, b) { return (b.tgl_iso || '').localeCompare(a.tgl_iso || ''); });

    return {
      success   : true,
      total     : results.length,
      truncated : results.length >= maxResults,
      results   : results
    };

  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// HELPER (BARU): Detect batch type — COIL/SHEET/WIP/FG/NG
// =========================================================================
function _detectBatchType(batchId, resolved) {
  if (resolved && resolved.source === 'GR') return 'COIL';

  var type = resolved ? String(resolved.type || '').toUpperCase() : '';
  if (type === 'SHEET') return 'SHEET';
  if (type === 'WIP')   return 'WIP';
  if (type === 'FGC' || type === 'FGS' || type === 'FCL' || type === 'FSL' || type === 'FSH') return 'FG';
  if (type === 'NG')    return 'NG';

  // Fallback: prefix detection dari Batch_ID
  var bid = String(batchId).toUpperCase();
  if (bid.match(/^GR/))              return 'COIL';
  if (bid.indexOf('SHT-') === 0)     return 'SHEET';
  if (bid.indexOf('WIP-') === 0)     return 'WIP';
  if (bid.indexOf('FC-') === 0)      return 'FG';
  if (bid.indexOf('NG-') === 0 || bid.indexOf('-NG') >= 0) return 'NG';

  return 'UNKNOWN';
}

// =========================================================================
// HELPER (BARU): Visibility per batch type
//   Tentukan posisi mana yang RELEVANT untuk ditampilkan
// =========================================================================
function _getVisibilityForType(batchType) {
  switch (batchType) {
    case 'COIL':
      return { STOK_COIL:true, STOK_SHEET:true, STOK_WIP:true, STOK_FG:true, STOK_NG:true, DELIVERED:true };
    case 'SHEET':
      return { STOK_COIL:false, STOK_SHEET:true, STOK_WIP:true, STOK_FG:true, STOK_NG:true, DELIVERED:true };
    case 'WIP':
      return { STOK_COIL:false, STOK_SHEET:false, STOK_WIP:true, STOK_FG:true, STOK_NG:true, DELIVERED:true };
    case 'FG':
      return { STOK_COIL:false, STOK_SHEET:false, STOK_WIP:false, STOK_FG:true, STOK_NG:false, DELIVERED:true };
    case 'NG':
      return { STOK_COIL:false, STOK_SHEET:false, STOK_WIP:false, STOK_FG:false, STOK_NG:true, DELIVERED:true };
    default:
      return { STOK_COIL:true, STOK_SHEET:true, STOK_WIP:true, STOK_FG:true, STOK_NG:true, DELIVERED:true };
  }
}

// =========================================================================
// HELPER (BARU): Derive status batch dari SPK
//   Cek apakah batch sedang jadi source di SPK aktif
// =========================================================================
function _deriveBatchStatus(batchId, cache) {
  var spkIdxs = cache.spk.byBatch[batchId] || [];
  if (spkIdxs.length === 0) return { status: 'Available', ref: '' };

  var sI = cache.spk.idx;
  // Prioritas: RUNNING > ANTRIAN > Available
  var runningSpk = null, antrianSpk = null;
  for (var i = 0; i < spkIdxs.length; i++) {
    var rs = cache.spk.rows[spkIdxs[i]];
    var status = String(rs[sI['Status']] || '').toUpperCase();
    var spkNo  = String(rs[sI['SPK_No']] || '');
    var mcNo   = String(rs[sI['MC_No']] || '');

    if (status === 'RUNNING') { runningSpk = { spkNo: spkNo, mcNo: mcNo }; break; }
    if (status === 'ANTRIAN' || status === 'OPEN') {
      if (!antrianSpk) antrianSpk = { spkNo: spkNo, mcNo: mcNo };
    }
  }

  if (runningSpk) {
    return { status: 'Running di ' + (runningSpk.mcNo || runningSpk.spkNo), ref: runningSpk.spkNo };
  }
  if (antrianSpk) {
    return { status: 'Reserved (' + antrianSpk.spkNo + ')', ref: antrianSpk.spkNo };
  }
  return { status: 'Available', ref: '' };
}

// =========================================================================
// HELPER (BARU): Cek alert vs threshold
// =========================================================================
function _checkAlert(metricName, actualValue, cache) {
  var idx = cache.thr.byMetric[metricName];
  if (idx === undefined) {
    return {
      metric: metricName, description: metricName,
      actual: actualValue, threshold_warning: null, threshold_alert: null,
      unit: '', level: 'OK', message: 'Threshold belum dikonfigurasi'
    };
  }
  var row = cache.thr.rows[idx];
  var tI  = cache.thr.idx;

  var active = String(row[tI['Active']] || '').toUpperCase() === 'TRUE';
  if (!active) {
    return {
      metric: metricName, description: String(row[tI['Description']] || metricName),
      actual: actualValue, threshold_warning: null, threshold_alert: null,
      unit: String(row[tI['Unit']] || ''),
      level: 'OK', message: 'Metric tidak aktif'
    };
  }

  var direction = String(row[tI['Direction']] || '>').trim();
  var warn      = parseFloat(row[tI['Warning_Threshold']]) || 0;
  var alert     = parseFloat(row[tI['Alert_Threshold']])   || 0;
  var unit      = String(row[tI['Unit']] || '');
  var desc      = String(row[tI['Description']] || metricName);
  var actual    = parseFloat(actualValue) || 0;

  var level = 'OK';
  if (direction === '>') {
    if (actual > alert)      level = 'CRITICAL';
    else if (actual > warn)  level = 'WARNING';
  } else if (direction === '<') {
    if (actual < alert)      level = 'CRITICAL';
    else if (actual < warn)  level = 'WARNING';
  }

  var actualStr = actual.toFixed(2) + (unit === 'percent' ? '%' : ' ' + unit);
  var levelMsg = level === 'OK' ? 'Dalam batas normal' :
                 level === 'WARNING' ? 'Perlu perhatian' : 'Kritis — investigasi sekarang';

  return {
    metric: metricName, description: desc,
    actual: actual, threshold_warning: warn, threshold_alert: alert,
    unit: unit, level: level,
    message: levelMsg + ' (actual: ' + actualStr + ')'
  };
}

// =========================================================================
// PUBLIC API 4 (BARU): getBatchMassBalance(batchId)
//   Hitung mass balance untuk 1 batch (semua level).
//   Auto-detect batch level dari Batch_ID prefix.
// =========================================================================
function getBatchMassBalance(batchId) {
  try {
    if (!batchId) throw new Error('batchId kosong');
    var startTs = new Date().getTime();

    var cache = _loadTraceData();
    var b = String(batchId).trim();

    // ---- 1. Resolve current batch ----
    var current = _resolveBatch(b, cache);
    if (!current.found) {
      return { success: false, message: 'Batch tidak ditemukan: ' + b };
    }

    var batchType = _detectBatchType(b, current);
    var visibility = _getVisibilityForType(batchType);

    // ---- 2. Get parent info (ancestors[0] kalau ada) ----
    var parentBatchId = current.source_batch || null;
    var parentDescription = null;
    if (parentBatchId) {
      var parentResolved = _resolveBatch(parentBatchId, cache);
      if (parentResolved.found) {
        parentDescription = parentBatchId + (parentResolved.description ? ' — ' + parentResolved.description : '');
      }
    }

    // ---- 3. Determine root batch ----
    var rootBatchId = current.root_batch || b;

    // ---- 4. Get all descendants (BFS) ----
    var descendantIds = [];
    var queue = [{ batch: b, level: 0 }];
    var seen = {}; seen[b] = true;
    var safety = 0;
    while (queue.length > 0 && safety < 5000) {
      var item = queue.shift();
      var childIdxs = cache.tl.bySrc[item.batch] || [];
      for (var k = 0; k < childIdxs.length; k++) {
        var cr = cache.tl.rows[childIdxs[k]];
        var childBatch = String(cr[cache.tl.idx['Batch_ID']] || '').trim();
        if (!childBatch || seen[childBatch]) continue;
        seen[childBatch] = true;
        descendantIds.push(childBatch);
        queue.push({ batch: childBatch, level: item.level + 1 });
      }
      safety++;
    }

    // ---- 5. allBatchIds = self + descendants ----
    var allBatchIds = [b].concat(descendantIds);

    // ---- 6. Collect positions ----
    var rawPositions = {
      STOK_COIL:  { kg:0, qty:0, details:[] },
      STOK_SHEET: { kg:0, qty:0, details:[] },
      STOK_WIP:   { kg:0, qty:0, details:[] },
      STOK_FG:    { kg:0, qty:0, details:[] },
      STOK_NG:    { kg:0, qty:0, details:[] },
      DELIVERED:  { kg:0, qty:0, details:[] }
    };

    function lookupStock(cacheKey, posKey, sheetLabel, qtyUnit) {
      var sh = cache[cacheKey];
      if (!sh || !sh.byBatch) return;

      // 🟢 Fix Mass Balance — pakai FISIK (Qty_In - Qty_Done/Delv), bukan Qty_Avail
      // Karena Qty_Avail sudah dikurangi Keep+Prod → miss material yg sedang di-keep tapi belum jadi output baru
      var iIn_Q  = sh.idx['Qty_In'];
      var iIn_K  = sh.idx['KG_In']  !== undefined ? sh.idx['KG_In']  : sh.idx['Kg_In'];

      // FG pakai Delv (delivered ke customer), sheet lain pakai Done (output sudah lahir jadi batch baru)
      var iOut_Q, iOut_K;
      if (cacheKey === 'fg') {
        iOut_Q = sh.idx['Qty_Delv'];
        iOut_K = sh.idx['KG_Delv'] !== undefined ? sh.idx['KG_Delv'] : sh.idx['Kg_Delv'];
      } else {
        iOut_Q = sh.idx['Qty_Done'];
        iOut_K = sh.idx['KG_Done'] !== undefined ? sh.idx['KG_Done'] : sh.idx['Kg_Done'];
      }

      var iItm  = sh.idx['Item_Code'];
      var iDesc = sh.idx['Description'];
      var iOwn  = sh.idx['Owner'];

      // Breakdown fields (untuk detail display kalau frontend butuh)
      var iQA   = sh.idx['Qty_Avail'];
      var iKA   = sh.idx['KG_Avail'] !== undefined ? sh.idx['KG_Avail'] : sh.idx['Kg_Avail'];
      var iQK   = sh.idx['Qty_Keep'];
      var iKK   = sh.idx['KG_Keep']  !== undefined ? sh.idx['KG_Keep']  : sh.idx['Kg_Keep'];
      var iQP   = sh.idx['Qty_Prod'];
      var iKP   = sh.idx['KG_Prod']  !== undefined ? sh.idx['KG_Prod']  : sh.idx['Kg_Prod'];

      for (var i = 0; i < allBatchIds.length; i++) {
        var bid = allBatchIds[i];
        var rowIdx = sh.byBatch[bid];
        if (rowIdx === undefined) continue;
        var row = sh.rows[rowIdx];

        var qtyIn  = iIn_Q  !== undefined ? (parseFloat(row[iIn_Q])  || 0) : 0;
        var kgIn   = iIn_K  !== undefined ? (parseFloat(row[iIn_K])  || 0) : 0;
        var qtyOut = iOut_Q !== undefined ? (parseFloat(row[iOut_Q]) || 0) : 0;
        var kgOut  = iOut_K !== undefined ? (parseFloat(row[iOut_K]) || 0) : 0;

        // 🟢 FISIK = In - Out (Done untuk Coil/Sheet/WIP/NG, Delv untuk FG)
        var qty = qtyIn - qtyOut;
        var kg  = kgIn  - kgOut;

        // Skip batch dengan fisik = 0 di detail (tapi tetap dihitung di total)
        if (qty <= 0 && kg <= 0) continue;

        rawPositions[posKey].kg  += kg;
        rawPositions[posKey].qty += qty;

        var statusInfo = _deriveBatchStatus(bid, cache);

        // 🟢 Optional breakdown untuk detail view
        var breakdown = {
          avail_qty: iQA !== undefined ? (parseFloat(row[iQA]) || 0) : 0,
          avail_kg : iKA !== undefined ? (parseFloat(row[iKA]) || 0) : 0,
          keep_qty : iQK !== undefined ? (parseFloat(row[iQK]) || 0) : 0,
          keep_kg  : iKK !== undefined ? (parseFloat(row[iKK]) || 0) : 0,
          prod_qty : iQP !== undefined ? (parseFloat(row[iQP]) || 0) : 0,
          prod_kg  : iKP !== undefined ? (parseFloat(row[iKP]) || 0) : 0,
          done_qty : qtyOut,
          done_kg  : kgOut
        };

        rawPositions[posKey].details.push({
          batch_id    : bid,
          description : iDesc !== undefined ? String(row[iDesc] || '') : '',
          item_code   : iItm  !== undefined ? String(row[iItm]  || '') : '',
          owner       : iOwn  !== undefined ? String(row[iOwn]  || '') : '',
          qty         : qty,           // fisik
          qty_unit    : qtyUnit,
          kg          : kg,            // fisik
          kg_unit     : 'Kg',
          lokasi      : sheetLabel,
          status      : statusInfo.status,
          breakdown   : breakdown       // 🟢 tersedia kalau frontend butuh (Keep, Prod, Avail, Done)
        });
      }
    }

    lookupStock('coil', 'STOK_COIL',  'Stok_Coil',  'Kg');   // Coil unit = Kg
    lookupStock('sht',  'STOK_SHEET', 'Stok_Sheet', 'Sht');
    lookupStock('wip',  'STOK_WIP',   'Stok_WIP',   'Sht');
    lookupStock('fg',   'STOK_FG',    'Stok_FG',    'Sht');
    lookupStock('ng',   'STOK_NG',    'Stok_NG',    'Sht');

    // Delivered lookup
    var dI = cache.dv.idx;
    var tz = Session.getScriptTimeZone();
    for (var ad = 0; ad < allBatchIds.length; ad++) {
      var dbid = allBatchIds[ad];
      var dvIdxs = cache.dv.byBatch[dbid] || [];
      for (var dv = 0; dv < dvIdxs.length; dv++) {
        var rd = cache.dv.rows[dvIdxs[dv]];
        var dQty = parseFloat(rd[dI['Delv_Q']])  || 0;
        var dKg  = parseFloat(rd[dI['Delv_KG']]) || 0;
        var dTgl = rd[dI['Tanggal']];

        rawPositions.DELIVERED.kg  += dKg;
        rawPositions.DELIVERED.qty += dQty;

        rawPositions.DELIVERED.details.push({
          batch_id    : dbid,
          sj_no       : String(rd[dI['SJ_No']]  || ''),
          tgl         : dTgl instanceof Date ? dTgl.toISOString() : '',
          tgl_label   : dTgl instanceof Date ? Utilities.formatDate(dTgl, tz, 'dd MMM yy') : '-',
          cust        : String(rd[dI['Cust']]   || ''),
          so_no       : String(rd[dI['SO_No']]  || ''),
          description : dI['Description'] !== undefined ? String(rd[dI['Description']] || '') : '',
          item_code   : String(rd[dI['Item_Code']] || ''),
          qty         : dQty,
          qty_unit    : 'Sht',
          kg          : dKg,
          kg_unit     : 'Kg',
          lokasi      : 'Dikirim ke Customer',
          status      : 'Shipped',
          armada      : String(rd[dI['No_Armada']] || ''),
          driver      : String(rd[dI['Driver']]    || '')
        });
      }
    }

    // ---- 7. Build positions array (urutan fixed) ----
    var initialKg  = current.kg  || 0;
    var initialQty = current.qty || 0;

    var posMeta = [
      { key:'STOK_COIL',  label:'Stok Coil',         icon:'🟦' },
      { key:'STOK_SHEET', label:'Stok Sheet',        icon:'🟩' },
      { key:'STOK_WIP',   label:'Stok WIP',          icon:'🟪' },
      { key:'STOK_FG',    label:'Stok FG',           icon:'🟧' },
      { key:'STOK_NG',    label:'Stok NG',           icon:'🟥' },
      { key:'DELIVERED',  label:'Sudah Delivered',   icon:'🚚' }
    ];

    var positions = [];
    var totalTercatatKg = 0;
    var totalTercatatQty = 0;
    for (var pi = 0; pi < posMeta.length; pi++) {
      var pm = posMeta[pi];
      var rp = rawPositions[pm.key];
      var pct = initialKg > 0 ? (rp.kg / initialKg * 100) : 0;
      positions.push({
        key     : pm.key,
        label   : pm.label,
        icon    : pm.icon,
        kg      : rp.kg,
        qty     : rp.qty,
        pct     : Math.round(pct * 100) / 100,
        visible : !!visibility[pm.key],
        details : rp.details
      });
      totalTercatatKg += rp.kg;
      totalTercatatQty += rp.qty;
    }

    // ---- 8. Summary ----
    var ngKg = rawPositions.STOK_NG.kg;
    var delvKg = rawPositions.DELIVERED.kg;
    var selisihKg = totalTercatatKg - initialKg;
    var selisihPct = initialKg > 0 ? (selisihKg / initialKg * 100) : 0;

    // Yield = (tercatat - NG) / initial * 100
    var yieldPct = 0;
    if (initialKg > 0) {
      yieldPct = ((totalTercatatKg - ngKg) / initialKg) * 100;
    }

    var ngRatePct = initialKg > 0 ? (ngKg / initialKg * 100) : 0;
    var delveredPct = initialKg > 0 ? (delvKg / initialKg * 100) : 0;
    var inhousePct  = initialKg > 0 ? ((totalTercatatKg - delvKg) / initialKg * 100) : 0;

    var summary = {
      total_tercatat_kg  : Math.round(totalTercatatKg * 100) / 100,
      total_tercatat_qty : totalTercatatQty,
      total_tercatat_pct : Math.round(totalTercatatKg / (initialKg || 1) * 10000) / 100,
      selisih_kg         : Math.round(selisihKg * 100) / 100,
      selisih_pct        : Math.round(selisihPct * 100) / 100,
      yield_pct          : Math.round(yieldPct * 100) / 100,
      ng_rate_pct        : Math.round(ngRatePct * 100) / 100,
      delivered_pct      : Math.round(delveredPct * 100) / 100,
      inhouse_pct        : Math.round(inhousePct * 100) / 100
    };

    // ---- 9. Alerts ----
    var alerts = [];
    // Mass Balance — pakai absolute(selisih_pct) karena bisa minus
    alerts.push(_checkAlert('MASS_BALANCE_PCT', Math.abs(selisihPct), cache));
    alerts.push(_checkAlert('NG_RATE_PCT',      ngRatePct,            cache));

    // Yield — pakai metric sesuai batch type
    if (batchType === 'COIL') {
      alerts.push(_checkAlert('YIELD_COIL_PCT',  yieldPct, cache));
    } else if (batchType === 'SHEET') {
      alerts.push(_checkAlert('YIELD_SHEET_PCT', yieldPct, cache));
    } else if (batchType === 'WIP') {
      alerts.push(_checkAlert('YIELD_WIP_PCT',   yieldPct, cache));
    }

    // ---- 10. Overall status (terparah dari alerts) ----
    var overallStatus = 'OK';
    var overallMsg = 'Semua metric dalam batas normal';
    for (var al = 0; al < alerts.length; al++) {
      if (alerts[al].level === 'CRITICAL') { overallStatus = 'CRITICAL'; overallMsg = 'Ada metric kritis — perlu investigasi'; break; }
      if (alerts[al].level === 'WARNING')  { overallStatus = 'WARNING';  overallMsg = 'Ada metric warning — perhatikan'; }
    }

    // ---- 11. Source info (Supplier, PO, DO) ----
    var supplier = '', noPo = '', noDo = '';
    if (batchType === 'COIL') {
      supplier = current.supplier;
      noPo     = current.no_po;
      noDo     = current.no_do;
    } else {
      // Untuk sub-batch, ambil dari root coil
      var rootResolved = _resolveBatch(rootBatchId, cache);
      if (rootResolved.found && rootResolved.source === 'GR') {
        supplier = rootResolved.supplier;
        noPo     = rootResolved.no_po;
        noDo     = rootResolved.no_do;
      } else {
        supplier = current.supplier;
        noPo     = current.no_po;
        noDo     = current.no_do;
      }
    }

    // ---- 12. Format tgl_in ----
    var tglIn = current.tgl_buat;
    var tglInLabel = '-';
    if (tglIn instanceof Date) {
      tglInLabel = Utilities.formatDate(tglIn, tz, 'dd MMM yyyy');
    }

    var endTs = new Date().getTime();

    return {
      success         : true,
      batch_id        : b,
      batch_type      : batchType,
      batch_level     : current.level || 0,
      item_code       : current.item_code || '',
      description     : current.description || '',
      spec            : current.spec || '',
      t               : current.t || null,
      p               : current.p || null,
      l               : current.l || null,
      supplier        : supplier,
      no_po           : noPo,
      no_do           : noDo,
      tgl_in          : tglInLabel,
      owner           : current.owner || '',
      parent_batch_id : parentBatchId,
      parent_description : parentDescription,
      root_batch_id   : rootBatchId,
      initial_kg      : Math.round(initialKg * 100) / 100,
      initial_qty     : initialQty,
      initial_qty_unit: batchType === 'COIL' ? 'Kg' : 'Sht',
      initial_kg_unit : 'Kg',
      positions       : positions,
      summary         : summary,
      alerts          : alerts,
      overall_status  : overallStatus,
      overall_message : overallMsg,
      total_descendants : descendantIds.length,
      generated_at    : new Date().toISOString(),
      query_duration_ms : endTs - startTs
    };

  } catch (e) {
    return { success: false, message: e.message || String(e), stack: e.stack || '' };
  }
}

// =========================================================================
// PUBLIC API 5 (BARU): getThresholds()
//   Ambil semua threshold dari M_Threshold (untuk display di UI)
// =========================================================================
function getThresholds() {
  try {
    var cache = _loadTraceData();
    var tI = cache.thr.idx;
    var result = [];
    for (var i = 0; i < cache.thr.rows.length; i++) {
      var r = cache.thr.rows[i];
      var name = String(r[tI['Metric_Name']] || '').trim();
      if (!name) continue;
      result.push({
        metric_name       : name,
        description       : String(r[tI['Description']] || ''),
        direction         : String(r[tI['Direction']] || '>'),
        warning_threshold : parseFloat(r[tI['Warning_Threshold']]) || 0,
        alert_threshold   : parseFloat(r[tI['Alert_Threshold']]) || 0,
        unit              : String(r[tI['Unit']] || ''),
        active            : String(r[tI['Active']] || '').toUpperCase() === 'TRUE'
      });
    }
    return { success: true, total: result.length, thresholds: result };
  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

// =========================================================================
// UTILITY (D4): setupNewSheets_T2()
// =========================================================================
function setupNewSheets_T2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var schemas = {
    'Trace_Parents'    : ['Trace_ID', 'Tgl', 'Child_Batch', 'Parent_Batch', 'Qty_Consumed', 'KG_Consumed', 'Note'],
    'Stock_Movement'   : ['Mov_ID', 'Tgl', 'Batch_ID', 'Item_Code', 'From_Loc', 'To_Loc', 'Qty', 'KG', 'Reason', 'Created_By', 'Note'],
    'Stock_Adjustment' : ['Adj_ID', 'Tgl', 'Batch_ID', 'Loc', 'Item_Code', 'Qty_Before', 'Qty_After', 'KG_Before', 'KG_After', 'Reason', 'Approved_By', 'Note'],
    'Repack_Log'       : ['Repack_ID', 'Tgl', 'Source_Batches', 'Child_Batches', 'Item_Code', 'Reason', 'Created_By', 'Note'],
    'Return_CS'        : ['Ret_ID', 'Tgl', 'SJ_Ref', 'SO_Ref', 'Cust', 'Batch_ID', 'Item_Code', 'Qty', 'KG', 'Reason', 'Action', 'Created_By', 'Note'],
    'Return_PO'        : ['Ret_ID', 'Tgl', 'GR_Ref', 'PO_Ref', 'Vendor', 'Batch_ID', 'Item_Code', 'Qty', 'KG', 'Reason', 'Created_By', 'Note'],
    'Opname_Bulanan'   : ['Period', 'Batch_ID', 'Loc', 'Item_Code', 'Qty_Snapshot', 'KG_Snapshot', 'Locked_DT', 'Locked_By'],
    'M_ITEM_VENDOR'    : ['Item_Code', 'Vendor_Code', 'Lead_Time_Days', 'MOQ', 'Multiples', 'Is_Default', 'Note'],
    'M_Forecast_Param' : ['Item_Code', 'Method', 'Alpha', 'Period_Lookback', 'Seasonal_Factor', 'Note'],
    'Forecast'         : ['Period', 'Item_Code', 'Forecast_Qty', 'Forecast_KG', 'Method', 'Generated_DT'],
    'MRP_Recommend'    : ['Rec_ID', 'Generated_DT', 'Item_Code', 'Current_Stock', 'Shortage_Qty', 'Recommended_Qty', 'Required_Date', 'Suggested_Vendor', 'Status', 'Approved_By', 'Approved_DT', 'Note']
  };

  var created = [], skipped = [];
  for (var name in schemas) {
    if (ss.getSheetByName(name)) {
      skipped.push(name);
      continue;
    }
    var sh = ss.insertSheet(name);
    var cols = schemas[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#f1f5f9');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, cols.length);
    created.push(name);
  }

  var msg = 'Setup selesai. Created: ' + created.length + ' sheet, Skipped: ' + skipped.length + '.\n\n' +
            (created.length > 0 ? 'Created:\n  - ' + created.join('\n  - ') + '\n\n' : '') +
            (skipped.length > 0 ? 'Skipped (sudah ada):\n  - ' + skipped.join('\n  - ') : '');

  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
  return { success: true, created: created, skipped: skipped };
}

// =========================================================================
// TEST FUNCTIONS
// =========================================================================
function _test_getBatchTrace() {
  var batchId = 'GR260013-A';
  var result = getBatchTrace(batchId);
  Logger.log('Success: ' + result.success);
  Logger.log('Meta: ' + JSON.stringify(result.meta, null, 2));
}

function _test_getBatchMutasi() {
  var batchId = 'GR260013-A';
  var result = getBatchMutasi(batchId);
  Logger.log('Total events: ' + result.total);
  if (result.events) {
    result.events.forEach(function(e, i) {
      Logger.log((i+1) + '. [' + e.tgl_label + '] ' + e.event_label + ' | ' + e.qty + ' / ' + e.kg + ' kg');
    });
  }
}

function _test_searchBatchByFilter() {
  var result = searchBatchByFilter({ owner: 'DRC', loc: 'COIL', limit: 10 });
  Logger.log('Total: ' + result.total);
}

// ⬇️ TEST BARU
function _test_getBatchMassBalance() {
  var batchId = 'GR260013-A'; // ⚠️ Ganti ke batch yang ada di data lo
  var result = getBatchMassBalance(batchId);
  Logger.log('Success: ' + result.success);
  if (!result.success) { Logger.log('Error: ' + result.message); return; }
  Logger.log('Batch: ' + result.batch_id + ' (' + result.batch_type + ')');
  Logger.log('Description: ' + result.description);
  Logger.log('Initial: ' + result.initial_kg + ' Kg / ' + result.initial_qty + ' ' + result.initial_qty_unit);
  Logger.log('Total descendants: ' + result.total_descendants);
  Logger.log('Total Tercatat: ' + result.summary.total_tercatat_kg + ' kg (' + result.summary.total_tercatat_pct + '%)');
  Logger.log('Selisih: ' + result.summary.selisih_kg + ' kg (' + result.summary.selisih_pct + '%)');
  Logger.log('Yield: ' + result.summary.yield_pct + '%');
  Logger.log('NG Rate: ' + result.summary.ng_rate_pct + '%');
  Logger.log('Overall Status: ' + result.overall_status + ' — ' + result.overall_message);
  Logger.log('--- POSITIONS ---');
  for (var i = 0; i < result.positions.length; i++) {
    var p = result.positions[i];
    if (!p.visible && p.kg === 0) continue;
    Logger.log('  ' + p.icon + ' ' + p.label + ': ' + p.qty + ' ' + (p.details[0] ? p.details[0].qty_unit : 'Sht') + ' / ' + p.kg + ' kg (' + p.pct + '%) — ' + p.details.length + ' detail batch');
  }
  Logger.log('--- ALERTS ---');
  for (var a = 0; a < result.alerts.length; a++) {
    Logger.log('  ' + result.alerts[a].level + ' | ' + result.alerts[a].description + ': ' + result.alerts[a].message);
  }
  Logger.log('Duration: ' + result.query_duration_ms + ' ms');
}

function _test_getThresholds() {
  var result = getThresholds();
  Logger.log('Success: ' + result.success);
  if (result.success) {
    Logger.log('Total: ' + result.total);
    result.thresholds.forEach(function(t) {
      Logger.log('  ' + t.metric_name + ' | ' + t.direction + ' ' + t.warning_threshold + '/' + t.alert_threshold + ' ' + t.unit + ' | Active: ' + t.active);
    });
  }
}

function debugMassBalanceWIP() {
  const batchId = 'GR260013-A'; // isi batch id yang dipakai Mass Balance
  
  const cache = _loadTraceData();
  
  // 1. Cek batch di Trace_Log
  const b = String(batchId).trim();
  const current = _resolveBatch(b, cache);
  Logger.log('=== 1. RESOLVE BATCH ===');
  Logger.log('found: ' + current.found);
  Logger.log('type: ' + current.type);
  Logger.log('kg: ' + current.kg);
  Logger.log('root_batch: ' + current.root_batch);
  
  // 2. Detect type
  const batchType = _detectBatchType(b, current);
  Logger.log('=== 2. DETECT TYPE ===');
  Logger.log('batchType: ' + batchType);
  
  // 3. Visibility
  const visibility = _getVisibilityForType(batchType);
  Logger.log('=== 3. VISIBILITY ===');
  Logger.log(JSON.stringify(visibility));
  
  // 4. BFS descendants
  var descendantIds = [];
  var queue = [{ batch: b, level: 0 }];
  var seen = {}; seen[b] = true;
  var safety = 0;
  while (queue.length > 0 && safety < 5000) {
    var item = queue.shift();
    var childIdxs = cache.tl.bySrc[item.batch] || [];
    Logger.log('  BFS ' + item.batch + ' → children: ' + JSON.stringify(childIdxs.map(function(i){ 
      return String(cache.tl.rows[i][cache.tl.idx['Batch_ID']] || '').trim();
    })));
    for (var k = 0; k < childIdxs.length; k++) {
      var cr = cache.tl.rows[childIdxs[k]];
      var childBatch = String(cr[cache.tl.idx['Batch_ID']] || '').trim();
      if (!childBatch || seen[childBatch]) continue;
      seen[childBatch] = true;
      descendantIds.push(childBatch);
      queue.push({ batch: childBatch, level: item.level + 1 });
    }
    safety++;
  }
  const allBatchIds = [b].concat(descendantIds);
  Logger.log('=== 4. ALL BATCH IDS ===');
  Logger.log(JSON.stringify(allBatchIds));
  
  // 5. Check Stok_WIP
  Logger.log('=== 5. Stok_WIP header ===');
  Logger.log('idx: ' + JSON.stringify(Object.keys(cache.wip.idx)));
  Logger.log('byBatch keys: ' + JSON.stringify(Object.keys(cache.wip.byBatch)));
  
  // 6. Cari batch WIP di Stok_WIP
  const iQA = cache.wip.idx['Qty_Avail'];
  const iKA = cache.wip.idx['KG_Avail'] !== undefined ? cache.wip.idx['KG_Avail'] : cache.wip.idx['Kg_Avail'];
  Logger.log('=== 6. Column indexes ===');
  Logger.log('iQA (Qty_Avail): ' + iQA);
  Logger.log('iKA (KG_Avail): ' + iKA);
  
  Logger.log('=== 7. Match check ===');
  for (var i = 0; i < allBatchIds.length; i++) {
    var bid = allBatchIds[i];
    var rowIdx = cache.wip.byBatch[bid];
    if (rowIdx === undefined) {
      Logger.log('  ' + bid + ' → NOT FOUND in cache.wip.byBatch');
    } else {
      var row = cache.wip.rows[rowIdx];
      var qty = iQA !== undefined ? (parseFloat(row[iQA]) || 0) : 0;
      var kg  = iKA !== undefined ? (parseFloat(row[iKA]) || 0) : 0;
      Logger.log('  ' + bid + ' → row ' + rowIdx + ', qty=' + qty + ', kg=' + kg);
    }
  }
}