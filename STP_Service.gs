/* =========================================================================
 * STP_SERVICE.GS — Stamping Production Request Module
 *
 * Modul input Request Stamping dari admin Stamping ke sheet STP_REQ.
 * Sheet STP_REQ menggunakan formula MAP/XLOOKUP untuk auto-fill kolom:
 *   Description, Spec, T, P, L, UoM, KG_Req, Qty_Fulfill, Qty_Sisa, Status
 * GAS HANYA menulis kolom input murni.
 *
 * Namespace: semua public function prefix `stp*`, private prefix `_stp_*`.
 *
 * PUBLIC FUNCTIONS:
 *   stpGetInitData()              → {items, customers} untuk form dropdown
 *   stpGetList(filter)            → grouped list dgn filter status/periode/search
 *   stpSaveNew(payload)           → create new STP (multi-item)
 *   stpUpdateHeader(payload)      → edit Cust/Periode/Schedule/Owner/Priority
 *   stpUpdateItem(payload)        → edit Qty_Req / NOTE / Ref_SPK per row
 *   stpAddItem(payload)           → tambah item baru ke STP existing
 *   stpDeleteItem(payload)        → hapus 1 row item (guard: belum ada SPK)
 *   stpCancelAll(payload)         → hapus semua row STP (guard: semua item belum SPK)
 *
 * GUARD RULES:
 *   - Edit Qty_Req / Delete item: block kalau ada SPK dgn SO_Ref=STP_No AND
 *     Item_Code=item (status ≠ CANCELLED).
 *   - Cancel STP: block kalau ada SPK linked apapun (item mana saja).
 *   - Edit header, edit NOTE, edit Ref_SPK: selalu boleh.
 * ========================================================================= */

// ===== KONSTAN =====
var STP_SHEET_NAME = 'STP_REQ';

// =========================================================================
// 1. INIT DATA — master untuk form dropdown
// =========================================================================
function stpGetInitData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var items = [], customers = [];

    // ─── M_ITEM ───
    var itemSheet = ss.getSheetByName('M_ITEM');
    if (itemSheet) {
      var iData = itemSheet.getDataRange().getValues();
      var iH = iData[0].map(function(h){ return String(h).trim(); });
      var iIC   = iH.indexOf('Item_Code');
      var iDesc = iH.indexOf('Description');
      var iSpec = iH.indexOf('Spec');
      var iT    = iH.indexOf('T');
      var iP    = iH.indexOf('P');
      var iL    = iH.indexOf('L');
      var iUoM  = iH.indexOf('Unit of Measure');
      if (iUoM < 0) iUoM = iH.indexOf('UoM_MC');
      var iWg   = iH.indexOf('Wg/Pce Cust');
      if (iWg < 0) iWg = iH.indexOf('Wg/Pce FC');
      var iCust = iH.indexOf('CUSTOMER');

      for (var r = 1; r < iData.length; r++) {
        var code = String(iData[r][iIC] || '').trim();
        if (!code) continue;
        items.push({
          item_code   : code,
          description : String(iData[r][iDesc] || '').trim(),
          spec        : String(iData[r][iSpec] || '').trim(),
          t           : parseFloat(iData[r][iT] || 0) || 0,
          p           : parseFloat(iData[r][iP] || 0) || 0,
          l           : (function(v){ var n=parseFloat(v); return isNaN(n) ? String(v||'').trim() : n; })(iData[r][iL]),
          uom         : String(iData[r][iUoM] || 'Pcs').trim(),
          wg_pce      : parseFloat(iData[r][iWg] || 0) || 0,
          cust_ref    : iCust >= 0 ? String(iData[r][iCust] || '').trim() : ''
        });
      }
    }

    // ─── M_CUST ───
    var custSheet = ss.getSheetByName('M_CUST');
    if (custSheet) {
      var cData = custSheet.getDataRange().getValues();
      var cH = cData[0].map(function(h){ return String(h).trim(); });
      var cCode = cH.indexOf('C_CODE');
      var cInit = cH.indexOf('C_INITIAL');
      var cName = cH.indexOf('CUSTOMER_NAME');
      for (var r2 = 1; r2 < cData.length; r2++) {
        var code2 = String(cData[r2][cCode] || '').trim();
        if (!code2) continue;
        customers.push({
          code    : code2,
          initial : cInit >= 0 ? String(cData[r2][cInit] || '').trim() : '',
          nama    : cName >= 0 ? String(cData[r2][cName] || '').trim() : ''
        });
      }
    }

    return { success: true, items: items, customers: customers };
  } catch (e) {
    return { success: false, message: 'Error stpGetInitData: ' + e.toString() };
  }
}

// =========================================================================
// 2. GET LIST — grouped by STP_No, dgn filter
// =========================================================================
function stpGetList(filter) {
  try {
    filter = filter || {};
    var statusF = String(filter.status || 'ALL').toUpperCase(); // ALL / OPEN / FULFILLED
    var periodF = String(filter.periode || '').trim();
    var searchF = String(filter.search || '').trim().toLowerCase();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(STP_SHEET_NAME);
    if (!sh) return { success: true, groups: [] };
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, groups: [] };

    var H = data[0].map(function(h){ return String(h).trim(); });
    var idx = {
      stp     : H.indexOf('STP_No'),
      tgl     : H.indexOf('Tgl_Input'),
      cust    : H.indexOf('Cust'),
      per     : H.indexOf('Periode'),
      sch     : H.indexOf('Schedule_Date'),
      item    : H.indexOf('Item_Code'),
      desc    : H.indexOf('Description'),
      spec    : H.indexOf('Spec'),
      t       : H.indexOf('T'),
      p       : H.indexOf('P'),
      l       : H.indexOf('L'),
      uom     : H.indexOf('UoM'),
      qtyReq  : H.indexOf('Qty_Req'),
      kgReq   : H.indexOf('KG_Req'),
      qtyFul  : H.indexOf('Qty_Fulfill'),
      qtySisa : H.indexOf('Qty_Sisa'),
      stat    : H.indexOf('Status'),
      ref     : H.indexOf('Ref_SPK'),
      own     : H.indexOf('Owner_Used'),
      pri     : H.indexOf('Priority'),
      note    : H.indexOf('NOTE'),
      by      : H.indexOf('Created_By')
    };

    var tz = Session.getScriptTimeZone();
    var groups = {};
    var order = [];

    for (var r = 1; r < data.length; r++) {
      var stpNo = String(data[r][idx.stp] || '').trim();
      if (!stpNo) continue;

      var item = {
        item_code   : String(data[r][idx.item] || '').trim(),
        description : String(data[r][idx.desc] || '').trim(),
        spec        : String(data[r][idx.spec] || '').trim(),
        t           : parseFloat(data[r][idx.t] || 0) || 0,
        p           : parseFloat(data[r][idx.p] || 0) || 0,
        l           : String(data[r][idx.l] || '').trim(),
        uom         : String(data[r][idx.uom] || '').trim(),
        qty_req     : parseFloat(data[r][idx.qtyReq] || 0) || 0,
        kg_req      : parseFloat(data[r][idx.kgReq] || 0) || 0,
        qty_fulfill : parseFloat(data[r][idx.qtyFul] || 0) || 0,
        qty_sisa    : parseFloat(data[r][idx.qtySisa] || 0) || 0,
        status      : String(data[r][idx.stat] || 'OPEN').toUpperCase(),
        ref_spk     : String(data[r][idx.ref] || '').trim(),
        note        : String(data[r][idx.note] || '').trim()
      };

      if (!groups[stpNo]) {
        var tglVal = data[r][idx.tgl];
        var schVal = data[r][idx.sch];
        groups[stpNo] = {
          stp_no       : stpNo,
          tgl_input    : tglVal instanceof Date ? Utilities.formatDate(tglVal, tz, 'yyyy-MM-dd HH:mm') : String(tglVal || ''),
          tgl_raw      : tglVal instanceof Date ? tglVal.getTime() : 0,
          cust         : String(data[r][idx.cust] || '').trim(),
          periode      : String(data[r][idx.per] || '').trim(),
          schedule     : schVal instanceof Date ? Utilities.formatDate(schVal, tz, 'yyyy-MM-dd') : String(schVal || ''),
          schedule_raw : schVal instanceof Date ? schVal.getTime() : 0,
          owner_used   : String(data[r][idx.own] || '').trim(),
          priority     : String(data[r][idx.pri] || '').trim(),
          created_by   : String(data[r][idx.by] || '').trim(),
          items        : []
        };
        order.push(stpNo);
      }
      groups[stpNo].items.push(item);
    }

    // Aggregate + filter
    var out = [];
    for (var i = 0; i < order.length; i++) {
      var g = groups[order[i]];
      var totalReq = 0, totalFul = 0, totalSisa = 0, openCnt = 0;
      for (var j = 0; j < g.items.length; j++) {
        totalReq  += g.items[j].qty_req;
        totalFul  += g.items[j].qty_fulfill;
        totalSisa += g.items[j].qty_sisa;
        if (g.items[j].status === 'OPEN') openCnt++;
      }
      g.total_qty_req  = totalReq;
      g.total_qty_ful  = totalFul;
      g.total_qty_sisa = totalSisa;
      g.item_count     = g.items.length;
      g.open_count     = openCnt;
      g.overall_status = openCnt > 0 ? 'OPEN' : 'FULFILLED';

      // Filter status
      if (statusF !== 'ALL' && g.overall_status !== statusF) continue;
      // Filter periode
      if (periodF && g.periode !== periodF) continue;
      // Filter search (STP_No / Cust / Item_Code / Description)
      if (searchF) {
        var hay = (g.stp_no + ' ' + g.cust + ' ' +
                   g.items.map(function(x){ return x.item_code + ' ' + x.description; }).join(' ')
                  ).toLowerCase();
        if (hay.indexOf(searchF) === -1) continue;
      }
      out.push(g);
    }

    // Sort: latest first
    out.sort(function(a, b){ return b.tgl_raw - a.tgl_raw; });

    return { success: true, groups: out };
  } catch (e) {
    return { success: false, message: 'Error stpGetList: ' + e.toString() };
  }
}

// =========================================================================
// 3. PRIVATE: generate STP_No — scan sheet, max+1 per tahun
// =========================================================================
function _stp_generateNo(data, H) {
  var iSTP = H.indexOf('STP_No');
  var yy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yy');
  var re = new RegExp('^STP-' + yy + '(\\d{4})$');
  var maxSeq = 0;
  for (var i = 1; i < data.length; i++) {
    var v = String(data[i][iSTP] || '').trim();
    var m = v.match(re);
    if (m) {
      var s = parseInt(m[1], 10);
      if (s > maxSeq) maxSeq = s;
    }
  }
  return 'STP-' + yy + String(maxSeq + 1).padStart(4, '0');
}

// =========================================================================
// 4. PRIVATE: cek apakah item punya SPK linked (guard edit/delete)
//    itemCode kosong → cek stpNo level (semua item)
// =========================================================================
function _stp_hasLinkedSpk(stpNo, itemCode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SPK');
  if (!sh) return false;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return false;
  var H = data[0].map(function(h){ return String(h).trim(); });
  var iSoRef = H.indexOf('SO_Ref');
  var iItem  = H.indexOf('Item_Code');
  var iStat  = H.indexOf('Status');
  if (iSoRef < 0 || iItem < 0) return false;

  var target = String(stpNo).trim();
  var targetItem = String(itemCode || '').trim();
  for (var i = 1; i < data.length; i++) {
    var sref = String(data[i][iSoRef] || '').trim();
    if (sref !== target) continue;
    if (targetItem) {
      var ic = String(data[i][iItem] || '').trim();
      if (ic !== targetItem) continue;
    }
    var st = iStat >= 0 ? String(data[i][iStat] || '').toUpperCase() : '';
    if (st === 'CANCELLED') continue;
    return true;
  }
  return false;
}

// =========================================================================
// 5. SAVE NEW STP (multi-item)
//    payload = { header:{cust, periode, schedule_date, owner_used, priority, note},
//                items:[{item_code, qty_req, note?, ref_spk?}, ...] }
// =========================================================================
function stpSaveNew(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload || !payload.header) return { success: false, message: 'Payload header kosong' };
    if (!Array.isArray(payload.items) || payload.items.length === 0) {
      return { success: false, message: 'Minimal 1 item wajib diisi' };
    }
    var head  = payload.header;
    var items = payload.items;

    // ─── Validasi header ───
    if (!head.cust || !String(head.cust).trim()) {
      return { success: false, message: 'Customer wajib diisi' };
    }
    if (!head.periode || !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(head.periode).trim())) {
      return { success: false, message: 'Periode wajib format YYYY-MM (contoh: 2026-07)' };
    }
    if (!head.schedule_date) return { success: false, message: 'Schedule Date wajib diisi' };

    // ─── Validasi item ───
    var seenItems = {};
    for (var v = 0; v < items.length; v++) {
      var it = items[v];
      var ic = String(it.item_code || '').trim();
      if (!ic) return { success: false, message: 'Item baris #' + (v+1) + ' belum pilih Item_Code' };
      if (seenItems[ic]) return { success: false, message: 'Item ' + ic + ' duplikat di baris #' + (v+1) };
      seenItems[ic] = true;
      var q = parseFloat(it.qty_req);
      if (!q || q <= 0) {
        return { success: false, message: 'Item ' + ic + ' Qty_Req harus > 0' };
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(STP_SHEET_NAME);
    if (!sh) return { success: false, message: 'Sheet ' + STP_SHEET_NAME + ' tidak ditemukan' };

    var data = sh.getDataRange().getValues();
    var H = data[0].map(function(h){ return String(h).trim(); });
    var iSTP = H.indexOf('STP_No');
    if (iSTP < 0) return { success: false, message: 'Kolom STP_No tidak ditemukan di sheet' };

    // Generate STP_No
    var stpNo = _stp_generateNo(data, H);

    // Cari last row non-empty (anti-bug ARRAY FORMULA)
    var lastRow = 1;
    for (var d = data.length - 1; d >= 1; d--) {
      if (String(data[d][iSTP] || '').trim() !== '') { lastRow = d + 1; break; }
    }
    var startRow = lastRow + 1;

    var now = new Date();
    var createdBy = '';
    try { createdBy = Session.getActiveUser().getEmail() || ''; } catch (er) { createdBy = ''; }

    for (var i2 = 0; i2 < items.length; i2++) {
      var it2 = items[i2];
      var rowMap = {
        'STP_No'        : stpNo,
        'Tgl_Input'     : now,
        'Cust'          : String(head.cust).trim(),
        'Periode'       : String(head.periode).trim(),
        'Schedule_Date' : new Date(head.schedule_date),
        'Item_Code'     : String(it2.item_code).trim(),
        'Qty_Req'       : parseFloat(it2.qty_req),
        'Ref_SPK'       : String(it2.ref_spk || '').trim(),
        'Owner_Used'    : String(head.owner_used || 'FC').trim().toUpperCase(),
        'Priority'      : String(head.priority || 'Normal').trim(),
        'NOTE'          : String(it2.note || head.note || '').trim(),
        'Created_By'    : createdBy
      };
      var targetRow = startRow + i2;
      H.forEach(function(colName, colIdx) {
        if (rowMap[colName] !== undefined) {
          sh.getRange(targetRow, colIdx + 1).setValue(rowMap[colName]);
        }
      });
    }

    SpreadsheetApp.flush();
    return {
      success : true,
      stp_no  : stpNo,
      count   : items.length,
      message : 'STP ' + stpNo + ' berhasil dibuat (' + items.length + ' item)'
    };
  } catch (e) {
    return { success: false, message: 'Error stpSaveNew: ' + e.toString() };
  } finally { lock.releaseLock(); }
}

// =========================================================================
// 6. UPDATE HEADER — apply ke SEMUA row STP_No tsb
//    payload = { stp_no, cust?, periode?, schedule_date?, owner_used?, priority? }
// =========================================================================
function stpUpdateHeader(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload || !payload.stp_no) return { success: false, message: 'stp_no wajib' };
    var stpNo = String(payload.stp_no).trim();

    if (payload.periode && !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(payload.periode).trim())) {
      return { success: false, message: 'Periode format harus YYYY-MM' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(STP_SHEET_NAME);
    if (!sh) return { success: false, message: 'Sheet STP_REQ tidak ditemukan' };

    var data = sh.getDataRange().getValues();
    var H = data[0].map(function(h){ return String(h).trim(); });
    var iSTP = H.indexOf('STP_No');

    var cols = {};
    if (payload.cust          !== undefined && payload.cust          !== '') cols['Cust']          = String(payload.cust).trim();
    if (payload.periode       !== undefined && payload.periode       !== '') cols['Periode']       = String(payload.periode).trim();
    if (payload.schedule_date !== undefined && payload.schedule_date !== '') cols['Schedule_Date'] = new Date(payload.schedule_date);
    if (payload.owner_used    !== undefined && payload.owner_used    !== '') cols['Owner_Used']    = String(payload.owner_used).trim().toUpperCase();
    if (payload.priority      !== undefined && payload.priority      !== '') cols['Priority']      = String(payload.priority).trim();

    if (Object.keys(cols).length === 0) {
      return { success: false, message: 'Tidak ada field yg diupdate' };
    }

    var updated = 0;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iSTP] || '').trim() !== stpNo) continue;
      Object.keys(cols).forEach(function(colName){
        var ci = H.indexOf(colName);
        if (ci < 0) return;
        sh.getRange(r + 1, ci + 1).setValue(cols[colName]);
      });
      updated++;
    }

    if (updated === 0) return { success: false, message: 'STP ' + stpNo + ' tidak ditemukan' };

    SpreadsheetApp.flush();
    return { success: true, message: 'Header STP ' + stpNo + ' diupdate (' + updated + ' row)' };
  } catch (e) {
    return { success: false, message: 'Error stpUpdateHeader: ' + e.toString() };
  } finally { lock.releaseLock(); }
}

// =========================================================================
// 7. UPDATE ITEM (Qty_Req / NOTE / Ref_SPK)
//    payload = { stp_no, item_code, qty_req?, note?, ref_spk? }
//    GUARD: qty_req hanya boleh diedit kalau item BELUM ada SPK linked
// =========================================================================
function stpUpdateItem(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload || !payload.stp_no || !payload.item_code) {
      return { success: false, message: 'stp_no & item_code wajib' };
    }
    var stpNo    = String(payload.stp_no).trim();
    var itemCode = String(payload.item_code).trim();

    // GUARD Qty_Req
    if (payload.qty_req !== undefined) {
      var q = parseFloat(payload.qty_req);
      if (!q || q <= 0) return { success: false, message: 'Qty_Req harus > 0' };
      if (_stp_hasLinkedSpk(stpNo, itemCode)) {
        return { success: false, message: 'Item ' + itemCode + ' sudah ada SPK linked, Qty_Req dikunci' };
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(STP_SHEET_NAME);
    var data = sh.getDataRange().getValues();
    var H = data[0].map(function(h){ return String(h).trim(); });
    var iSTP  = H.indexOf('STP_No');
    var iItem = H.indexOf('Item_Code');
    var iQty  = H.indexOf('Qty_Req');
    var iNote = H.indexOf('NOTE');
    var iRef  = H.indexOf('Ref_SPK');

    var targetRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iSTP]  || '').trim() === stpNo &&
          String(data[r][iItem] || '').trim() === itemCode) {
        targetRow = r + 1;
        break;
      }
    }
    if (targetRow < 0) return { success: false, message: 'Item ' + itemCode + ' di STP ' + stpNo + ' tidak ditemukan' };

    if (payload.qty_req !== undefined && iQty >= 0) {
      sh.getRange(targetRow, iQty + 1).setValue(parseFloat(payload.qty_req));
    }
    if (payload.note !== undefined && iNote >= 0) {
      sh.getRange(targetRow, iNote + 1).setValue(String(payload.note));
    }
    if (payload.ref_spk !== undefined && iRef >= 0) {
      sh.getRange(targetRow, iRef + 1).setValue(String(payload.ref_spk));
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Item ' + itemCode + ' diupdate' };
  } catch (e) {
    return { success: false, message: 'Error stpUpdateItem: ' + e.toString() };
  } finally { lock.releaseLock(); }
}

// =========================================================================
// 8. ADD ITEM baru ke STP existing
//    payload = { stp_no, item_code, qty_req, note?, ref_spk? }
//    Header value (Cust/Periode/Schedule/Owner/Priority) diambil dari row
//    pertama STP_No tsb (biar konsisten).
// =========================================================================
function stpAddItem(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload || !payload.stp_no || !payload.item_code) {
      return { success: false, message: 'stp_no & item_code wajib' };
    }
    var stpNo    = String(payload.stp_no).trim();
    var itemCode = String(payload.item_code).trim();
    var qty      = parseFloat(payload.qty_req);
    if (!qty || qty <= 0) return { success: false, message: 'Qty_Req harus > 0' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(STP_SHEET_NAME);
    var data = sh.getDataRange().getValues();
    var H = data[0].map(function(h){ return String(h).trim(); });
    var iSTP  = H.indexOf('STP_No');
    var iItem = H.indexOf('Item_Code');

    // Cari row template + cek duplikat
    var tmplRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iSTP] || '').trim() === stpNo) {
        if (String(data[r][iItem] || '').trim() === itemCode) {
          return { success: false, message: 'Item ' + itemCode + ' sudah ada di STP ' + stpNo };
        }
        if (tmplRow < 0) tmplRow = r;
      }
    }
    if (tmplRow < 0) return { success: false, message: 'STP ' + stpNo + ' tidak ditemukan' };

    // Cari last non-empty row
    var lastRow = 1;
    for (var d = data.length - 1; d >= 1; d--) {
      if (String(data[d][iSTP] || '').trim() !== '') { lastRow = d + 1; break; }
    }
    var newRow = lastRow + 1;

    var iTgl  = H.indexOf('Tgl_Input');
    var iCust = H.indexOf('Cust');
    var iPer  = H.indexOf('Periode');
    var iSch  = H.indexOf('Schedule_Date');
    var iOwn  = H.indexOf('Owner_Used');
    var iPri  = H.indexOf('Priority');
    var iQty  = H.indexOf('Qty_Req');
    var iRef  = H.indexOf('Ref_SPK');
    var iNote = H.indexOf('NOTE');
    var iBy   = H.indexOf('Created_By');

    var createdBy = '';
    try { createdBy = Session.getActiveUser().getEmail() || ''; } catch (er) { createdBy = ''; }

    var writes = [
      [iSTP,  stpNo],
      [iTgl,  new Date()],
      [iCust, data[tmplRow][iCust]],
      [iPer,  data[tmplRow][iPer]],
      [iSch,  data[tmplRow][iSch]],
      [iItem, itemCode],
      [iQty,  qty],
      [iRef,  String(payload.ref_spk || '').trim()],
      [iOwn,  data[tmplRow][iOwn]],
      [iPri,  data[tmplRow][iPri]],
      [iNote, String(payload.note || '').trim()],
      [iBy,   createdBy]
    ];
    for (var w = 0; w < writes.length; w++) {
      if (writes[w][0] >= 0) {
        sh.getRange(newRow, writes[w][0] + 1).setValue(writes[w][1]);
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Item ' + itemCode + ' ditambahkan ke STP ' + stpNo };
  } catch (e) {
    return { success: false, message: 'Error stpAddItem: ' + e.toString() };
  } finally { lock.releaseLock(); }
}

// =========================================================================
// 9. DELETE ITEM (1 row)
//    GUARD: block kalau item punya SPK linked.
//    GUARD: block kalau ini item TERAKHIR (redirect ke stpCancelAll).
// =========================================================================
function stpDeleteItem(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload || !payload.stp_no || !payload.item_code) {
      return { success: false, message: 'stp_no & item_code wajib' };
    }
    var stpNo    = String(payload.stp_no).trim();
    var itemCode = String(payload.item_code).trim();

    if (_stp_hasLinkedSpk(stpNo, itemCode)) {
      return { success: false, message: 'Item ' + itemCode + ' sudah ada SPK linked, tidak boleh dihapus' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(STP_SHEET_NAME);
    var data = sh.getDataRange().getValues();
    var H = data[0].map(function(h){ return String(h).trim(); });
    var iSTP  = H.indexOf('STP_No');
    var iItem = H.indexOf('Item_Code');

    var stpCount = 0;
    var targetRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iSTP] || '').trim() === stpNo) {
        stpCount++;
        if (String(data[r][iItem] || '').trim() === itemCode) {
          targetRow = r + 1;
        }
      }
    }
    if (targetRow < 0) return { success: false, message: 'Item tidak ditemukan' };
    if (stpCount === 1) {
      return { success: false, message: 'Ini item terakhir di STP ' + stpNo + '. Pakai Cancel STP untuk hapus.' };
    }

    sh.deleteRow(targetRow);
    SpreadsheetApp.flush();
    return { success: true, message: 'Item ' + itemCode + ' dihapus dari STP ' + stpNo };
  } catch (e) {
    return { success: false, message: 'Error stpDeleteItem: ' + e.toString() };
  } finally { lock.releaseLock(); }
}

// =========================================================================
// 10. CANCEL — hapus SEMUA row STP_No tsb
//     GUARD: block kalau ada SPK linked (item mana saja)
// =========================================================================
function stpCancelAll(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload || !payload.stp_no) return { success: false, message: 'stp_no wajib' };
    var stpNo = String(payload.stp_no).trim();

    if (_stp_hasLinkedSpk(stpNo, '')) {
      return { success: false, message: 'STP ' + stpNo + ' sudah ada SPK linked, tidak boleh di-cancel' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(STP_SHEET_NAME);
    var data = sh.getDataRange().getValues();
    var H = data[0].map(function(h){ return String(h).trim(); });
    var iSTP = H.indexOf('STP_No');

    var rowsToDelete = [];
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iSTP] || '').trim() === stpNo) rowsToDelete.push(r + 1);
    }
    if (rowsToDelete.length === 0) {
      return { success: false, message: 'STP ' + stpNo + ' tidak ditemukan' };
    }

    // Hapus dari bawah biar index tidak geser
    for (var i = rowsToDelete.length - 1; i >= 0; i--) {
      sh.deleteRow(rowsToDelete[i]);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'STP ' + stpNo + ' dihapus (' + rowsToDelete.length + ' row)'
    };
  } catch (e) {
    return { success: false, message: 'Error stpCancelAll: ' + e.toString() };
  } finally { lock.releaseLock(); }
}