function _test_massBalance_sprint1B() {
  var res = getLiveStockData();
  if (!res.success) { Logger.log('ERR: ' + res.message); return; }
  
  Logger.log('=== SUMMARY per kategori ===');
  ['coil','sheet','wip','fg','ng'].forEach(function(cat) {
    var s = res.summary[cat];
    Logger.log(cat.toUpperCase() + ': batch=' + s.total_batch + 
      ' | kg_in=' + Math.round(s.kg_in) + 
      ' | kg_keep=' + Math.round(s.kg_keep) + 
      ' | kg_konsumsi=' + Math.round(s.kg_konsumsi) + 
      ' | kg_avail=' + Math.round(s.kg_avail) + 
      ' | selisih warn/crit=' + s.selisih_warn_count + '/' + s.selisih_crit_count);
  });
  
  Logger.log('\n=== SAMPLE 3 batch dengan selisih ===');
  var batches = [].concat(res.batches.coil, res.batches.sheet, res.batches.wip, res.batches.fg, res.batches.ng);
  var flagged = batches.filter(function(b) { return b.selisih_level; }).slice(0, 3);
  flagged.forEach(function(b) {
    Logger.log(b.batch_id + ' [' + b.selisih_level + ']: selisih=' + b.selisih_kg + 
      ' | in=' + b.kg_in + ' keep=' + b.kg_keep + ' kons=' + b.kg_konsumsi + ' av=' + b.kg_avail);
  });
  
  Logger.log('\n=== SAMPLE 3 batch WIP dengan vendor enriched ===');
  var wipWithVendor = res.batches.wip.filter(function(b) { return b.supplier; }).slice(0, 3);
  wipWithVendor.forEach(function(b) {
    Logger.log(b.batch_id + ': ' + b.supplier + ' · ' + b.no_coil);
  });
}

/* ============================================================================
 * 🔍 DIAGNOSTIC — Mulai_DT / Selesai_DT / Duration Column Bug
 * Paste ke Apps Script Editor → pilih function → Run → View Logs
 * ============================================================================ */
function DIAG_bugMulaiSelesai() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SPK');
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();

  Logger.log('════════════════════════════════════════════════════');
  Logger.log('DIAG SPK — ' + new Date().toISOString());
  Logger.log('Sheet: SPK | LastRow: ' + lastRow + ' | LastCol: ' + lastCol);
  Logger.log('Timezone: ' + Session.getScriptTimeZone());
  Logger.log('Spreadsheet locale: ' + ss.getSpreadsheetLocale());

  // ─── 1. HEADER CHECK: raw bytes untuk detect hidden chars ────────────────
  Logger.log('\n─── 1. HEADER ROW RAW CHECK ───');
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var wantedCols = ['Status', 'Mulai_DT', 'Selesai_DT', 'Duration', 'Owner'];
  wantedCols.forEach(function(name){
    var idx = -1;
    for (var i = 0; i < headerRow.length; i++) {
      if (String(headerRow[i]) === name) { idx = i; break; }
    }
    if (idx === -1) {
      // try trim
      for (var i = 0; i < headerRow.length; i++) {
        if (String(headerRow[i]).trim() === name) { idx = i; break; }
      }
    }
    if (idx === -1) {
      Logger.log('  ❌ NOT FOUND: "' + name + '"');
    } else {
      var raw = String(headerRow[idx]);
      var codes = [];
      for (var c = 0; c < raw.length; c++) codes.push(raw.charCodeAt(c));
      Logger.log('  ✅ "' + name + '" @ col ' + (idx+1) + ' | raw="' + raw + '" | len=' + raw.length + ' | charCodes=[' + codes.join(',') + ']');
    }
  });

  // ─── 2. CELL VALUE + FORMAT untuk row Header + row 2 ─────────────────────
  Logger.log('\n─── 2. CELL VALUE & FORMAT (sample rows) ───');
  var iMulai   = headerRow.map(function(h){return String(h).trim();}).indexOf('Mulai_DT');
  var iSelesai = headerRow.map(function(h){return String(h).trim();}).indexOf('Selesai_DT');
  var iDur     = headerRow.map(function(h){return String(h).trim();}).indexOf('Duration');
  var iSpk     = headerRow.map(function(h){return String(h).trim();}).indexOf('SPK_No');
  var iStat    = headerRow.map(function(h){return String(h).trim();}).indexOf('Status');

  // Cari 3 row: HEADER RUNNING pertama, OUT DONE pertama, OUT DONE dengan Selesai isi
  var allData = sh.getRange(2, 1, Math.min(500, lastRow-1), lastCol).getValues();
  var sampleRows = [];
  for (var i = 0; i < allData.length; i++) {
    var st = String(allData[i][iStat] || '').toUpperCase();
    if (st === 'RUNNING' || st === 'DONE') {
      sampleRows.push({ rowIdx: i+2, spk: allData[i][iSpk], status: st });
      if (sampleRows.length >= 5) break;
    }
  }
  sampleRows.forEach(function(s){
    var rMulai   = sh.getRange(s.rowIdx, iMulai+1);
    var rSelesai = sh.getRange(s.rowIdx, iSelesai+1);
    var rDur     = iDur >= 0 ? sh.getRange(s.rowIdx, iDur+1) : null;
    var vMulai   = rMulai.getValue();
    var vSelesai = rSelesai.getValue();
    Logger.log('  Row ' + s.rowIdx + ' | ' + s.spk + ' [' + s.status + ']');
    Logger.log('    Mulai_DT   : value=' + JSON.stringify(vMulai) + ' | type=' + typeof vMulai + ' | isDate=' + (vMulai instanceof Date) + ' | format="' + rMulai.getNumberFormat() + '" | formula="' + rMulai.getFormula() + '"');
    Logger.log('    Selesai_DT : value=' + JSON.stringify(vSelesai) + ' | type=' + typeof vSelesai + ' | isDate=' + (vSelesai instanceof Date) + ' | format="' + rSelesai.getNumberFormat() + '" | formula="' + rSelesai.getFormula() + '"');
    if (rDur) {
      var vDur = rDur.getValue();
      Logger.log('    Duration   : value=' + JSON.stringify(vDur) + ' | format="' + rDur.getNumberFormat() + '" | formula="' + rDur.getFormula() + '"');
    }
  });

  // ─── 3. TEST WRITE: coba tulis Date ke cell test ─────────────────────────
  Logger.log('\n─── 3. TEST WRITE new Date() ───');
  var testRow = lastRow + 1;
  var testCol = iMulai + 1;
  try {
    var beforeFmt = sh.getRange(testRow, testCol).getNumberFormat();
    sh.getRange(testRow, testCol).setValue(new Date());
    SpreadsheetApp.flush();
    var afterVal = sh.getRange(testRow, testCol).getValue();
    var afterFmt = sh.getRange(testRow, testCol).getNumberFormat();
    Logger.log('  Test row ' + testRow + ', col ' + testCol + ' (Mulai_DT)');
    Logger.log('  Format sebelum write: "' + beforeFmt + '"');
    Logger.log('  Value setelah write : ' + JSON.stringify(afterVal) + ' | isDate=' + (afterVal instanceof Date));
    Logger.log('  Format setelah write: "' + afterFmt + '"');
    // Cleanup
    sh.getRange(testRow, testCol).clearContent();
  } catch (e) {
    Logger.log('  ❌ TEST WRITE FAILED: ' + e.toString());
  }

  // ─── 4. TRIGGERS (installed) ─────────────────────────────────────────────
  Logger.log('\n─── 4. INSTALLED TRIGGERS ───');
  var trigs = ScriptApp.getProjectTriggers();
  if (trigs.length === 0) {
    Logger.log('  (no installed triggers)');
  } else {
    trigs.forEach(function(t){
      Logger.log('  - handler="' + t.getHandlerFunction() + '" | event=' + t.getEventType() + ' | source=' + t.getTriggerSource());
    });
  }

  // ─── 5. Check for column with formula (in case Mulai_DT / Selesai_DT ARRAYFORMULA'd) ─
  Logger.log('\n─── 5. FORMULA CHECK on header row (row 1) ───');
  [iMulai, iSelesai, iDur].forEach(function(idx, i){
    if (idx < 0) return;
    var label = ['Mulai_DT', 'Selesai_DT', 'Duration'][i];
    var f = sh.getRange(1, idx+1).getFormula();
    var v = sh.getRange(1, idx+1).getValue();
    Logger.log('  ' + label + ' header cell: value="' + v + '" | formula="' + f + '"');
    // Also check row 2 for ARRAYFORMULA output
    var f2 = sh.getRange(2, idx+1).getFormula();
    Logger.log('    Row 2 formula: "' + f2 + '"');
  });

  Logger.log('\n════════════════════════════════════════════════════');
  Logger.log('DIAG SELESAI. Copy semua log di atas & kirim balik.');
}




/* ============================================================================
 * 🔍 DIAG-2 — Target row CTL-260026 (row-level inspection)
 * ============================================================================ */
function DIAG_targetCTL260026() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SPK');
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();

  Logger.log('════════════════════════════════════════════════════');
  Logger.log('DIAG-2 TARGET CTL-260026 — ' + new Date().toISOString());

  var allData = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = allData[0].map(function(h){ return String(h).trim(); });
  var iSpk = headers.indexOf('SPK_No');
  var iType = headers.indexOf('SPK_Type');
  var iParent = headers.indexOf('Parent_SPK');
  var iStatus = headers.indexOf('Status');
  var iMulai = headers.indexOf('Mulai_DT');
  var iSelesai = headers.indexOf('Selesai_DT');
  var iDur = headers.indexOf('Duration');
  var iTglBuat = headers.indexOf('Tgl_Buat');
  var iQtyAct = headers.indexOf('Qty_Actual');

  // Cari semua row yg SPK_No mengandung "CTL-260026"
  var targetRows = [];
  for (var r = 1; r < allData.length; r++) {
    var spk = String(allData[r][iSpk] || '');
    if (spk.indexOf('CTL-260026') >= 0) {
      targetRows.push(r+1); // 1-based row number
    }
  }
  Logger.log('Found ' + targetRows.length + ' rows matching CTL-260026:');

  targetRows.forEach(function(rowNum){
    var arrIdx = rowNum - 1;
    var spk = allData[arrIdx][iSpk];
    var typ = allData[arrIdx][iType];
    var st  = allData[arrIdx][iStatus];
    var tgl = allData[arrIdx][iTglBuat];
    var qty = allData[arrIdx][iQtyAct];

    Logger.log('\n─── Row ' + rowNum + ' | ' + spk + ' | ' + typ + ' | Status=' + st + ' | Qty_Actual=' + qty + ' ───');

    // Deep inspect Mulai_DT
    var cMulai = sh.getRange(rowNum, iMulai+1);
    var vMulai = cMulai.getValue();
    var dvMulai = cMulai.getDisplayValue();
    var fMulai = cMulai.getNumberFormat();
    var fmMulai = cMulai.getFormula();
    Logger.log('  Mulai_DT   : value=' + JSON.stringify(vMulai) + ' | display="' + dvMulai + '" | type=' + typeof vMulai + ' | isDate=' + (vMulai instanceof Date) + ' | format="' + fMulai + '" | formula="' + fmMulai + '"');

    // Deep inspect Selesai_DT
    var cSelesai = sh.getRange(rowNum, iSelesai+1);
    var vSelesai = cSelesai.getValue();
    var dvSelesai = cSelesai.getDisplayValue();
    var fSelesai = cSelesai.getNumberFormat();
    var fmSelesai = cSelesai.getFormula();
    Logger.log('  Selesai_DT : value=' + JSON.stringify(vSelesai) + ' | display="' + dvSelesai + '" | type=' + typeof vSelesai + ' | isDate=' + (vSelesai instanceof Date) + ' | format="' + fSelesai + '" | formula="' + fmSelesai + '"');

    // Duration
    if (iDur >= 0) {
      var cDur = sh.getRange(rowNum, iDur+1);
      Logger.log('  Duration   : value=' + JSON.stringify(cDur.getValue()) + ' | display="' + cDur.getDisplayValue() + '" | format="' + cDur.getNumberFormat() + '"');
    }

    // Tgl_Buat (buat compare, ini juga Date field)
    Logger.log('  Tgl_Buat   : value=' + JSON.stringify(tgl) + ' | isDate=' + (tgl instanceof Date));
  });

  // Compare dengan row LAMA yang berhasil
  Logger.log('\n─── BANDING dgn row lama (row 6 = CTL-260001) ───');
  var oldRow = 6;
  var cOldMulai = sh.getRange(oldRow, iMulai+1);
  var cOldSelesai = sh.getRange(oldRow, iSelesai+1);
  Logger.log('  Row 6 Mulai_DT   : value=' + JSON.stringify(cOldMulai.getValue()) + ' | display="' + cOldMulai.getDisplayValue() + '" | format="' + cOldMulai.getNumberFormat() + '"');
  Logger.log('  Row 6 Selesai_DT : value=' + JSON.stringify(cOldSelesai.getValue()) + ' | display="' + cOldSelesai.getDisplayValue() + '" | format="' + cOldSelesai.getNumberFormat() + '"');

  Logger.log('\n════════════════════════════════════════════════════');
}



/* ============================================================================
 * 🔍 DIAG-3 — Test hipotesis format conflict
 * ============================================================================ */
function DIAG_testFormatFix() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SPK');
  
  // Target row 611 (CTL-260026 HEADER, Mulai_DT kosong dengan format 0.###)
  var testRow = 611;
  var mulaiCol = 19; // Kolom S = Mulai_DT
  
  Logger.log('════════════════════════════════════════════════════');
  Logger.log('TEST 1: setValue tanpa clear format (kondisi bug sekarang)');
  var c = sh.getRange(testRow, mulaiCol);
  Logger.log('  Before: value=' + JSON.stringify(c.getValue()) + ' | format="' + c.getNumberFormat() + '"');
  
  c.setValue(new Date());
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  
  Logger.log('  After setValue: value=' + JSON.stringify(c.getValue()) + ' | display="' + c.getDisplayValue() + '" | format="' + c.getNumberFormat() + '" | isDate=' + (c.getValue() instanceof Date));
  
  // Rollback
  c.setValue('');
  c.setNumberFormat('0.###############');
  SpreadsheetApp.flush();
  
  Logger.log('\nTEST 2: clearFormat DULU, baru setValue');
  var c2 = sh.getRange(testRow, mulaiCol);
  Logger.log('  Before: value=' + JSON.stringify(c2.getValue()) + ' | format="' + c2.getNumberFormat() + '"');
  
  c2.setNumberFormat('d mmm yy H:mm');  // Explicit Date format
  c2.setValue(new Date());
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  
  Logger.log('  After setNumberFormat+setValue: value=' + JSON.stringify(c2.getValue()) + ' | display="' + c2.getDisplayValue() + '" | format="' + c2.getNumberFormat() + '" | isDate=' + (c2.getValue() instanceof Date));
  
  // Rollback
  c2.setValue('');
  c2.setNumberFormat('0.###############');
  SpreadsheetApp.flush();
  
  Logger.log('\n════════════════════════════════════════════════════');
  Logger.log('Kedua test dirollback. Data awal restored.');
}

/* ============================================================================
 * 🔍 DIAG-4 — Hunt untuk protection / validation / merge yang block write
 * ============================================================================ */
function DIAG_findBlocker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SPK');

  Logger.log('════════════════════════════════════════════════════');
  Logger.log('DIAG-4 — Hunt blocker on Mulai_DT / Selesai_DT columns');

  // ─── A. SHEET-LEVEL PROTECTIONS ──────────────────────────────────────────
  Logger.log('\n─── A. SHEET PROTECTIONS ───');
  var sheetProtections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  Logger.log('  Sheet-level protections: ' + sheetProtections.length);
  sheetProtections.forEach(function(p, i){
    Logger.log('  [' + i + '] desc="' + p.getDescription() + '" | editable=' + p.canEdit() + ' | canDomainEdit=' + p.canDomainEdit());
  });

  // ─── B. RANGE-LEVEL PROTECTIONS ──────────────────────────────────────────
  Logger.log('\n─── B. RANGE PROTECTIONS ───');
  var rangeProtections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  Logger.log('  Range-level protections: ' + rangeProtections.length);
  rangeProtections.forEach(function(p, i){
    var r = p.getRange();
    Logger.log('  [' + i + '] range=' + r.getA1Notation() + ' | desc="' + p.getDescription() + '" | canEdit=' + p.canEdit());
  });

  // ─── C. DATA VALIDATION on target cells ──────────────────────────────────
  Logger.log('\n─── C. DATA VALIDATION on target cells ───');
  var checkCells = [
    { row: 611, col: 19, label: 'CTL-260026 HEADER Mulai_DT' },
    { row: 611, col: 20, label: 'CTL-260026 HEADER Selesai_DT' },
    { row: 612, col: 19, label: 'CTL-260026-01 Mulai_DT' },
    { row: 612, col: 20, label: 'CTL-260026-01 Selesai_DT' },
    { row: 615, col: 20, label: 'CTL-260026-02 Selesai_DT (works!)' },
    { row: 6,   col: 19, label: 'CTL-260001 Mulai_DT (works!)' },
    { row: 800, col: 19, label: 'Row 800 (empty test row - worked!)' }
  ];
  checkCells.forEach(function(cc){
    var c = sh.getRange(cc.row, cc.col);
    var dv = c.getDataValidation();
    if (dv) {
      Logger.log('  ⚠️ ' + cc.label + ' @ R' + cc.row + 'C' + cc.col + ' — HAS validation: criteria=' + dv.getCriteriaType() + ' | allowInvalid=' + dv.getAllowInvalid());
    } else {
      Logger.log('  ✅ ' + cc.label + ' @ R' + cc.row + 'C' + cc.col + ' — no validation');
    }
    // Also check merged
    var m = c.getMergedRanges();
    if (m.length > 0) {
      Logger.log('     ⚠️ MERGED with: ' + m.map(function(r){return r.getA1Notation();}).join(', '));
    }
  });

  // ─── D. Test on ROW 800 (blank row) — should work ────────────────────────
  Logger.log('\n─── D. Control test on row 800 (should work) ───');
  var ctrl = sh.getRange(800, 19);
  ctrl.clearContent();
  ctrl.setNumberFormat('');
  SpreadsheetApp.flush();
  ctrl.setValue(new Date());
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  var ctrlVal = ctrl.getValue();
  Logger.log('  Row 800 col 19 after setValue: value=' + JSON.stringify(ctrlVal) + ' | isDate=' + (ctrlVal instanceof Date) + ' | format="' + ctrl.getNumberFormat() + '"');
  ctrl.clearContent();
  ctrl.setNumberFormat('0.###############');
  SpreadsheetApp.flush();

  // ─── E. AGGRESSIVE fix test on row 611:19 (clearContent + clearFormat + setValue) ─
  Logger.log('\n─── E. Aggressive fix test on row 611:19 ───');
  var tgt = sh.getRange(611, 19);
  var oldVal = tgt.getValue();
  var oldFmt = tgt.getNumberFormat();
  Logger.log('  Before: value=' + JSON.stringify(oldVal) + ' | format="' + oldFmt + '"');

  tgt.clearContent();
  tgt.clearFormat();
  SpreadsheetApp.flush();
  Utilities.sleep(200);
  Logger.log('  After clearContent+clearFormat: value=' + JSON.stringify(tgt.getValue()) + ' | format="' + tgt.getNumberFormat() + '"');

  tgt.setValue(new Date());
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  var newVal = tgt.getValue();
  Logger.log('  After setValue(new Date()): value=' + JSON.stringify(newVal) + ' | isDate=' + (newVal instanceof Date) + ' | display="' + tgt.getDisplayValue() + '" | format="' + tgt.getNumberFormat() + '"');

  // Rollback
  tgt.setValue(oldVal === '' ? '' : oldVal);
  tgt.setNumberFormat(oldFmt);
  SpreadsheetApp.flush();
  Logger.log('  Rolled back to original state.');

  Logger.log('\n════════════════════════════════════════════════════');
}