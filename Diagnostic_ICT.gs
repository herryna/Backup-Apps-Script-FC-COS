/**
 * ═══════════════════════════════════════════════════════════════
 * DIAGNOSTIC REKAP_ICT AUDIT — READ ONLY
 * ═══════════════════════════════════════════════════════════════
 * Fungsi: simulasi impact backend FG-guard patch tanpa deploy.
 * Sheet production TIDAK diubah. Output ke sheet DIAG_ICT_Audit
 * (auto-create). Boleh dihapus manual setelah review.
 * ═══════════════════════════════════════════════════════════════
 */
function diagnostic_RekapICT_Audit() {
  var ss = SpreadsheetApp.getActive();
  var spkSht = ss.getSheetByName('SPK');
  var ictSht = ss.getSheetByName('Rekap_ICT');
  if (!spkSht || !ictSht) {
    SpreadsheetApp.getUi().alert('Sheet SPK atau Rekap_ICT tidak ditemukan.');
    return;
  }

  var spkData = spkSht.getDataRange().getValues();
  var ictData = ictSht.getDataRange().getValues();
  var spkHdr  = spkData[0];
  var ictHdr  = ictData[0];

  // SPK indices
  var iSpk    = spkHdr.indexOf('SPK_No');
  var iType   = spkHdr.indexOf('SPK_Type');
  var iStatus = spkHdr.indexOf('Status');
  var iTgt    = spkHdr.indexOf('Target_Loc');
  var iOwn    = spkHdr.indexOf('Owner');
  var iOwnU   = spkHdr.indexOf('Owner_Used');
  var iItem   = spkHdr.indexOf('Item_Code');
  var iSel    = spkHdr.indexOf('Selesai_DT');

  // ICT indices
  var jSpk   = ictHdr.indexOf('SPK_No');
  var jTgl   = ictHdr.indexOf('Tgl_Transfer');
  var jItem  = ictHdr.indexOf('Item_Code');
  var jDesc  = ictHdr.indexOf('Description');
  var jDari  = ictHdr.indexOf('Dari_Owner');
  var jKe    = ictHdr.indexOf('Ke_Owner');
  var jQty   = ictHdr.indexOf('Qty_Sht');
  var jKg    = ictHdr.indexOf('Qty_KG');

  // Build SPK lookup
  var spkMap = {};
  for (var i = 1; i < spkData.length; i++) {
    var no = String(spkData[i][iSpk] || '').trim();
    if (!no) continue;
    spkMap[no] = {
      type   : String(spkData[i][iType]   || '').trim(),
      status : String(spkData[i][iStatus] || '').trim(),
      tgt    : String(spkData[i][iTgt]    || '').trim(),
      owner  : String(spkData[i][iOwn]    || '').trim(),
      ownerU : String(spkData[i][iOwnU]   || '').trim(),
      item   : String(spkData[i][iItem]   || '').trim(),
      done   : iSel !== -1 ? spkData[i][iSel] : ''
    };
  }

  // Helper: is FG target?
  function isFGTarget(t) {
    if (!t) return false;
    return (t === 'Stok_FG' || t.indexOf('FG_') === 0);
  }

  // Scan ICT rows
  var wouldSkip = [];      // ICT existing yang bakal ke-skip patch
  var wouldFire = [];      // ICT existing yang aman (target FG)
  var noMatch   = [];      // ICT dengan SPK_No gak ada di SPK sheet

  for (var j = 1; j < ictData.length; j++) {
    var ictSpk = String(ictData[j][jSpk] || '').trim();
    if (!ictSpk) continue;
    var s = spkMap[ictSpk];
    if (!s) {
      noMatch.push({
        row: j+1, spk_no: ictSpk,
        tgl: ictData[j][jTgl],
        dari: String(ictData[j][jDari]),
        ke  : String(ictData[j][jKe]),
        qty : ictData[j][jQty],
        kg  : ictData[j][jKg]
      });
      continue;
    }
    var isFG = isFGTarget(s.tgt);
    var record = {
      row     : j+1,
      spk_no  : ictSpk,
      spk_type: s.type,
      tgt     : s.tgt,
      dari    : String(ictData[j][jDari]),
      ke      : String(ictData[j][jKe]),
      qty     : ictData[j][jQty],
      kg      : ictData[j][jKg],
      tgl     : ictData[j][jTgl]
    };
    if (isFG) wouldFire.push(record); else wouldSkip.push(record);
  }

  // Scan SPK: DONE + Owner≠Owner_Used + FG → check apakah ICT sudah ada
  var ictSpkSet = {};
  for (var k = 1; k < ictData.length; k++) {
    ictSpkSet[String(ictData[k][jSpk] || '').trim()] = true;
  }
  var missingIct = [];
  Object.keys(spkMap).forEach(function(no) {
    var s = spkMap[no];
    if (s.status !== 'DONE') return;
    if (['CTL-OUT','SHR-OUT','SLT-OUT','ALLOC-OUT'].indexOf(s.type) === -1) return;
    if (!s.owner || !s.ownerU) return;
    if (s.owner.toUpperCase() === s.ownerU.toUpperCase()) return;
    if (!isFGTarget(s.tgt)) return;
    if (!ictSpkSet[no]) {
      missingIct.push({
        spk_no: no, type: s.type, tgt: s.tgt,
        owner: s.owner, ownerU: s.ownerU, done: s.done
      });
    }
  });

  // Write diagnostic sheet
  var diagName = 'DIAG_ICT_Audit';
  var diag = ss.getSheetByName(diagName);
  if (diag) ss.deleteSheet(diag);
  diag = ss.insertSheet(diagName);

  var out = [];
  out.push(['DIAGNOSTIC REKAP_ICT AUDIT — READ ONLY', '', '', '', '', '', '', '']);
  out.push(['Run at: ' + new Date().toString(), '', '', '', '', '', '', '']);
  out.push(['', '', '', '', '', '', '', '']);
  out.push(['📊 RINGKASAN', '', '', '', '', '', '', '']);
  out.push(['Total ICT existing', ictData.length - 1, '', '', '', '', '', '']);
  out.push(['  ✅ Aman (target=FG, tetap fire)', wouldFire.length, '', '', '', '', '', '']);
  out.push(['  🔴 Ke-SKIP setelah patch (target≠FG)', wouldSkip.length, '', '', '', '', '', '']);
  out.push(['  ⚠️ SPK_No tidak ketemu di SPK sheet', noMatch.length, '', '', '', '', '', '']);
  out.push(['Missing ICT (SPK DONE-FG-cross tapi belum ada ICT)', missingIct.length, '', '', '', '', '', '']);
  out.push(['', '', '', '', '', '', '', '']);

  // Section 1: Would Skip
  out.push(['🔴 ICT YANG BAKAL DI-SKIP SETELAH PATCH (target ≠ FG)', '', '', '', '', '', '', '']);
  out.push(['ICT_Row', 'Tgl_Transfer', 'SPK_No', 'SPK_Type', 'Target_Loc', 'Dari→Ke', 'Qty', 'KG']);
  wouldSkip.forEach(function(r) {
    out.push([r.row, r.tgl, r.spk_no, r.spk_type, r.tgt, r.dari + '→' + r.ke, r.qty, r.kg]);
  });
  out.push(['', '', '', '', '', '', '', '']);

  // Section 2: Missing ICT
  out.push(['📋 SPK DONE-FG-CROSS TAPI TIDAK ADA ROW ICT (potensi bug lama)', '', '', '', '', '', '', '']);
  out.push(['SPK_No', 'SPK_Type', 'Target_Loc', 'Owner→Owner_Used', 'Selesai_DT', '', '', '']);
  missingIct.forEach(function(m) {
    out.push([m.spk_no, m.type, m.tgt, m.owner + '→' + m.ownerU, m.done, '', '', '']);
  });
  out.push(['', '', '', '', '', '', '', '']);

  // Section 3: ICT with no matching SPK
  out.push(['⚠️ ICT DENGAN SPK_NO TIDAK KETEMU (data anomaly)', '', '', '', '', '', '', '']);
  out.push(['ICT_Row', 'Tgl_Transfer', 'SPK_No', 'Dari→Ke', 'Qty', 'KG', '', '']);
  noMatch.forEach(function(r) {
    out.push([r.row, r.tgl, r.spk_no, r.dari + '→' + r.ke, r.qty, r.kg, '', '']);
  });

  diag.getRange(1, 1, out.length, 8).setValues(out);
  diag.setFrozenRows(3);
  diag.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  diag.autoResizeColumns(1, 8);

  SpreadsheetApp.getUi().alert(
    'Diagnostic selesai!\n\n' +
    '📊 Total ICT: ' + (ictData.length - 1) + '\n' +
    '🔴 Bakal ke-skip: ' + wouldSkip.length + '\n' +
    '📋 Missing ICT: ' + missingIct.length + '\n' +
    '⚠️ No SPK match: ' + noMatch.length + '\n\n' +
    'Detail ada di sheet: ' + diagName
  );
}
/**
 * Inspect 1 SPK secara detail — READ ONLY
 * Ubah SPK_NO_TARGET di bawah kalau mau cek SPK lain
 */
function diagnostic_InspectSPK() {
  var SPK_NO_TARGET = 'SHR-260106-01';   // ← Ganti kalau mau cek SPK lain
  
  var ss = SpreadsheetApp.getActive();
  var spkSht = ss.getSheetByName('SPK');
  var ictSht = ss.getSheetByName('Rekap_ICT');
  var miSht  = ss.getSheetByName('M_ITEM');
  
  var spkData = spkSht.getDataRange().getValues();
  var hdr = spkData[0];
  
  var found = null;
  for (var i = 1; i < spkData.length; i++) {
    if (String(spkData[i][hdr.indexOf('SPK_No')]).trim() === SPK_NO_TARGET) {
      found = { rowIdx: i + 1, data: spkData[i] };
      break;
    }
  }
  
  if (!found) {
    SpreadsheetApp.getUi().alert('SPK ' + SPK_NO_TARGET + ' tidak ditemukan.');
    return;
  }
  
  // Cek Rekap_ICT: ada gak row untuk SPK ini?
  var ictData = ictSht.getDataRange().getValues();
  var ictHdr  = ictData[0];
  var jSpk = ictHdr.indexOf('SPK_No');
  var existingIct = [];
  for (var j = 1; j < ictData.length; j++) {
    if (String(ictData[j][jSpk]).trim() === SPK_NO_TARGET) {
      existingIct.push({ rowIdx: j + 1, data: ictData[j] });
    }
  }
  
  // Lookup M_ITEM description
  var itemCode = String(found.data[hdr.indexOf('Item_Code')] || '').trim();
  var miData = miSht.getDataRange().getValues();
  var miHdr  = miData[0];
  var miItemIdx = miHdr.indexOf('Item_Code');
  var miDescIdx = miHdr.indexOf('Description');
  var descLookup = '';
  for (var k = 1; k < miData.length; k++) {
    if (String(miData[k][miItemIdx]).trim() === itemCode) {
      descLookup = String(miData[k][miDescIdx]).trim();
      break;
    }
  }
  
  // Build report
  var lines = [];
  lines.push('═══════════════════════════════════════');
  lines.push('INSPECT SPK: ' + SPK_NO_TARGET);
  lines.push('Row di sheet SPK: ' + found.rowIdx);
  lines.push('═══════════════════════════════════════');
  lines.push('');
  lines.push('📋 DATA SPK (semua field non-blank):');
  hdr.forEach(function(h, idx) {
    var v = found.data[idx];
    if (v !== '' && v !== null && v !== undefined) {
      lines.push('  ' + h + ' : ' + v);
    }
  });
  lines.push('');
  lines.push('🎯 KEY FIELDS UNTUK ICT:');
  lines.push('  Selesai_DT     : ' + found.data[hdr.indexOf('Selesai_DT')]);
  lines.push('  Item_Code      : ' + itemCode);
  lines.push('  Description    : ' + descLookup + '  (dari M_ITEM)');
  lines.push('  Owner          : ' + found.data[hdr.indexOf('Owner')]);
  lines.push('  Owner_Used     : ' + found.data[hdr.indexOf('Owner_Used')]);
  lines.push('  Target_Loc     : ' + found.data[hdr.indexOf('Target_Loc')]);
  lines.push('  Qty_Actual     : ' + found.data[hdr.indexOf('Qty_Actual')]);
  lines.push('  KG_Actual      : ' + found.data[hdr.indexOf('KG_Actual')]);
  lines.push('  Status         : ' + found.data[hdr.indexOf('Status')]);
  lines.push('');
  lines.push('📊 EXISTING ROW DI Rekap_ICT: ' + existingIct.length);
  if (existingIct.length > 0) {
    existingIct.forEach(function(e) {
      lines.push('  Row ' + e.rowIdx + ': ' + e.data.join(' | '));
    });
  } else {
    lines.push('  (tidak ada — confirmed missing)');
  }
  lines.push('');
  lines.push('═══════════════════════════════════════');
  lines.push('SUGGESTED BACKFILL VALUE (kalau perlu):');
  lines.push('═══════════════════════════════════════');
  lines.push('  Tgl_Transfer : ' + found.data[hdr.indexOf('Selesai_DT')]);
  lines.push('  SPK_No       : ' + SPK_NO_TARGET);
  lines.push('  Item_Code    : ' + itemCode);
  lines.push('  Description  : ' + descLookup);
  lines.push('  Dari_Owner   : ' + found.data[hdr.indexOf('Owner')]);
  lines.push('  Ke_Owner     : ' + found.data[hdr.indexOf('Owner_Used')]);
  lines.push('  Qty_Sht      : ' + found.data[hdr.indexOf('Qty_Actual')]);
  lines.push('  Qty_KG       : ' + found.data[hdr.indexOf('KG_Actual')]);
  
  Logger.log(lines.join('\n'));
  
  // Show alert with summary
  SpreadsheetApp.getUi().alert(
    'Inspect selesai — cek Execution log untuk detail lengkap.\n\n' +
    'SPK Row #: ' + found.rowIdx + '\n' +
    'Item_Code: ' + itemCode + '\n' +
    'Owner→Owner_Used: ' + found.data[hdr.indexOf('Owner')] + '→' + found.data[hdr.indexOf('Owner_Used')] + '\n' +
    'Target_Loc: ' + found.data[hdr.indexOf('Target_Loc')] + '\n' +
    'Qty: ' + found.data[hdr.indexOf('Qty_Actual')] + ' / KG: ' + found.data[hdr.indexOf('KG_Actual')] + '\n\n' +
    'Existing ICT: ' + existingIct.length
  );
}

/**
 * ═══════════════════════════════════════════════════════════════
 * BACKFILL 1 ROW REKAP_ICT — SHR-260106-01
 * ═══════════════════════════════════════════════════════════════
 * Reason: Owner_Used diedit manual ke DRC setelah Board DONE.
 * writeRekapICT tidak fire ulang → ICT missing.
 *
 * 4 SAFETY CHECK sebelum insert:
 *   1. Verify SPK exists & Status=DONE
 *   2. Verify Owner≠Owner_Used & Target=FG
 *   3. Verify belum ada ICT row untuk SPK ini
 *   4. UI confirmation dialog (Yes/No)
 * ═══════════════════════════════════════════════════════════════
 */
function backfill_ICT_SHR_260106_01() {
  var SPK_NO = 'SHR-260106-01';
  
  // ── Values (from user confirmation) ──
  var PAYLOAD = {
    tgl_transfer : null,              // will auto-fill from Selesai_DT
    spk_no       : SPK_NO,
    item_code    : '2-PBI-033',
    description  : 'SPHC-PO 2,3X90,5X630 mm',
    dari_owner   : 'FC',
    ke_owner     : 'DRC',
    qty_sht      : 100,
    qty_kg       : 103.00
  };
  
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var spkSht = ss.getSheetByName('SPK');
  var ictSht = ss.getSheetByName('Rekap_ICT');
  
  if (!spkSht || !ictSht) {
    ui.alert('❌ Sheet SPK atau Rekap_ICT tidak ditemukan.');
    return;
  }
  
  // ══════════════════════════════════════════════════
  // SAFETY CHECK 1: Verify SPK exists & Status=DONE
  // ══════════════════════════════════════════════════
  var spkData = spkSht.getDataRange().getValues();
  var spkHdr  = spkData[0];
  var iSpk    = spkHdr.indexOf('SPK_No');
  var iType   = spkHdr.indexOf('SPK_Type');
  var iStatus = spkHdr.indexOf('Status');
  var iTgt    = spkHdr.indexOf('Target_Loc');
  var iOwn    = spkHdr.indexOf('Owner');
  var iOwnU   = spkHdr.indexOf('Owner_Used');
  var iItem   = spkHdr.indexOf('Item_Code');
  var iQtyAct = spkHdr.indexOf('Qty_Actual');
  var iKgAct  = spkHdr.indexOf('KG_Actual');
  var iSel    = spkHdr.indexOf('Selesai_DT');
  
  var spkRow = null;
  for (var i = 1; i < spkData.length; i++) {
    if (String(spkData[i][iSpk]).trim() === SPK_NO) {
      spkRow = spkData[i];
      break;
    }
  }
  
  if (!spkRow) {
    ui.alert('❌ SAFETY CHECK 1 FAILED\n\nSPK ' + SPK_NO + ' tidak ditemukan di sheet SPK.');
    return;
  }
  
  var spkStatus = String(spkRow[iStatus] || '').trim().toUpperCase();
  var spkType   = String(spkRow[iType]   || '').trim();
  if (spkStatus !== 'DONE') {
    ui.alert('❌ SAFETY CHECK 1 FAILED\n\nSPK ' + SPK_NO + ' Status = "' + spkStatus + '", bukan DONE.\n\nHanya SPK dengan Status=DONE yang boleh di-backfill ICT.');
    return;
  }
  
  // ══════════════════════════════════════════════════
  // SAFETY CHECK 2: Verify Owner≠Owner_Used & Target=FG
  // ══════════════════════════════════════════════════
  var spkOwner  = String(spkRow[iOwn]  || '').trim();
  var spkOwnerU = String(spkRow[iOwnU] || '').trim();
  var spkTgt    = String(spkRow[iTgt]  || '').trim();
  var spkItem   = String(spkRow[iItem] || '').trim();
  var spkQty    = Number(spkRow[iQtyAct]) || 0;
  var spkKg     = Number(spkRow[iKgAct])  || 0;
  var spkSel    = spkRow[iSel];
  
  var isFG = (spkTgt === 'Stok_FG' || spkTgt.indexOf('FG_') === 0);
  if (spkOwner.toUpperCase() === spkOwnerU.toUpperCase()) {
    ui.alert('❌ SAFETY CHECK 2 FAILED\n\nOwner (' + spkOwner + ') = Owner_Used (' + spkOwnerU + ').\nTidak ada cross-billing, tidak perlu backfill ICT.');
    return;
  }
  if (!isFG) {
    ui.alert('❌ SAFETY CHECK 2 FAILED\n\nTarget_Loc = "' + spkTgt + '" bukan FG.\nRule: ICT hanya untuk hasil FG.');
    return;
  }
  
  // Cross-check payload vs SPK actual
  var mismatches = [];
  if (spkItem !== PAYLOAD.item_code)     mismatches.push('Item_Code: SPK=' + spkItem + ' vs Payload=' + PAYLOAD.item_code);
  if (spkOwner !== PAYLOAD.dari_owner)   mismatches.push('Owner: SPK=' + spkOwner + ' vs Payload=' + PAYLOAD.dari_owner);
  if (spkOwnerU !== PAYLOAD.ke_owner)    mismatches.push('Owner_Used: SPK=' + spkOwnerU + ' vs Payload=' + PAYLOAD.ke_owner);
  if (Math.abs(spkQty - PAYLOAD.qty_sht) > 0.01) mismatches.push('Qty: SPK=' + spkQty + ' vs Payload=' + PAYLOAD.qty_sht);
  if (Math.abs(spkKg  - PAYLOAD.qty_kg)  > 0.01) mismatches.push('KG: SPK=' + spkKg + ' vs Payload=' + PAYLOAD.qty_kg);
  
  if (mismatches.length > 0) {
    var msgM = '⚠️ PAYLOAD MISMATCH DENGAN SPK ACTUAL:\n\n' + mismatches.join('\n') + '\n\nBatalkan dan review payload.';
    ui.alert(msgM);
    return;
  }
  
  // Auto-fill tgl_transfer dari Selesai_DT
  PAYLOAD.tgl_transfer = spkSel || new Date();
  
  // ══════════════════════════════════════════════════
  // SAFETY CHECK 3: Belum ada ICT row untuk SPK ini
  // ══════════════════════════════════════════════════
  var ictData = ictSht.getDataRange().getValues();
  var ictHdr  = ictData[0];
  var jSpk    = ictHdr.indexOf('SPK_No');
  var existingCount = 0;
  for (var k = 1; k < ictData.length; k++) {
    if (String(ictData[k][jSpk]).trim() === SPK_NO) existingCount++;
  }
  if (existingCount > 0) {
    ui.alert('❌ SAFETY CHECK 3 FAILED\n\nSudah ada ' + existingCount + ' row di Rekap_ICT untuk SPK ' + SPK_NO + '.\n\nBackfill dibatalkan untuk cegah duplikat.');
    return;
  }
  
  // ══════════════════════════════════════════════════
  // SAFETY CHECK 4: UI Confirmation Dialog
  // ══════════════════════════════════════════════════
  var previewLines = [
    '📋 PREVIEW ROW YANG AKAN DI-INSERT KE Rekap_ICT:',
    '',
    '  Tgl_Transfer : ' + PAYLOAD.tgl_transfer,
    '  SPK_No       : ' + PAYLOAD.spk_no,
    '  Item_Code    : ' + PAYLOAD.item_code,
    '  Description  : ' + PAYLOAD.description,
    '  Dari_Owner   : ' + PAYLOAD.dari_owner,
    '  Ke_Owner     : ' + PAYLOAD.ke_owner,
    '  Qty_Sht      : ' + PAYLOAD.qty_sht,
    '  Qty_KG       : ' + PAYLOAD.qty_kg,
    '',
    'Semua safety check LULUS. Insert row ini?'
  ];
  
  var response = ui.alert(
    'Konfirmasi Backfill Rekap_ICT',
    previewLines.join('\n'),
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Backfill dibatalkan oleh user.');
    return;
  }
  
  // ══════════════════════════════════════════════════
  // EXECUTE INSERT
  // ══════════════════════════════════════════════════
  var newRow = ictHdr.map(function(h) {
    switch(h) {
      case 'Tgl_Transfer': return PAYLOAD.tgl_transfer;
      case 'SPK_No'      : return PAYLOAD.spk_no;
      case 'Item_Code'   : return PAYLOAD.item_code;
      case 'Description' : return PAYLOAD.description;
      case 'Dari_Owner'  : return PAYLOAD.dari_owner;
      case 'Ke_Owner'    : return PAYLOAD.ke_owner;
      case 'Qty_Sht'     : return PAYLOAD.qty_sht;
      case 'Qty_KG'      : return PAYLOAD.qty_kg;
      default            : return '';
    }
  });
  
  ictSht.appendRow(newRow);
  SpreadsheetApp.flush();
  
  Logger.log('BACKFILL SUCCESS: ' + SPK_NO + ' inserted to Rekap_ICT row ' + ictSht.getLastRow());
  
  ui.alert('✅ SUKSES\n\n1 row berhasil di-insert ke Rekap_ICT (row ' + ictSht.getLastRow() + ').\n\nSPK: ' + SPK_NO + '\nDari: ' + PAYLOAD.dari_owner + ' → Ke: ' + PAYLOAD.ke_owner + '\nQty: ' + PAYLOAD.qty_sht + ' / KG: ' + PAYLOAD.qty_kg);
}