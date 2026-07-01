/* =========================================================================
 * DASHBOARD_SERVICE.GS — Executive Dashboard Aggregator
 * Single round-trip endpoint untuk Operations Dashboard.
 * 
 * Public:
 *   getDashboardData(periodYYYYMM)
 *     periodYYYYMM: '2026-07' format. Default = current month.
 *     return: { period, ppic, production, warehouse, charts }
 * 
 *   getDashboardPeriodOptions()
 *     return: array of YYYY-MM dari 2026-01 sampai bulan ini
 * 
 * Catatan:
 *   - Warehouse data SELALU real-time (tidak terpengaruh period)
 *   - PPIC achievement Period-aware (SO carry-over dari bulan lalu jika OPEN)
 *   - Production aggregation HEADER only (skip OUT)
 *   - Aging threshold: Coil/Sheet >60d, WIP >30d, FG >90d
 * ========================================================================= */

// =========================================================================
// 1. ENTRY POINTS
// =========================================================================
function getDashboardData(periodInput) {
  var period = _dashNormPeriod(periodInput);
  var range  = _dashRangeFromPeriod(period);

  var spkRows  = _dashReadSheet('SPK');
  var soRows   = _dashReadSheet('SO');
  var delvRows = _dashReadSheet('DELV');
  var stock    = _dashReadStockSheets();

  return {
    period:     _dashPeriodInfo(period, range),
    ppic:       _dashCalcPpic(period, range, soRows, delvRows),
    production: _dashCalcProduction(range, spkRows),
    warehouse:  _dashCalcWarehouse(stock),
    charts:     _dashBuildCharts(period, range, soRows, delvRows)
  };
}

function getDashboardPeriodOptions() {
  // List YYYY-MM dari 2026-01 sampai bulan berjalan, descending (terbaru di atas)
  var startYear = 2026, startMonth = 1;
  var now = new Date();
  var endYear = now.getFullYear(), endMonth = now.getMonth() + 1;
  var list = [];
  for (var y = endYear; y >= startYear; y--) {
    var monthEnd = (y === endYear) ? endMonth : 12;
    var monthStart = (y === startYear) ? startMonth : 1;
    for (var m = monthEnd; m >= monthStart; m--) {
      list.push(y + '-' + String(m).padStart(2, '0'));
    }
  }
  return list;
}


// =========================================================================
// 2. DATE & PERIOD HELPERS
// =========================================================================
function _dashNormPeriod(input) {
  // Default ke bulan ini kalau input kosong/invalid
  if (input && /^\d{4}-(0[1-9]|1[0-2])$/.test(String(input))) {
    return String(input);
  }
  var now = new Date();
  return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
}

function _dashRangeFromPeriod(period) {
  var parts = period.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var from = new Date(y, m, 1, 0, 0, 0, 0);
  var to   = new Date(y, m + 1, 0, 23, 59, 59, 999); // last day of month
  return { from: from, to: to };
}

function _dashPeriodInfo(period, range) {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  // Hitung minggu berjalan (1-4) untuk period jika period == bulan ini
  var weekNum = null;
  var isCurrent = (
    range.from.getFullYear() === now.getFullYear() &&
    range.from.getMonth() === now.getMonth()
  );
  if (isCurrent) {
    weekNum = Math.min(Math.ceil(now.getDate() / 7), 4);
  }
  return {
    period:     period,
    label:      Utilities.formatDate(range.from, tz, 'MMMM yyyy'),
    from:       Utilities.formatDate(range.from, tz, 'dd MMM yyyy'),
    to:         Utilities.formatDate(range.to,   tz, 'dd MMM yyyy'),
    is_current: isCurrent,
    week_num:   weekNum,       // null kalau bukan bulan berjalan
    expected_pct: weekNum ? weekNum * 25 : 100  // W1=25, W2=50, W3=75, W4=100
  };
}

function _dashInRange(d, range) {
  if (!d) return false;
  if (!(d instanceof Date)) { d = new Date(d); if (isNaN(d.getTime())) return false; }
  return d >= range.from && d <= range.to;
}

function _dashDaysSince(d) {
  if (!d) return null;
  if (!(d instanceof Date)) d = new Date(d);
  if (isNaN(d.getTime())) return null;
  return Math.floor((new Date() - d) / 86400000);
}


// =========================================================================
// 3. SHEET READERS
// =========================================================================
function _dashReadSheet(sheetName) {
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

function _dashReadStockSheets() {
  return {
    coil:  _dashReadSheet('Stok_Coil'),
    sheet: _dashReadSheet('Stok_Sheet'),
    wip:   _dashReadSheet('Stok_WIP'),
    fg:    _dashReadSheet('Stok_FG')
  };
}


// =========================================================================
// 4. PPIC ACTIVE SO LOGIC (core logic Period-aware)
// =========================================================================
function _dashIsActiveSO(soRow, periodYYYYMM) {
  // Return true jika SO ini "aktif" di period yang sedang dilihat:
  //   1. SO_Period == period (SO bulan ini)
  //   2. SO_Period < period AND STATUS == OPEN (carry-over dari bulan lalu)
  var soPeriod = String(soRow.SO_Period || '').trim();
  var status   = String(soRow.STATUS || soRow.Status || '').toUpperCase();
  
  if (status === 'CANCELLED') return false;
  if (!soPeriod) return false;  // SO tanpa period di-skip
  
  if (soPeriod === periodYYYYMM) return true;
  if (soPeriod < periodYYYYMM && status === 'OPEN') return true;
  return false;
}

function _dashCalcPpic(period, range, soRows, delvRows) {
  // SO Aktif di period ini
  var activeSos = soRows.filter(function(s){ return _dashIsActiveSO(s, period); });
  
  // Total SO KG aktif
  var totalSoKg = 0;
  activeSos.forEach(function(s) {
    totalSoKg += parseFloat(s.SO_KG || (parseFloat(s.SO_Q || 0) * parseFloat(s['Wg/Pce FC'] || 0)));
  });
  
  // Delta vs prev period (bulan sebelumnya — sama logic carry-over)
  var prevPeriod = _dashPrevPeriod(period);
  var prevActive = soRows.filter(function(s){ return _dashIsActiveSO(s, prevPeriod); });
  var prevSoKg = 0;
  prevActive.forEach(function(s) {
    prevSoKg += parseFloat(s.SO_KG || (parseFloat(s.SO_Q || 0) * parseFloat(s['Wg/Pce FC'] || 0)));
  });
  var deltaPct = prevSoKg > 0 ? Math.round((totalSoKg - prevSoKg) / prevSoKg * 100) : (totalSoKg > 0 ? 100 : 0);
  
  // Delivered KG di period ini — DELV with Tanggal in range
  var deliveredKg = 0;
  delvRows.forEach(function(d) {
    if (!_dashInRange(d.Tanggal, range)) return;
    deliveredKg += parseFloat(d.Delv_KG || 0);
  });
  
  // Backlog KG = SUM(SO_KG - Delv_KG) untuk SO aktif STATUS = OPEN
  var backlogKg = 0;
  activeSos.forEach(function(s) {
    var status = String(s.STATUS || s.Status || '').toUpperCase();
    if (status !== 'OPEN') return;
    var soKg = parseFloat(s.SO_KG || (parseFloat(s.SO_Q || 0) * parseFloat(s['Wg/Pce FC'] || 0)));
    var delKg = parseFloat(s.Delv_KG || 0);
    var sisa = soKg - delKg;
    if (sisa > 0) backlogKg += sisa;
  });
  
  // Completion % = Delivered / Total SO KG aktif
  var completionPct = totalSoKg > 0 ? Math.round(deliveredKg / totalSoKg * 100 * 10) / 10 : 0;
  
  // Behind Schedule list
  var now = new Date(), tz = Session.getScriptTimeZone();
  var behind = [], seen = {};
  activeSos.forEach(function(s) {
    var status = String(s.STATUS || s.Status || '').toUpperCase();
    if (status !== 'OPEN') return;
    var sched = s.SCHEDULE_DATE;
    if (!(sched instanceof Date)) sched = new Date(sched);
    if (isNaN(sched.getTime()) || sched >= now) return;
    var soKg = parseFloat(s.SO_KG || (parseFloat(s.SO_Q || 0) * parseFloat(s['Wg/Pce FC'] || 0)));
    var delKg = parseFloat(s.Delv_KG || 0);
    var sisaKg = soKg - delKg;
    if (sisaKg <= 0) return;
    var no = String(s.SO_No || '').trim();
    if (!no || seen[no]) return;
    seen[no] = true;
    behind.push({
      so_no: no,
      cust: String(s.Cust || ''),
      sched: Utilities.formatDate(sched, tz, 'dd MMM yyyy'),
      overdue_days: Math.floor((now - sched) / 86400000),
      sisa_kg: Math.round(sisaKg)
    });
  });
  behind.sort(function(a, b){ return b.overdue_days - a.overdue_days; });
  
  return {
    total_so:        { kg: Math.round(totalSoKg), delta_pct: deltaPct },
    delivered:       { kg: Math.round(deliveredKg) },
    backlog:         { kg: Math.round(backlogKg) },
    completion:      { pct: completionPct },
    behind_schedule: behind
  };
}

function _dashPrevPeriod(period) {
  var parts = period.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  m -= 1;
  if (m === 0) { m = 12; y -= 1; }
  return y + '-' + String(m).padStart(2, '0');
}


// =========================================================================
// 5. CHART DATA BUILDERS (per customer)
// =========================================================================
function _dashBuildCharts(period, range, soRows, delvRows) {
  // Aggregate per customer untuk SO aktif period ini
  var activeSos = soRows.filter(function(s){ return _dashIsActiveSO(s, period); });
  
  // Map: cust → { soKg, delvKg, backlogOnTimeKg, backlogOverdueKg }
  var byCust = {};
  
  // Initialize from active SO
  activeSos.forEach(function(s) {
    var cust = String(s.Cust || '').trim();
    if (!cust) return;
    if (!byCust[cust]) byCust[cust] = { soKg: 0, delvKg: 0, ontimeKg: 0, overdueKg: 0 };
    
    var soKg = parseFloat(s.SO_KG || (parseFloat(s.SO_Q || 0) * parseFloat(s['Wg/Pce FC'] || 0)));
    var delKg = parseFloat(s.Delv_KG || 0);
    var status = String(s.STATUS || s.Status || '').toUpperCase();
    
    byCust[cust].soKg += soKg;
    
    if (status === 'OPEN') {
      var sisa = soKg - delKg;
      if (sisa <= 0) return;
      var sched = s.SCHEDULE_DATE;
      if (!(sched instanceof Date)) sched = new Date(sched);
      var now = new Date();
      if (!isNaN(sched.getTime()) && sched < now) {
        byCust[cust].overdueKg += sisa;
      } else {
        byCust[cust].ontimeKg += sisa;
      }
    }
  });
  
  // Delivered KG di period ini — sum dari DELV by Cust
  delvRows.forEach(function(d) {
    if (!_dashInRange(d.Tanggal, range)) return;
    var cust = String(d.Cust || '').trim();
    if (!cust || !byCust[cust]) return;  // Only count delivery untuk customer yg punya SO aktif
    byCust[cust].delvKg += parseFloat(d.Delv_KG || 0);
  });
  
  // Build achievement array (ASC by pct — critical di kiri)
  var achievement = Object.keys(byCust).map(function(c) {
    var x = byCust[c];
    var pct = x.soKg > 0 ? (x.delvKg / x.soKg * 100) : 0;
    return {
      cust: c,
      pct:  Math.round(pct * 10) / 10,
      so_kg: Math.round(x.soKg),
      delv_kg: Math.round(x.delvKg)
    };
  }).filter(function(x){ return x.so_kg > 0; });  // Hide customer tanpa SO
  
  achievement.sort(function(a, b){ return a.pct - b.pct; });
  
  // Build backlog array (DESC by overdue — critical di kiri)
  var backlog = Object.keys(byCust).map(function(c) {
    var x = byCust[c];
    return {
      cust: c,
      ontime_ton:  Math.round(x.ontimeKg / 1000 * 100) / 100,
      overdue_ton: Math.round(x.overdueKg / 1000 * 100) / 100  // positif (akan dirender minus di frontend)
    };
  }).filter(function(x){ return (x.ontime_ton > 0 || x.overdue_ton > 0); });
  
  backlog.sort(function(a, b){ return b.overdue_ton - a.overdue_ton; });
  
  return {
    achievement: achievement,
    backlog:     backlog
  };
}


// =========================================================================
// 6. PRODUCTION KPI
// =========================================================================
function _dashCalcProduction(range, spkRows) {
  var done = spkRows.filter(function(s) {
    var type = String(s.SPK_Type || '').toUpperCase();
    if (type.indexOf('OUT') >= 0) return false;
    if (type.indexOf('ALO') >= 0 || type.indexOf('ALLOC') >= 0) return false;
    if (String(s.Status || '').toUpperCase() !== 'DONE') return false;
    return _dashInRange(s.Selesai_DT, range);
  });

  var outputKg = 0, planKg = 0, ngKg = 0;
  var byProcess = { CTL: 0, SHR: 0, SLT: 0 };
  var byMachine = {};

  done.forEach(function(s) {
    var actual = parseFloat(s.KG_Actual || 0);
    var target = parseFloat(s.KG_Target || 0);
    var ng     = parseFloat(s.KG_NG     || 0);
    var type   = String(s.SPK_Type || '').toUpperCase();
    var mc     = String(s.MC_No    || '').trim();

    outputKg += actual;
    planKg   += target;
    ngKg     += ng;

    if      (type.indexOf('CTL') >= 0) byProcess.CTL += actual;
    else if (type.indexOf('SHR') >= 0) byProcess.SHR += actual;
    else if (type.indexOf('SLT') >= 0) byProcess.SLT += actual;

    if (mc) {
      if (!byMachine[mc]) byMachine[mc] = { actual: 0, target: 0 };
      byMachine[mc].actual += actual;
      byMachine[mc].target += target;
    }
  });

  var perMachine = [];
  Object.keys(byMachine).sort().forEach(function(mc) {
    var m = byMachine[mc];
    perMachine.push({
      mc: mc,
      actual: Math.round(m.actual),
      target: Math.round(m.target),
      pct:    m.target > 0 ? Math.round(m.actual / m.target * 100) : 0
    });
  });

  return {
    output:         { kg: Math.round(outputKg) },
    plan_vs_actual: {
      plan_kg:   Math.round(planKg),
      actual_kg: Math.round(outputKg),
      pct:       planKg > 0 ? Math.round(outputKg / planKg * 100) : 0
    },
    ng: {
      kg:  Math.round(ngKg),
      pct: outputKg > 0 ? Math.round(ngKg / outputKg * 1000) / 10 : 0
    },
    by_process:  { CTL: Math.round(byProcess.CTL), SHR: Math.round(byProcess.SHR), SLT: Math.round(byProcess.SLT) },
    per_machine: perMachine
  };
}


// =========================================================================
// 7. WAREHOUSE KPI (always real-time)
// =========================================================================
function _dashCalcWarehouse(stock) {
  var fg = stock.fg || [];
  return {
    coil:     _dashAgingBucket(stock.coil,  'Tgl_Masuk',  'KG_Avail', 60),
    sheet:    _dashAgingBucket(stock.sheet, 'Tgl_Masuk',  'Kg_Avail', 60),
    wip:      _dashAgingBucket(stock.wip,   'Tgl_Masuk',  'KG_Avail', 30),
    fg_cust:  _dashAgingBucket(fg.filter(function(r){ return String(r.Target_Loc||'').trim() === 'FG_Cust'; }),         'Tgl_Output', 'KG_Avail', 90),
    fg_stamp: _dashAgingBucket(fg.filter(function(r){ return String(r.Target_Loc||'').trim() === 'FG_RM_Stamping'; }),  'Tgl_Output', 'KG_Avail', 90)
  };
}

function _dashAgingBucket(rows, dateField, kgField, deadThreshold) {
  var total = 0, lt30 = 0, m30_60 = 0, m60_90 = 0, gt90 = 0, dead = 0, batches = 0;
  rows.forEach(function(r) {
    var kg = parseFloat(r[kgField] || 0);
    if (kg <= 0) return;
    var days = _dashDaysSince(r[dateField]);
    if (days === null) return;
    total += kg;
    batches++;
    if      (days < 30) lt30   += kg;
    else if (days < 60) m30_60 += kg;
    else if (days < 90) m60_90 += kg;
    else                gt90   += kg;
    if (days > deadThreshold) dead += kg;
  });
  return {
    total_kg: Math.round(total),
    batches:  batches,
    aging: {
      lt30:   Math.round(lt30),
      m30_60: Math.round(m30_60),
      m60_90: Math.round(m60_90),
      gt90:   Math.round(gt90),
      dead:   Math.round(dead)
    },
    dead_threshold: deadThreshold
  };
}