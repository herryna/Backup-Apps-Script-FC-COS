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