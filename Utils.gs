/* =========================================================================
 * UTILS.GS — Generic getter functions dipakai banyak modul
 * - getSheetData(name)     : ambil mentah isi sheet sebagai array of object
 * - getSPKList(filter)     : ambil SPK dengan optional filter
 * - getStockLive()         : ambil 4 sheet stock sekaligus (1 round trip)
 * Single source of truth — tidak boleh ada duplikat di file lain.
 * ========================================================================= */

function getSheetData(sheetName) {
  var sheet = getSheet(sheetName);
  if (!sheet) return []; 
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Jika cuma ada header, kembalikan kosong
  
  // Trim spasi tersembunyi di semua header
  var headers = data[0].map(function(h) {
    return String(h).trim();
  });
  
  var result = [];
  
  // Looping dari baris ke-2 sampai bawah
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // FITUR KEBAL: Abaikan jika baris tersebut benar-benar kosong semua
    if(row.join("").trim() === "") continue; 
    
    var obj = {};
    headers.forEach(function(h, idx) {
      var val = row[idx];

      // Konversi Date object ke string agar tidak error di JSON
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy');
      }

      // Trim string values
      if (typeof val === 'string') {
        val = val.trim();
      }

      obj[h] = val;
    });
    result.push(obj);
  }
  
  return result;
}


function getSPKList(filter) {
  const sheet = getSheet(SHEET_NAMES.SPK);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data[0].map(function(h){ return String(h).trim(); });
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach(function(h, idx) { 
      let val = data[i][idx];
      // Date object → ISO string untuk JSON-safe transfer
      if (val instanceof Date) val = val.toISOString();
      row[h] = val; 
    });
    
    if (filter) {
      if (filter.status  && row.Status   !== filter.status)         continue;
      if (filter.type    && !String(row.SPK_Type||'').includes(filter.type)) continue;
      if (filter.machine && row.MC_No    !== filter.machine)        continue;
    }
    result.push(row);
  }
  return result;
}


function getStockLive() {
  return {
    coil  : getSheetData(SHEET_NAMES.STOK_COIL),
    sheet : getSheetData(SHEET_NAMES.STOK_SHEET),
    wip   : getSheetData(SHEET_NAMES.STOK_WIP),
    fg    : getSheetData(SHEET_NAMES.STOK_FG)
  };
}