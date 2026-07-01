/**
 * Fungsi untuk mengambil data Master CT dari M_ITEM dan menghitung durasi SPK
 * @param {string} itemCode - Kode barang dari SPK
 * @param {number} qtyTarget - Jumlah lembar target (Qty_Target)
 * @param {number} kgTarget - Jumlah berat target (KG_Target)
 * @return {object} Objek berisi Plan_Setup, Plan_Run, dan Total_Durasi
 */
function hitungRencanaDurasi(itemCode, qtyTarget, kgTarget) {
  // 1. Definisikan default jika data tidak ditemukan (Jaring Pengaman)
  var hasil = {
    planSetup: 15, // default 15 menit
    planRun: 0,
    totalDurasi: 15
  };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetItem = ss.getSheetByName("M_ITEM");
    if (!sheetItem) return hasil; // Jika sheet M_ITEM tidak ada, pakai default
    
    var dataItem = sheetItem.getDataRange().getValues();
    
    // Cari index kolom di M_ITEM berdasarkan header Anda
    var header = dataItem[0];
    var idxItemCode = header.indexOf("Item_Code"); 
    var idxSetup = header.indexOf("Setup_Menit");
    var idxTargetCT = header.indexOf("Target_CT_Jam");
    var idxUomMC = header.indexOf("UoM_MC");
    
    // Loop untuk mencari Item_Code yang cocok
    for (var i = 1; i < dataItem.length; i++) {
      if (dataItem[i][idxItemCode] === itemCode) {
        var setupMenit = Number(dataItem[i][idxSetup]) || 0;
        var targetCT = Number(dataItem[i][idxTargetCT]) || 0;
        var uomMC = String(dataItem[i][idxUomMC]).toUpperCase().trim();
        
        hasil.planSetup = setupMenit;
        
        // 2. Logika Hitung Waktu Running berdasarkan UoM_MC
        if (targetCT > 0) {
          if (uomMC === "SHT" || uomMC === "LBR") {
            // Jika basis mesin adalah lembaran
            hasil.planRun = Math.ceil((qtyTarget / targetCT) * 60);
          } else if (uomMC === "KG") {
            // Jika basis mesin adalah tonase/kilogram (untuk CTL / SLT)
            hasil.planRun = Math.ceil((kgTarget / targetCT) * 60);
          }
        }
        break; // Berhenti loop jika sudah ketemu
      }
    }
    
    // Hitung total durasi
    hasil.totalDurasi = hasil.planSetup + hasil.planRun;
    return hasil;
    
  } catch(e) {
    Logger.log("Error hitung durasi: " + e.toString());
    return hasil; // Kembalikan default jika ada error script
  }
}

/**
 * Fungsi Pintar untuk Menambahkan Menit sesuai Jam Kerja Pabrik
 * Mengabaikan Istirahat, Sabtu-Minggu, dan Libur Nasional
 */
function addWorkingMinutes(startDate, durationMinutes) {
  var d = new Date(startDate.getTime());
  
  // DAFTAR HARI LIBUR NASIONAL (Format: YYYY-MM-DD)
  // Bisa ditambahkan manual setiap pergantian tahun jika ada tanggal merah baru
  var liburNasional = [
    "2026-05-14", 
    "2026-05-24", 
    "2026-06-01", 
    "2026-08-17"  
  ];

  function isHoliday(dateObj) {
    var tz = Session.getScriptTimeZone();
    var ymd = Utilities.formatDate(dateObj, tz, "yyyy-MM-dd");
    return liburNasional.indexOf(ymd) !== -1;
  }

  function getNextWorkingMinute(dateObj) {
    while (true) {
      var day = dateObj.getDay(); // 0=Minggu, 1=Senin ... 5=Jumat, 6=Sabtu
      var h = dateObj.getHours();
      var m = dateObj.getMinutes();
      var t = h * 100 + m; // Format angka waktu (07:30 = 730)

      // 1. Cek Libur / Weekend (Sabtu & Minggu)
      if (day === 0 || day === 6 || isHoliday(dateObj)) {
        dateObj.setDate(dateObj.getDate() + 1); 
        dateObj.setHours(7, 30, 0, 0); 
        continue;
      }

      // 2. Cek Jam Kerja SENIN s/d KAMIS
      if (day >= 1 && day <= 4) {
        if (t >= 730 && t < 1140) return; // Waktu Kerja Pagi (07:30 - 11:40)
        
        if (t >= 1140 && t < 1240) {
          dateObj.setHours(12, 40, 0, 0); // Sedang Istirahat, lompat ke 12:40
          continue;
        }
        
        if (t >= 1240 && t < 1630) return; // Waktu Kerja Siang (12:40 - 16:30)
        
        if (t >= 1630) {
          dateObj.setDate(dateObj.getDate() + 1); // Sudah pulang, lompat ke besok pagi
          dateObj.setHours(7, 30, 0, 0);
          continue;
        }
        
        if (t < 730) {
          dateObj.setHours(7, 30, 0, 0); // Terlalu pagi, paskan ke jam masuk
          continue;
        }
      }

      // 3. Cek Jam Kerja JUMAT
      if (day === 5) {
        if (t >= 730 && t < 1130) return; // Kerja Pagi Jumat (07:30 - 11:30)
        
        if (t >= 1130 && t < 1300) {
          dateObj.setHours(13, 0, 0, 0); // Istirahat Jumatan, lompat ke 13:00
          continue;
        }
        
        if (t >= 1300 && t < 1700) return; // Kerja Siang Jumat (13:00 - 17:00)
        
        if (t >= 1700) {
          dateObj.setDate(dateObj.getDate() + 1); // Pulang Jumat, lompat ke Sabtu (nanti disaring di atas)
          dateObj.setHours(7, 30, 0, 0);
          continue;
        }
        
        if (t < 730) {
          dateObj.setHours(7, 30, 0, 0); 
          continue;
        }
      }
    }
  }

  // Validasi awal agar waktu mulai jatuh di dalam jam aktif kerja
  getNextWorkingMinute(d);

  // Jalankan penambahan menit demi menit
  while (durationMinutes > 0) {
    d.setMinutes(d.getMinutes() + 1);
    getNextWorkingMinute(d);
    durationMinutes--;
  }
  
  return d;
}

/**
 * Fungsi Utama untuk Estafet Jadwal — BATCH WRITE v2
 * PERUBAHAN: LockService dihapus dari fungsi ini.
 * Lock diurus oleh caller (savePlanningBoard, saveSPK_CTL, dll).
 * Memanggil lock dari dalam fungsi yang sudah lock = silent fail.
 */
function kalkulasiEstimasiWaktu() {
  // ── TIDAK ADA LockService di sini ──
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName("SPK");
    var data    = sheet.getDataRange().getValues(); // READ SEKALI
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var numRows = data.length;

    var iSpk         = headers.indexOf("SPK_No");
    var iMc          = headers.indexOf("MC_No");
    var iType        = headers.indexOf("SPK_Type");
    var iParent      = headers.indexOf("Parent_SPK");
    var iStatus      = headers.indexOf("Status");
    var iDurasi      = headers.indexOf("Total_Durasi_Menit");
    var iEstMulai    = headers.indexOf("Estimasi_Jam_Mulai");
    var iEstSelesai  = headers.indexOf("Estimasi_Jam_Selesai");
    var iBaseMulai   = headers.indexOf("Baseline_Mulai");
    var iBaseSelesai = headers.indexOf("Baseline_Selesai");
    var iPrio        = headers.indexOf("Priority");
    var iPlanSeq     = headers.indexOf("Plan_Seq");

    if (iMc === -1 || iEstMulai === -1 || iType === -1 || iSpk === -1) return;

    // ── LANGKAH 1: Bangun map referensi ──
    var parentMap = {};
    var typeMap   = {};
    for (var i = 1; i < data.length; i++) {
      var spkNo   = String(data[i][iSpk]    || '').trim();
      var spkType = String(data[i][iType]   || '').trim();
      var parent  = String(data[i][iParent] || '').trim();
      if (spkNo) { parentMap[spkNo] = parent; typeMap[spkNo] = spkType; }
    }

    // ── LANGKAH 2: Dependency SHR-HEADER → CTL-HEADER ──
    var shrCtlDep = {};
    for (var spk in typeMap) {
      if (typeMap[spk] === 'SHR-HEADER') {
        var ctlOutNo    = parentMap[spk]      || '';
        var ctlHeaderNo = parentMap[ctlOutNo] || '';
        if (typeMap[ctlOutNo] === 'CTL-OUT' && typeMap[ctlHeaderNo] === 'CTL-HEADER') {
          shrCtlDep[spk] = ctlHeaderNo;
        }
      }
    }

    // ── LANGKAH 3: Kelompokkan antrian aktif per mesin (HEADER only) ──
    var mesinMap = {};
    for (var i = 1; i < data.length; i++) {
      var spkType = String(data[i][iType]   || '').toUpperCase().trim();
      var status  = String(data[i][iStatus] || '').toUpperCase();
      var mc      = String(data[i][iMc]     || '').trim();
      var durasi  = Number(data[i][iDurasi]) || 0;
      var spkNo   = String(data[i][iSpk]    || '').trim();

      if (spkType.indexOf('-HEADER') === -1) continue;
      if (status !== 'ANTRIAN' && status !== 'RUNNING') continue;
      if (!mc || durasi <= 0 || !spkNo) continue;

      var planSeq = iPlanSeq !== -1 ? (Number(data[i][iPlanSeq]) || 0) : 0;
      if (!mesinMap[mc]) mesinMap[mc] = [];
      mesinMap[mc].push({
        rowIdx      : i + 1,
        spkNo       : spkNo,
        status      : status,
        prio        : String(data[i][iPrio] || '').toUpperCase(),
        planSeq     : planSeq,
        durasi      : durasi,
        hasBaseline : String(data[i][iBaseMulai] || '').trim() !== '',
        earliestStart: null
      });
    }

    // ── LANGKAH 4: Sort mesin — CTL dulu, lalu SHR, SLT ──
    var mesinKeys = Object.keys(mesinMap);
    mesinKeys.sort(function(a, b) {
      function urutan(mc) {
        var u = mc.toUpperCase();
        if (u.indexOf('CTL') !== -1) return 0;
        if (u.indexOf('SHR') !== -1) return 1;
        if (u.indexOf('SLT') !== -1) return 2;
        return 3;
      }
      return urutan(a) - urutan(b);
    });

    var wktSekarang       = new Date();
    var tz                = Session.getScriptTimeZone();
    var completionTimeMap = {};

    // ── LANGKAH 5: Greedy Scheduling — HITUNG DI MEMORY ──
    for (var mIdx = 0; mIdx < mesinKeys.length; mIdx++) {
      var mc      = mesinKeys[mIdx];
      var antrean = mesinMap[mc];

      antrean.sort(function(a, b) {
        if (a.status === 'RUNNING' && b.status !== 'RUNNING') return -1;
        if (b.status === 'RUNNING' && a.status !== 'RUNNING') return  1;
        var aSeq = a.planSeq, bSeq = b.planSeq;
        if (aSeq > 0 && bSeq > 0) return aSeq - bSeq;
        if (aSeq > 0 && bSeq === 0) return -1;
        if (aSeq === 0 && bSeq > 0) return  1;
        var wA = (a.prio === 'URGENT') ? 2 : (a.prio === 'HIGH' ? 1 : 0);
        var wB = (b.prio === 'URGENT') ? 2 : (b.prio === 'HIGH' ? 1 : 0);
        return wB - wA;
      });

      antrean.forEach(function(item) {
        var ctlParentNo = shrCtlDep[item.spkNo] || null;
        if (ctlParentNo && completionTimeMap[ctlParentNo]) {
          item.earliestStart = completionTimeMap[ctlParentNo];
        }
      });

      var remaining      = antrean.slice();
      var waktuMesinFree = new Date(wktSekarang.getTime());

      while (remaining.length > 0) {
        var tersedia = remaining.filter(function(item) {
          return !item.earliestStart || item.earliestStart <= waktuMesinFree;
        });

        var pilihan;
        if (tersedia.length > 0) {
          pilihan = tersedia[0];
        } else {
          var nextTime = null;
          remaining.forEach(function(item) {
            if (item.earliestStart && (!nextTime || item.earliestStart < nextTime)) {
              nextTime = item.earliestStart;
            }
          });
          if (!nextTime) break;
          waktuMesinFree = new Date(nextTime.getTime());
          continue;
        }

        var tMulaiAkurat = addWorkingMinutes(waktuMesinFree, 0);
        var tSelesai     = addWorkingMinutes(tMulaiAkurat, pilihan.durasi);

        var formatMulai   = Utilities.formatDate(tMulaiAkurat, tz, "dd MMM HH:mm");
        var formatSelesai = Utilities.formatDate(tSelesai,     tz, "dd MMM HH:mm");

        // Simpan ke memory
        var dIdx = pilihan.rowIdx - 1;
        data[dIdx][iEstMulai]   = formatMulai;
        data[dIdx][iEstSelesai] = formatSelesai;

        if (!pilihan.hasBaseline && iBaseMulai !== -1 && iBaseSelesai !== -1) {
          data[dIdx][iBaseMulai]   = formatMulai;
          data[dIdx][iBaseSelesai] = formatSelesai;
        }

        completionTimeMap[pilihan.spkNo] = tSelesai;
        waktuMesinFree = tSelesai;
        remaining = remaining.filter(function(it) { return it.spkNo !== pilihan.spkNo; });
      }
    }

    // ── LANGKAH 6: BATCH WRITE — 4 API calls total ──
    if (numRows > 1) {
      sheet.getRange(1, iEstMulai   + 1, numRows, 1)
           .setValues(data.map(function(r) { return [r[iEstMulai]]; }));
      sheet.getRange(1, iEstSelesai + 1, numRows, 1)
           .setValues(data.map(function(r) { return [r[iEstSelesai]]; }));

      if (iBaseMulai !== -1 && iBaseSelesai !== -1) {
        sheet.getRange(1, iBaseMulai   + 1, numRows, 1)
             .setValues(data.map(function(r) { return [r[iBaseMulai]]; }));
        sheet.getRange(1, iBaseSelesai + 1, numRows, 1)
             .setValues(data.map(function(r) { return [r[iBaseSelesai]]; }));
      }

      SpreadsheetApp.flush(); // ✅ Pastikan semua tulis committed
    }

  } catch (e) {
    Logger.log("Error Kalkulasi Jadwal: " + e.message);
  }
  // ── TIDAK ADA finally lock.releaseLock() ──
}