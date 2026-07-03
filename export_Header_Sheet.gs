function exportHeadersAndFormulasToTxt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  let txtContent = "DATA SHEET, HEADER KOLOM, DAN FORMULA\n";
  txtContent += "========================================\n\n";

  // Looping untuk membaca setiap sheet
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastCol = sheet.getLastColumn();
    let headerList = [];

    // Cek isi header dari kolom 1 sampai kolom terakhir
    if (lastCol > 0) {
      const range = sheet.getRange(1, 1, 1, lastCol);
      const values = range.getValues()[0];
      const formulas = range.getFormulas()[0];
      
      for (let i = 0; i < lastCol; i++) {
        let val = values[i].toString().trim();
        let form = formulas[i].toString().trim();
        
        if (val !== "" || form !== "") {
          if (form !== "") {
            headerList.push(`${val}  --->  [Formula: ${form}]`);
          } else {
            headerList.push(`${val}`);
          }
        }
      }
    }

    // Hanya tulis data ke file TXT JIKA headerList tidak kosong
    if (headerList.length > 0) {
      txtContent += `Nama Sheet : ${sheetName}\n`;
      txtContent += `Header     : \n  - ${headerList.join("\n  - ")}\n`;
      txtContent += "----------------------------------------\n";
    }
  });

  // --- KUNCI PERUBAHAN ADA DI SINI ---
  
  // 1. Ambil waktu saat ini dan ubah formatnya
  // Format: Tahun-Bulan-Tanggal Jam.Menit (contoh: 2026-07-01 23.57)
  const zonaWaktu = Session.getScriptTimeZone();
  const tglExport = Utilities.formatDate(new Date(), zonaWaktu, "yyyy-MM-dd HH.mm");

  // 2. Tentukan nama file yang akan dibuat dengan tambahan tanggal
  const namaFile = ss.getName() + " - Struktur Database - " + tglExport + ".txt";

  // Buat file TXT tersebut dan simpan ke Google Drive
  DriveApp.createFile(namaFile, txtContent, MimeType.PLAIN_TEXT);
  
  // Tampilkan notifikasi Toast (muncul di pojok kanan bawah, tidak bikin timeout)
  SpreadsheetApp.getActiveSpreadsheet().toast("File TXT beserta tanggal berhasil dibuat di Google Drive!", "Berhasil", 10);
}