/**
 * Script ini menangani request POST yang masuk ke Web App (Google Apps Script).
 * Bertindak sebagai REST API untuk menyimpan data ke Google Spreadsheet.
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var timestamp = payload.timestamp;
    var noCoin = payload.no_coin;
    var kodeAlat = payload.kode_alat;
    var status = "Dipinjam";
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data_Peminjaman");
    
    if (!sheet) {
      sheet = ss.insertSheet("Data_Peminjaman");
      sheet.appendRow(["Timestamp", "No Coin", "Kode Alat", "Status"]);
    }
    
    sheet.appendRow([timestamp, noCoin, kodeAlat, status]);
    
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success", 
      "message": "Data peminjaman berhasil ditambahkan"
    })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Menangani request GET (mengembalikan data JSON).
 * Ini akan otomatis membolehkan CORS dari browser karena GAS Web App me-redirect output.
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPeminjaman = ss.getSheetByName("Data_Peminjaman");
    var sheetMahasiswa = ss.getSheetByName("Data_Mahasiswa");
    
    if (!sheetPeminjaman) {
      return ContentService.createTextOutput(JSON.stringify({"status": "success", "data": []}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var dataPeminjaman = sheetPeminjaman.getDataRange().getValues();
    var dataMahasiswa = sheetMahasiswa ? sheetMahasiswa.getDataRange().getValues() : [];
    
    // Buat map (kamus) untuk mahasiswa: No Coin -> Nama Lengkap
    var mhsMap = {};
    if (dataMahasiswa.length > 1) {
      for (var i = 1; i < dataMahasiswa.length; i++) {
        var noCoin = dataMahasiswa[i][0];
        var nama = dataMahasiswa[i][1];
        if (noCoin) {
          mhsMap[noCoin] = nama;
        }
      }
    }
    
    var activeLoans = [];
    if (dataPeminjaman.length > 1) {
      // Loop dari baris 2 (indeks 1) ke bawah
      for (var j = 1; j < dataPeminjaman.length; j++) {
        var status = dataPeminjaman[j][3];
        if (status === "Dipinjam") {
          var coin = dataPeminjaman[j][1];
          activeLoans.push({
            timestamp: dataPeminjaman[j][0],
            no_coin: coin,
            nama: mhsMap[coin] || "Nama tidak ditemukan",
            kode_alat: dataPeminjaman[j][2]
          });
        }
      }
    }
    
    // Urutkan data terbaru di atas (reverse)
    activeLoans.reverse();
    
    return ContentService.createTextOutput(JSON.stringify({"status": "success", "data": activeLoans}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
