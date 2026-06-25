/**
 * Sistem terpusat menggunakan doGet agar tidak ada kendala CORS (Cross-Origin) dari browser.
 * Semua request (Pinjam, Kembali, dan Ambil Data) akan masuk ke sini melalui URL parameter.
 */
function doGet(e) {
  // CORS Headers
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPeminjaman = ss.getSheetByName("Data_Peminjaman");
    var sheetMahasiswa = ss.getSheetByName("Data_Mahasiswa");
    
    // Pastikan sheet ada
    if (!sheetPeminjaman) {
      sheetPeminjaman = ss.insertSheet("Data_Peminjaman");
      sheetPeminjaman.appendRow(["Timestamp Pinjam", "No Coin", "Nama Peminjam", "Kode Alat", "Status", "Timestamp Kembali"]);
    }
    
    var action = e.parameter.action;
    
    // ==========================================
    // ACTION: PINJAM ALAT
    // ==========================================
    if (action === "pinjam") {
      var timestamp = e.parameter.timestamp;
      var noCoin = e.parameter.no_coin;
      var kodeAlat = e.parameter.kode_alat;
      var status = "Dipinjam";
      
      var namaMhs = null;
      if (sheetMahasiswa) {
        var dataMhs = sheetMahasiswa.getDataRange().getValues();
        for (var i = 1; i < dataMhs.length; i++) {
          if (dataMhs[i][0] == noCoin) {
            namaMhs = dataMhs[i][1];
            break;
          }
        }
      }
      
      if (!namaMhs) {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", 
          "message": "Koin tidak valid. Data mahasiswa tidak ditemukan."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      sheetPeminjaman.appendRow([timestamp, noCoin, namaMhs, kodeAlat, status, ""]);
      
      return ContentService.createTextOutput(JSON.stringify({
        "status": "success", 
        "message": "Data peminjaman berhasil ditambahkan"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ==========================================
    // ACTION: KEMBALIKAN ALAT
    // ==========================================
    else if (action === "kembali") {
      var noCoin = e.parameter.no_coin;
      var kodeAlat = e.parameter.kode_alat;
      var timestampKembali = e.parameter.timestamp_kembali;
      
      var data = sheetPeminjaman.getDataRange().getValues();
      var updated = false;
      
      if (data.length > 1) {
        var sheetHeaders = data[0];
        
        var colCoin = sheetHeaders.indexOf("No Coin") !== -1 ? sheetHeaders.indexOf("No Coin") : 1;
        var colKode = sheetHeaders.indexOf("Kode Alat") !== -1 ? sheetHeaders.indexOf("Kode Alat") : 3;
        var colStatus = sheetHeaders.indexOf("Status") !== -1 ? sheetHeaders.indexOf("Status") : 4;
        var colTimeKembali = sheetHeaders.indexOf("Timestamp Kembali");
        if (colTimeKembali === -1) colTimeKembali = sheetHeaders.length; 
        
        var targetCoin = String(noCoin).trim();
        var targetKode = String(kodeAlat).trim();

        // Loop dari bawah ke atas
        for (var j = data.length - 1; j > 0; j--) {
          var rowCoin = String(data[j][colCoin]).trim();
          var rowKode = String(data[j][colKode]).trim();
          var rowStatus = String(data[j][colStatus]).trim();
          
          if (rowStatus === "Dipinjam" && rowCoin === targetCoin && rowKode === targetKode) {
            sheetPeminjaman.getRange(j + 1, colStatus + 1).setValue("Kembali");
            sheetPeminjaman.getRange(j + 1, colTimeKembali + 1).setValue(timestampKembali);
            updated = true;
            break;
          }
        }
      }
      
      if (updated) {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "success", 
          "message": "Alat berhasil dikembalikan"
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", 
          "message": "Gagal: Alat tersebut tidak sedang Anda pinjam."
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // ==========================================
    // ACTION: AMBIL DATA (DEFAULT)
    // ==========================================
    else {
      var dataPeminjaman = sheetPeminjaman.getDataRange().getValues();
      var activeLoans = [];
      
      if (dataPeminjaman.length > 1) {
        var sheetHeaders = dataPeminjaman[0];
        
        var colTime = sheetHeaders.indexOf("Timestamp Pinjam");
        if (colTime === -1) colTime = sheetHeaders.indexOf("Timestamp");
        if (colTime === -1) colTime = 0;
        
        var colCoin = sheetHeaders.indexOf("No Coin") !== -1 ? sheetHeaders.indexOf("No Coin") : 1;
        var colKode = sheetHeaders.indexOf("Kode Alat") !== -1 ? sheetHeaders.indexOf("Kode Alat") : 3;
        var colStatus = sheetHeaders.indexOf("Status") !== -1 ? sheetHeaders.indexOf("Status") : 4;
        
        var colNama = sheetHeaders.findIndex(function(h) { return h.toString().toLowerCase().indexOf("nama") !== -1; });
        
        var mhsMap = {};
        if (sheetMahasiswa) {
          var dataMahasiswa = sheetMahasiswa.getDataRange().getValues();
          for (var i = 1; i < dataMahasiswa.length; i++) {
            if (dataMahasiswa[i][0]) mhsMap[dataMahasiswa[i][0]] = dataMahasiswa[i][1];
          }
        }
        
        for (var j = 1; j < dataPeminjaman.length; j++) {
          var status = String(dataPeminjaman[j][colStatus]).trim();
          
          if (status === "Dipinjam") {
            var coin = dataPeminjaman[j][colCoin];
            var kodeAlat = dataPeminjaman[j][colKode];
            var timestampPinjam = dataPeminjaman[j][colTime];
            
            var namaPeminjam = "";
            if (colNama !== -1 && dataPeminjaman[j][colNama]) {
              namaPeminjam = dataPeminjaman[j][colNama];
            } else {
              namaPeminjam = mhsMap[coin] || "Nama tidak diketahui";
            }
            
            activeLoans.push({
              timestamp_pinjam: timestampPinjam,
              no_coin: coin,
              nama: namaPeminjam,
              kode_alat: kodeAlat
            });
          }
        }
      }
      
      activeLoans.reverse();
      
      return ContentService.createTextOutput(JSON.stringify({"status": "success", "data": activeLoans}))
        .setMimeType(ContentService.MimeType.JSON);
    }
      
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Fallback jika ter-panggil via POST (meski kita sekarang pakai GET sepenuhnya)
 */
function doPost(e) {
  return doGet(e);
}
