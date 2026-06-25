/**
 * Script ini menangani request POST yang masuk ke Web App.
 */
function doPost(e) {
  // CORS Headers Helper
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || "pinjam"; 
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPeminjaman = ss.getSheetByName("Data_Peminjaman");
    var sheetMahasiswa = ss.getSheetByName("Data_Mahasiswa");
    
    if (!sheetPeminjaman) {
      sheetPeminjaman = ss.insertSheet("Data_Peminjaman");
      sheetPeminjaman.appendRow(["Timestamp Pinjam", "No Coin", "Nama Peminjam", "Kode Alat", "Status", "Timestamp Kembali"]);
    }
    
    if (action === "pinjam") {
      var timestamp = payload.timestamp;
      var noCoin = payload.no_coin;
      var kodeAlat = payload.kode_alat;
      var status = "Dipinjam";
      
      // Validasi Koin: Cari nama peminjam dari Data_Mahasiswa
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
      
      // Jika koin tidak ditemukan, tolak peminjaman
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
      
    } else if (action === "kembali") {
      var timestampPinjam = payload.timestamp_pinjam; 
      var noCoin = payload.no_coin;
      var kodeAlat = payload.kode_alat;
      var timestampKembali = payload.timestamp_kembali;
      
      var data = sheetPeminjaman.getDataRange().getValues();
      var updated = false;
      
      if (data.length > 1) {
        var sheetHeaders = data[0];
        
        var colTime = sheetHeaders.indexOf("Timestamp Pinjam");
        if (colTime === -1) colTime = sheetHeaders.indexOf("Timestamp");
        if (colTime === -1) colTime = 0;
        
        var colCoin = sheetHeaders.indexOf("No Coin") !== -1 ? sheetHeaders.indexOf("No Coin") : 1;
        var colKode = sheetHeaders.indexOf("Kode Alat") !== -1 ? sheetHeaders.indexOf("Kode Alat") : 3;
        var colStatus = sheetHeaders.indexOf("Status") !== -1 ? sheetHeaders.indexOf("Status") : 4;
        var colTimeKembali = sheetHeaders.indexOf("Timestamp Kembali");
        if (colTimeKembali === -1) colTimeKembali = sheetHeaders.length; 
        
        for (var j = data.length - 1; j > 0; j--) {
          var rowTime = data[j][colTime];
          var rowCoin = data[j][colCoin];
          var rowKode = data[j][colKode];
          var rowStatus = data[j][colStatus];
          
          if (rowStatus === "Dipinjam" && rowCoin == noCoin && rowKode == kodeAlat && rowTime == timestampPinjam) {
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
          "message": "Data peminjaman tidak ditemukan"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Menangani request GET (mengembalikan data JSON ke modal).
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
        var status = dataPeminjaman[j][colStatus];
        
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
      
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
