/**
 * Script ini menangani request POST yang masuk ke Web App.
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || "pinjam"; 
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPeminjaman = ss.getSheetByName("Data_Peminjaman");
    var sheetMahasiswa = ss.getSheetByName("Data_Mahasiswa");
    
    // Pastikan sheet ada. Jika belum ada, buat baru dengan header lengkap
    if (!sheetPeminjaman) {
      sheetPeminjaman = ss.insertSheet("Data_Peminjaman");
      sheetPeminjaman.appendRow(["Timestamp Pinjam", "No Coin", "Nama Peminjam", "Kode Alat", "Status", "Timestamp Kembali"]);
    }
    
    if (action === "pinjam") {
      var timestamp = payload.timestamp;
      var noCoin = payload.no_coin;
      var kodeAlat = payload.kode_alat;
      var status = "Dipinjam";
      
      // Cari nama peminjam dari Data_Mahasiswa berdasarkan No Coin
      var namaMhs = "Nama tidak ditemukan";
      if (sheetMahasiswa) {
        var dataMhs = sheetMahasiswa.getDataRange().getValues();
        for (var i = 1; i < dataMhs.length; i++) {
          if (dataMhs[i][0] == noCoin) {
            namaMhs = dataMhs[i][1];
            break;
          }
        }
      }
      
      // Menambahkan data. Urutan kolom: 
      // A: Timestamp Pinjam, B: No Coin, C: Nama Peminjam, D: Kode Alat, E: Status, F: Timestamp Kembali
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
      
      // Loop dari bawah ke atas (karena data terbaru ada di bawah)
      for (var j = data.length - 1; j > 0; j--) {
        // Kita gunakan asumsi urutan kolom:
        // 0: Timestamp Pinjam, 1: No Coin, 2: Nama Peminjam, 3: Kode Alat, 4: Status, 5: Timestamp Kembali
        var rowTime = data[j][0];
        var rowCoin = data[j][1];
        var rowKode = data[j][3];
        var rowStatus = data[j][4];
        
        // Pengecekan ekstra untuk sheet lama yang kolomnya masih (Timestamp, No Coin, Kode Alat, Status)
        // Kita cek panjang baris
        if (data[0].indexOf("Nama Peminjam") === -1) {
          // Format lama: 0: Timestamp, 1: No Coin, 2: Kode Alat, 3: Status
          rowTime = data[j][0];
          rowCoin = data[j][1];
          rowKode = data[j][2];
          rowStatus = data[j][3];
          
          if (rowStatus === "Dipinjam" && rowCoin == noCoin && rowKode == kodeAlat && rowTime == timestampPinjam) {
            sheetPeminjaman.getRange(j + 1, 4).setValue("Kembali");
            // Set Timestamp kembali di kolom baru (misal F)
            sheetPeminjaman.getRange(j + 1, 6).setValue(timestampKembali);
            updated = true;
            break;
          }
        } else {
          // Format baru
          if (rowStatus === "Dipinjam" && rowCoin == noCoin && rowKode == kodeAlat && rowTime == timestampPinjam) {
            // Ubah Status di kolom E (indeks 5) menjadi Kembali
            sheetPeminjaman.getRange(j + 1, 5).setValue("Kembali");
            // Set Timestamp Kembali di kolom F (indeks 6)
            sheetPeminjaman.getRange(j + 1, 6).setValue(timestampKembali);
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
    
    if (!sheetPeminjaman) {
      return ContentService.createTextOutput(JSON.stringify({"status": "success", "data": []}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var dataPeminjaman = sheetPeminjaman.getDataRange().getValues();
    var activeLoans = [];
    
    if (dataPeminjaman.length > 1) {
      var isFormatLama = dataPeminjaman[0].indexOf("Nama Peminjam") === -1;
      
      // Ambil sheet Mahasiswa hanya untuk format lama yang nggak ada kolom Nama Peminjam
      var mhsMap = {};
      if (isFormatLama) {
        var sheetMahasiswa = ss.getSheetByName("Data_Mahasiswa");
        if (sheetMahasiswa) {
          var dataMahasiswa = sheetMahasiswa.getDataRange().getValues();
          for (var i = 1; i < dataMahasiswa.length; i++) {
            if (dataMahasiswa[i][0]) mhsMap[dataMahasiswa[i][0]] = dataMahasiswa[i][1];
          }
        }
      }
      
      for (var j = 1; j < dataPeminjaman.length; j++) {
        var status = isFormatLama ? dataPeminjaman[j][3] : dataPeminjaman[j][4];
        
        if (status === "Dipinjam") {
          var coin = dataPeminjaman[j][1];
          var kodeAlat = isFormatLama ? dataPeminjaman[j][2] : dataPeminjaman[j][3];
          var namaPeminjam = isFormatLama ? (mhsMap[coin] || "Nama tidak ditemukan") : dataPeminjaman[j][2];
          var timestampPinjam = dataPeminjaman[j][0];
          
          activeLoans.push({
            timestamp_pinjam: timestampPinjam,
            no_coin: coin,
            nama: namaPeminjam,
            kode_alat: kodeAlat
          });
        }
      }
    }
    
    // Urutkan data terbaru di atas
    activeLoans.reverse();
    
    return ContentService.createTextOutput(JSON.stringify({"status": "success", "data": activeLoans}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
