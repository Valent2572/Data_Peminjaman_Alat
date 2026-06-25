/**
 * Sistem terpusat menggunakan doGet agar tidak ada kendala CORS dari browser.
 */
function doGet(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPeminjaman = ss.getSheetByName("Data_Peminjaman");
    var sheetMahasiswa = ss.getSheetByName("Data_Mahasiswa");
    var sheetAlat = ss.getSheetByName("Data_Alat");
    
    if (!sheetPeminjaman) {
      sheetPeminjaman = ss.insertSheet("Data_Peminjaman");
      sheetPeminjaman.appendRow(["Timestamp Pinjam", "No Coin", "Nama Lengkap Peminjam", "Kode Alat", "Nama Detil Alat", "Status", "Timestamp Kembali"]);
    }
    
    // Cache Alat
    var alatMap = {};
    if (sheetAlat) {
      var dataAlat = sheetAlat.getDataRange().getValues();
      if (dataAlat.length > 0) {
        var alatHeaders = dataAlat[0];
        var colKodeAlat = -1, colNamaAlat = -1;
        for (var h = 0; h < alatHeaders.length; h++) {
          var text = String(alatHeaders[h]).toUpperCase();
          if (text.indexOf("KODE") !== -1) colKodeAlat = h;
          if (text.indexOf("NAMA") !== -1 && text.indexOf("DETIL") !== -1) colNamaAlat = h;
        }
        if (colKodeAlat === -1) colKodeAlat = alatHeaders.length >= 6 ? 5 : 0; 
        if (colNamaAlat === -1) colNamaAlat = colKodeAlat + 1;
        
        for (var i = 1; i < dataAlat.length; i++) {
          if (dataAlat[i][colKodeAlat]) {
            alatMap[String(dataAlat[i][colKodeAlat]).trim()] = String(dataAlat[i][colNamaAlat] || "").trim();
          }
        }
      }
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
      
      // Validasi Mahasiswa
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
      
      // Validasi Alat
      var kodeAlatStr = String(kodeAlat).trim();
      var namaDetilAlat = alatMap[kodeAlatStr];
      if (!namaDetilAlat) {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", 
          "message": "Kode alat tidak dikenali di sheet Data_Alat."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Kita pakai logika cerdas untuk append row agar menyesuaikan urutan kolom saat ini
      var headP = sheetPeminjaman.getDataRange().getValues()[0];
      var newRow = new Array(headP.length);
      for(var k=0; k<newRow.length; k++) newRow[k] = ""; // Inisialisasi kosong
      
      var idxTime = headP.indexOf("Timestamp Pinjam"); if(idxTime===-1) idxTime=0;
      var idxCoin = headP.indexOf("No Coin"); if(idxCoin===-1) idxCoin=1;
      var idxNamaM = headP.findIndex(function(h) { return h.toString().toLowerCase().indexOf("nama") !== -1 && h.toString().toLowerCase().indexOf("alat") === -1; }); if(idxNamaM===-1) idxNamaM=2;
      var idxKode = headP.indexOf("Kode Alat"); if(idxKode===-1) idxKode=3;
      var idxNamaA = headP.indexOf("Nama Detil Alat"); 
      var idxStatus = headP.indexOf("Status"); if(idxStatus===-1) idxStatus=headP.length; // Taruh di belakang kalau ga ada
      
      newRow[idxTime] = timestamp;
      newRow[idxCoin] = noCoin;
      newRow[idxNamaM] = namaMhs;
      newRow[idxKode] = kodeAlat;
      if (idxNamaA !== -1) {
        newRow[idxNamaA] = namaDetilAlat;
      } else {
        // Kalau belum ada kolom Nama Detil Alat, taruh setelah Kode Alat saja (resiko menggeser, maka disarankan tambah kolom)
        newRow.push(namaDetilAlat);
      }
      newRow[idxStatus] = status;
      
      sheetPeminjaman.appendRow(newRow);
      
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
        var colNamaAlatSheet = sheetHeaders.indexOf("Nama Detil Alat");
        var colStatus = sheetHeaders.indexOf("Status");
        
        var colNama = sheetHeaders.findIndex(function(h) { return h.toString().toLowerCase().indexOf("nama") !== -1 && h.toString().toLowerCase().indexOf("alat") === -1; });
        
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
            
            var nDetilAlat = (colNamaAlatSheet !== -1) ? dataPeminjaman[j][colNamaAlatSheet] : alatMap[kodeAlat];
            
            activeLoans.push({
              timestamp_pinjam: timestampPinjam,
              no_coin: coin,
              nama: namaPeminjam,
              kode_alat: kodeAlat,
              nama_alat: nDetilAlat || "Alat tidak terdaftar"
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

function doPost(e) {
  return doGet(e);
}
