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
    var sheetAlat = ss.getSheetByName("Data_Alat");
    
    if (!sheetPeminjaman) {
      sheetPeminjaman = ss.insertSheet("Data_Peminjaman");
      sheetPeminjaman.appendRow(["Timestamp Pinjam", "No Koin", "Nama Lengkap Peminjam", "Kode Alat", "Nama Detil Alat", "Status", "Timestamp Kembali"]);
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
            alatMap[String(dataAlat[i][colKodeAlat]).trim().toUpperCase()] = String(dataAlat[i][colNamaAlat] || "").trim();
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
      var noCoin = String(e.parameter.no_coin).trim();
      var kodeAlat = String(e.parameter.kode_alat).trim();
      var status = "Dipinjam";
      
      // Validasi Mahasiswa dari 4 Sheet (Tk.1 sampai Tk.4)
      var sheetNames = ["Data_Mahasiswa Tk.1", "Data_Mahasiswa Tk.2", "Data_Mahasiswa Tk.3", "Data_Mahasiswa Tk.4"];
      var namaMhs = null;
      var tingkatMhs = "";
      
      for (var s = 0; s < sheetNames.length; s++) {
        var sheetMhs = ss.getSheetByName(sheetNames[s]);
        if (sheetMhs) {
          var dataMhs = sheetMhs.getDataRange().getValues();
          for (var i = 1; i < dataMhs.length; i++) {
            if (String(dataMhs[i][0]).trim() === noCoin) {
              namaMhs = String(dataMhs[i][1]).trim(); // Ambil nama
              tingkatMhs = "Tk." + (s + 1); // Penanda tingkat
              break;
            }
          }
        }
        if (namaMhs) break; // Berhenti mencari jika sudah ketemu
      }
      
      if (!namaMhs) {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", 
          "message": "Koin tidak valid. Data mahasiswa tidak ditemukan di tingkat manapun."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Gabungkan nama dengan tingkatnya (Contoh: "Adelia Putri Rahmadhani (Tk.1)")
      var namaLengkapDenganTingkat = namaMhs + " (" + tingkatMhs + ")";
      
      // Validasi Alat
      var namaDetilAlat = alatMap[kodeAlat.toUpperCase()];
      if (!namaDetilAlat) {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", 
          "message": "Kode alat tidak dikenali di sheet Data_Alat."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      var headP = sheetPeminjaman.getDataRange().getValues()[0];
      var newRow = new Array(headP.length);
      for(var k=0; k<newRow.length; k++) newRow[k] = ""; 
      
      var idxTime = headP.indexOf("Timestamp Pinjam"); if(idxTime===-1) idxTime=0;
      var idxCoin = headP.findIndex(function(h) { return String(h).toUpperCase().indexOf("KOIN") !== -1 || String(h).toUpperCase().indexOf("COIN") !== -1; }); if(idxCoin===-1) idxCoin=1;
      var idxNamaM = headP.findIndex(function(h) { return String(h).toUpperCase().indexOf("NAMA") !== -1 && String(h).toUpperCase().indexOf("ALAT") === -1; }); if(idxNamaM===-1) idxNamaM=2;
      var idxKode = headP.findIndex(function(h) { return String(h).toUpperCase().indexOf("KODE ALAT") !== -1; }); if(idxKode===-1) idxKode=3;
      var idxNamaA = headP.findIndex(function(h) { return String(h).toUpperCase().indexOf("DETIL ALAT") !== -1 || String(h).toUpperCase().indexOf("NAMA ALAT") !== -1; }); 
      var idxStatus = headP.findIndex(function(h) { return String(h).toUpperCase() === "STATUS"; }); if(idxStatus===-1) idxStatus=headP.length; 
      
      newRow[idxTime] = timestamp;
      newRow[idxCoin] = noCoin;
      newRow[idxNamaM] = namaLengkapDenganTingkat;
      newRow[idxKode] = kodeAlat;
      if (idxNamaA !== -1) {
        newRow[idxNamaA] = namaDetilAlat;
      } else {
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
      var noCoin = String(e.parameter.no_coin).trim();
      var kodeAlat = String(e.parameter.kode_alat).trim();
      var timestampKembali = e.parameter.timestamp_kembali;
      
      var data = sheetPeminjaman.getDataRange().getValues();
      var updated = false;
      
      if (data.length > 1) {
        var sheetHeaders = data[0];
        
        var colCoin = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KOIN") !== -1 || String(h).toUpperCase().indexOf("COIN") !== -1; }); if(colCoin===-1) colCoin=1;
        var colKode = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KODE ALAT") !== -1; }); if(colKode===-1) colKode=3;
        var colStatus = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase() === "STATUS"; }); if(colStatus===-1) colStatus=4;
        var colTimeKembali = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KEMBALI") !== -1; });
        if (colTimeKembali === -1) colTimeKembali = sheetHeaders.length; 
        
        var targetCoin = noCoin;
        var targetKode = kodeAlat.toUpperCase();

        for (var j = data.length - 1; j > 0; j--) {
          var rowCoin = String(data[j][colCoin]).trim();
          var rowKode = String(data[j][colKode]).trim().toUpperCase();
          var rowStatus = String(data[j][colStatus]).trim().toUpperCase();
          
          if (rowStatus === "DIPINJAM" && rowCoin === targetCoin && rowKode === targetKode) {
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
        
        var colTime = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("TIMESTAMP PINJAM") !== -1; }); if(colTime===-1) colTime=0;
        var colCoin = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KOIN") !== -1 || String(h).toUpperCase().indexOf("COIN") !== -1; }); if(colCoin===-1) colCoin=1;
        var colKode = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KODE ALAT") !== -1; }); if(colKode===-1) colKode=3;
        var colNamaAlatSheet = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("DETIL ALAT") !== -1 || String(h).toUpperCase().indexOf("NAMA ALAT") !== -1; });
        var colStatus = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase() === "STATUS"; }); if(colStatus===-1) colStatus=4;
        
        var colNama = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("NAMA") !== -1 && String(h).toUpperCase().indexOf("ALAT") === -1; });
        
        // Build MhsMap dari 4 Sheet untuk fallback
        var mhsMap = {};
        var sheetNames = ["Data_Mahasiswa Tk.1", "Data_Mahasiswa Tk.2", "Data_Mahasiswa Tk.3", "Data_Mahasiswa Tk.4"];
        for (var s = 0; s < sheetNames.length; s++) {
          var sMhs = ss.getSheetByName(sheetNames[s]);
          if (sMhs) {
            var dMhs = sMhs.getDataRange().getValues();
            for (var i = 1; i < dMhs.length; i++) {
              if (dMhs[i][0]) mhsMap[String(dMhs[i][0]).trim()] = String(dMhs[i][1]).trim() + " (Tk." + (s + 1) + ")";
            }
          }
        }
        
        for (var j = 1; j < dataPeminjaman.length; j++) {
          var status = String(dataPeminjaman[j][colStatus]).trim().toUpperCase();
          
          if (status === "DIPINJAM") {
            var coin = String(dataPeminjaman[j][colCoin]).trim();
            var kodeAlat = String(dataPeminjaman[j][colKode]).trim();
            var timestampPinjam = dataPeminjaman[j][colTime];
            
            var namaPeminjam = "";
            if (colNama !== -1 && dataPeminjaman[j][colNama]) {
              namaPeminjam = dataPeminjaman[j][colNama];
            } else {
              namaPeminjam = mhsMap[coin] || "Nama tidak diketahui";
            }
            
            var nDetilAlat = (colNamaAlatSheet !== -1 && dataPeminjaman[j][colNamaAlatSheet]) ? dataPeminjaman[j][colNamaAlatSheet] : alatMap[kodeAlat.toUpperCase()];
            
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
