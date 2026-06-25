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
      sheetPeminjaman.appendRow(["Timestamp Pinjam", "No Koin", "Nama Lengkap Peminjam", "NIM", "Kode Alat", "Nama Detil Alat", "Status", "Timestamp Kembali"]);
    }
    
    function sortPeminjamanData(sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) return; 
      
      var sHeaders = data[0];
      var rows = data.slice(1);
      
      var colStatus = sHeaders.findIndex(function(h) { return String(h).toUpperCase() === "STATUS"; });
      var colTimeKembali = sHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KEMBALI") !== -1; });
      var colTimePinjam = sHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("PINJAM") !== -1; });
      
      if (colStatus === -1) colStatus = sHeaders.length - 2; 
      
      function parseTs(ts) {
        if (!ts) return 0;
        var parts = String(ts).split(" - ");
        if (parts.length !== 2) return 0;
        var timeP = parts[0].split(":");
        var dateP = parts[1].split("-");
        if (dateP.length !== 3 || timeP.length !== 3) return 0;
        return new Date(dateP[2], dateP[1]-1, dateP[0], timeP[0], timeP[1], timeP[2]).getTime();
      }
      
      rows.sort(function(a, b) {
        var statusA = String(a[colStatus] || "").trim().toUpperCase();
        var statusB = String(b[colStatus] || "").trim().toUpperCase();
        
        var weightA = (statusA === "DIPINJAM") ? 0 : 1;
        var weightB = (statusB === "DIPINJAM") ? 0 : 1;
        
        if (weightA !== weightB) {
          return weightA - weightB; 
        }
        
        if (weightA === 1) { 
          var timeA = parseTs(colTimeKembali !== -1 ? a[colTimeKembali] : "");
          var timeB = parseTs(colTimeKembali !== -1 ? b[colTimeKembali] : "");
          return timeB - timeA; 
        }
        
        var pTimeA = parseTs(colTimePinjam !== -1 ? a[colTimePinjam] : "");
        var pTimeB = parseTs(colTimePinjam !== -1 ? b[colTimePinjam] : "");
        return pTimeB - pTimeA; 
      });
      
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    
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
      var noCoin = String(e.parameter.no_coin || "").trim();
      var kodeAlat = String(e.parameter.kode_alat).trim();
      var tipePeminjam = String(e.parameter.tipe_peminjam || "mahasiswa").toLowerCase();
      var namaInstruktur = String(e.parameter.nama_instruktur || "").trim();
      var status = "Dipinjam";
      
      var namaLengkapDenganTingkat = "";
      var nimMhs = "-";
      
      if (tipePeminjam === "instruktur") {
        if (!namaInstruktur) {
          return ContentService.createTextOutput(JSON.stringify({
            "status": "error", 
            "message": "Nama Instruktur tidak boleh kosong."
          })).setMimeType(ContentService.MimeType.JSON);
        }
        namaLengkapDenganTingkat = namaInstruktur + " (Instruktur)";
        noCoin = "-";
      } else {
        // OPTIMISASI: Baca sheet secara terarah berdasarkan simbol jika memungkinkan
        var targetSheets = [];
        if (noCoin.indexOf("δ") !== -1) targetSheets.push("Data_Mahasiswa Tk.1");
        else if (noCoin.indexOf("γ") !== -1) targetSheets.push("Data_Mahasiswa Tk.2");
        else if (noCoin.indexOf("β") !== -1) targetSheets.push("Data_Mahasiswa Tk.3");
        else if (noCoin.indexOf("α") !== -1) targetSheets.push("Data_Mahasiswa Tk.4");
        else {
          // Jika tidak dikenali, baca semua sebagai cadangan
          targetSheets = ["Data_Mahasiswa Tk.1", "Data_Mahasiswa Tk.2", "Data_Mahasiswa Tk.3", "Data_Mahasiswa Tk.4"];
        }

        var namaMhs = null;
        var tingkatMhs = "";
        
        for (var s = 0; s < targetSheets.length; s++) {
          var sheetMhs = ss.getSheetByName(targetSheets[s]);
          if (sheetMhs) {
            var dataMhs = sheetMhs.getDataRange().getValues();
            for (var i = 1; i < dataMhs.length; i++) {
              if (String(dataMhs[i][0]).trim() === noCoin) {
                namaMhs = String(dataMhs[i][1]).trim();
                nimMhs = String(dataMhs[i][2] || "-").trim(); // NIM di kolom C
                
                // Set tingkat sesuai nama sheet
                if (targetSheets[s].indexOf("Tk.1") !== -1) tingkatMhs = "Tk.1";
                else if (targetSheets[s].indexOf("Tk.2") !== -1) tingkatMhs = "Tk.2";
                else if (targetSheets[s].indexOf("Tk.3") !== -1) tingkatMhs = "Tk.3";
                else if (targetSheets[s].indexOf("Tk.4") !== -1) tingkatMhs = "Tk.4";
                break;
              }
            }
          }
          if (namaMhs) break;
        }
        
        if (!namaMhs) {
          return ContentService.createTextOutput(JSON.stringify({
            "status": "error", 
            "message": "Koin tidak valid. Data mahasiswa tidak ditemukan."
          })).setMimeType(ContentService.MimeType.JSON);
        }
        namaLengkapDenganTingkat = namaMhs + " (" + tingkatMhs + ")";
      }
      
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
      var idxNamaM = headP.findIndex(function(h) { return String(h).toUpperCase().indexOf("NAMA") !== -1 && String(h).toUpperCase().indexOf("ALAT") === -1 && String(h).toUpperCase().indexOf("DETIL") === -1; }); if(idxNamaM===-1) idxNamaM=2;
      var idxNim = headP.findIndex(function(h) { return String(h).toUpperCase() === "NIM"; });
      var idxKode = headP.findIndex(function(h) { return String(h).toUpperCase().indexOf("KODE ALAT") !== -1; }); if(idxKode===-1) idxKode=3;
      var idxNamaA = headP.findIndex(function(h) { return String(h).toUpperCase().indexOf("DETIL ALAT") !== -1 || String(h).toUpperCase().indexOf("NAMA ALAT") !== -1; }); 
      var idxStatus = headP.findIndex(function(h) { return String(h).toUpperCase() === "STATUS"; }); if(idxStatus===-1) idxStatus=headP.length; 
      
      newRow[idxTime] = timestamp;
      newRow[idxCoin] = noCoin;
      newRow[idxNamaM] = namaLengkapDenganTingkat;
      if (idxNim !== -1) newRow[idxNim] = nimMhs; 
      
      if (idxKode !== -1) {
        newRow[idxKode] = kodeAlat;
      }
      
      if (idxNamaA !== -1) {
        newRow[idxNamaA] = (idxKode === -1) ? (namaDetilAlat + " (" + kodeAlat + ")") : namaDetilAlat;
      } else {
        if (idxNim === -1) newRow.push(nimMhs);
        newRow.push((idxKode === -1) ? (namaDetilAlat + " (" + kodeAlat + ")") : namaDetilAlat);
      }
      
      newRow[idxStatus] = status;
      
      sheetPeminjaman.appendRow(newRow);
      
      sortPeminjamanData(sheetPeminjaman);
      
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
        var colKode = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KODE ALAT") !== -1; });
        var colNamaA = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("DETIL ALAT") !== -1 || String(h).toUpperCase().indexOf("NAMA ALAT") !== -1; });
        var colStatus = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase() === "STATUS"; }); if(colStatus===-1) colStatus=4;
        var colTimeKembali = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KEMBALI") !== -1; });
        if (colTimeKembali === -1) colTimeKembali = sheetHeaders.length; 
        
        var targetCoin = noCoin;
        var targetKode = kodeAlat.toUpperCase();

        for (var j = data.length - 1; j > 0; j--) {
          var rowCoin = String(data[j][colCoin]).trim();
          var rowStatus = String(data[j][colStatus]).trim().toUpperCase();
          
          var matchKode = false;
          if (colKode !== -1) {
            matchKode = (String(data[j][colKode]).trim().toUpperCase() === targetKode);
          } else if (colNamaA !== -1) {
            matchKode = (String(data[j][colNamaA]).toUpperCase().indexOf("(" + targetKode + ")") !== -1);
          }
          
          if (rowStatus === "DIPINJAM" && rowCoin === targetCoin && matchKode) {
            sheetPeminjaman.getRange(j + 1, colStatus + 1).setValue("Kembali");
            sheetPeminjaman.getRange(j + 1, colTimeKembali + 1).setValue(timestampKembali);
            updated = true;
            break;
          }
        }
      }
      
      if (updated) {
        sortPeminjamanData(sheetPeminjaman);
        
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
        var colKode = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("KODE ALAT") !== -1; });
        var colNamaAlatSheet = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("DETIL ALAT") !== -1 || String(h).toUpperCase().indexOf("NAMA ALAT") !== -1; });
        var colStatus = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase() === "STATUS"; }); if(colStatus===-1) colStatus=4;
        var colNama = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase().indexOf("NAMA") !== -1 && String(h).toUpperCase().indexOf("ALAT") === -1 && String(h).toUpperCase().indexOf("DETIL") === -1; });
        var colNim = sheetHeaders.findIndex(function(h) { return String(h).toUpperCase() === "NIM"; });
        
        // OPTIMISASI BESAR: 
        // Dihapus blok kode mhsMap yang membaca 4 sheet Mahasiswa setiap kali memuat web.
        // Nama peminjam sudah direkam di sheet Peminjaman, jadi tinggal dibaca saja.
        
        for (var j = 1; j < dataPeminjaman.length; j++) {
          var status = String(dataPeminjaman[j][colStatus]).trim().toUpperCase();
          
          if (status === "DIPINJAM") {
            var coin = String(dataPeminjaman[j][colCoin]).trim();
            var timestampPinjam = dataPeminjaman[j][colTime];
            
            var kodeAlat = (colKode !== -1) ? String(dataPeminjaman[j][colKode]).trim() : "";
            
            var namaPeminjam = "Nama tidak diketahui";
            if (colNama !== -1 && dataPeminjaman[j][colNama]) {
              namaPeminjam = dataPeminjaman[j][colNama]; // Langsung ambil dari sheet peminjaman!
            }
            
            var nimPeminjam = "-";
            if (colNim !== -1 && dataPeminjaman[j][colNim]) {
              nimPeminjam = dataPeminjaman[j][colNim];
            }
            
            var nDetilAlat = (colNamaAlatSheet !== -1 && dataPeminjaman[j][colNamaAlatSheet]) ? dataPeminjaman[j][colNamaAlatSheet] : alatMap[kodeAlat.toUpperCase()];
            
            if (colKode === -1 && nDetilAlat) {
              var match = String(nDetilAlat).match(/\((.*?)\)$/);
              if (match) {
                kodeAlat = match[1];
                nDetilAlat = String(nDetilAlat).replace(/\s*\(.*?\)$/, ''); 
              }
            }
            
            activeLoans.push({
              timestamp_pinjam: timestampPinjam,
              no_coin: coin,
              nama: namaPeminjam,
              nim: nimPeminjam,
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
