/**
 * Script ini menangani request POST yang masuk ke Web App (Google Apps Script).
 * Bertindak sebagai REST API untuk menyimpan data ke Google Spreadsheet.
 */
function doPost(e) {
  try {
    // Mem-parsing body request
    // Meskipun Content-Type dari client mungkin 'text/plain' (karena mode no-cors), 
    // isinya adalah JSON string, sehingga kita bisa mem-parsingnya.
    var payload = JSON.parse(e.postData.contents);
    
    var timestamp = payload.timestamp;
    var noCoin = payload.no_coin;
    var kodeAlat = payload.kode_alat;
    var status = "Dipinjam";
    
    // Mendapatkan Spreadsheet aktif (asumsi script ini terikat pada file Spreadsheet yang sama).
    // Jika tidak terikat, gunakan SpreadsheetApp.openById("ID_SPREADSHEET_ANDA");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Membuka sheet "Data_Peminjaman"
    var sheet = ss.getSheetByName("Data_Peminjaman");
    
    // Jika sheet belum ada, kita buat sheet tersebut (sebagai fallback/keamanan)
    if (!sheet) {
      sheet = ss.insertSheet("Data_Peminjaman");
      // Menambahkan header default
      sheet.appendRow(["Timestamp", "No Coin", "Kode Alat", "Status"]);
    }
    
    // Menambahkan baris baru dengan data peminjaman di baris paling bawah
    sheet.appendRow([timestamp, noCoin, kodeAlat, status]);
    
    // Response yang dikembalikan. 
    // (Perlu diingat: pada fetch dengan mode 'no-cors' dari klien, response body ini tidak akan dapat dibaca oleh JS di browser.
    // Namun response ini tetap berguna jika dites dengan tool seperti Postman atau cURL)
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success", 
      "message": "Data peminjaman berhasil ditambahkan"
    })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // Menangani error dan mencatatnya
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Fungsi doGet digunakan jika ada pengguna yang membuka URL Web App lewat browser (opsional).
 */
function doGet(e) {
  return ContentService.createTextOutput("Web App Aktif. Gunakan metode POST untuk mengirim data peminjaman.");
}
