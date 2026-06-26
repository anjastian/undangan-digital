/**
 * SCRIPT GOOGLE APPS SCRIPT UNTUK UNDANGAN PERNIKAHAN DIGITAL
 * 
 * Petunjuk Penggunaan:
 * 1. Buat Google Sheet baru.
 * 2. Buat 3 lembar kerja (Tab) dengan nama: "Guests", "RSVPs", dan "Wishes".
 * 3. Baris pertama (Header) masing-masing sheet:
 *    - Sheet "Guests": Name | Phone | Status | AccessCount | FirstAccessTime | LastAccessTime
 *    - Sheet "RSVPs": Name | Phone | Attendance | GuestsCount | SubmissionTime
 *    - Sheet "Wishes": Name | Wish | Timestamp
 * 4. Buka Ekstensi > Apps Script.
 * 5. Hapus semua kode bawaan dan tempel kode di bawah ini.
 * 6. Klik "Terapkan" / "Deploy" > "Terapkan Baru" / "New Deployment".
 * 7. Pilih tipe: "Aplikasi Web" / "Web App".
 * 8. Jalankan sebagai: "Saya" (akun Google Anda).
 * 9. Siapa yang memiliki akses: "Siapa saja" / "Anyone".
 * 10. Salin URL Aplikasi Web yang diberikan untuk dimasukkan ke halaman Admin undangan Anda.
 */

// Mengatasi CORS Policy agar web undangan dapat mengirim data ke Google Sheets
function doGet(e) {
  var action = e.parameter.action;
  
  if (action === "getWishes") {
    return handleGetWishes();
  } else if (action === "getDashboardData") {
    return handleGetDashboardData();
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Apps Script aktif!" }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    // Fallback jika dikirim lewat form-urlencoded
    data = e.parameter;
  }
  
  var action = data.action;
  
  if (action === "logAccess") {
    return handleLogAccess(data);
  } else if (action === "submitRSVP") {
    return handleSubmitRSVP(data);
  } else if (action === "submitWish") {
    return handleSubmitWish(data);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Action tidak dikenal!" }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

// 1. TANDA AKSES TAMU (NOTIFIKASI / LOGGING BILA TAMU MEMBUKA WEB)
function handleLogAccess(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Guests");
  if (!sheet) {
    return createErrorResponse("Sheet 'Guests' tidak ditemukan");
  }
  
  var name = data.name || "Tamu Tanpa Nama";
  var phone = data.phone || "-";
  var now = new Date();
  
  var rows = sheet.getDataRange().getValues();
  var found = false;
  var rowIndex = -1;
  
  // Mencari apakah tamu sudah ada di daftar
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] == name && (phone === "-" || rows[i][1] == phone)) {
      found = true;
      rowIndex = i + 1; // Baris di sheet 1-indexed
      break;
    }
  }
  
  if (found) {
    // Update data tamu lama
    var accessCount = parseInt(rows[rowIndex - 1][3] || 0) + 1;
    sheet.getRange(rowIndex, 3).setValue("Accessed"); // Status
    sheet.getRange(rowIndex, 4).setValue(accessCount); // AccessCount
    sheet.getRange(rowIndex, 6).setValue(now); // LastAccessTime
  } else {
    // Tambah tamu baru jika belum terdaftar (misal link disebar manual tanpa admin)
    sheet.appendRow([name, phone, "Accessed", 1, now, now]);
  }
  
  return createSuccessResponse("Akses berhasil dicatat");
}

// 2. MENYIMPAN DATA RSVP
function handleSubmitRSVP(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RSVPs");
  if (!sheet) {
    return createErrorResponse("Sheet 'RSVPs' tidak ditemukan");
  }
  
  var name = data.name || "Tanpa Nama";
  var phone = data.phone || "-";
  var attendance = data.attendance || "Tidak Hadir";
  var guestsCount = data.guestsCount || 0;
  var now = new Date();
  
  // Masukkan data RSVP ke sheet
  sheet.appendRow([name, phone, attendance, guestsCount, now]);
  
  // Update status di daftar tamu jika terdaftar
  var guestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Guests");
  if (guestSheet) {
    var rows = guestSheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] == name) {
        guestSheet.getRange(i + 1, 3).setValue("RSVP: " + attendance);
        break;
      }
    }
  }
  
  return createSuccessResponse("RSVP berhasil disimpan");
}

// 3. MENYIMPAN UCAPAN & DOA (WISHES)
function handleSubmitWish(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Wishes");
  if (!sheet) {
    return createErrorResponse("Sheet 'Wishes' tidak ditemukan");
  }
  
  var name = data.name || "Anonim";
  var wish = data.wish || "";
  var now = new Date();
  
  if (wish.trim() === "") {
    return createErrorResponse("Ucapan tidak boleh kosong");
  }
  
  sheet.appendRow([name, wish, now]);
  return createSuccessResponse("Ucapan berhasil disimpan");
}

// 4. MENGAMBIL DAFTAR UCAPAN UNTUK DITAMPILKAN DI WEB
function handleGetWishes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Wishes");
  var wishes = [];
  
  if (sheet) {
    var rows = sheet.getDataRange().getValues();
    // Looping dari bawah ke atas agar ucapan terbaru muncul paling dulu
    for (var i = rows.length - 1; i >= 1; i--) {
      wishes.push({
        name: rows[i][0],
        wish: rows[i][1],
        timestamp: rows[i][2]
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: wishes }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

// 5. MENGAMBIL DATA MONITORING UNTUK DASHBOARD ADMIN
function handleGetDashboardData() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var guestSheet = spreadSheet.getSheetByName("Guests");
  var rsvpSheet = spreadSheet.getSheetByName("RSVPs");
  var wishSheet = spreadSheet.getSheetByName("Wishes");
  
  var guests = [];
  var rsvps = [];
  var wishes = [];
  
  if (guestSheet) {
    var rows = guestSheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      guests.push({
        name: rows[i][0],
        phone: rows[i][1],
        status: rows[i][2],
        accessCount: rows[i][3],
        firstAccess: rows[i][4],
        lastAccess: rows[i][5]
      });
    }
  }
  
  if (rsvpSheet) {
    var rows = rsvpSheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      rsvps.push({
        name: rows[i][0],
        phone: rows[i][1],
        attendance: rows[i][2],
        guestsCount: rows[i][3],
        timestamp: rows[i][4]
      });
    }
  }
  
  if (wishSheet) {
    var rows = wishSheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      wishes.push({
        name: rows[i][0],
        wish: rows[i][1],
        timestamp: rows[i][2]
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: {
      guests: guests,
      rsvps: rsvps,
      wishes: wishes
    }
  }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

// Fungsi bantu untuk response
function createSuccessResponse(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: msg }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

function createErrorResponse(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: msg }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}
