/**
 * GOOGLE APPS SCRIPT BACKEND FOR DOKUMEN SISWA
 * Spreadsheet ID: 1ViRbaN_tNOPD4M8kNbwdxKEXYzhD4DVLn5MWf10KVsI
 * Sheet Tab Name: Data
 * 
 * Kolom yang digunakan:
 * - Kolom A (1): NIS (Pencarian)
 * - Kolom B (2): Identitas / Nama Siswa
 * - Kolom C (3): Link Dokumen PDF
 * - Kolom D (4): Status Persetujuan ("SETUJU CETAK" / "PERLU EDIT KEMBALI")
 * - Kolom E (5): Alasan Edit Kembali (jika ada)
 * - Kolom F (6): Waktu Konfirmasi (Timestamp)
 */

const SPREADSHEET_ID = '1ViRbaN_tNOPD4M8kNbwdxKEXYzhD4DVLn5MWf10KVsI';
const SHEET_NAME = 'Data';

/**
 * FUNGSI OTOMATIS MEMBUAT & MEMFORMAT DATABASE GOOGLE SHEET
 * Jalankan fungsi ini 1x dari Apps Script Editor untuk menyiapkan header & data contoh.
 */
function setupDatabase() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error('Spreadsheet tidak ditemukan. Pastikan SPREADSHEET_ID valid.');
  }

  // Cari atau buat tab 'Data'
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // 1. Buat Header Kolom (Baris 1)
  var headers = [
    ['NIS', 'Identitas / Nama Siswa', 'Link Dokumen PDF', 'Status Persetujuan', 'Alasan Edit Kembali', 'Waktu Konfirmasi']
  ];
  
  var headerRange = sheet.getRange(1, 1, 1, 6);
  headerRange.setValues(headers);

  // Styling Visual Header
  headerRange.setBackground('#0284c7'); // Biru Brand
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // 2. Isi Data Contoh / Dummy (Jika baris masih kosong)
  if (sheet.getLastRow() <= 1) {
    var sampleData = [
      ['1001', 'Ahmad Rizky Pratama', 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf', '', '', ''],
      ['1002', 'Siti Nurhaliza', 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf', '', '', ''],
      ['1003', 'Budi Santoso', 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf', '', '', '']
    ];
    sheet.getRange(2, 1, sampleData.length, 6).setValues(sampleData);
  }

  // 3. Buat Validasi Data (Dropdown) pada Kolom D (Status Persetujuan)
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['SETUJU CETAK', 'PERLU EDIT KEMBALI'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(rule);

  // 4. Atur Lebar Kolom agar Rapih
  sheet.setColumnWidth(1, 110); // NIS
  sheet.setColumnWidth(2, 220); // Identitas / Nama
  sheet.setColumnWidth(3, 350); // Link PDF
  sheet.setColumnWidth(4, 180); // Status
  sheet.setColumnWidth(5, 280); // Alasan
  sheet.setColumnWidth(6, 180); // Waktu Konfirmasi

  // Bekukan Baris Header (Freeze Top Row)
  sheet.setFrozenRows(1);

  Logger.log('Database Google Sheet berhasil disiapkan!');
  return 'Database "Data" berhasil dibuat dan disiapkan!';
}

/**
 * Menerima kiriman data HTTP POST dari aplikasi web frontend
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Tunggu maksimal 10 detik untuk menghindari konflik baris saat pengiriman bersamaan
  lock.tryLock(10000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse('error', 'Tidak ada data yang dikirimkan.');
    }

    var requestData = JSON.parse(e.postData.contents);
    var targetNis = String(requestData.nis || '').trim();
    var newStatus = String(requestData.status || '').trim();
    var newAlasan = String(requestData.alasan || '').trim();
    var timestamp = new Date();

    if (!targetNis) {
      return createJsonResponse('error', 'NIS tidak valid.');
    }

    // Buka spreadsheet dan sheet 'Data'
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return createJsonResponse('error', 'Sheet "Data" tidak ditemukan.');
    }

    // Ambil seluruh data baris
    var values = sheet.getDataRange().getValues();
    var foundIndex = -1;

    // Cari baris yang sesuai dengan NIS (Kolom A / index 0)
    for (var i = 1; i < values.length; i++) {
      var currentNis = String(values[i][0]).trim();
      if (currentNis.toLowerCase() === targetNis.toLowerCase()) {
        foundIndex = i + 1; // Konversi index array ke nomor baris spreadsheet (1-based index)
        break;
      }
    }

    if (foundIndex === -1) {
      return createJsonResponse('error', 'NIS "' + targetNis + '" tidak ditemukan.');
    }

    // Update data ke Google Sheet:
    // Baris: foundIndex, Kolom 4 (Kolom D) = Status
    // Baris: foundIndex, Kolom 5 (Kolom E) = Alasan
    // Baris: foundIndex, Kolom 6 (Kolom F) = Timestamp
    sheet.getRange(foundIndex, 4).setValue(newStatus);
    sheet.getRange(foundIndex, 5).setValue(newAlasan);
    sheet.getRange(foundIndex, 6).setValue(timestamp);

    return createJsonResponse('success', 'Status dan alasan berhasil diperbarui.');

  } catch (err) {
    return createJsonResponse('error', 'Terjadi kesalahan server: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Endpoint GET opsional untuk verifikasi kesehatan Web App
 */
function doGet(e) {
  return ContentService.createTextOutput("Backend Web App Cek Dokumen Siswa Aktif.")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Helper untuk membuat response JSON dengan standar CORS
 */
function createJsonResponse(result, message) {
  var output = JSON.stringify({
    result: result,
    message: message
  });

  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
