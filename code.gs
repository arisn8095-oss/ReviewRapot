/**
 * GOOGLE APPS SCRIPT BACKEND FOR DOKUMEN SISWA
 * Spreadsheet ID: 1ViRbaN_tNOPD4M8kNbwdxKEXYzhD4DVLn5MWf10KVsI
 * Sheet Tab Name: Data2
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
const SHEET_NAME = 'Data2';

/**
 * FUNGSI OTOMATIS MEMBUAT & MEMFORMAT DATABASE GOOGLE SHEET
 * Jalankan fungsi ini 1x dari Apps Script Editor untuk menyiapkan header tab Data2 & data contoh.
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

  // Cari atau buat tab 'Data2'
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

  // Visual Styling Header
  headerRange.setBackground('#0284c7'); // Biru Brand
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // 2. Isi Data Contoh / Dummy (Jika baris masih kosong)
  if (sheet.getLastRow() <= 1) {
    var sampleData = [
      ['1001', 'Ahmad Rizky Pratama', 'https://raw.githubusercontent.com/mozilla/pdf.js/ba2edeae/web/compressed.tracemonkey-pldi-09.pdf', '', '', ''],
      ['1002', 'Siti Nurhaliza', 'https://raw.githubusercontent.com/mozilla/pdf.js/ba2edeae/web/compressed.tracemonkey-pldi-09.pdf', 'PERLU EDIT KEMBALI', 'Mohon perbaiki tanggal lahir pada lembar kedua menjadi 12 Mei 2008.', ''],
      ['1003', 'Budi Santoso', 'https://raw.githubusercontent.com/mozilla/pdf.js/ba2edeae/web/compressed.tracemonkey-pldi-09.pdf', 'SETUJU CETAK', '', '']
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
  sheet.setColumnWidth(2, 230); // Identitas / Nama
  sheet.setColumnWidth(3, 350); // Link Dokumen PDF
  sheet.setColumnWidth(4, 190); // Status Persetujuan
  sheet.setColumnWidth(5, 300); // Alasan Edit Kembali
  sheet.setColumnWidth(6, 190); // Waktu Konfirmasi

  // Bekukan Baris Header
  sheet.setFrozenRows(1);

  Logger.log('Database tab Data2 berhasil disiapkan!');
  return 'Database tab "Data2" berhasil dibuat dan disiapkan!';
}

/**
 * Menerima kiriman data HTTP POST dari aplikasi web frontend
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Tunggu maksimal 10 detik untuk menghindari konflik penulisan bersamaan
  lock.tryLock(10000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse('error', 'Tidak ada payload data yang diterima.');
    }

    var requestData = JSON.parse(e.postData.contents);
    var targetNis = String(requestData.nis || '').trim();
    var newStatus = String(requestData.status || '').trim();
    var newAlasan = String(requestData.alasan || '').trim();
    var timestamp = new Date();

    if (!targetNis) {
      return createJsonResponse('error', 'NIS siswa wajib diisi.');
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return createJsonResponse('error', 'Sheet "' + SHEET_NAME + '" tidak ditemukan.');
    }

    var values = sheet.getDataRange().getValues();
    var foundIndex = -1;

    // Cari baris yang sesuai dengan NIS (Kolom A / index 0)
    for (var i = 1; i < values.length; i++) {
      var currentNis = String(values[i][0]).trim();
      if (currentNis.toLowerCase() === targetNis.toLowerCase()) {
        foundIndex = i + 1; // Konversi index array ke baris spreadsheet
        break;
      }
    }

    if (foundIndex === -1) {
      return createJsonResponse('error', 'NIS "' + targetNis + '" tidak ditemukan di tab ' + SHEET_NAME);
    }

    // Simpan pembaruan status, alasan, dan waktu konfirmasi
    sheet.getRange(foundIndex, 4).setValue(newStatus);
    sheet.getRange(foundIndex, 5).setValue(newAlasan);
    sheet.getRange(foundIndex, 6).setValue(timestamp);

    return createJsonResponse('success', 'Status untuk NIS ' + targetNis + ' berhasil diperbarui.');

  } catch (err) {
    return createJsonResponse('error', 'Kesalahan server: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Endpoint GET untuk pengecekan status server Web App
 */
function doGet(e) {
  return ContentService.createTextOutput("Backend Web App Dokumen Siswa (Tab Data2) Aktif dan Berjalan.")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Helper untuk membuat response JSON
 */
function createJsonResponse(result, message) {
  var output = JSON.stringify({
    result: result,
    message: message
  });

  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
