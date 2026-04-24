// Google Apps Script untuk Menangani Form HKI
// Simpan script ini di editor Google Apps Script yang terhubung ke Google Spreadsheet Anda.

// ID Folder Google Drive tempat menyimpan Foto KTP
// Ganti dengan ID Folder Anda (Lihat dari URL folder di Drive)
const DRIVE_FOLDER_ID = '1HVV0TFLsJj1WoB_zxCOpFvq9s8jtLhvO'; 

// Nama Sheet tempat data akan disimpan
// Ganti dengan nama Sheet yang ada di Spreadsheet Anda (misal: 'Sheet1')
const SHEET_NAME = 'Sheet1';

function doPost(e) {
  try {
    // Membaca payload dari request POST
    Logger.log("doPost dipanggil. Raw contents: " + (e.postData ? e.postData.contents.substring(0, 200) : "NULL"));
    const data = JSON.parse(e.postData.contents);
    Logger.log("Data diterima untuk NIK: " + data.nik + ", Nama: " + data.nama);
    
    // Membuka Spreadsheet yang aktif (tempat script ini berada)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Jika sheet tidak ditemukan, gunakan sheet pertama sebagai fallback
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // Header pengecekan (Jika baris pertama kosong, buat header)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "NIK", "Nama Lengkap", "Email", "No. Telepon", 
        "Jenis Kelamin", "NPWP", "Alamat", "Negara", "Provinsi", 
        "Kabupaten/Kota", "Kecamatan", "Kelurahan", "Kode Pos",
        "Jenis Permohonan", "Jenis Ciptaan", "Sub-Jenis Ciptaan", 
        "Tanggal Diumumkan", "Judul", "Uraian Singkat", "Data Anggota", "URL KTP"
      ]);
      // Styling header
      sheet.getRange(1, 1, 1, 22).setFontWeight("bold").setBackground("#E3000F").setFontColor("white");
      sheet.setFrozenRows(1);
    }
    
    // Menyiapkan folder Drive
    let folder;
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      Logger.log("Folder Drive ditemukan: " + folder.getName());
    } catch (folderError) {
      Logger.log("Folder error: " + folderError.toString() + " — fallback ke root.");
      // Jika folder gagal diakses, simpan ke root (hanya sebagai fallback)
      folder = DriveApp.getRootFolder();
    }
    
    // Mendekode file base64 menjadi blob gambar
    Logger.log("Mendekode file: " + data.fileName + ", mimeType: " + data.mimeType);
    const decodedFile = Utilities.base64Decode(data.fileData);
    const blob = Utilities.newBlob(decodedFile, data.mimeType, data.judul + "_" + data.fileName);
    
    // Menyimpan blob ke folder Drive
    const file = folder.createFile(blob);
    Logger.log("File berhasil disimpan ke Drive: " + file.getUrl());
    
    // Mendapatkan URL file
    const fileUrl = file.getUrl();
    
    // Menyisipkan data pemohon utama ke baris baru di Spreadsheet
    sheet.appendRow([
      new Date(),           // Timestamp
      "'" + data.nik,
      data.nama,
      data.email,
      "'" + data.telepon,
      data.jenisKelamin,
      "'" + data.npwp,
      data.alamat,
      data.negara,
      data.provinsi,
      data.kota,
      data.kecamatan,
      data.kelurahan,
      "'" + data.kodePos,
      data.jenisPermohonan,
      data.jenisCiptaan,
      data.subJenisCiptaan,
      data.tanggalDiumumkan,
      data.judul,
      data.uraianSingkat,
      "Pemohon Utama",
      fileUrl
    ]);

    // Menyisipkan data anggota tambahan, masing-masing satu baris
    if (data.anggotaList && data.anggotaList.length > 0) {
      data.anggotaList.forEach((anggota, idx) => {
        sheet.appendRow([
          new Date(),
          "'" + anggota.nik,
          anggota.nama,
          anggota.email,
          "'" + anggota.telepon,
          anggota.jenisKelamin,
          "'" + (anggota.npwp || '-'),
          anggota.alamat,
          "Indonesia",
          anggota.provinsi,
          anggota.kota,
          anggota.kecamatan,
          anggota.kelurahan,
          '-',
          data.jenisPermohonan,
          data.jenisCiptaan,
          data.subJenisCiptaan,
          data.tanggalDiumumkan,
          data.judul,
          data.uraianSingkat,
          "Anggota " + (idx + 1),
          fileUrl
        ]);
      });
    }
    
    Logger.log("Data berhasil di-append ke sheet.");
    
    // Mengembalikan respons sukses ke client
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data berhasil disimpan.'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // Mengembalikan respons error jika terjadi kesalahan
    Logger.log("ERROR: " + error.toString() + "\nStack: " + error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Opsi pre-flight request untuk menangani CORS jika diperlukan
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.JSON);
}

// Verifikasi apakah data NIK sudah tersimpan
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'verify') {
      const nik = e.parameter.nik;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
      
      if (sheet.getLastRow() <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ found: false }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      // Cari NIK di kolom B (kolom 2), skip header
      const nikColumn = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
      const found = nikColumn.some(row => String(row[0]).replace("'", "") === String(nik));
      
      return ContentService.createTextOutput(JSON.stringify({ found: found }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
