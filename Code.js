/**
 * ============================================================================
 * SISTEM ARSIP SURAT PGRI KECAMATAN TEWEH SELATAN (SIAS PGRI-TESEL)
 * Backend Google Apps Script (Code.gs) - Versi Update
 * ============================================================================
 */

const CONFIG = {
  // 1. Tempelkan link Google Sheet Surat Keluar Anda di bawah ini:
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1HtFZkob_PJmToEsizaxZv8c87SnchXTHL9KBZ2uz_a4/edit?usp=sharing',
  
  // 2. Tempelkan link Folder Google Drive untuk SURAT MASUK di bawah ini:
  DRIVE_FOLDER_MASUK_URL: 'https://drive.google.com/drive/folders/1I-1R6k3DW5ad0w5XDU5u8Q2kZqsEcrXI?usp=drive_link',

  // 3. Tempelkan link Folder Google Drive untuk SURAT KELUAR di bawah ini:
  DRIVE_FOLDER_KELUAR_URL: 'https://drive.google.com/drive/folders/1k-3fwaJ7dcfAOX5p5N8S7M5Vt9ZL-cvC?usp=drive_link',
  
  // 4. Nama Sheet pada Google Spreadsheet
  SHEET_SURAT_KELUAR: 'Sheet1',
  SHEET_SURAT_MASUK: 'Surat Masuk',
  
  // 5. Daftar Akun Pengguna (Silakan sesuaikan username & password di sini)
  USERS: {
    'admin': { password: '123', nama: 'Admin PGRI Teweh Selatan', jabatan: 'Sekretariat PGRI' },
    'pgri': { password: 'pgri2026', nama: 'Pengurus PGRI', jabatan: 'Kec. Teweh Selatan' }
  }
};

/**
 * Menampilkan halaman utama aplikasi web
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('E-Arsip Surat | PGRI Kec. Teweh Selatan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper: Mengekstrak kode identifikasi dari tautan Google Drive / Google Sheets
 */
function extractIdFromUrl_(urlOrId) {
  if (!urlOrId) return '';
  const match = String(urlOrId).match(/[-\w]{25,}/);
  return match ? match[0] : String(urlOrId).trim();
}

/**
 * Helper: Membuka objek Spreadsheet
 */
function getSpreadsheet_() {
  const ssId = extractIdFromUrl_(CONFIG.SPREADSHEET_URL);
  if (ssId && !ssId.includes('TEMPEL_LINK')) {
    return SpreadsheetApp.openById(ssId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Helper: Membuka Folder Drive Surat Masuk
 */
function getFolderMasuk_() {
  const folderId = extractIdFromUrl_(CONFIG.DRIVE_FOLDER_MASUK_URL);
  if (folderId && !folderId.includes('TEMPEL_LINK')) {
    return DriveApp.getFolderById(folderId);
  }
  return DriveApp.getRootFolder();
}

/**
 * Helper: Membuka Folder Drive Surat Keluar
 */
function getFolderKeluar_() {
  const folderId = extractIdFromUrl_(CONFIG.DRIVE_FOLDER_KELUAR_URL);
  if (folderId && !folderId.includes('TEMPEL_LINK')) {
    return DriveApp.getFolderById(folderId);
  }
  return DriveApp.getRootFolder();
}

/**
 * Mengambil file logo pgri.jpg jika sudah diunggah ke Folder Drive
 */
function getLogoBase64() {
  try {
    const folder = getFolderMasuk_();
    const files = folder.getFilesByName('pgri.jpg');
    if (files.hasNext()) {
      const file = files.next();
      const base64 = Utilities.base64Encode(file.getBlob().getBytes());
      return 'data:' + file.getMimeType() + ';base64,' + base64;
    }
  } catch (e) {
    // Gunakan fallback SVG logo PGRI di frontend jika file belum ada
  }
  return null;
}

/**
 * Proses Verifikasi Login
 */
function verifyLogin(username, password) {
  const cleanUser = String(username || '').trim().toLowerCase();
  const cleanPass = String(password || '').trim();

  if (CONFIG.USERS[cleanUser] && CONFIG.USERS[cleanUser].password === cleanPass) {
    return {
      success: true,
      user: {
        username: cleanUser,
        nama: CONFIG.USERS[cleanUser].nama,
        jabatan: CONFIG.USERS[cleanUser].jabatan
      }
    };
  }
  return {
    success: false,
    message: 'Username atau password tidak sesuai. Silakan periksa kembali.'
  };
}

/**
 * Mengambil seluruh data awal untuk Dashboard, Surat Masuk, dan Surat Keluar
 */
function getInitialData() {
  try {
    const ss = getSpreadsheet_();
    
    // 1. Siapkan Sheet Surat Keluar (Sheet1)
    let sheetKeluar = ss.getSheetByName(CONFIG.SHEET_SURAT_KELUAR);
    if (!sheetKeluar) {
      sheetKeluar = ss.getSheets()[0] || ss.insertSheet(CONFIG.SHEET_SURAT_KELUAR);
    }
    if (sheetKeluar.getLastRow() === 0) {
      sheetKeluar.appendRow([
        'No Urut',
        'Nomor Surat Keluar',
        'Tentang / Perihal Surat',
        'Jenis Surat',
        'Tujuan Surat',
        'Bulan (Romawi)',
        'Tahun',
        'Tanggal Input',
        'Nama File',
        'Link Preview',
        'Link Download / View'
      ]);
      sheetKeluar.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#0F5132').setFontColor('#FFFFFF');
    }

    // 2. Siapkan Sheet Surat Masuk
    let sheetMasuk = ss.getSheetByName(CONFIG.SHEET_SURAT_MASUK);
    if (!sheetMasuk) {
      sheetMasuk = ss.insertSheet(CONFIG.SHEET_SURAT_MASUK);
    }
    if (sheetMasuk.getLastRow() === 0) {
      sheetMasuk.appendRow([
        'Waktu Upload',
        'Nomor Surat Masuk',
        'Tanggal Surat',
        'Pengirim / Asal Surat',
        'Perihal / Tentang',
        'Nama File',
        'Link Preview',
        'Link Download / View'
      ]);
      sheetMasuk.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#0F5132').setFontColor('#FFFFFF');
    }

    // Baca Data Surat Keluar
    const dataKeluar = [];
    if (sheetKeluar.getLastRow() > 1) {
      const maxCols = Math.max(11, sheetKeluar.getLastColumn());
      const rowsK = sheetKeluar.getRange(2, 1, sheetKeluar.getLastRow() - 1, maxCols).getDisplayValues();
      rowsK.forEach((r, idx) => {
        if (r[1] || r[2]) {
          dataKeluar.push({
            rowNumber: idx + 2,
            noUrut: r[0] || '',
            nomorSurat: r[1] || '',
            tentang: r[2] || '',
            jenis: r[3] || '',
            tujuan: r[4] || '-',
            bulan: r[5] || '',
            tahun: r[6] || '',
            tanggalInput: r[7] || '',
            namaFile: r[8] || '',
            previewUrl: r[9] || '',
            viewUrl: r[10] || ''
          });
        }
      });
    }

    // Baca Data Surat Masuk
    const dataMasuk = [];
    if (sheetMasuk.getLastRow() > 1) {
      const rowsM = sheetMasuk.getRange(2, 1, sheetMasuk.getLastRow() - 1, 8).getDisplayValues();
      rowsM.forEach((r, idx) => {
        if (r[1] || r[4]) {
          dataMasuk.push({
            rowNumber: idx + 2,
            waktuUpload: r[0] || '',
            nomorSurat: r[1] || '',
            tanggalSurat: r[2] || '',
            pengirim: r[3] || '',
            perihal: r[4] || '',
            namaFile: r[5] || '',
            previewUrl: r[6] || '',
            viewUrl: r[7] || ''
          });
        }
      });
    }

    // Hitung nomor urut berikutnya untuk Surat Keluar
    const nextNumber = dataKeluar.length + 1;
    const formattedNextNumber = String(nextNumber).padStart(2, '0');

    return {
      success: true,
      suratKeluar: dataKeluar.reverse(),
      suratMasuk: dataMasuk.reverse(),
      nextUrut: formattedNextNumber
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal memuat data: ' + err.message
    };
  }
}

/**
 * Menyimpan Data Surat Keluar ke Google Sheet & Upload Dokumen ke Folder Drive Surat Keluar
 */
function saveSuratKeluar(payload) {
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName(CONFIG.SHEET_SURAT_KELUAR) || ss.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'No Urut',
        'Nomor Surat Keluar',
        'Tentang / Perihal Surat',
        'Jenis Surat',
        'Tujuan Surat',
        'Bulan (Romawi)',
        'Tahun',
        'Tanggal Input',
        'Nama File',
        'Link Preview',
        'Link Download / View'
      ]);
    }

    let finalFileName = '-';
    let previewUrl = '';
    let fileUrl = '';

    // Jika terdapat file surat keluar yang diunggah, simpan ke Folder Drive Surat Keluar
    if (payload.base64Data && payload.fileName) {
      const folderKeluar = getFolderKeluar_();
      const contentType = payload.mimeType || 'application/pdf';
      const bytes = Utilities.base64Decode(payload.base64Data);
      const cleanNomor = String(payload.nomorSurat || 'SURAT_KELUAR').replace(/[^a-zA-Z0-9-_]/g, '_');
      finalFileName = cleanNomor + '_' + payload.fileName;
      const blob = Utilities.newBlob(bytes, contentType, finalFileName);

      const driveFile = folderKeluar.createFile(blob);
      driveFile.setDescription('Surat Keluar: ' + payload.nomorSurat + ' - ' + payload.tentang);

      fileUrl = driveFile.getUrl();
      previewUrl = fileUrl.replace(/\/view.*$/, '/preview');
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    sheet.appendRow([
      "'" + payload.noUrut,
      payload.nomorSurat,
      payload.tentang,
      payload.jenis,
      payload.tujuan || '-',
      payload.bulan,
      payload.tahun,
      timestamp,
      finalFileName,
      previewUrl,
      fileUrl
    ]);

    return {
      success: true,
      message: payload.base64Data
        ? 'Surat keluar berhasil diunggah ke Drive & dicatat ke Google Sheet!'
        : 'Nomor & data surat keluar berhasil disimpan ke Google Sheet!',
      fileData: {
        nomorSurat: payload.nomorSurat,
        tentang: payload.tentang,
        namaFile: finalFileName,
        previewUrl: previewUrl,
        viewUrl: fileUrl
      }
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal menyimpan surat keluar: ' + err.message
    };
  }
}

/**
 * Mengunggah File Surat Masuk ke Folder Google Drive & Mencatatnya ke Sheet
 */
function uploadSuratMasuk(payload) {
  try {
    const folderMasuk = getFolderMasuk_();
    
    // Decode file dari Base64
    const contentType = payload.mimeType || 'application/pdf';
    const bytes = Utilities.base64Decode(payload.base64Data);
    const cleanNomor = String(payload.nomorSurat || 'SURAT_MASUK').replace(/[^a-zA-Z0-9-_]/g, '_');
    const finalFileName = cleanNomor + '_' + payload.fileName;
    const blob = Utilities.newBlob(bytes, contentType, finalFileName);

    // Simpan ke Folder Google Drive Surat Masuk
    const driveFile = folderMasuk.createFile(blob);
    driveFile.setDescription('Surat Masuk: ' + payload.nomorSurat + ' - ' + payload.perihal);

    const fileUrl = driveFile.getUrl();
    const previewUrl = fileUrl.replace(/\/view.*$/, '/preview');

    // Catat ke Sheet Surat Masuk
    const ss = getSpreadsheet_();
    let sheetMasuk = ss.getSheetByName(CONFIG.SHEET_SURAT_MASUK);
    if (!sheetMasuk) {
      sheetMasuk = ss.insertSheet(CONFIG.SHEET_SURAT_MASUK);
      sheetMasuk.appendRow([
        'Waktu Upload',
        'Nomor Surat Masuk',
        'Tanggal Surat',
        'Pengirim / Asal Surat',
        'Perihal / Tentang',
        'Nama File',
        'Link Preview',
        'Link Download / View'
      ]);
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    sheetMasuk.appendRow([
      timestamp,
      payload.nomorSurat,
      payload.tanggalSurat || '-',
      payload.pengirim || '-',
      payload.perihal || '-',
      finalFileName,
      previewUrl,
      fileUrl
    ]);

    return {
      success: true,
      message: 'Surat masuk berhasil diunggah ke Google Drive!',
      fileData: {
        nomorSurat: payload.nomorSurat,
        perihal: payload.perihal,
        namaFile: finalFileName,
        previewUrl: previewUrl,
        viewUrl: fileUrl
      }
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal mengunggah surat masuk: ' + err.message
    };
  }
}