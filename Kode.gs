// Variabel Konfigurasi
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1BuY4yZ_CPzcJoUh6l1fEfMuQVQtCskctzppnBC_PDlU/';
const ADMIN_EMAIL = 'tubagus.budi@thopasindo.com'; 
const DRIVE_FOLDER_ID = '16T9zbcfXDLGo63zdAq0mZC-mELVqH9Gk'; 
const SENDER_NAME = 'No Reply - Sistem ATK';

// NAMA-NAMA SHEET
const ATK_LIST_SHEET_NAME = 'Daftar ATK';
const REQUEST_LOG_SHEET_NAME = 'Permintaan ATK';
const KARYAWAN_SHEET_NAME = 'Data Karyawan'; // Sheet baru untuk akun
const CUTI_SHEET_NAME = 'Data Cuti';

// =======================================================
// 1. SISTEM ROUTING HALAMAN
// =======================================================
function doGet(e) {
  try {
    var page = e.parameter.page;

    // DEFAULT ROUTING: Jika tidak ada parameter, paksa buka halaman Login
    if (!page || page == 'login') {
      return HtmlService.createTemplateFromFile('Login')
        .evaluate().setTitle('Login Portal HRIS')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    // Halaman Menu Utama / Dasbor Karyawan
    else if (page == 'menu') {
      return HtmlService.createTemplateFromFile('MenuUtama')
        .evaluate().setTitle('Portal HR & GA')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    // Halaman Dasbor Admin
    else if (page == 'admin') {
      return HtmlService.createTemplateFromFile('Dasbor')
        .evaluate().setTitle('Halaman Admin HRIS')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    // Halaman Form ATK, Cuti, dan Riwayat tetap sama
    else if (page == 'Form') {
      return HtmlService.createTemplateFromFile('Form')
        .evaluate().setTitle('Form Permintaan ATK').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    else if (page == 'FormCuti') {
      return HtmlService.createTemplateFromFile('FormCuti')
        .evaluate().setTitle('Form Pengajuan Cuti').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    else if (page == 'riwayat') {
      return HtmlService.createTemplateFromFile('Riwayat')
        .evaluate().setTitle('Riwayat Permintaan ATK').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    else {
      return HtmlService.createHtmlOutput("<h2>Halaman '" + page + "' tidak ditemukan.</h2>");
    }
  } catch (err) {
    return HtmlService.createHtmlOutput("<h2>Terjadi Kesalahan: " + err.message + "</h2>");
  }
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// =======================================================
// 2. FUNGSI AUTENTIKASI (LOGIN)
// =======================================================
function prosesLogin(email, password) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sistem Error: Sheet "Data Karyawan" tidak ditemukan.' };
    }

    const data = sheet.getDataRange().getValues();
    
    // Looping dari baris kedua (index 1) untuk melewati header
    for (let i = 1; i < data.length; i++) {
      let dbEmail = data[i][2].toString().trim().toLowerCase();
      let dbPassword = data[i][3].toString().trim();
      let inputEmail = email.trim().toLowerCase();
      
      // Jika email dan password cocok
      if (dbEmail === inputEmail && dbPassword === password) {
        return {
          success: true,
          userData: {
            id: data[i][0],
            nama: data[i][1],
            email: dbEmail,
            peran: data[i][4],     // Admin atau Karyawan
            jabatan: data[i][5],
            divisi: data[i][6],
            sisaCuti: data[i][7]
          }
        };
      }
    }
    
    return { success: false, message: 'Email atau Password salah!' };
    
  } catch (error) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + error.message };
  }
}

// =======================================================
// 3. FUNGSI ATK (Tetap Sama Seperti Sebelumnya)
// =======================================================
function getAtkItems() {
  const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(ATK_LIST_SHEET_NAME);
  const range = sheet.getRange('A2:B' + sheet.getLastRow());
  return range.getValues().map(row => ({ name: row[0], stock: parseInt(row[1]) || 0 })).filter(item => item.name && item.name.trim() !== '');
}

function submitRequest(formData) {
  try {
    const logSheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME);
    const atkSheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(ATK_LIST_SHEET_NAME);
    const atkData = atkSheet.getRange('A2:B' + atkSheet.getLastRow()).getValues();
    const timestamp = new Date();
    const requestId = `REQ-${timestamp.getTime()}`;
    const itemsRequested = [];

    for (const item of formData.requestedItems) {
      const stockInfo = atkData.find(row => row[0] === item.name);
      if (!stockInfo || item.quantity > stockInfo[1]) throw new Error(`Stok untuk ${item.name} tidak mencukupi.`);
      itemsRequested.push(`- ${item.name} (Jumlah: ${item.quantity})`);
    }

    formData.requestedItems.forEach(item => {
      logSheet.appendRow([requestId, timestamp, formData.employeeName, formData.employeeEmail, formData.department, item.name, item.quantity, 'Menunggu Persetujuan', '']);
    });
    
    MailApp.sendEmail({ to: ADMIN_EMAIL, subject: `Permintaan ATK Baru: ${requestId}`, body: `Ada permintaan ATK baru dari ${formData.employeeName}. Silakan cek Dasbor Admin.`, name: SENDER_NAME });
    sendSubmissionEmail(formData.employeeEmail, formData.employeeName, itemsRequested);

    return { success: true, message: 'Permintaan berhasil dikirim dan menunggu persetujuan.' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getUserRequests(email) {
  const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const requests = [];
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][3] && data[i][3].toString().trim().toLowerCase() === email.trim().toLowerCase()) {
      requests.push({ rowNum: i + 1, timestamp: new Date(data[i][1]).toISOString(), itemName: data[i][5], quantity: data[i][6], status: data[i][7], proofLink: data[i][8] || null });
    }
  }
  return requests;
}

function sendSubmissionEmail(toEmail, name, items) {
  if (!toEmail || toEmail === '') return;
  const subject = `[Sistem ATK] Permintaan Berhasil Diajukan`;
  let body = `Yth. ${name},\n\nPermintaan ATK Anda telah berhasil diajukan dan sedang menunggu persetujuan General Affairs.\n\nDetail:\n${items.join('\n')}\n\nPesan ini dikirim otomatis.`;
  try { MailApp.sendEmail({ to: toEmail, subject: subject, body: body, name: SENDER_NAME }); } catch (e) {}
}

// =======================================================
// 4. FUNGSI CUTI (Tetap Sama Sementara)
// =======================================================
function simpanDataCuti(dataForm) {
  try {
    var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(CUTI_SHEET_NAME);
    if (!sheet) return { status: 'error', message: 'Sheet "Data Cuti" tidak ditemukan di Spreadsheet!' };
    
    var timestamp = new Date();
    var status = "Menunggu Approval Atasan/HRD"; 
    var rowData = [timestamp, dataForm.nama, dataForm.jabatan, dataForm.divisi, dataForm.penempatan, dataForm.jenis_cuti, dataForm.tanggal_mulai, dataForm.tanggal_selesai, dataForm.lama_cuti, dataForm.alamat_cuti, dataForm.no_hp, dataForm.pengganti, status];
    
    sheet.appendRow(rowData);
    return { status: 'success', message: 'Pengajuan cuti berhasil dikirim! Silakan tunggu persetujuan.' };
  } catch (error) {
    return { status: 'error', message: 'Terjadi kesalahan sistem: ' + error.toString() };
  }
}
