// Variabel Konfigurasi
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1BuY4yZ_CPzcJoUh6l1fEfMuQVQtCskctzppnBC_PDlU/';
const ADMIN_EMAIL = 'tubagus.budi@thopasindo.com'; 
const SENDER_NAME = 'HRIS & Logistic Portal';

// NAMA-NAMA SHEET
const ATK_LIST_SHEET_NAME = 'Daftar ATK';
const REQUEST_LOG_SHEET_NAME = 'Permintaan ATK';
const KARYAWAN_SHEET_NAME = 'Data Karyawan';
const CUTI_SHEET_NAME = 'Data Cuti';

// FUNGSI AUTO-CREATE SHEET UNTUK PENGUMUMAN & KALENDER
function checkAndCreateSheets() {
  const ss = SpreadsheetApp.openByUrl(SHEET_URL);
  if(!ss.getSheetByName('Pengumuman')) {
    let sh = ss.insertSheet('Pengumuman');
    sh.appendRow(['Judul Pengumuman', 'Tanggal Berlaku', 'Isi Pengumuman', 'Warna']);
  }
  if(!ss.getSheetByName('Libur Nasional')) {
    let sh = ss.insertSheet('Libur Nasional');
    sh.appendRow(['Tanggal', 'Nama Libur', 'Keterangan', 'Status']);
  }
}

function safeDateString(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString();
  return val.toString();
}

function hashPassword(password) {
  if (!password) return "";
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) { hashVal += 256; }
    if (hashVal.toString(16).length == 1) { txtHash += '0'; }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function doGet(e) {
  try {
    checkAndCreateSheets(); // Auto-inisialisasi database
    var page = e.parameter.page;
    if (!page || page == 'login') return HtmlService.createTemplateFromFile('Login').evaluate().setTitle('Login Portal HRIS').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    else if (page == 'menu') return HtmlService.createTemplateFromFile('MenuUtama').evaluate().setTitle('Portal HR & GA').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    else if (page == 'admin') return HtmlService.createTemplateFromFile('Dasbor').evaluate().setTitle('Halaman Admin HRIS').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    else if (page == 'Form') return HtmlService.createTemplateFromFile('Form').evaluate().setTitle('Form Permintaan ATK').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    else if (page == 'FormCuti') return HtmlService.createTemplateFromFile('FormCuti').evaluate().setTitle('Form Pengajuan Cuti').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    else if (page == 'riwayat') return HtmlService.createTemplateFromFile('Riwayat').evaluate().setTitle('Riwayat Permintaan ATK').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    else return HtmlService.createHtmlOutput("<h2>Halaman '" + page + "' tidak ditemukan.</h2>");
  } catch (err) { return HtmlService.createHtmlOutput("<h2>Terjadi Kesalahan: " + err.message + "</h2>"); }
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

// ==========================================
// 2. FUNGSI AUTENTIKASI & USER
// ==========================================
function prosesLogin(email, password) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const inputHash = hashPassword(password);
    for (let i = 1; i < data.length; i++) {
      let dbEmail = data[i][2] ? data[i][2].toString().trim().toLowerCase() : '';
      let dbPassword = data[i][3] ? data[i][3].toString().trim() : '';
      if (dbEmail === email.trim().toLowerCase()) {
        if (dbPassword === inputHash || dbPassword === password) {
          return { success: true, userData: { id: data[i][0], nama: data[i][1], email: dbEmail, peran: data[i][4], jabatan: data[i][5], divisi: data[i][6], sisaCuti: data[i][7] } };
        }
      }
    }
    return { success: false, message: 'Email atau Password salah!' };
  } catch (error) { return { success: false, message: error.message }; }
}

function getSisaCuti(email) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const targetEmail = (email || '').toString().trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      let dbEmail = data[i][2] ? data[i][2].toString().trim().toLowerCase() : '';
      if (dbEmail === targetEmail) return data[i][7];
    }
    return '-';
  } catch(e) { return '-'; }
}

function getEmailByName(nama) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const targetName = nama.trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      let dbName = data[i][1] ? data[i][1].toString().trim().toLowerCase() : '';
      if (dbName === targetName) return data[i][2].toString().trim();
    }
  } catch(e) {}
  return null;
}

function sendCutiEmail(toEmail, nama, status, jenisCuti, tglMulai, tglSelesai, alasan) {
    if (!toEmail || toEmail === '') return;
    const subject = `[HRIS] Status Pengajuan Cuti Anda: ${status.split(':')[0]}`;
    let htmlBody = `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;"><div style="background-color: #4f46e5; color: white; padding: 20px; text-align: center;"><h2 style="margin: 0;">Notifikasi Pengajuan Cuti</h2></div><div style="padding: 20px; color: #374151;"><p>Halo <b>${nama}</b>,</p><p>Pengajuan cuti Anda telah di-update dengan detail sebagai berikut:</p><table style="width: 100%; border-collapse: collapse; margin: 20px 0;"><tr><td style="padding: 8px; border-bottom: 1px solid #e5e7eb; width: 40%; color: #6b7280;">Jenis Cuti</td><td style="padding: 8px; border-bottom: 1px solid #e5e7eb; font-weight: bold;">${jenisCuti}</td></tr><tr><td style="padding: 8px; border-bottom: 1px solid #e5e7eb; color: #6b7280;">Tanggal</td><td style="padding: 8px; border-bottom: 1px solid #e5e7eb; font-weight: bold;">${new Date(tglMulai).toLocaleDateString('id-ID')} s/d ${new Date(tglSelesai).toLocaleDateString('id-ID')}</td></tr><tr><td style="padding: 8px; border-bottom: 1px solid #e5e7eb; color: #6b7280;">Status Keputusan</td><td style="padding: 8px; border-bottom: 1px solid #e5e7eb; font-weight: bold; color: ${status.includes('Disetujui') ? '#16a34a' : '#dc2626'};">${status}</td></tr></table><p style="font-size: 12px; color: #9ca3af; margin-top: 30px; text-align: center;">Email ini dibuat secara otomatis oleh sistem HRIS. Mohon tidak membalas email ini.</p></div></div>`;
    try { MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: htmlBody, name: SENDER_NAME }); } catch (e) {}
}

function simpanDataCuti(dataForm) {
  try {
    SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(CUTI_SHEET_NAME).appendRow([
      new Date(), dataForm.nama, dataForm.jabatan, dataForm.divisi, dataForm.penempatan, 
      dataForm.jenis_cuti, dataForm.tanggal_mulai, dataForm.tanggal_selesai, dataForm.lama_cuti, 
      dataForm.alamat_cuti, dataForm.no_hp, dataForm.pengganti, "Menunggu Persetujuan"
    ]);
    try { MailApp.sendEmail({ to: ADMIN_EMAIL, subject: `[HRIS] Pengajuan Cuti Baru - ${dataForm.nama}`, body: `Ada pengajuan cuti baru dari ${dataForm.nama}. Silakan login ke Portal HRIS Admin.`, name: SENDER_NAME }); } catch(e){}
    return { status: 'success', message: 'Pengajuan cuti berhasil dikirim!' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function submitRequest(formData) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME);
    if(!sheet) throw new Error("Sheet Permintaan ATK tidak ditemukan.");
    const timestamp = new Date();
    const requestId = `REQ-${timestamp.getTime()}`;
    formData.requestedItems.forEach(item => { sheet.appendRow([requestId, timestamp, formData.employeeName, formData.employeeEmail, formData.department, item.name, item.quantity, 'Menunggu Persetujuan', '']); });
    try { MailApp.sendEmail({ to: ADMIN_EMAIL, subject: `[Logistik] Permintaan Barang Baru: ${formData.employeeName}`, body: `Ada permintaan ATK baru dari ${formData.employeeName}. Silakan cek Dasbor Admin.`, name: SENDER_NAME }); } catch(e){}
    return { success: true, message: 'Permintaan ATK berhasil dikirim.' };
  } catch (error) { return { success: false, message: error.message }; }
}

function getAtkItems() { return SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(ATK_LIST_SHEET_NAME).getRange('A2:B' + SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(ATK_LIST_SHEET_NAME).getLastRow()).getValues().map(row => ({ name: row[0], stock: parseInt(row[1]) || 0 })).filter(item => item.name && item.name.trim() !== ''); }

function getUserRequests(email) { 
  try {
      const data = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME).getDataRange().getValues(); 
      const requests = []; 
      const targetEmail = (email || '').toString().trim().toLowerCase();
      for (let i = data.length - 1; i > 0; i--) { 
        let sheetEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : '';
        if (sheetEmail === targetEmail) { requests.push({ rowNum: i + 1, timestamp: safeDateString(data[i][1]), itemName: data[i][5], quantity: data[i][6], status: data[i][7], proofLink: data[i][8] || null }); } 
      } 
      return requests; 
  } catch(e) { return []; }
}

function getUserCuti(nama) { 
  try {
      const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(CUTI_SHEET_NAME);
      if(!sheet) return [];
      const data = sheet.getDataRange().getValues(); 
      const requests = []; 
      const targetName = (nama || '').toString().trim().toLowerCase();
      for (let i = data.length - 1; i > 0; i--) { 
        let sheetName = data[i][1] ? data[i][1].toString().trim().toLowerCase() : '';
        if (sheetName === targetName) { 
          let st = data[i][12] ? data[i][12].toString() : 'Menunggu Persetujuan';
          requests.push({ timestamp: safeDateString(data[i][0]), nama: data[i][1], jabatan: data[i][2], divisi: data[i][3], penempatan: data[i][4], jenis: data[i][5], tglMulai: safeDateString(data[i][6]), tglSelesai: safeDateString(data[i][7]), lama: data[i][8], alamat: data[i][9], noHp: data[i][10], pengganti: data[i][11], status: st }); 
        } 
      } 
      return requests; 
  } catch(e) { return []; }
}

function confirmPickupWithPhoto(fileData, rowNum) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64Data.split(',')[1]), fileData.mimeType || 'image/png', fileData.fileName);
    const fileUrl = folder.createFile(blob).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW).getUrl();
    sheet.getRange(rowNum, 8).setValue('Sudah Diambil'); sheet.getRange(rowNum, 9).setValue(fileUrl); 
    return { success: true, newStatus: 'Sudah Diambil', proofLink: fileUrl };
  } catch (e) { return { success: false, message: 'Gagal menyimpan foto: ' + e.message }; }
}

function getAllCuti() { const data = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(CUTI_SHEET_NAME).getDataRange().getValues(); const requests = []; for (let i = data.length - 1; i > 0; i--) { requests.push({ rowNum: i + 1, timestamp: safeDateString(data[i][0]), nama: data[i][1], divisi: data[i][3], jenis: data[i][5], tglMulai: safeDateString(data[i][6]), tglSelesai: safeDateString(data[i][7]), lama: data[i][8], alasan: data[i][9], status: data[i][12] }); } return requests; }

function processCutiApproval(rowNum, isApproved, reason, namaKaryawan, lamaCuti) { try { const sheetCuti = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(CUTI_SHEET_NAME); const dataCuti = sheetCuti.getDataRange().getValues(); const status = isApproved ? 'Disetujui' : `Ditolak: ${reason}`; sheetCuti.getRange(rowNum, 13).setValue(status); const reqData = dataCuti[rowNum - 1]; const jenisCuti = reqData[5]; const tglMulai = reqData[6]; const tglSelesai = reqData[7]; if (isApproved && namaKaryawan && lamaCuti) { const sheetKaryawan = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME); const dataKar = sheetKaryawan.getDataRange().getValues(); const targetName = namaKaryawan.trim().toLowerCase(); for(let i = 1; i < dataKar.length; i++){ let sheetKarName = dataKar[i][1] ? dataKar[i][1].toString().trim().toLowerCase() : ''; if(sheetKarName === targetName) { let sisaLama = parseInt(dataKar[i][7]) || 0; sheetKaryawan.getRange(i+1, 8).setValue(sisaLama - parseInt(lamaCuti)); break; } } } const emailKaryawan = getEmailByName(namaKaryawan); if(emailKaryawan) sendCutiEmail(emailKaryawan, namaKaryawan, status, jenisCuti, tglMulai, tglSelesai, reason); return { success: true, message: 'Status Cuti diupdate.' }; } catch (e) { return { success: false, message: e.message }; } }

function getKaryawanData() { const data = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME).getDataRange().getValues(); const karyawan = []; for(let i=1; i<data.length; i++) { karyawan.push({ rowNum: i+1, id: data[i][0], nama: data[i][1], email: data[i][2], peran: data[i][4], jabatan: data[i][5], divisi: data[i][6], sisaCuti: data[i][7] }); } return karyawan; }
function simpanKaryawanBaru(data) { try { let securePassword = hashPassword(data.password); SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME).appendRow([data.id, data.nama, data.email, securePassword, data.peran, data.jabatan, data.divisi, data.sisaCuti]); return {success: true, message: 'Karyawan ditambahkan.'}; } catch(e) { return {success: false, message: e.message}; } }
function updateKaryawan(data) { try { const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME); const sheetData = sheet.getDataRange().getValues(); for (let i = 1; i < sheetData.length; i++) { if (sheetData[i][0].toString() === data.id.toString()) { sheet.getRange(i + 1, 2).setValue(data.nama); sheet.getRange(i + 1, 3).setValue(data.email); if (data.password && data.password.trim() !== '') { sheet.getRange(i + 1, 4).setValue(hashPassword(data.password)); } sheet.getRange(i + 1, 5).setValue(data.peran); sheet.getRange(i + 1, 6).setValue(data.jabatan); sheet.getRange(i + 1, 7).setValue(data.divisi); sheet.getRange(i + 1, 8).setValue(data.sisaCuti); return { success: true, message: 'Data diupdate.' }; } } return { success: false, message: 'Karyawan tidak ditemukan.' }; } catch(e) { return { success: false, message: e.message }; } }
function hapusKaryawan(id) { try { const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0].toString() === id.toString()) { sheet.deleteRow(i + 1); return { success: true, message: 'Karyawan dihapus.' }; } } return { success: false, message: 'Karyawan tidak ditemukan.' }; } catch(e) { return { success: false, message: e.message }; } }
function getPendingRequests() { const data = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME).getDataRange().getValues(); const requests = []; for (let i = 1; i < data.length; i++) { if (data[i][7] === 'Menunggu Persetujuan') { requests.push({ rowNum: i + 1, timestamp: safeDateString(data[i][1]), employeeName: data[i][2], department: data[i][4], itemName: data[i][5], quantity: data[i][6] }); } } return requests; }
function processApproval(rowNum, isApproved, reason) { try { SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME).getRange(rowNum, 8).setValue(isApproved ? 'Disetujui' : `Ditolak: ${reason}`); return { success: true, message: 'Status diupdate.' }; } catch (e) { return { success: false, message: e.message }; } }
function getConfirmedProofRecords() { const data = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME).getDataRange().getValues(); const records = []; for (let i = data.length - 1; i > 0; i--) { if (data[i][7] === 'Sudah Diambil' || (data[i][8] && data[i][8].toString().includes('http'))) { records.push({ rowNum: i + 1, timestamp: safeDateString(data[i][1]), employeeName: data[i][2], department: data[i][4], itemName: data[i][5], quantity: data[i][6], proofLink: data[i][8] }); } } return records; }
function addStock(name, qty) { try { const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(ATK_LIST_SHEET_NAME); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === name) { const currentStock = parseInt(data[i][1]) || 0; sheet.getRange(i + 1, 2).setValue(currentStock + parseInt(qty)); return { success: true, message: `Stok ${name} ditambah.` }; } } return { success: false, message: 'Barang tak ditemukan.' }; } catch (e) { return { success: false, message: e.message }; } }
function addNewItem(name, stock, limit) { try { SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(ATK_LIST_SHEET_NAME).appendRow([name, stock, limit]); return { success: true, message: 'Barang baru ditambahkan.' }; } catch (e) { return { success: false, message: e.message }; } }

function getDashboardData() {
  try {
    const atkData = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(ATK_LIST_SHEET_NAME).getDataRange().getValues();
    const reqData = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME).getDataRange().getValues();
    const cutiData = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(CUTI_SHEET_NAME).getDataRange().getValues();
    const karData = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME).getDataRange().getValues();
    
    const lowStockItems = [];
    for(let i = 1; i < atkData.length; i++) { let stock = parseInt(atkData[i][1]) || 0; let threshold = parseInt(atkData[i][2]) || 5; if (stock <= threshold && atkData[i][0]) lowStockItems.push({ name: atkData[i][0], stock: stock, threshold: threshold }); }
    let itemCounts = {}, deptCounts = {};
    for (let i = 1; i < reqData.length; i++) { let dept = reqData[i][4], item = reqData[i][5], qty = parseInt(reqData[i][6]) || 0; if (item) itemCounts[item] = (itemCounts[item] || 0) + qty; if (dept) deptCounts[dept] = (deptCounts[dept] || 0) + 1;  }
    const topItems = Object.entries(itemCounts).sort((a, b) => b[1] - a[1]).slice(0, 5); const topDepts = Object.entries(deptCounts).sort((a, b) => b[1] - a[1]);

    const totalKaryawan = karData.length > 1 ? karData.length - 1 : 0;
    let totalMenungguCuti = 0; let sedangCutiCount = 0; let cutiMendatang = []; let jenisCutiCount = {}; let employeeLeaveFreq = {}; 
    let trendData = {}; for (let m = 5; m >= 0; m--) { let d = new Date(); d.setMonth(d.getMonth() - m); let monthName = d.toLocaleString('id-ID', { month: 'short' }); trendData[monthName] = 0; }
    let now = new Date(); now.setHours(0,0,0,0); 

    for (let i = 1; i < cutiData.length; i++) {
        let timestampRaw = cutiData[i][0]; let nama = cutiData[i][1]; let status = cutiData[i][12] ? cutiData[i][12].toString() : ''; let jenis = cutiData[i][5]; let tglMulaiRaw = cutiData[i][6]; let tglSelesaiRaw = cutiData[i][7];
        if (status.includes('Menunggu')) totalMenungguCuti++;
        if(jenis && !status.includes('Ditolak')) { jenisCutiCount[jenis] = (jenisCutiCount[jenis] || 0) + 1; if(nama) employeeLeaveFreq[nama] = (employeeLeaveFreq[nama] || 0) + 1; }
        if (timestampRaw) { let tsDate = new Date(timestampRaw); let mName = tsDate.toLocaleString('id-ID', { month: 'short' }); if (trendData[mName] !== undefined) trendData[mName]++; }
        if (status === 'Disetujui' && tglMulaiRaw && tglSelesaiRaw) {
            let tglMulai = new Date(tglMulaiRaw); tglMulai.setHours(0,0,0,0); let tglSelesai = new Date(tglSelesaiRaw); tglSelesai.setHours(23,59,59,999);
            if (now >= tglMulai && now <= tglSelesai) sedangCutiCount++; else if (tglMulai > now) cutiMendatang.push({ nama: nama, jenis: jenis, tglMulai: safeDateString(tglMulaiRaw), tglSelesai: safeDateString(tglSelesaiRaw), lama: cutiData[i][8] });
        }
    }
    
    let topEmpName = "-"; let topEmpFreq = 0;
    for (let emp in employeeLeaveFreq) { if (employeeLeaveFreq[emp] > topEmpFreq) { topEmpFreq = employeeLeaveFreq[emp]; topEmpName = emp; } }
    cutiMendatang.sort((a, b) => new Date(a.tglMulai) - new Date(b.tglMulai));
    
    return { lowStockItems: lowStockItems, topItems: topItems, topDepts: topDepts, cutiStats: { totalKaryawan: totalKaryawan, pengajuanBaru: totalMenungguCuti, sedangCuti: sedangCutiCount, jadwalMendatang: cutiMendatang.slice(0, 5), distribusi: Object.entries(jenisCutiCount), topEmployee: { name: topEmpName, count: topEmpFreq }, trendLabels: Object.keys(trendData), trendValues: Object.values(trendData) } };
  } catch (e) { return { lowStockItems: [], topItems: [], topDepts: [], cutiStats: {totalKaryawan: 0, pengajuanBaru: 0, sedangCuti: 0, jadwalMendatang: [], distribusi: [], topEmployee: { name: '-', count: 0 }, trendLabels: [], trendValues: []} }; }
}

function downloadReport(start, end) {
  try {
    const data = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(REQUEST_LOG_SHEET_NAME).getDataRange().getDisplayValues();
    let csvData = data[0].join(',') + '\n'; 
    const startDate = start ? new Date(start) : null; const endDate = end ? new Date(end) : null;
    if(endDate) endDate.setHours(23,59,59);
    for(let i = 1; i < data.length; i++) {
      let rowDate = new Date(data[i][1]); let include = true;
      if (startDate && rowDate < startDate) include = false; if (endDate && rowDate > endDate) include = false;
      if(include) { let row = data[i].map(cell => `"${cell.toString().replace(/"/g, '""')}"`); csvData += row.join(',') + '\n'; }
    }
    return { success: true, csv: csvData };
  } catch (e) { return { success: false, message: e.message }; }
}

function ubahPassword(email, oldPassword, newPassword) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(KARYAWAN_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const inputOldHash = hashPassword(oldPassword);
    const newHash = hashPassword(newPassword);

    for (let i = 1; i < data.length; i++) {
      if (data[i][2].toString().trim().toLowerCase() === email.trim().toLowerCase()) {
        let dbPassword = data[i][3].toString().trim();
        if (dbPassword === inputOldHash || dbPassword === oldPassword) {
          sheet.getRange(i + 1, 4).setValue(newHash); 
          return { success: true, message: 'Password berhasil diperbarui! Silakan login kembali.' };
        } else {
          return { success: false, message: 'Password lama salah.' };
        }
      }
    }
    return { success: false, message: 'Akun tidak ditemukan.' };
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.message }; }
}

// =================================================================
// 6. FUNGSI UNTUK PENGUMUMAN & KALENDER (DIPERBAIKI APINYA)
// =================================================================
function getHomeData() {
  try {
    const sheetPengumuman = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Pengumuman');
    const sheetLibur = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Libur Nasional');
    let pengumuman = [];
    if (sheetPengumuman) {
      const dataP = sheetPengumuman.getDataRange().getDisplayValues();
      for(let i = 1; i < dataP.length; i++) { if(dataP[i][0]) { pengumuman.push({ judul: dataP[i][0], tanggal: dataP[i][1], isi: dataP[i][2], warna: dataP[i][3] || 'Biru' }); } }
    }
    let libur = [];
    if (sheetLibur) {
      const dataL = sheetLibur.getDataRange().getValues();
      for(let i = 1; i < dataL.length; i++) { if(dataL[i][0]) { libur.push({ tanggal: safeDateString(dataL[i][0]), nama: dataL[i][1], keterangan: dataL[i][2], status: dataL[i][3] || 'Libur' }); } }
    }
    return { success: true, pengumuman: pengumuman, libur: libur };
  } catch(e) { return { success: false, message: e.message }; }
}

function getAdminInfoData() {
    const sheetP = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Pengumuman');
    const sheetL = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Libur Nasional');
    let pList = [], lList = [];
    if(sheetP) { const dP = sheetP.getDataRange().getDisplayValues(); for(let i=1; i<dP.length; i++) { if(dP[i][0]) pList.push({rowNum: i+1, judul: dP[i][0], tanggal: dP[i][1], isi: dP[i][2], warna: dP[i][3]}); } }
    if(sheetL) { const dL = sheetL.getDataRange().getValues(); for(let i=1; i<dL.length; i++) { if(dL[i][0]) lList.push({rowNum: i+1, tanggal: safeDateString(dL[i][0]), nama: dL[i][1], keterangan: dL[i][2], status: dL[i][3]}); } }
    return { pengumuman: pList, libur: lList };
}

function simpanPengumuman(j, t, i, w) { SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Pengumuman').appendRow([j, t, i, w]); return {success:true}; }
function hapusPengumuman(row) { SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Pengumuman').deleteRow(row); return {success:true}; }
function simpanLibur(t, n, k, s) { SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Libur Nasional').appendRow([t, n, k, s]); return {success:true}; }
function hapusLibur(row) { SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Libur Nasional').deleteRow(row); return {success:true}; }

// FUNGSI TARIK DATA API (DIPERBAIKI)
function syncLiburNasional(year) {
  try {
    // Menggunakan API baru yang lebih stabil dan lengkap
    const url = 'https://api-harilibur.vercel.app/api?year=' + year;
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    
    if(response.getResponseCode() !== 200) {
        throw new Error("Penyedia API sedang down/gangguan (404/500)");
    }
    
    const json = JSON.parse(response.getContentText());
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName('Libur Nasional');
    
    // Hapus seluruh data lama yang ada di sheet (Kecuali baris 1 header)
    if(sheet.getLastRow() > 1) { 
        sheet.getRange(2, 1, sheet.getLastRow()-1, 4).clearContent(); 
    }
    
    // Masukkan data baru dari pemerintah
    json.forEach(item => {
      // is_national_holiday adalah boolean yang menandakan hari merah resmi
      if(item.is_national_holiday) {
         let isCuti = item.holiday_name.toLowerCase().includes('cuti bersama');
         let statusLabel = isCuti ? 'Cuti Bersama' : 'Libur Nasional';
         let badgeLabel = isCuti ? 'Cuti' : 'Libur';
         
         // item.holiday_date formatnya "2025-01-01"
         sheet.appendRow([item.holiday_date, item.holiday_name, statusLabel, badgeLabel]);
      }
    });
    
    return {success: true, message: 'Berhasil menyinkronkan jadwal libur untuk tahun ' + year};
  } catch(e) {
    return {success: false, message: 'Gagal Menarik API: ' + e.message + '. Silakan tambah manual.'};
  }
}
