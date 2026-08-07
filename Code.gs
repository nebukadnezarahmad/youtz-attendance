// ══════════════════════════════════════════════════════════════
// Youtz Group Corporation — SMART ATTENDANCE: Google Apps Script Backend
// Deploy sebagai Web App (Execute as: Me, Access: Anyone)
// ══════════════════════════════════════════════════════════════

// ── KONFIGURASI ──
// Ganti SPREADSHEET_ID dengan ID spreadsheet Anda
// ID ada di URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
const SPREADSHEET_ID = '1UNZmzQtS8u1Lw1kzA1s275S-lVcHX0odBXLWXK6jU5Y';
const DRIVE_FOLDER_NAME = 'Youtz-Attendance-Photos';
// Sheet penampung draft harian (target & laporan sebelum absen pulang)
const DRAFT_SHEET_NAME = 'Draft Harian';

// ── MAIN HANDLER ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = findOrCreateSheet(ss, data.staffName);
    
    if (data.type === 'masuk') {
      return handleMasuk(sheet, data);
    } else if (data.type === 'pulang') {
      return handlePulang(sheet, data);
    } else if (data.type === 'draft') {
      return handleSaveDraft(ss, data);
    }
    
    return jsonResponse({ success: false, message: 'Tipe tidak dikenali' });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'history' && e.parameter.staff) {
    return handleHistory(e.parameter.staff);
  }
  if (e && e.parameter && e.parameter.action === 'draft' && e.parameter.staff) {
    return handleGetDraft(e.parameter.staff);
  }
  return jsonResponse({ success: true, message: 'Youtz Attendance API aktif.' });
}

// ── GET HISTORY ──
function handleHistory(staffName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(staffName);
    if (!sheet) {
      return jsonResponse({ success: true, data: [] });
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return jsonResponse({ success: true, data: [] });
    
    // Ambil kolom: Tanggal (A:1) s.d Foto Pulang (K:11)
    const range = sheet.getRange(2, 1, lastRow - 1, 11);
    const rows = range.getDisplayValues();
    
    const history = [];
    const bulanIndo = {
      'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4, 'Mei': 5, 'Juni': 6,
      'Juli': 7, 'Agustus': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12,
      'January': 1, 'February': 2, 'March': 3, 'May': 5, 'June': 6, 'July': 7,
      'August': 8, 'October': 10, 'December': 12
    };
    
    for (let i = 0; i < rows.length; i++) {
      const tanggalTeks = rows[i][0]; // "07 August 2026"
      if (!tanggalTeks) continue;
      
      const parts = tanggalTeks.split(' ');
      if (parts.length === 3) {
        const yyyy = parts[2];
        const mm = bulanIndo[parts[1]] || 1;
        const dd = parseInt(parts[0], 10);
        const key = `${yyyy}-${mm}-${dd}`;
        
        history.push({
          key: key,
          in: rows[i][2], // Jam Masuk
          out: rows[i][3], // Jam Pulang
          status: rows[i][4], // Status
          target: rows[i][7], // Target Kerja
          log: rows[i][8], // Laporan Pekerjaan
          fotoIn: drivePreviewUrl(rows[i][9]),  // J: Foto Masuk
          fotoOut: drivePreviewUrl(rows[i][10]) // K: Foto Pulang
        });
      }
    }
    
    return jsonResponse({ success: true, data: history });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

// ── DRAFT HARIAN: SIMPAN ──
// Draft disimpan di sheet terpisah agar tidak mengganggu baris absensi.
// Satu baris per staf per tanggal (ditimpa setiap kali draft disimpan).
function handleSaveDraft(ss, data) {
  try {
    const sheet = findOrCreateDraftSheet(ss);
    const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
    const row = findDraftRow(sheet, data.staffName, today);
    const values = [
      data.staffName,                 // A: Staff
      today,                          // B: Tanggal
      formatTargetText(data.targets), // C: Target Kerja
      data.laporan || '',             // D: Laporan Pekerjaan
      data.draftTgtTime || '',        // E: Jam Draft Target
      data.draftLogTime || '',        // F: Jam Draft Laporan
      data.lanjut ? 'YA' : '',        // G: Sudah Lanjut Absen Pulang
      Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy HH:mm:ss') // H
    ];

    if (row > 0) {
      sheet.getRange(row, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
    return jsonResponse({ success: true, message: 'Draft tersimpan.' });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

// ── DRAFT HARIAN: AMBIL ──
function handleGetDraft(staffName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DRAFT_SHEET_NAME);
    if (!sheet) return jsonResponse({ success: true, data: null });

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return jsonResponse({ success: true, data: null });

    const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
    const rows = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();

    for (let i = rows.length - 1; i >= 0; i--) {
      if (rows[i][0] === staffName && rows[i][1] === today) {
        // Hari kerja sudah ditutup dengan absen pulang — tidak ada draft aktif
        if (rows[i][6] === 'SELESAI') return jsonResponse({ success: true, data: null });
        return jsonResponse({
          success: true,
          data: {
            target: rows[i][2],
            laporan: rows[i][3],
            draftTgtTime: rows[i][4],
            draftLogTime: rows[i][5],
            lanjut: rows[i][6] === 'YA'
          }
        });
      }
    }
    return jsonResponse({ success: true, data: null });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

// ── HELPER: Tandai draft hari ini sudah selesai (setelah absen pulang) ──
function markDraftSelesai(ss, staffName) {
  try {
    const sheet = ss.getSheetByName(DRAFT_SHEET_NAME);
    if (!sheet) return;
    const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
    const row = findDraftRow(sheet, staffName, today);
    if (row > 0) sheet.getRange(row, 7).setValue('SELESAI');
  } catch (err) {
    Logger.log('markDraftSelesai: ' + err);
  }
}

// ── HELPER: Cari baris draft milik staf pada tanggal tertentu ──
function findDraftRow(sheet, staffName, tanggal) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === staffName && rows[i][1] === tanggal) return i + 2;
  }
  return -1;
}

// ── HELPER: Buat sheet draft bila belum ada ──
function findOrCreateDraftSheet(ss) {
  let sheet = ss.getSheetByName(DRAFT_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(DRAFT_SHEET_NAME);
  const headers = ['Staff', 'Tanggal', 'Target Kerja', 'Laporan Pekerjaan', 'Jam Draft Target', 'Jam Draft Laporan', 'Lanjut Absen Pulang', 'Terakhir Diperbarui'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#CB2E72');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 120);  // Staff
  sheet.setColumnWidth(2, 140);  // Tanggal
  sheet.setColumnWidth(3, 280);  // Target
  sheet.setColumnWidth(4, 320);  // Laporan
  sheet.setColumnWidth(5, 130);  // Jam draft target
  sheet.setColumnWidth(6, 130);  // Jam draft laporan
  sheet.setColumnWidth(7, 150);  // Lanjut
  sheet.setColumnWidth(8, 180);  // Terakhir diperbarui
  sheet.setFrozenRows(1);
  return sheet;
}

// ── HELPER: Susun teks target ("2/3 selesai" + daftar bercentang) ──
function formatTargetText(targets) {
  if (!targets || !targets.length) return '';
  const done = targets.filter(function (t) { return t.done; }).length;
  let text = done + '/' + targets.length + ' selesai';
  targets.forEach(function (t, i) {
    text += '\n' + (i + 1) + '. ' + (t.done ? '[v] ' : '[ ] ') + t.label;
  });
  return text;
}

// ── ABSEN MASUK ──
function handleMasuk(sheet, data) {
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
  const hari = getHariIndonesia(new Date());
  
  // Cek apakah sudah absen hari ini
  const existingRow = findTodayRow(sheet);
  if (existingRow > 0) {
    return jsonResponse({ success: false, message: 'Sudah absen masuk hari ini.' });
  }
  
  // Upload foto masuk ke Drive
  let fotoMasukLink = '';
  if (data.fotoMasuk) {
    fotoMasukLink = uploadPhoto(data.fotoMasuk, data.staffName + '_masuk_' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss'));
  }
  
  // Tentukan status
  const jamMasuk = data.jamMasuk || '';
  const jam = parseInt(jamMasuk.split(':')[0] || '0');
  const menit = parseInt(jamMasuk.split(':')[1] || '0');
  let status = 'On Time';
  if (jam > 10 || (jam === 10 && menit > 15)) {
    status = 'Terlambat';
  } else if (jam === 10 && menit >= 1 && menit <= 15) {
    status = 'Terlambat (toleransi)';
  }
  
  // Tambah row baru
  sheet.appendRow([
    today,           // A: Tanggal
    hari,            // B: Hari
    jamMasuk,        // C: Jam Masuk
    '',              // D: Jam Pulang (diisi nanti)
    status,          // E: Status
    '',              // F: Total Jam Kerja (dihitung nanti)
    data.lokasi || '',  // G: Lokasi
    '',              // H: Target Kerja (diisi nanti)
    '',              // I: Laporan Pekerjaan (diisi nanti)
    fotoMasukLink,   // J: Foto Masuk
    ''               // K: Foto Pulang (diisi nanti)
  ]);
  
  return jsonResponse({ success: true, message: 'Absen masuk berhasil dicatat.' });
}

// ── ABSEN PULANG ──
function handlePulang(sheet, data) {
  const row = findTodayRow(sheet);
  if (row <= 0) {
    return jsonResponse({ success: false, message: 'Belum absen masuk hari ini.' });
  }
  
  // Upload foto pulang ke Drive
  let fotoPulangLink = '';
  if (data.fotoPulang) {
    fotoPulangLink = uploadPhoto(data.fotoPulang, data.staffName + '_pulang_' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss'));
  }
  
  // Hitung total jam kerja
  const jamMasuk = sheet.getRange(row, 3).getValue();
  const jamPulang = data.jamPulang || '';
  const totalJam = hitungDurasi(jamMasuk, jamPulang);
  
  // Format target
  const targetText = formatTargetText(data.targets);

  // Update row
  sheet.getRange(row, 4).setValue(jamPulang);          // D: Jam Pulang
  sheet.getRange(row, 6).setValue(totalJam);            // F: Total Jam Kerja
  sheet.getRange(row, 8).setValue(targetText);          // H: Target Kerja
  sheet.getRange(row, 9).setValue(data.laporan || '');  // I: Laporan Pekerjaan
  sheet.getRange(row, 11).setValue(fotoPulangLink);     // K: Foto Pulang

  // Tandai draft hari ini sudah selesai agar tidak dipulihkan lagi
  markDraftSelesai(sheet.getParent(), data.staffName);

  return jsonResponse({ success: true, message: 'Absen pulang berhasil dicatat.' });
}

// ── HELPER: Cari atau buat sheet per karyawan ──
function findOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Buat header
    const headers = ['Tanggal', 'Hari', 'Jam Masuk', 'Jam Pulang', 'Status', 'Total Jam Kerja', 'Lokasi', 'Target Kerja', 'Laporan Pekerjaan', 'Foto Masuk', 'Foto Pulang'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Styling header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#007ED4');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    headerRange.setHorizontalAlignment('center');
    
    // Set lebar kolom
    sheet.setColumnWidth(1, 140);  // Tanggal
    sheet.setColumnWidth(2, 80);   // Hari
    sheet.setColumnWidth(3, 100);  // Jam Masuk
    sheet.setColumnWidth(4, 100);  // Jam Pulang
    sheet.setColumnWidth(5, 90);   // Status
    sheet.setColumnWidth(6, 140);  // Total Jam
    sheet.setColumnWidth(7, 120);  // Lokasi
    sheet.setColumnWidth(8, 250);  // Target
    sheet.setColumnWidth(9, 300);  // Laporan
    sheet.setColumnWidth(10, 200); // Foto Masuk
    sheet.setColumnWidth(11, 200); // Foto Pulang
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── HELPER: Cari row hari ini ──
function findTodayRow(sheet) {
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === today) return i + 2; // +2 karena offset header dan 0-index
  }
  return -1;
}

// ── HELPER: Ubah link Drive menjadi URL thumbnail yang bisa dipakai <img> ──
function drivePreviewUrl(url) {
  if (!url) return '';
  const m = url.toString().match(/[-\w]{25,}/);
  return m ? 'https://drive.google.com/thumbnail?id=' + m[0] + '&sz=w400' : '';
}

// ── HELPER: Upload foto ke Google Drive ──
function uploadPhoto(base64Data, fileName) {
  try {
    // Cari atau buat folder
    let folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
    }
    
    // Decode base64
    const parts = base64Data.split(',');
    const mimeMatch = parts[0].match(/data:(.*?);/);
    const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    const decoded = Utilities.base64Decode(parts[1] || parts[0]);
    const blob = Utilities.newBlob(decoded, mime, fileName + '.jpg');
    
    // Simpan ke Drive
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (err) {
    Logger.log('Upload error: ' + err);
    return 'Upload gagal: ' + err.toString();
  }
}

// ── HELPER: Hitung durasi kerja ──
function hitungDurasi(masuk, pulang) {
  if (!masuk || !pulang) return '-';
  const pm = masuk.toString().split(':').map(Number);
  const pp = pulang.toString().split(':').map(Number);
  let diffSec = ((pp[0] || 0) * 3600 + (pp[1] || 0) * 60 + (pp[2] || 0)) - ((pm[0] || 0) * 3600 + (pm[1] || 0) * 60 + (pm[2] || 0));
  if (diffSec < 0) diffSec += 86400;
  const h = Math.floor(diffSec / 3600);
  const m = Math.floor((diffSec % 3600) / 60);
  return h + ' Jam ' + m + ' Menit';
}

// ── HELPER: Nama hari dalam Bahasa Indonesia ──
function getHariIndonesia(date) {
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  return days[date.getDay()];
}

// ── HELPER: JSON Response ──
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
