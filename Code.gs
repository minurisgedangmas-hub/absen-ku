const SETTINGS_SHEET = 'Pengaturan';
const CLASSES_SHEET = 'Daftar_Kelas';
const STUDENTS_SHEET = 'Data_Siswa';
const ABSENCE_LOGS_SHEET = 'Absen_Manual';
const ATTENDANCE_SHEET = 'Kehadiran';
const WA_API_KEY = '9ed61aa8151df31bd0f9718b82f067462938f0a5cd3cfd312dac18069d34019b';
// ALTERNATIF ENDPOINT (Coba satu per satu jika gagal):
// 1. 'https://gateway.pdmhadirq.cloud/api/send-message'  <-- Sedang dicoba
// 2. 'https://gateway.pdmhadirq.cloud/send-message'
// 3. 'https://mywa.pdmhadirq.my.id/send-message'
const WA_ENDPOINT = 'https://gateway.pdmhadirq.cloud/api/send-message'; // Menggunakan /api/ sebagai default yang lebih standar
const WA_FALLBACK_ENDPOINT = 'https://gateway.pdmhadirq.cloud/send-message';

/**
 * Robust sheet retrieval that tries exact match, then replaces underscores with spaces and vice versa,
 * and finally tries case-insensitive matching.
 */
function getSheetResilient(name) {
  if (!name) {
    Logger.log('getSheetResilient called with null or undefined name');
    return null;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Try exact match
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  
  // 2. Try replacing underscores with spaces
  const withSpaces = name.replace(/_/g, ' ');
  sheet = ss.getSheetByName(withSpaces);
  if (sheet) return sheet;
  
  // 3. Try replacing spaces with underscores
  const withUnderscores = name.replace(/\s+/g, '_');
  sheet = ss.getSheetByName(withUnderscores);
  if (sheet) return sheet;
  
  // 4. Case-insensitive search
  const allSheets = ss.getSheets();
  const lowerName = name.toLowerCase().replace(/_/g, ' ');
  for (let i = 0; i < allSheets.length; i++) {
    const sName = allSheets[i].getName().toLowerCase().replace(/_/g, ' ');
    if (sName === lowerName) return allSheets[i];
  }
  
  return null;
}

/**
 * Serves the web application.
 */
function doGet() {
  const settings = getAppSettings();
  const template = HtmlService.createTemplateFromFile('index');
  template.initialSettings = settings;
  
  return template.evaluate()
    .setTitle(settings.org_name || 'Presensi Digital Pro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Optimization: Only initialize if critical sheets are missing
  if (!getSheetResilient(SETTINGS_SHEET) || !getSheetResilient(STUDENTS_SHEET)) {
    initializeSystem();
  }
  
  let settings = {
    org_name: 'MI NURUL ISLAM GEDANGMAS',
    logo_url: '',
    background_url: '',
    footer_text: 'Selamat Datang di Sistem Presensi Digital Pro. Harap melakukan scan QR Code dengan tertib.',
    classes: []
  };

  // Load Settings from Sheet
  const settingsSheet = getSheetResilient(SETTINGS_SHEET);
  if (settingsSheet) {
    const data = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const key = (data[i][0] || "").toString().trim();
      let val = data[i][1];
      if (key && (val !== undefined && val !== null)) {
        if (key === 'logo_url' || key === 'background_url' || key === 'banner_url') {
          val = convertDriveLink(val);
        } else if ((key === 'jam_masuk' || key === 'jam_pulang' || key.startsWith('jam_pulang_')) && val instanceof Date) {
          // Format Date object to HH:mm string
          val = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), 'HH:mm');
        }
        settings[key] = val;
      }
    }
  }

  // Load Classes from Sheet
  const classesSheet = getSheetResilient(CLASSES_SHEET);
  if (classesSheet) {
    const data = classesSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
       if (data[i][0]) settings.classes.push(data[i][0].toString());
    }
  }

  return settings;
}

/**
 * Initializes all necessary sheets and dummy data if they don't exist.
 * Can be run manually from Script Editor.
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Initialize Settings (Pengaturan)
  let settingsSheet = getSheetResilient(SETTINGS_SHEET);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET);
    settingsSheet.appendRow(['Meta Key', 'Value', 'Deskripsi']);
  }

  const data = settingsSheet.getDataRange().getValues();
  const existingKeys = data.map(row => row[0]);
  const requiredSettings = [
    ['org_name', 'MI NURUL ISLAM GEDANGMAS', 'Nama Lembaga'],
    ['logo_url', '', 'Link Logo (Drive/URL)'],
    ['background_url', '', 'Link Background (Drive/URL)'],
    ['footer_text', 'Selamat Datang di Sistem Presensi Digital Pro. Harap melakukan scan QR Code dengan tertib.', 'Teks Berjalan'],
    ['jam_masuk', '07:00', 'Batas Awal Jam Masuk'],
    ['jam_pulang', '13:00', 'Batas Awal Jam Pulang'],
    ['hari_libur', 'Minggu', 'Hari Libur (Pisahkan dengan koma jika lebih dari satu)'],
    ['wa_sender', '6285187232455', 'Nomor WA Pengirim (Gateway)'],
    ['wa_api_key', WA_API_KEY, 'API Key dari Dashboard MyWA'],
    ['wa_endpoint', WA_ENDPOINT, 'Endpoint API WhatsApp Gateway'],
    ['wa_template_masuk', '*NOTIFIKASI PRESENSI*', 'Template WA Masuk'],
    ['wa_template_pulang', '*NOTIFIKASI PRESENSI*\n\nAssalamualaikum Ayah/Bunda,\n\nAlhamdulillah, putra/putri Anda:\nNama: *{{nama}}*\nKelas: *{{kelas}}*\n\nTELAH BERHASIL melakukan *Absen Pulang* pada pukul *{{waktu}}*.\n\nTerima kasih.\n*{{lembaga}}*', 'Template WA Pulang'],
    ['wa_template_sakit', '*NOTIFIKASI UPDATE*\n\nAssalamualaikum Ayah/Bunda,\n\nKami menginformasikan bahwa ananda *{{nama}}* hari ini terdata *SAKIT*. Semoga lekas sembuh.\n\nTerima kasih.\n*{{lembaga}}*', 'Template WA Sakit'],
    ['wa_template_izin', '*NOTIFIKASI UPDATE*\n\nAssalamualaikum Ayah/Bunda,\n\nKami menginformasikan bahwa ananda *{{nama}}* hari ini terdata *IZIN*.\n\nTerima kasih.\n*{{lembaga}}*', 'Template WA Izin'],
    ['wa_template_alpha', '*Peringatan Presensi*\n\nAssalamualaikum Ayah/Bunda,\n\nKami informasikan bahwa ananda *{{nama}}* hari ini *TIDAK ADA KETERANGAN* (Alpha) hingga jam masuk berakhir.\n\nMohon konfirmasinya.\n*{{lembaga}}*', 'Template WA Alpha'],
    ['banner_url', 'https://images.unsplash.com/photo-1546410531-bb4caa6b424d?q=80&w=2071&auto=format&fit=crop', 'URL Foto Dashboard Atas'],
    ['banner_caption', 'Mencetak Generasi Cerdas dan Berakhlak Karimah', 'Keterangan Foto Dashboard'],
    ['admin_user', 'admin', 'Username Login Admin (Full Akses)'],
    ['admin_pass', '12345', 'Password Login Admin'],
    ['piket_user', 'piket', 'Username Login Petugas Piket (Hanya Absensi)'],
    ['piket_pass', 'piket123', 'Password Login Petugas Piket'],
    ['jam_pulang_khusus', '', 'Override jam pulang per tanggal (Format: YYYY-MM-DD=HH:mm, pisahkan dengan koma)'],
    ['jam_pulang_senin_kamis', '12:00', 'Jam Pulang Hari Senin-Kamis'],
    ['jam_pulang_jumat', '10:30', 'Jam Pulang Hari Jumat'],
    ['jam_pulang_sabtu', '11:30', 'Jam Pulang Hari Sabtu']
  ];
  
  requiredSettings.forEach(s => {
    if (!existingKeys.includes(s[0])) {
      settingsSheet.appendRow(s);
    }
  });

  // 2. Initialize Classes (Daftar_Kelas)
  let classesSheet = getSheetResilient(CLASSES_SHEET);
  if (!classesSheet || classesSheet.getLastRow() <= 1) {
    classesSheet = classesSheet || ss.insertSheet(CLASSES_SHEET);
    if (classesSheet.getLastRow() <= 1) {
      classesSheet.clear();
      classesSheet.appendRow(['Nama Kelas']);
      ['Kelas 1', 'Kelas 2', 'Kelas 3', 'Kelas 4', 'Kelas 5', 'Kelas 6'].forEach(c => classesSheet.appendRow([c]));
    }
  }

  // 3. Initialize Students (Data_Siswa)
  let studentSheet = getSheetResilient(STUDENTS_SHEET);
  if (!studentSheet || studentSheet.getLastRow() <= 1) {
    studentSheet = studentSheet || ss.insertSheet(STUDENTS_SHEET);
    if (studentSheet.getLastRow() <= 1) {
      studentSheet.clear();
      studentSheet.appendRow(['ID Siswa', 'Nama', 'Kelas', 'No HP Orang Tua']);
      const dummyStudents = [
        ['ID001', 'Ahmad Saputra', 'Kelas 1', '6281234567801'],
        ['ID002', 'Budi Santoso', 'Kelas 1', '6281234567802'],
        ['ID003', 'Citra Lestari', 'Kelas 2', '6281234567803'],
        ['ID004', 'Dedi Wijaya', 'Kelas 2', '6281234567804'],
        ['ID005', 'Eko Prasetyo', 'Kelas 3', '6281234567805'],
        ['ID006', 'Fani Rahmawati', 'Kelas 3', '6281234567806'],
        ['ID007', 'Gita Permata', 'Kelas 4', '6281234567807'],
        ['ID008', 'Hadi Sucipto', 'Kelas 4', '6281234567808'],
        ['ID009', 'Indah Kusuma', 'Kelas 5', '6281234567809'],
        ['ID010', 'Joko Susilo', 'Kelas 5', '6281234567810'],
        ['ID011', 'Kania Dewi', 'Kelas 6', '6281234567811'],
        ['ID012', 'Lucky Pratama', 'Kelas 6', '6281234567812']
      ];
      dummyStudents.forEach(s => studentSheet.appendRow(s));
    }
  }

  // 4. Initialize Absence Logs (Absen_Manual)
  let absenceSheet = getSheetResilient(ABSENCE_LOGS_SHEET);
  if (!absenceSheet) {
    absenceSheet = ss.insertSheet(ABSENCE_LOGS_SHEET);
    absenceSheet.appendRow(['Timestamp', 'ID Siswa', 'Status', 'Alasan', 'Notifikasi (WA Status)']);
  } else {
    // Add missing column for notifications if existing
    const data = absenceSheet.getDataRange().getValues();
    if (data[0].length < 5) {
      absenceSheet.getRange(1, 5).setValue('Notifikasi (WA Status)');
    }
  }

  // 5. Initialize Attendance (Kehadiran - Combined Masuk & Pulang)
  let attSheet = getSheetResilient(ATTENDANCE_SHEET);
  if (!attSheet) {
    attSheet = ss.insertSheet(ATTENDANCE_SHEET);
    attSheet.appendRow(['Tanggal', 'ID SISWA', 'NAMA SISWA', 'KELAS', 'MASUK', 'JAM', 'PULANG', 'JAM']);
  }
}

/**
 * Adds a custom menu to the spreadsheet.
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🚀 Presensi')
        .addItem('Kirim Notif Sakit/Izin (Cek Log)', 'processAbsenceLog')
        .addItem('Kirim Pengingat Alpha (Scan)', 'sendAlphaReminders')
        .addSeparator()
        .addItem('Buat Rekapan Bulanan', 'generateMonthlyReport')
        .addItem('Inisialisasi Sistem', 'initializeSystem')
        .addSeparator()
        .addItem('Hapus Sheet Lama (Pembersihan)', 'cleanupOldSheets')
        .addToUi();
  } catch (e) {
    // Silent catch for web app context where UI is unavailable
    Logger.log('UI Menu skipped: ' + e.toString());
  }
}

/**
 * Cleanup old/unused sheets (English names and Gallery).
 */
function cleanupOldSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const oldSheets = [
    'Settings', 
    'Classes', 
    'Gallery', 
    'Students', 
    'Absence_Logs', 
    'Attendance_In', 
    'Attendance_Out',
    'Galeri',
    'Izin_Sakit'
  ];
  
  const response = ui.alert('Pembersihan Sheet', 'Apakah Anda yakin ingin menghapus sheet lama (English names & Galeri)? Pastikan data penting sudah dipindahkan ke sheet baru (Siswa, Kehadiran, dll).', ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) return;

  let deletedCount = 0;
  oldSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      try {
        ss.deleteSheet(sheet);
        deletedCount++;
      } catch (e) {
        Logger.log(`Gagal menghapus ${name}: ${e.message}`);
      }
    }
  });

  ui.alert(`Berhasil menghapus ${deletedCount} sheet lama.`);
}



/**
 * Converts a standard Google Drive sharing link to a direct image link.
 * Supports various formats including /file/d/ID/view, open?id=ID, etc.
 */
function convertDriveLink(url) {
  if (!url) return url;
  
  // Robust regex to extract Drive File ID (usually 33 characters)
  const driveRegex = /(?:id=|\/d\/|file\/d\/)([a-zA-Z0-9_-]{25,})/;
  const match = url.match(driveRegex);
  
  if (match && match[1]) {
    const fileId = match[1];
    // Return direct thumbnail link which is more reliable for embedding
    return `https://drive.google.com/thumbnail?id=${fileId}&sz=w500`;
  }
  
  return url;
}

/**
 * Processes attendance based on scanned QR data (Student ID).
 */
function processAttendance(studentId, type, className) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10 seconds wait for individual scan
  } catch (e) {
    return { success: false, message: 'Server sibuk, silakan scan ulang.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ATTENDANCE_SHEET) || ss.insertSheet(ATTENDANCE_SHEET);
    const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
    const appSettings = getAppSettings();
    
    // Check Holiday
    const holidayCheck = isHoliday(appSettings.hari_libur);
    if (holidayCheck.isHoliday) {
      lock.releaseLock();
      return { success: false, message: `Hari ini adalah hari libur (${holidayCheck.reason}).` };
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Tanggal', 'ID SISWA', 'NAMA SISWA', 'KELAS', 'MASUK', 'JAM', 'PULANG', 'JAM']);
    }
    
    let studentName = 'Unknown';
    let parentPhone = '';
    let studentClass = 'Unknown';
    
    if (studentSheet) {
      const data = studentSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === studentId.toString()) {
          studentName = data[i][1];
          studentClass = data[i][2];
          parentPhone = data[i][3];
          break;
        }
      }
    }
    
    const timestamp = new Date();
    const timeStr = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), 'HH:mm:ss');
    const shortTime = timeStr.substring(0, 5); // HH:mm
    
    let status = 'Hadir';
    if (type === 'Masuk' && appSettings.jam_masuk) {
      let jamMasuk = appSettings.jam_masuk.toString().substring(0, 5);
      if (isTimeGreater(shortTime, jamMasuk)) status = 'Terlambat';
    } else if (type === 'Pulang' && appSettings.jam_pulang) {
      let jamPulang = appSettings.jam_pulang.toString();
      const dayNameEn = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), 'EEEE');
      
      if (['Monday', 'Tuesday', 'Wednesday', 'Thursday'].includes(dayNameEn)) {
        jamPulang = appSettings.jam_pulang_senin_kamis || jamPulang;
      } else if (dayNameEn === 'Friday') {
        jamPulang = appSettings.jam_pulang_jumat || jamPulang;
      } else if (dayNameEn === 'Saturday') {
        jamPulang = appSettings.jam_pulang_sabtu || jamPulang;
      }

      const dateStr = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      const overrides = (appSettings.jam_pulang_khusus || '').split(',').map(o => o.trim());
      for (let o of overrides) {
        if (o.includes('=')) {
          const [oDate, oTime] = o.split('=');
          if (oDate === dateStr) {
            jamPulang = oTime;
            break;
          }
        }
      }
      if (jamPulang.length > 5) jamPulang = jamPulang.substring(0, 5);
      if (isTimeGreater(jamPulang, shortTime)) status = 'Pulang Awal';
    }

    const today = new Date();
    today.setHours(0,0,0,0);
    const attDataAll = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = attDataAll.length - 1; i >= 1; i--) { // Search from bottom for speed
      const rowDate = new Date(attDataAll[i][0]);
      if (rowDate instanceof Date) {
        rowDate.setHours(0,0,0,0);
        if (rowDate.getTime() === today.getTime() && attDataAll[i][1].toString() === studentId.toString()) {
          rowIndex = i + 1;
          break;
        }
      }
      // Speed up: Stop searching if rows are more than 500 rows deep
      if (attDataAll.length > 200 && (attDataAll.length - i) > 500) break; 
    }

    if (rowIndex > 0) {
      if (type === 'Masuk') {
        sheet.getRange(rowIndex, 5, 1, 2).setValues([[status, shortTime]]);
      } else {
        sheet.getRange(rowIndex, 7, 1, 2).setValues([[status, shortTime]]);
      }
    } else {
      const newRow = [timestamp, studentId, studentName, studentClass, "", "", "", ""];
      if (type === 'Masuk') { newRow[4] = status; newRow[5] = shortTime; }
      else { newRow[6] = status; newRow[7] = shortTime; }
      sheet.appendRow(newRow);
    }
    
    lock.releaseLock(); 

    // Trigger WhatsApp Notification
    if (parentPhone) {
      // Rule: For "Pulang" mode, only send WA if student was present in the morning
      let shouldSendWA = true;
      if (type === 'Pulang') {
        const morningStatus = rowIndex > 0 ? attDataAll[rowIndex-1][4] : "";
        // Jika tidak ada status pagi, tapi sekarang diabsen, tetap kirim jika ini adalah kehadiran manual/bulk
        // Namun jika ini scan QR, biasanya hanya kirim jika sudah ada record pagi.
        // Kita buat lebih fleksibel: kirim jika memang ada status masuk ATAU jika ini record baru (mungkin lupa absen pagi)
        if (!morningStatus && rowIndex > 0) shouldSendWA = false; 
      }

      if (shouldSendWA) {
        const orgName = appSettings.org_name || 'Sekolah';
        const statusLabel = status === 'Hadir' ? 'TEPAT WAKTU' : status.toUpperCase();
        let templateKey = (type && type.toLowerCase() === 'pulang') ? 'wa_template_pulang' : 'wa_template_masuk';
        let template = appSettings[templateKey] || '';
        
        let msg = template
          .replace(/{{nama}}/g, studentName)
          .replace(/{{kelas}}/g, studentClass)
          .replace(/{{waktu}}/g, shortTime)
          .replace(/{{status}}/g, statusLabel)
          .replace(/{{lembaga}}/g, orgName);
        
        if (!msg || !template) {
          msg = `*NOTIFIKASI PRESENSI*\n\nAlhamdulillah, ananda *${studentName}* telah berhasil *Absen ${type}* pada pukul *${shortTime}*.\n\n*${orgName}*`;
        }

        const waRes = sendWhatsAppNotification(parentPhone, msg, appSettings.wa_sender);
        
        return {
          success: true,
          name: studentName,
          type: type || 'Masuk',
          time: timeStr,
          status: status,
          waStatus: waRes ? (waRes.status || waRes.success ? 'Terkirim' : (waRes.msg || waRes.message || 'Gagal')) : 'Gagal'
        };
      }
    }
    
    return { 
      success: true, 
      name: studentName, 
      type: type || 'Masuk', 
      time: timeStr, 
      status: status, 
      waStatus: 'No Phone' 
    };
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    Logger.log('Server Error: ' + e.toString());
    return { success: false, message: 'Kesalahan Server: ' + e.message };
  }
}


/**
 * FUNGSI INTEGRASI APPSHEET:
 * Fungsi ini akan mendeteksi ketika AppSheet menambahkan baris baru di Spreadsheet.
 * Jika baris ditambahkan, fungsi ini akan otomatis mengisi Nama, Kelas, 
 * dan mengirim notifikasi WhatsApp menggunakan data dari AppSheet.
 * 
 * CARA AKTIFKAN: Di Editor Apps Script, klik ikon Jam (Triggers) -> Add Trigger -> 
 * Pilih "handleAppSheetSync", Event Source: "From spreadsheet", Event Type: "On change".
 */
function handleAppSheetSync(e) {
  // Hanya proses jika ada penambahan baris baru (INSERT_ROW)
  if (!e || e.changeType !== 'INSERT_ROW') return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const rowData = sheet.getRange(lastRow, 1, 1, 8).getValues()[0];
  const studentId = rowData[1]; // Ambil ID dari Kolom B
  
  if (!studentId) return;

  // Deteksi mode dari AppSheet (E=Masuk, G=Pulang)
  let mode = "Masuk";
  if (rowData[6]) mode = "Pulang"; 

  // Panggil logika pengisian data & kirim WhatsApp
  processAttendance(studentId.toString(), mode);
}

/**
 * FUNGSI TEST: Jalankan fungsi ini untuk mengetes WhatsApp Gateway.
 * Ganti nomor di bawah dengan nomor HP Anda untuk mencoba.
 */
function testKirimWA() {
  const nomorTest = '6285335115241'; // <-- GANTI DENGAN NOMOR ANDA
  const pesanTest = 'Tes koneksi WhatsApp Gateway dari Sistem Presensi Digital.';
  const settings = getAppSettings();
  
  Logger.log('Memulai test pengiriman WA...');
  const hasil = sendWhatsAppNotification(nomorTest, pesanTest, settings.wa_sender);
  Logger.log('Hasil Test: ' + JSON.stringify(hasil));
  
  if (hasil.status || hasil.success) {
    Logger.log('BERHASIL! Koneksi WA Gateway sudah oke.');
  } else {
    Logger.log('--- TIPS PERBAIKAN ---');
    if (settings.wa_sender === '6285233235194') {
      Logger.log('PENTING: Anda masih menggunakan NOMOR SENDER CONTOH (6285233235194). Silakan ganti dengan NOMOR HP ANDA di Sheet Pengaturan.');
    }
    if (!settings.wa_api_key) {
      Logger.log('PENTING: wa_api_key di Sheet Pengaturan masih kosong.');
    }
    Logger.log('Gagal: ' + (hasil.msg || 'Periksa API Key dan status Connected di Dashboard MyWA.'));
  }
}

/**
 * Sends a message via MyWA Gateway using JSON POST (Standard Dokumentasi).
 * JANGAN JALANKAN FUNGSI INI SECARA LANGSUNG (pilih testKirimWA untuk mengetes).
 */
/**
 * Sends a message via MyWA Gateway using JSON POST (Standard Dokumentasi).
 * JANGAN JALANKAN FUNGSI INI SECARA LANGSUNG (pilih testKirimWA untuk mengetes).
 */
/**
 * Sends a message via WA Gateway with enhanced robustness and logging.
 * Tries multiple common API patterns used by Indonesian WA Gateways (MPWA, WBM, etc.)
 */
function sendWhatsAppNotification(phone, message, sender, optionalSettings) {
  if (!phone || !message) {
    Logger.log('WA Error: Nomor HP atau Pesan Kosong.');
    return { status: false, msg: 'Nomor atau Pesan Kosong' };
  }

  const settings = optionalSettings || getAppSettings();
  const apiKey = (settings.wa_api_key || WA_API_KEY || '').toString().trim();
  const waSender = (sender || settings.wa_sender || '').toString().trim();
  
  if (!waSender) {
    Logger.log('WA Error: Nomor Pengirim (wa_sender) belum diatur.');
    return { status: false, msg: 'Nomor Pengirim Kosong' };
  }

  let cleanPhone = phone.toString().replace(/[^0-9]/g, '');
  if (cleanPhone.startsWith('0')) {
    cleanPhone = '62' + cleanPhone.substring(1);
  } else if (!cleanPhone.startsWith('62')) {
    cleanPhone = '62' + cleanPhone;
  }
  
  const primaryEndpoint = (settings.wa_endpoint || WA_ENDPOINT || '').toString().trim();
  const endpoints = [
    primaryEndpoint,
    primaryEndpoint.replace('/api/send-message', '/send-message'), // Try root if /api/ fails
    WA_FALLBACK_ENDPOINT
  ].filter((v, i, a) => v && a.indexOf(v) === i); // Unique endpoints

  // Common variations for body keys
  const bodyVariations = [
    { 'api_key': apiKey, 'sender': waSender, 'number': cleanPhone, 'message': message },
    { 'api_key': apiKey, 'sender': waSender, 'receiver': cleanPhone, 'message': message }, // some use receiver
    { 'api_key': apiKey, 'sender': waSender, 'number': cleanPhone, 'message': message, 'type': 'text' }
  ];
  
  let lastError = '';
  
  for (let url of endpoints) {
    for (let body of bodyVariations) {
      Logger.log(`WA Attempt: To=${cleanPhone} via ${url} with keys: ${Object.keys(body).join(',')}`);
      
      try {
        // 1. Try JSON
        const optionsJson = {
          'method': 'post',
          'contentType': 'application/json',
          'payload': JSON.stringify(body),
          'muteHttpExceptions': true,
          'connectTimeout': 10000
        };
        const resJson = UrlFetchApp.fetch(url, optionsJson);
        if (checkSuccess(resJson)) {
          Logger.log(`WA Success (JSON): ${cleanPhone}`);
          return { status: true, msg: 'Sent' };
        }
        
        // 2. Try Form-data (some gateways require this)
        const optionsForm = {
          'method': 'post',
          'payload': body,
          'muteHttpExceptions': true,
          'connectTimeout': 10000
        };
        const resForm = UrlFetchApp.fetch(url, optionsForm);
        if (checkSuccess(resForm)) {
          Logger.log(`WA Success (Form): ${cleanPhone}`);
          return { status: true, msg: 'Sent' };
        }
        
        lastError = `HTTP ${resJson.getResponseCode()}: ${resJson.getContentText()}`;
      } catch (e) {
        lastError = e.toString();
        Logger.log(`WA Attempt Error: ${lastError}`);
      }
    }
  }

  Logger.log(`WA Final Failure for ${cleanPhone}: ${lastError}`);
  return { status: false, msg: lastError };
}

/** Helper to check success response */
function checkSuccess(res) {
  if (!res) return false;
  const code = res.getResponseCode();
  const text = res.getContentText();
  
  if (code !== 200) return false;
  
  // Robust check: if response contains success-indicator words, trust it
  const lowerText = text.toLowerCase();
  if (lowerText.includes('success') || lowerText.includes('sent') || lowerText.includes('ok') || lowerText.includes('"status":true')) {
    return true;
  }

  try {
    const p = JSON.parse(text);
    return (p.status === true || p.success === true || p.status === 'success' || 
            (p.message && p.message.toLowerCase().includes('success')));
  } catch (e) {
    return false;
  }
}

/** Function for Bulk WA */
function prepareWARequest(url, apiKey, sender, phone, message) {
  let cleanPhone = phone.toString().replace(/[^0-9]/g, '');
  if (cleanPhone.startsWith('0')) { 
    cleanPhone = '62' + cleanPhone.substring(1); 
  } else if (!cleanPhone.startsWith('62')) { 
    cleanPhone = '62' + cleanPhone; 
  }

  return {
    url: url,
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      api_key: apiKey,
      sender: sender.toString().trim(),
      number: cleanPhone,
      message: message,
      type: 'text'
    }),
    muteHttpExceptions: true
  };
}


/**
 * Fungsi Diagnostik: Jalankan ini untuk mengetahui apa yang salah
 */
function waDiagnostic() {
  const settings = getAppSettings();
  Logger.log('--- DIAGNOSTIK WHATSAPP ---');
  Logger.log('Nomor Pengirim: ' + settings.wa_sender);
  Logger.log('API Key: ' + settings.wa_api_key.substring(0,5) + '...');
  Logger.log('Endpoint: ' + settings.wa_endpoint);
  
  const testNum = '6285335115241';
  const res = sendWhatsAppNotification(testNum, 'Test Diagnostik Sistem Presensi', settings.wa_sender);
  Logger.log('Hasil: ' + JSON.stringify(res));
  
  if (res.msg && res.msg.includes('Check your connection')) {
    Logger.log('KESIMPULAN: Server MyWA tidak bisa kontak HP Anda. Mohon Disconnect dan Scan QR ulang di MyWA Dashboard.');
  }
}

/**
 * Fetches dashboard statistics.
 */
function getDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = getSheetResilient(ATTENDANCE_SHEET);
  const studentSheet = getSheetResilient(STUDENTS_SHEET);
  const manualSheet = getSheetResilient(ABSENCE_LOGS_SHEET);
  
  let stats = {
    totalSiswa: 0,
    masuk: 0,
    izin: 0,
    sakit: 0,
    alpha: 0,
    belumAbsen: 0,
    pulang: 0
  };
  
  if (studentSheet) {
    stats.totalSiswa = Math.max(0, studentSheet.getLastRow() - 1);
  }
  
  const today = new Date();
  const todayStr = today.toDateString();
  
  if (attSheet && attSheet.getLastRow() > 1) {
    const data = attSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];
      if (rowDate instanceof Date && rowDate.toDateString() === todayStr) {
        if (data[i][4]) stats.masuk++; 
        if (data[i][6]) stats.pulang++; 
      }
    }
  }
  
  if (manualSheet && manualSheet.getLastRow() > 1) {
    const data = manualSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];
      if (rowDate instanceof Date && rowDate.toDateString() === todayStr) {
        const status = (data[i][2] || "").toString().toUpperCase();
        if (status === 'IZIN') stats.izin++;
        if (status === 'SAKIT') stats.sakit++;
        if (status === 'ALPHA') stats.alpha++;
      }
    }
  }
  
  stats.belumAbsen = stats.totalSiswa - (stats.masuk + stats.izin + stats.sakit + stats.alpha);
  if (stats.belumAbsen < 0) stats.belumAbsen = 0;
  
  return stats;
}

/**
 * Checks if a date is a holiday based on settings.
 */
function isHoliday(holidaySetting) {
  if (!holidaySetting) return { isHoliday: false };
  
  const today = new Date();
  const dayName = Utilities.formatDate(today, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'EEEE'); // e.g., 'Monday'
  const dateStr = Utilities.formatDate(today, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  
  // Indonesian Day Names mapping
  const dayMapping = {
    'Monday': 'Senin',
    'Tuesday': 'Selasa',
    'Wednesday': 'Rabu',
    'Thursday': 'Kamis',
    'Friday': 'Jumat',
    'Saturday': 'Sabtu',
    'Sunday': 'Minggu'
  };
  const todayIndo = dayMapping[dayName] || dayName;
  
  const holidays = holidaySetting.split(',').map(h => h.trim());
  
  if (holidays.includes(todayIndo)) {
    return { isHoliday: true, reason: todayIndo };
  }
  
  if (holidays.includes(dateStr)) {
    return { isHoliday: true, reason: dateStr };
  }
  
  return { isHoliday: false };
}

/**
 * Helper to compare two HH:mm strings.
 * Returns true if t1 > t2.
 */
function isTimeGreater(t1, t2) {
  const [h1, m1] = t1.split(':').map(Number);
  const [h2, m2] = t2.split(':').map(Number);
  return (h1 * 60 + m1) > (h2 * 60 + m2);
}

/**
 * GENERATE MONTHLY RECAP
 * Generates a summary sheet for attendance.
 */
function generateMonthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. Get Month/Year input
  const response = ui.prompt('Buat Rekapan Bulanan', 'Masukkan Bulan dan Tahun (Format: MM-YYYY, contoh: 02-2026):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const input = response.getResponseText().trim();
  const [month, year] = input.split('-').map(Number);
  
  if (!month || !year || month < 1 || month > 12) {
    ui.alert('Format salah! Gunakan MM-YYYY.');
    return;
  }

  const monthName = getMonthNameIndo(month);
  const targetSheetName = `Rekap_${monthName}_${year}`;
  let reportSheet = ss.getSheetByName(targetSheetName);
  
  if (reportSheet) {
    const confirm = ui.alert('Sheet sudah ada. Timpa?', ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(targetSheetName);
  }

  // Header
  reportSheet.appendRow(['No', 'ID Siswa', 'Nama', 'Kelas', 'Hadir', 'Terlambat', 'Sakit', 'Izin', 'Alpha', 'Total Hari Efektif']);
  
  // Styling Header
  reportSheet.getRange(1, 1, 1, 10).setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');

  // 2. Load Data
  const students = ss.getSheetByName(STUDENTS_SHEET).getDataRange().getValues();
  const attData = ss.getSheetByName(ATTENDANCE_SHEET) ? ss.getSheetByName(ATTENDANCE_SHEET).getDataRange().getValues() : [];
  const absenceLogs = ss.getSheetByName(ABSENCE_LOGS_SHEET) ? ss.getSheetByName(ABSENCE_LOGS_SHEET).getDataRange().getValues() : [];

  // Filter attendance & absence by month/year (Student is present if Masuk status exists)
  const filteredAtt = attData.filter(row => {
    if (!(row[0] instanceof Date)) return false;
    return (row[0].getMonth() + 1) === month && row[0].getFullYear() === year && row[4] !== "";
  });
  
  const filteredAbs = absenceLogs.filter(row => {
    if (!(row[0] instanceof Date)) return false;
    return (row[0].getMonth() + 1) === month && row[0].getFullYear() === year;
  });

  // Calculate Effective Days (Number of unique days in Attendance_In for this month)
  const uniqueDays = [...new Set(filteredAtt.map(row => row[0].toDateString()))];
  const totalDays = uniqueDays.length;

  // 3. Process each student
  let rowIdx = 1;
  for (let i = 1; i < students.length; i++) {
    const sId = students[i][0].toString();
    const sName = students[i][1];
    const sClass = students[i][2];
    
    // Counts
    let hadirCount = 0;
    let telatCount = 0;
    let sakitCount = 0;
    let izinCount = 0;
    
    // Check Attendance
    filteredAtt.forEach(att => {
      if (att[1].toString() === sId) {
        hadirCount++;
        if (att[5] === 'Terlambat') telatCount++;
      }
    });
    
    // Check Absence Logs
    filteredAbs.forEach(abs => {
      if (abs[1].toString() === sId) {
        const status = abs[2].toString().toUpperCase();
        if (status === 'SAKIT') sakitCount++;
        if (status === 'IZIN') izinCount++;
      }
    });
    
    // Alpha = Total Days - (Hadir + Sakit + Izin)
    let alphaCount = totalDays - (hadirCount + sakitCount + izinCount);
    if (alphaCount < 0) alphaCount = 0;

    reportSheet.appendRow([rowIdx++, sId, sName, sClass, hadirCount, telatCount, sakitCount, izinCount, alphaCount, totalDays]);
  }

  ui.alert('Berhasil! Rekapan bulanan telah dibuat di sheet ' + targetSheetName);
}

/**
 * Helper to get Indo Month Name
 */
function getMonthNameIndo(m) {
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return months[m - 1];
}

/**
 * PROCESS ABSENCE LOGS
 * Sends WA for manual entries (Sakit/Izin) in Absence_Logs.
 */
function processAbsenceLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABSENCE_LOGS_SHEET);
  const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
  const appSettings = getAppSettings();
  
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const students = studentSheet.getDataRange().getValues();
  const orgName = appSettings.org_name || 'Sekolah';
  
  let sentCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const sId = data[i][1].toString();
    const status = data[i][2].toString().toUpperCase();
    const notified = data[i][4];
    
    if (notified) continue; // Skip if already notified
    
    // Find Student Details
    let studentName = 'Unknown';
    let parentPhone = '';
    for (let j = 1; j < students.length; j++) {
      if (students[j][0].toString() === sId) {
        studentName = students[j][1];
        parentPhone = students[j][3];
        break;
      }
    }
    
    if (parentPhone) {
      // Perubahan: Sakit dan Izin tidak mengirim WA otomatis
      if (status === 'SAKIT' || status === 'IZIN') {
        sheet.getRange(i + 1, 5).setValue('Non-Aktif');
        continue;
      }

      const templateKey = status === 'ALPHA' ? 'wa_template_alpha' : ''; // Hanya Alpha (jika ada template)
      if (!templateKey) {
        sheet.getRange(i + 1, 5).setValue('Skip');
        continue;
      }

      const template = appSettings[templateKey] || '';
      
      let msg = template
        .replace(/{{nama}}/g, studentName)
        .replace(/{{lembaga}}/g, orgName);
        
      if (!template) {
        msg = `*Peringatan Presensi*\n\nAnanda *${studentName}* hari ini *TIDAK ADA KETERANGAN* (Alpha).\n\n*${orgName}*`;
      }
      
      const res = sendWhatsAppNotification(parentPhone, msg, appSettings.wa_sender);
      if (res.status || res.success) {
        sheet.getRange(i + 1, 5).setValue('Sent - ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'HH:mm'));
        sentCount++;
      } else {
        sheet.getRange(i + 1, 5).setValue('Failed');
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(`Berhasil mengirim ${sentCount} notifikasi.`);
}

/**
 * SEND ALPHA REMINDERS
 * Finds students who haven't checked in today and sends an Alpha notification.
 */
function sendAlphaReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
  const attSheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const absenceSheet = ss.getSheetByName(ABSENCE_LOGS_SHEET);
  const appSettings = getAppSettings();
  
  if (!studentSheet || !attSheet) return;
  
  const students = studentSheet.getDataRange().getValues();
  const attData = attSheet.getDataRange().getValues();
  const absenceData = absenceSheet ? absenceSheet.getDataRange().getValues() : [];
  
  const todayStr = new Date().toDateString();
  const orgName = appSettings.org_name || 'Sekolah';
  const apiKey = (appSettings.wa_api_key || WA_API_KEY || '').toString().trim();
  const waUrl = (appSettings.wa_endpoint || WA_ENDPOINT || '').toString().trim();
  const waSender = (appSettings.wa_sender || '').toString().trim();

  const checkedInToday = attData.filter(row => {
    return row[0] instanceof Date && row[0].toDateString() === todayStr && row[4] !== "";
  }).map(row => row[1].toString());
  
  const absenceToday = absenceData.filter(row => {
    return row[0] instanceof Date && row[0].toDateString() === todayStr;
  }).map(row => row[1].toString());

  const waRequests = [];
  
  for (let i = 1; i < students.length; i++) {
    const sId = students[i][0].toString();
    const sName = students[i][1];
    const parentPhone = students[i][3];
    
    if (!checkedInToday.includes(sId) && !absenceToday.includes(sId) && parentPhone && waUrl && apiKey && waSender) {
      const template = appSettings.wa_template_alpha || '';
      let msg = template.replace(/{{nama}}/g, sName).replace(/{{lembaga}}/g, orgName);
      if (!template) msg = `*Peringatan Presensi*\n\nAnanda *${sName}* belum melakukan absen hari ini.\n\n*${orgName}*`;
      
      waRequests.push(prepareWARequest(waUrl, apiKey, waSender, parentPhone, msg));
    }
  }
  
  if (waRequests.length > 0) {
    try {
      UrlFetchApp.fetchAll(waRequests);
    } catch (e) {
      Logger.log('Alpha Reminders Error: ' + e.toString());
    }
  }
  
  SpreadsheetApp.getUi().alert(`Berhasil mengirim ${waRequests.length} pengingat Alpha.`);
}


/**
 * Fetches all student data dynamically.
 */
function getStudentsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = getSheetResilient(STUDENTS_SHEET);
  
  if (!studentSheet) {
    Logger.log('Student sheet not found!');
    return { headers: [], data: [] };
  }

  const data = studentSheet.getDataRange().getValues();
  if (data.length <= 1) return { headers: [], data: [] };

  const headers = data[0];
  const students = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Skip empty rows
    
    const studentObj = {
      _raw: row
    };
    headers.forEach((header, index) => {
      const key = header.toString().toLowerCase().trim().replace(/\s+/g, '_');
      studentObj[key] = row[index];
    });
    // Explicitly map common fields if they exist
    if (studentObj.id_siswa) studentObj.id = studentObj.id_siswa;
    if (studentObj.nama_siswa || studentObj.nama) studentObj.name = studentObj.nama_siswa || studentObj.nama;
    
    students.push(studentObj);
  }

  return { headers: headers, data: students };
}

/**
 * Fetches attendance recap for a specific month and year.
 */
function getAttendanceRecap(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
  const attSheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const absenceSheet = ss.getSheetByName(ABSENCE_LOGS_SHEET);
  
  if (!studentSheet) return { recap: [], totalEffective: 0, monthName: getMonthNameIndo(month) };

  const students = studentSheet.getDataRange().getValues();
  const attData = attSheet ? attSheet.getDataRange().getValues() : [];
  const absenceLogs = absenceSheet ? absenceSheet.getDataRange().getValues() : [];
  
  // Filter attendance & absence by month/year (Presence is confirmed if Masuk status column is not empty)
  const filteredAtt = attData.filter(row => {
    if (!(row[0] instanceof Date)) return false;
    return (row[0].getMonth() + 1) === Number(month) && row[0].getFullYear() === Number(year) && row[4] !== "";
  });
  
  const filteredAbs = absenceLogs.filter(row => {
    if (!(row[0] instanceof Date)) return false;
    return (row[0].getMonth() + 1) === Number(month) && row[0].getFullYear() === Number(year);
  });

  const uniqueDays = [...new Set(filteredAtt.map(row => row[0].toDateString()))];
  const totalDays = uniqueDays.length;

  const recap = [];
  for (let i = 1; i < students.length; i++) {
    const sId = students[i][0].toString();
    const sName = students[i][1];
    const sClass = students[i][2];
    
    let hadirCount = 0;
    let telatCount = 0;
    let sakitCount = 0;
    let izinCount = 0;
    
    filteredAtt.forEach(att => {
      if (att[1].toString() === sId) {
        hadirCount++;
        if (att[4] === 'Terlambat') telatCount++;
      }
    });
    
    filteredAbs.forEach(abs => {
      if (abs[1].toString() === sId) {
        const status = abs[2].toString().toUpperCase();
        if (status === 'SAKIT') sakitCount++;
        if (status === 'IZIN') izinCount++;
      }
    });
    
    let alphaCount = totalDays - (hadirCount + sakitCount + izinCount);
    if (alphaCount < 0) alphaCount = 0;

    recap.push({
      id: sId,
      name: sName,
      class: sClass,
      hadir: hadirCount,
      telat: telatCount,
      sakit: sakitCount,
      izin: izinCount,
      alpha: alphaCount,
      totalEffective: totalDays
    });
  }
  
  return {
    recap: recap,
    totalEffective: totalDays,
    monthName: getMonthNameIndo(month)
  };
}

/**
 * Fetches attendance report for a specific date range.
 */
function getAttendanceReportByRange(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
  const attSheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const absenceSheet = ss.getSheetByName(ABSENCE_LOGS_SHEET);
  
  if (!studentSheet) return { recap: [], totalEffective: 0 };

  const students = studentSheet.getDataRange().getValues();
  const attData = attSheet ? attSheet.getDataRange().getValues() : [];
  const absenceLogs = absenceSheet ? absenceSheet.getDataRange().getValues() : [];
  
  const startDate = new Date(startDateStr);
  startDate.setHours(0,0,0,0);
  const endDate = new Date(endDateStr);
  endDate.setHours(23,59,59,999);

  // Filter attendance & absence by range
  const filteredAtt = attData.filter(row => {
    if (!(row[0] instanceof Date)) return false;
    const d = new Date(row[0]);
    return d >= startDate && d <= endDate && row[4] !== "";
  });
  
  const filteredAbs = absenceLogs.filter(row => {
    if (!(row[0] instanceof Date)) return false;
    const d = new Date(row[0]);
    return d >= startDate && d <= endDate;
  });

  const uniqueDays = [...new Set(filteredAtt.map(row => row[0].toDateString()))];
  const totalDays = uniqueDays.length;

  const recap = [];
  for (let i = 1; i < students.length; i++) {
    const sId = students[i][0].toString();
    const sName = students[i][1];
    const sClass = students[i][2];
    
    let hadirCount = 0;
    let telatCount = 0;
    let sakitCount = 0;
    let izinCount = 0;
    
    filteredAtt.forEach(att => {
      if (att[1].toString() === sId) {
        hadirCount++;
        if (att[4] === 'Terlambat') telatCount++;
      }
    });
    
    filteredAbs.forEach(abs => {
      if (abs[1].toString() === sId) {
        const status = abs[2].toString().toUpperCase();
        if (status === 'SAKIT') sakitCount++;
        if (status === 'IZIN') izinCount++;
      }
    });
    
    let alphaCount = totalDays - (hadirCount + sakitCount + izinCount);
    if (alphaCount < 0) alphaCount = 0;

    recap.push({
      id: sId,
      name: sName,
      class: sClass,
      hadir: hadirCount,
      telat: telatCount,
      sakit: sakitCount,
      izin: izinCount,
      alpha: alphaCount,
      totalEffective: totalDays
    });
  }
  
  return {
    recap: recap,
    totalEffective: totalDays
  };
}

/**
 * GENERATE DAILY MATRIX RECAP
 * Creates a formatted matrix-style overview (No, Name, 1...31, S, I, A)
 */
function generateDailyMatrixReport(month, year, className) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const monthNum = Number(month);
    const yearNum = Number(year);
    const monthName = getMonthNameIndo(monthNum);
    let targetSheetName = `Matrix_${monthName}_${yearNum}`;
    if (className) {
       targetSheetName = `Matrix_${className.replace(/\s+/g, '_')}_${monthName}_${yearNum}`;
    }
    
    let sheet = ss.getSheetByName(targetSheetName);
    if (sheet) {
      sheet.clear();
      sheet.clearFormats();
    } else {
      sheet = ss.insertSheet(targetSheetName);
    }
    
    let students = getStudentsData().data;
    if (className) {
      students = students.filter(s => s.class === className);
    }
    
    if (students.length === 0) {
      return { success: false, message: `Tidak ada data siswa untuk ${className || 'semua kelas'}.` };
    }
    
    // Sort students by name
    students.sort((a, b) => a.name.localeCompare(b.name));

    const lastDay = new Date(yearNum, monthNum, 0).getDate();
    
    // Header Row 1: Grouped Headers
    sheet.getRange(1, 1, 1, 2).merge().setValue('Data Siswa').setBackground('#1e293b').setFontColor('white');
    sheet.getRange(1, 3, 1, lastDay).merge().setValue(`Bulan ${monthName}`).setBackground('#334155').setFontColor('white');
    sheet.getRange(1, 3 + lastDay, 1, 3).merge().setValue('Jumlah').setBackground('#1e293b').setFontColor('white');
    
    // Header Row 2: Detailed Headers
    const headers = ['No', 'Nama Siswa'];
    for (let d = 1; d <= lastDay; d++) headers.push(d);
    headers.push('S', 'I', 'A');
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]).setBackground('#f1f5f9').setFontWeight('bold');
    
    // Load Attendance & Absence
    const attSheet = ss.getSheetByName(ATTENDANCE_SHEET);
    const absenceLogsSheet = ss.getSheetByName(ABSENCE_LOGS_SHEET);
    
    const attData = attSheet ? attSheet.getDataRange().getValues() : [];
    const absenceLogs = absenceLogsSheet ? absenceLogsSheet.getDataRange().getValues() : [];
    
    const attendanceMap = {};
    
    // Helper to process logs
    const processLogs = (data, isAbsence) => {
      if (data.length <= 1) return;
      
      data.forEach(row => {
        const rowDate = row[0];
        if (!(rowDate instanceof Date)) return;
        
        // Performance: Filter by Month and Year here
        if (rowDate.getMonth() + 1 !== monthNum || rowDate.getFullYear() !== yearNum) return;
        
        const sId = row[1].toString();
        const day = rowDate.getDate();
        if (!attendanceMap[sId]) attendanceMap[sId] = {};
        
        if (isAbsence) {
          const status = row[2] ? row[2].toString().toUpperCase().charAt(0) : 'A';
          attendanceMap[sId][day] = status;
        } else {
          // Mark 'H' (Hadir) if Masuk status exists (Col E)
          if (row[4] !== "") {
            attendanceMap[sId][day] = 'H';
          }
        }
      });
    };
    
    processLogs(attData, false);
    processLogs(absenceLogs, true);
    
    // Populate Rows
    const rows = [];
    students.forEach((s, idx) => {
      const row = [idx + 1, s.name];
      let sCount = 0, iCount = 0, aCount = 0;
      
      for (let d = 1; d <= lastDay; d++) {
        const status = (attendanceMap[s.id] && attendanceMap[s.id][d]) || 'A';
        row.push(status);
        if (status === 'S') sCount++;
        else if (status === 'I') iCount++;
        else if (status === 'A') aCount++;
      }
      row.push(sCount, iCount, aCount);
      rows.push(row);
    });
    
    if (rows.length > 0) {
      sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
    }
    
    // Global Formatting
    const fullRange = sheet.getRange(1, 1, rows.length + 2, headers.length);
    fullRange.setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    fullRange.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Arial');
    
    // specific formatting
    sheet.getRange(3, 2, rows.length, 1).setHorizontalAlignment('left'); // Names
    sheet.setFrozenColumns(2);
    sheet.setFrozenRows(2);
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    sheet.setColumnWidth(2, 250); // Set fixed width for name column
    
    return { success: true, sheetName: targetSheetName };
  } catch (e) {
    Logger.log('Matrix Report Error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Returns the direct download URL for a specific sheet as XLSX.
 */
function getDownloadUrl(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return null;
    
    const spreadsheetId = ss.getId();
    const sheetId = sheet.getSheetId();
    
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx&gid=${sheetId}`;
  } catch (e) {
    Logger.log('Export Error: ' + e.toString());
    return null;
  }
}
/**
 * SUBMIT MANUAL ATTENDANCE
 * Records manual attendance entry (A, S, I) and sends WA notification.
 */
function submitManualAttendance(studentId, statusCode, type) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10 seconds wait
  } catch (e) {
    return { success: false, message: 'Server sibuk, silakan coba lagi.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
    const absenceSheet = ss.getSheetByName(ABSENCE_LOGS_SHEET) || ss.insertSheet(ABSENCE_LOGS_SHEET);
    const appSettings = getAppSettings();
    
    if (!studentId || !statusCode) {
      lock.releaseLock();
      return { success: false, message: 'ID Siswa and Status are required.' };
    }

    // Map status code
    const statusMap = {
      'H': 'HADIR',
      'S': 'SAKIT',
      'I': 'IZIN',
      'A': 'ALPHA'
    };
    const status = statusMap[statusCode.toUpperCase()];
    if (!status) {
      lock.releaseLock();
      return { success: false, message: 'Invalid status code.' };
    }

    if (status === 'HADIR') {
      lock.releaseLock();
      return processAttendance(studentId, type || 'Masuk');
    }

    // Find Student Details
    let studentName = '';
    let parentPhone = '';
    const students = studentSheet.getDataRange().getValues();
    for (let i = students.length - 1; i >= 1; i--) { // Search from bottom
      if (students[i][0].toString() === studentId.toString()) {
        studentName = students[i][1];
        parentPhone = students[i][3];
        break;
      }
    }

    if (!studentName) {
      lock.releaseLock();
      return { success: false, message: 'ID Siswa tidak ditemukan.' };
    }

    // Check for existing records today (Avoid double entry)
    const today = new Date();
    today.setHours(0,0,0,0);
    const existingLogs = absenceSheet.getDataRange().getValues();
    for (let i = existingLogs.length - 1; i >= 1; i--) {
        const rowDate = new Date(existingLogs[i][0]);
        if (rowDate instanceof Date) {
            rowDate.setHours(0,0,0,0);
            if (rowDate.getTime() === today.getTime() && existingLogs[i][1].toString() === studentId.toString()) {
                lock.releaseLock();
                return { success: true, alreadyRecorded: true, name: studentName, status: existingLogs[i][2], waStatus: 'Sudah Tercatat' };
            }
        }
        if (existingLogs.length > 200 && (existingLogs.length - i) > 500) break;
    }
    
    // Also check Attendance In (Should not mark S/I/A if already Masuk)
    const attSheet = getSheetResilient(ATTENDANCE_SHEET);
    if (attSheet) {
        const attData = attSheet.getDataRange().getValues();
        for (let i = attData.length - 1; i >= 1; i--) {
            const rowDate = new Date(attData[i][0]);
            if (rowDate instanceof Date && rowDate.setHours(0,0,0,0) === today.getTime() && attData[i][1].toString() === studentId.toString() && attData[i][4]) {
                lock.releaseLock();
                return { success: false, message: `${studentName} sudah tercatat HADIR hari ini.` };
            }
            if (attData.length > 200 && (attData.length - i) > 500) break;
        }
    }

    const timestamp = new Date();
    const orgName = appSettings.org_name || 'Sekolah';
    
    // Append to Absence_Logs
    const newRow = [timestamp, studentId, status, 'Input Manual', 'Pending'];
    absenceSheet.appendRow(newRow);
    const lastRow = absenceSheet.getLastRow();
    
    lock.releaseLock();

    // Trigger WhatsApp Notification
    let waResult = 'Not Sent';
    
    // Send WA for Alpha ONLY
    if (parentPhone && status === 'ALPHA') {
      const templateKey = 'wa_template_alpha';
      const template = appSettings[templateKey] || '';
      let msg = template
        .replace(/{{nama}}/g, studentName)
        .replace(/{{lembaga}}/g, orgName);
        
      if (!template) {
        msg = `*NOTIFIKASI UPDATE*\n\nAssalamualaikum Ayah/Bunda,\n\nKami menginformasikan bahwa ananda *${studentName}* hari ini terdata *${status}*.\n\nTerima kasih.\n*${orgName}*`;
      }
      
      const res = sendWhatsAppNotification(parentPhone, msg, appSettings.wa_sender);
      if (res.status || res.success) {
        waResult = 'Sent - ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'HH:mm');
        absenceSheet.getRange(lastRow, 5).setValue(waResult);
      } else {
        waResult = 'Failed: ' + (res.msg || 'Unknown error');
        absenceSheet.getRange(lastRow, 5).setValue(waResult);
      }
    }

    return {
      success: true,
      name: studentName,
      status: status,
      waStatus: waResult
    };

  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    Logger.log('Manual Attendance Error: ' + e.toString());
    return { success: false, message: 'Server Error: ' + e.message };
  }
}

/**
 * SUBMIT BULK MANUAL ATTENDANCE
 * Records multiple entries at once to speed up manual processing.
 */
/**
 * SUBMIT BULK MANUAL ATTENDANCE
 * Records multiple entries at once to speed up manual processing.
 */
/**
 * SUBMIT BULK MANUAL ATTENDANCE
 * Records multiple entries at once to speed up manual processing.
 * Improved to batch sheet updates and WA notifications.
 */
function submitBulkAttendance(studentStatusMap, type) {
  // studentStatusMap format: { "studentId": "H/S/I/A" }
  const studentIds = Object.keys(studentStatusMap);
  if (studentIds.length === 0) {
    return { success: false, message: 'No students selected.' };
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (e) {
    return { success: false, message: 'Sistem sibuk (Lock Timeout).' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settings = getAppSettings();
    const timeZone = ss.getSpreadsheetTimeZone();
    const timestamp = new Date();
    const timeStr = Utilities.formatDate(timestamp, timeZone, 'HH:mm');
    const orgName = settings.org_name || 'Sekolah';
    const apiKey = (settings.wa_api_key || WA_API_KEY || '').toString().trim();
    const waSender = (settings.wa_sender || '').toString().trim();
    const waUrl = (settings.wa_endpoint || WA_ENDPOINT || '').toString().trim();
    
    const statusMap = { 'H': 'HADIR', 'S': 'SAKIT', 'I': 'IZIN', 'A': 'ALPHA' };
    
    // 1. Get Student Data (Batch)
    const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
    const studentsRaw = studentSheet.getDataRange().getValues();
    const studentInfoMap = {};
    for (let i = 1; i < studentsRaw.length; i++) {
      studentInfoMap[studentsRaw[i][0].toString()] = {
        name: studentsRaw[i][1],
        class: studentsRaw[i][2],
        phone: studentsRaw[i][3]
      };
    }

    const waRequests = [];
    let processedCount = 0;

    const attSheet = ss.getSheetByName(ATTENDANCE_SHEET) || ss.insertSheet(ATTENDANCE_SHEET);
    const absenceSheet = ss.getSheetByName(ABSENCE_LOGS_SHEET) || ss.insertSheet(ABSENCE_LOGS_SHEET);
    
    if (attSheet.getLastRow() === 0) {
      attSheet.appendRow(['Tanggal', 'ID SISWA', 'NAMA SISWA', 'KELAS', 'MASUK', 'JAM', 'PULANG', 'JAM']);
    }

    const today = new Date();
    today.setHours(0,0,0,0);
    const todayMillis = today.getTime();

    // Load ALL data once for today's check
    const attDataAll = attSheet.getDataRange().getValues();
    const absDataAll = absenceSheet.getDataRange().getValues();

    const existingAttMap = {}; // sId -> rowIndex (1-based)
    for (let i = 1; i < attDataAll.length; i++) {
      const d = new Date(attDataAll[i][0]);
      if (d instanceof Date && d.setHours(0,0,0,0) === todayMillis) {
        existingAttMap[attDataAll[i][1].toString()] = i + 1;
      }
    }

    const existingAbsMap = {}; // sId -> true
    for (let i = 1; i < absDataAll.length; i++) {
       const d = new Date(absDataAll[i][0]);
       if (d instanceof Date && d.setHours(0,0,0,0) === todayMillis) {
         existingAbsMap[absDataAll[i][1].toString()] = true;
       }
    }

    const appendsHadir = [];
    const appendsManual = [];
    
    // Track row updates to do them in batch or mini-batches
    const rowUpdatesAtt = {}; // rowIndex -> [status, time]

    studentIds.forEach(id => {
      const statusCode = studentStatusMap[id].toUpperCase();
      const statusLabel = statusMap[statusCode];
      const s = studentInfoMap[id];
      if (!s || !statusLabel) return;

      // Skip if already has THIS status today
      if (statusLabel === 'HADIR') {
        const rowIndex = existingAttMap[id.toString()];
        let finalStatus = 'Hadir';
        if (type === 'Masuk' && settings.jam_masuk) {
          if (isTimeGreater(timeStr, settings.jam_masuk.toString().substring(0, 5))) finalStatus = 'Terlambat';
        } else if (type === 'Pulang' && settings.jam_pulang) {
          if (isTimeGreater(settings.jam_pulang.toString().substring(0, 5), timeStr)) finalStatus = 'Pulang Awal';
        }

        if (rowIndex) {
          rowUpdatesAtt[rowIndex] = rowUpdatesAtt[rowIndex] || [null, null, null, null];
          if (type === 'Masuk') {
            rowUpdatesAtt[rowIndex][0] = finalStatus;
            rowUpdatesAtt[rowIndex][1] = timeStr;
          } else {
            rowUpdatesAtt[rowIndex][2] = finalStatus;
            rowUpdatesAtt[rowIndex][3] = timeStr;
          }
          processedCount++;
        } else {
          const newRow = [timestamp, id, s.name, s.class, "", "", "", ""];
          if (type === 'Masuk') { newRow[4] = finalStatus; newRow[5] = timeStr; }
          else { newRow[6] = finalStatus; newRow[7] = timeStr; }
          appendsHadir.push(newRow);
          processedCount++;
        }

        // WA for Hadir
        if (s.phone && waUrl && apiKey && waSender) {
          // Rule: For "Pulang" mode, only send WA if student was present in the morning (Hadir/Terlambat)
          let shouldSendWA = true;
          if (type === 'Pulang') {
            const morningStatus = rowIndex ? attDataAll[rowIndex-1][4] : "";
            // Perbaikan: Jika rowIndex tidak ada (record baru hari ini), 
            // artinya siswa tersebut baru pertama kali absen hari ini di jam pulang. Kita tetap izinkan kirim WA.
            // Kita hanya blok jika rowIndex ADA tapi MASUK-nya kosong (berarti dia memang hadirnya pulang saja tapi data baris sudah ada tanpa status masuk)
            if (rowIndex && !morningStatus) shouldSendWA = false;
          }

          if (shouldSendWA) {
            const tKey = (type === 'Pulang') ? 'wa_template_pulang' : 'wa_template_masuk';
            const t = settings[tKey] || '';
            let msg = t.replace(/{{nama}}/g, s.name).replace(/{{kelas}}/g, s.class).replace(/{{waktu}}/g, timeStr)
                       .replace(/{{status}}/g, finalStatus.toUpperCase()).replace(/{{lembaga}}/g, orgName);
            if (!t) msg = `*NOTIFIKASI PRESENSI*\n\nAnanda *${s.name}* telah absen *${type}* (${finalStatus}) pkl *${timeStr}*.\n\n*${orgName}*`;
            waRequests.push(prepareWARequest(waUrl, apiKey, waSender, s.phone, msg));
          }
        }

      } else {
        // SAKIT / IZIN / ALPHA
        if (existingAbsMap[id] || existingAttMap[id]) return; // Skip if any record exists

        let waLogStatus = 'No Notif';
        
        // WA for Manual (Hanya kirim untuk ALPHA sesuai permintaan user)
        if (s.phone && waUrl && apiKey && waSender && statusLabel === 'ALPHA') {
          const t = settings.wa_template_alpha || '';
          let msg = t.replace(/{{nama}}/g, s.name).replace(/{{kelas}}/g, s.class).replace(/{{waktu}}/g, timeStr)
                     .replace(/{{status}}/g, statusLabel).replace(/{{lembaga}}/g, orgName);
          if (!t) msg = `*NOTIFIKASI PRESENSI*\n\nInformasi bahwa ananda *${s.name}* hari ini terdata *${statusLabel}*.\n\n*${orgName}*`;
          waRequests.push(prepareWARequest(waUrl, apiKey, waSender, s.phone, msg));
          waLogStatus = 'Sent - ' + timeStr;
        }

        appendsManual.push([timestamp, id, statusLabel, 'Input Bulk', waLogStatus]);
        processedCount++;
      }
    });

    // 2. Execute Batch Updates to Sheet
    // HADIR: Updates for existing rows
    const updateIndices = Object.keys(rowUpdatesAtt);
    if (updateIndices.length > 0) {
      // Optimization: Group consecutive rows? Usually they aren't.
      // We'll just update each. For very high volume, we could write the whole "today range".
      updateIndices.forEach(idx => {
        const vals = rowUpdatesAtt[idx];
        if (vals[0]) attSheet.getRange(idx, 5, 1, 2).setValues([[vals[0], vals[1]]]);
        if (vals[2]) attSheet.getRange(idx, 7, 1, 2).setValues([[vals[2], vals[3]]]);
      });
    }
    // HADIR: New rows
    if (appendsHadir.length > 0) {
      attSheet.getRange(attSheet.getLastRow() + 1, 1, appendsHadir.length, 8).setValues(appendsHadir);
    }
    // MANUAL: New rows
    if (appendsManual.length > 0) {
      absenceSheet.getRange(absenceSheet.getLastRow() + 1, 1, appendsManual.length, 5).setValues(appendsManual);
    }

    // 3. Send WA parallel
    if (waRequests.length > 0) {
      try { UrlFetchApp.fetchAll(waRequests); } catch (e) { Logger.log('Bulk WA Error: ' + e.toString()); }
    }

    lock.releaseLock();
    return { success: true, count: processedCount, total: studentIds.length };

  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    Logger.log('Bulk Error: ' + e.toString());
    return { success: false, message: 'Kesalahan Server: ' + e.message };
  }
}


/**
 * Gets attendance status for all students for today.
 * Optimized to scan from the bottom for high performance.
 */
function getTodayAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = getSheetResilient(ATTENDANCE_SHEET);
  const absenceSheet = getSheetResilient(ABSENCE_LOGS_SHEET);
  
  const todayMillis = new Date().setHours(0,0,0,0);
  const attendance = {};
  
  if (attSheet) {
    const data = attSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const d = data[i][0];
      if (d instanceof Date && d.setHours(0,0,0,0) === todayMillis) {
        const sId = data[i][1].toString();
        attendance[sId] = attendance[sId] || { masuk: false, pulang: false, status: '' };
        if (data[i][4]) attendance[sId].masuk = true;
        if (data[i][6]) attendance[sId].pulang = true;
      } else if (d instanceof Date && d.getTime() < todayMillis) {
        if (data.length - i > 300) break; // Limit scan once we're past today
      }
    }
  }
  
  if (absenceSheet) {
    const data = absenceSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
        const d = data[i][0];
        if (d instanceof Date && d.setHours(0,0,0,0) === todayMillis) {
            const sId = data[i][1].toString();
            attendance[sId] = attendance[sId] || { masuk: false, pulang: false, status: '' };
            attendance[sId].masuk = true; 
            attendance[sId].status = data[i][2]; // SAKIT/IZIN/ALPHA
        } else if (d instanceof Date && d.getTime() < todayMillis) {
            if (data.length - i > 300) break;
        }
    }
  }
    
  return attendance;
}



