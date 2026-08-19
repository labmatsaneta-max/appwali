/**
 * STREAMING_CHUNK:Initializing Google Apps Script Database Setup Routine...
 * Menyiapkan Spreadsheet & Sheet Otomatis
 * Jalankan fungsi setupDatabase() pertama kali dari Editor Apps Script
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = [
    {
      name: 'SETTINGS',
      headers: ['key', 'value']
    },
    {
      name: 'USERS',
      headers: ['user_id', 'username', 'password', 'role', 'nama_lengkap', 'nisn_siswa', 'status_aktif']
    },
    {
      name: 'STUDENTS',
      headers: ['nisn', 'nik', 'nama_lengkap', 'jenis_kelamin', 'tempat_lahir', 'tanggal_lahir', 'kelas', 'nama_ayah', 'nama_ibu', 'no_hp_ortu', 'alamat']
    },
    {
      name: 'SCHEDULES',
      headers: ['jadwal_id', 'hari', 'jam_ke', 'waktu', 'mata_pelajaran', 'guru_pengampu', 'keterangan']
    },
    {
      name: 'ATTENDANCE',
      headers: ['absensi_id', 'tanggal', 'nisn', 'nama_siswa', 'kelas', 'status', 'keterangan']
    },
    {
      name: 'ANNOUNCEMENTS',
      headers: ['pengumuman_id', 'tanggal', 'judul', 'isi_pengumuman', 'kategori', 'penulis']
    },
    {
      name: 'MESSAGES',
      headers: ['pesan_id', 'nisn', 'pengirim_role', 'pengirim_nama', 'isi_pesan', 'waktu_kirim']
    },
    {
      name: 'DOCUMENTS',
      headers: ['berkas_id', 'nisn', 'nama_siswa', 'jenis_dokumen', 'nama_file', 'file_data', 'ukuran_kb', 'tanggal_upload', 'folder_url']
    }
  ];

  sheets.forEach(sh => {
    let sheet = ss.getSheetByName(sh.name);
    if (!sheet) {
      sheet = ss.insertSheet(sh.name);
      sheet.appendRow(sh.headers);
      sheet.getRange(1, 1, 1, sh.headers.length).setFontWeight('bold').setBackground('#f1f5f9');
    }
  });

  // Insert default settings if SETTINGS is empty
  const settingsSheet = ss.getSheetByName('SETTINGS');
  if (settingsSheet.getLastRow() === 1) {
    settingsSheet.appendRow(['namaSekolah', 'MIN 3 MADIUN']);
    settingsSheet.appendRow(['namaKelas', 'Kelas 1A']);
    settingsSheet.appendRow(['tahunAjaran', '2026/2027']);
    settingsSheet.appendRow(['semester', 'Semester 1 (Ganjil)']);
    settingsSheet.appendRow(['logoUrl', '']);
  }

  // Insert default admin if USERS is empty
  const userSheet = ss.getSheetByName('USERS');
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow(['USR-001', 'admin', 'admin@123', 'admin', 'Guru Wali Kelas 1A', '', 'AKTIF']);
  }

  return "Setup Database Selesai dan Berhasil!";
}

/**
 * STREAMING_CHUNK:Processing Login with Auto NISN Verification...
 */
function processLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('USERS');
  const studentSheet = ss.getSheetByName('STUDENTS');
  
  const userData = userSheet.getDataRange().getValues();
  
  // 1. Check USERS sheet for Admin accounts
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][1] === username && userData[i][2] === password && userData[i][6] === 'AKTIF') {
      return {
        success: true,
        user: {
          id: userData[i][0],
          username: userData[i][1],
          role: userData[i][3],
          name: userData[i][4],
          nisn: userData[i][5]
        }
      };
    }
  }

  // 2. Automatically check NISN from STUDENTS sheet for Wali Murid Login
  if (studentSheet) {
    const studentData = studentSheet.getDataRange().getValues();
    for (let j = 1; j < studentData.length; j++) {
      const nisn = studentData[j][0].toString();
      if (nisn === username && (password === username || password === 'member@123')) {
        return {
          success: true,
          user: {
            id: 'USR-' + nisn,
            username: nisn,
            role: 'member',
            name: 'Wali Murid - ' + studentData[j][2],
            nisn: nisn
          }
        };
      }
    }
  }

  return { success: false, message: 'Login gagal! NISN atau password tidak cocok.' };
}

/**
 * STREAMING_CHUNK:Managing File Storage with Folder Separation per Student...
 */
function saveDocumentData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DOCUMENTS');
  
  // Find or Create Parent Folder on Google Drive
  let parentFolder;
  const folders = DriveApp.getFoldersByName("Arsip Berkas MIN 3 Madiun");
  if (folders.hasNext()) {
    parentFolder = folders.next();
  } else {
    parentFolder = DriveApp.createFolder("Arsip Berkas MIN 3 Madiun");
  }

  // Find or Create Subfolder per Student ([NISN] - Nama)
  const studentFolderName = "[" + data.nisn + "] - " + data.namaSiswa;
  let studentFolder;
  const childFolders = parentFolder.getFoldersByName(studentFolderName);
  if (childFolders.hasNext()) {
    studentFolder = childFolders.next();
  } else {
    studentFolder = parentFolder.createFolder(studentFolderName);
  }

  // Append Record to Sheet
  sheet.appendRow([
    data.id, 
    data.nisn, 
    data.namaSiswa, 
    data.jenis, 
    data.namaFile, 
    data.fileData, 
    data.sizeKb, 
    data.date,
    studentFolder.getUrl()
  ]);

  return { 
    success: true, 
    message: 'Berkas berhasil disimpan di folder ' + studentFolderName,
    folderUrl: studentFolder.getUrl()
  };
}

/**
 * STREAMING_CHUNK:Handling API HTTP POST Requests...
 */
function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    let result = { success: false, message: 'Action not found' };

    if (action === 'LOGIN') {
      result = processLogin(contents.username, contents.password);
    } else if (action === 'GET_ALL_DATA') {
      result = getAllMasterData();
    } else if (action === 'SAVE_SETTINGS') {
      result = saveSettingsData(contents.data);
    } else if (action === 'SAVE_STUDENT') {
      result = saveStudentData(contents.data);
    } else if (action === 'SAVE_SCHEDULE') {
      result = saveScheduleData(contents.data);
    } else if (action === 'SAVE_ATTENDANCE') {
      result = saveAttendanceData(contents.data);
    } else if (action === 'SEND_MESSAGE') {
      result = saveMessageData(contents.data);
    } else if (action === 'UPLOAD_DOCUMENT') {
      result = saveDocumentData(contents.data);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getAllMasterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    success: true,
    schedules: getSheetDataAsObjects(ss.getSheetByName('SCHEDULES')),
    students: getSheetDataAsObjects(ss.getSheetByName('STUDENTS')),
    attendance: getSheetDataAsObjects(ss.getSheetByName('ATTENDANCE')),
    announcements: getSheetDataAsObjects(ss.getSheetByName('ANNOUNCEMENTS')),
    messages: getSheetDataAsObjects(ss.getSheetByName('MESSAGES')),
    documents: getSheetDataAsObjects(ss.getSheetByName('DOCUMENTS'))
  };
}

function getSheetDataAsObjects(sheet) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, idx) => {
      obj[h] = row[idx];
    });
    return obj;
  });
}

function saveSettingsData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SETTINGS');
  sheet.clearContents();
  sheet.appendRow(['key', 'value']);
  Object.keys(data).forEach(k => {
    sheet.appendRow([k, data[k]]);
  });
  return { success: true, message: 'Pengaturan sekolah berhasil disimpan' };
}

function saveStudentData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STUDENTS');
  sheet.appendRow([
    data.nisn, data.nik, data.nama, data.jk, data.tempatLahir, 
    data.tglLahir, data.kelas, data.ayah, data.ibu, data.hp, data.alamat
  ]);
  return { success: true, message: 'Data siswa berhasil disimpan' };
}

function saveScheduleData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SCHEDULES');
  sheet.appendRow([
    data.id, data.hari, data.jam, data.waktu, data.mapel, data.guru, data.ket
  ]);
  return { success: true, message: 'Jadwal berhasil disimpan' };
}

function saveAttendanceData(dataList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ATTENDANCE');
  dataList.forEach(item => {
    sheet.appendRow([
      Date.now(), item.date, item.nisn, item.name, '1A', item.status, item.ket
    ]);
  });
  return { success: true, message: 'Presensi berhasil dicatat' };
}

function saveMessageData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MESSAGES');
  sheet.appendRow([
    data.id, data.nisn, data.pengirimRole, data.pengirimNama, data.isi, data.waktu
  ]);
  return { success: true, message: 'Pesan terkirim' };
}
