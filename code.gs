/**
 * SISTEM PRESENSI SISWA TERINTEGRASI ORANG TUA
 * Google Apps Script Backend (code.gs)
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Presensi Sekolah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Mengambil spreadsheet aktif tempat skrip dipasang
function getDb() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    throw new Error("Pastikan skrip ini dijalankan di dalam Google Sheets (Ekstensi > Apps Script).");
  }
}

/**
 * Membuat tabel-tabel database jika belum ada dan mengisinya dengan data sampel
 */
function setupDatabase() {
  const ss = getDb();
  
  // 1. Membuat Sheet Kelas
  let sheetKelas = ss.getSheetByName('Kelas');
  if (!sheetKelas) {
    sheetKelas = ss.insertSheet('Kelas');
    sheetKelas.appendRow(['Nama Kelas']);
    sheetKelas.appendRow(['10-A']);
    sheetKelas.appendRow(['10-B']);
  }

  // 2. Membuat Sheet Siswa
  let sheetSiswa = ss.getSheetByName('Siswa');
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet('Siswa');
    sheetSiswa.appendRow(['ID Siswa', 'Nama Siswa', 'Kelas', 'Email Orang Tua']);
    // Data Sampel
    sheetSiswa.appendRow(['S101', 'Budi Santoso', '10-A', 'ortu.budi@email.com']);
    sheetSiswa.appendRow(['S102', 'Ani Wijaya', '10-A', 'ortu.ani@email.com']);
    sheetSiswa.appendRow(['S103', 'Candra Kirana', '10-B', 'ortu.candra@email.com']);
    sheetSiswa.appendRow(['S104', 'Dedi Kurniawan', '10-B', 'ortu.dedi@email.com']);
  }
  
  // 3. Membuat Sheet Akun
  let sheetAkun = ss.getSheetByName('Akun');
  if (!sheetAkun) {
    sheetAkun = ss.insertSheet('Akun');
    sheetAkun.appendRow(['Email', 'Nama', 'Peran', 'ID Siswa Terkait']);
    // Data Sampel
    sheetAkun.appendRow(['guru@sekolah.sch.id', 'Ibu Herlina (Wali Kelas)', 'Guru', '']);
    sheetAkun.appendRow(['ortu.budi@email.com', 'Pak Joko (Orang Tua Budi)', 'Orang Tua', 'S101']);
    sheetAkun.appendRow(['ortu.ani@email.com', 'Ibu Sari (Orang Tua Ani)', 'Orang Tua', 'S102']);
    sheetAkun.appendRow(['ortu.candra@email.com', 'Pak Rudi (Orang Tua Candra)', 'Orang Tua', 'S103']);
  }
  
  // 4. Membuat Sheet Kehadiran
  let sheetKehadiran = ss.getSheetByName('Kehadiran');
  if (!sheetKehadiran) {
    sheetKehadiran = ss.insertSheet('Kehadiran');
    sheetKehadiran.appendRow([
      'ID Presensi', 
      'Tanggal', 
      'ID Siswa', 
      'Nama Siswa', 
      'Kelas', 
      'Status', 
      'Keterangan', 
      'Waktu Input', 
      'Input Oleh'
    ]);
    
    // Data Sampel kehadiran masa lalu untuk simulasi dashboard orang tua
    const hariIni = new Date();
    const formatTanggal = (d) => d.toISOString().split('T')[0];
    
    let tgl1 = new Date(); tgl1.setDate(hariIni.getDate() - 2);
    let tgl2 = new Date(); tgl2.setDate(hariIni.getDate() - 1);
    
    sheetKehadiran.appendRow(['P001', formatTanggal(tgl1), 'S101', 'Budi Santoso', '10-A', 'Hadir', 'Tepat waktu', new Date().toISOString(), 'guru@sekolah.sch.id']);
    sheetKehadiran.appendRow(['P002', formatTanggal(tgl1), 'S102', 'Ani Wijaya', '10-A', 'Sakit', 'Demam', new Date().toISOString(), 'guru@sekolah.sch.id']);
    sheetKehadiran.appendRow(['P003', formatTanggal(tgl2), 'S101', 'Budi Santoso', '10-A', 'Hadir', 'Tepat waktu', new Date().toISOString(), 'guru@sekolah.sch.id']);
    sheetKehadiran.appendRow(['P004', formatTanggal(tgl2), 'S102', 'Ani Wijaya', '10-A', 'Izin', 'Acara keluarga', new Date().toISOString(), 'guru@sekolah.sch.id']);
  }
  
  return "Database berhasil diinisialisasi dengan data demo!";
}

/**
 * Login sederhana berdasarkan email yang terdaftar di Sheet 'Akun'
 */
function loginUser(email) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Akun');
  if (!sheet) return { success: false, message: "Database belum siap. Silakan klik tombol inisialisasi." };
  
  const data = sheet.getDataRange().getValues();
  email = email.trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email) {
      return {
        success: true,
        user: {
          email: data[i][0],
          nama: data[i][1],
          peran: data[i][2], // 'Guru' atau 'Orang Tua'
          idSiswa: data[i][3] // ID Siswa jika Orang Tua
        }
      };
    }
  }
  return { success: false, message: "Email tidak terdaftar di sistem." };
}

/**
 * Mengambil daftar kelas aktif dari Google Sheets
 */
function getClasses() {
  const ss = getDb();
  const sheet = ss.getSheetByName('Kelas');
  if (!sheet) return ['10-A', '10-B']; // fallback
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) list.push(data[i][0].toString());
  }
  return list;
}

/**
 * Menambahkan kelas baru ke database
 */
function tambahKelas(namaKelas) {
  try {
    const ss = getDb();
    let sheet = ss.getSheetByName('Kelas');
    if (!sheet) {
      sheet = ss.insertSheet('Kelas');
      sheet.appendRow(['Nama Kelas']);
    }
    
    namaKelas = namaKelas.trim().toUpperCase();
    if (!namaKelas) return { success: false, message: "Nama kelas tidak boleh kosong." };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toUpperCase() === namaKelas) {
        return { success: false, message: "Kelas '" + namaKelas + "' sudah ada di database!" };
      }
    }
    
    sheet.appendRow([namaKelas]);
    return { success: true, message: "Berhasil menambahkan kelas " + namaKelas };
  } catch (error) {
    return { success: false, message: "Gagal menyimpan kelas: " + error.message };
  }
}

/**
 * Mengambil daftar siswa berdasarkan kelas terpilih (Untuk Guru)
 */
function getSiswaByKelas(kelas) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Siswa');
  const data = sheet.getDataRange().getValues();
  const list = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === kelas) {
      list.push({
        idSiswa: data[i][0],
        namaSiswa: data[i][1],
        kelas: data[i][2]
      });
    }
  }
  return list;
}

/**
 * Mengambil seluruh daftar siswa master beserta relasi data orang tua
 */
function getAllSiswa() {
  const ss = getDb();
  const sheetSiswa = ss.getSheetByName('Siswa');
  if (!sheetSiswa) return [];
  const dataSiswa = sheetSiswa.getDataRange().getValues();
  
  const sheetAkun = ss.getSheetByName('Akun');
  const dataAkun = sheetAkun ? sheetAkun.getDataRange().getValues() : [];
  
  // Mapping untuk mengambil nama orang tua berdasarkan email
  const parentMap = {};
  for (let i = 1; i < dataAkun.length; i++) {
    if (dataAkun[i][2] === 'Orang Tua') {
      parentMap[dataAkun[i][0].toString().toLowerCase()] = dataAkun[i][1];
    }
  }
  
  const list = [];
  for (let i = 1; i < dataSiswa.length; i++) {
    const emailLower = dataSiswa[i][3].toString().toLowerCase();
    list.push({
      idSiswa: dataSiswa[i][0],
      namaSiswa: dataSiswa[i][1],
      kelas: dataSiswa[i][2],
      emailOrtu: dataSiswa[i][3],
      namaOrtu: parentMap[emailLower] || '-'
    });
  }
  return list;
}

/**
 * Menghapus data siswa dan akun akses orang tua terkait
 */
function hapusSiswa(idSiswa, emailOrtu) {
  try {
    const ss = getDb();
    
    // 1. Hapus dari sheet Siswa
    const sheetSiswa = ss.getSheetByName('Siswa');
    if (sheetSiswa) {
      const dataSiswa = sheetSiswa.getDataRange().getValues();
      for (let i = 1; i < dataSiswa.length; i++) {
        if (dataSiswa[i][0].toString().toUpperCase() === idSiswa.toUpperCase()) {
          sheetSiswa.deleteRow(i + 1);
          break;
        }
      }
    }
    
    // 2. Hapus dari sheet Akun
    const sheetAkun = ss.getSheetByName('Akun');
    if (sheetAkun && emailOrtu) {
      const dataAkun = sheetAkun.getDataRange().getValues();
      const emailLower = emailOrtu.trim().toLowerCase();
      for (let i = 1; i < dataAkun.length; i++) {
        if (dataAkun[i][0].toString().toLowerCase() === emailLower && dataAkun[i][2] === 'Orang Tua') {
          sheetAkun.deleteRow(i + 1);
          break;
        }
      }
    }
    
    return { success: true, message: "Data siswa '" + idSiswa + "' dan akun orang tua berhasil dihapus permanen!" };
  } catch (error) {
    return { success: false, message: "Gagal menghapus siswa: " + error.message };
  }
}

/**
 * Menyimpan data presensi baru dari Guru ke Google Sheets
 * Dan mengupdate jika presensi pada tanggal & siswa tersebut sudah ada
 */
function simpanPresensi(tanggal, kelas, records, emailGuru) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Kehadiran');
  const data = sheet.getDataRange().getValues();
  
  // Mencari indeks baris yang sudah ada untuk tanggal + ID Siswa yang sama agar bisa di-update
  const mapExistingRow = {};
  for (let i = 1; i < data.length; i++) {
    const tglSheet = data[i][1];
    let tglString = "";
    if (tglSheet instanceof Date) {
      tglString = tglSheet.toISOString().split('T')[0];
    } else {
      tglString = tglSheet.toString().split('T')[0];
    }
    
    const idSiswa = data[i][2];
    if (tglString === tanggal) {
      mapExistingRow[idSiswa] = i + 1;
    }
  }
  
  records.forEach(rec => {
    const waktuInput = new Date().toISOString();
    const existingRow = mapExistingRow[rec.idSiswa];
    
    if (existingRow) {
      // Update data yang sudah ada
      sheet.getRange(existingRow, 6).setValue(rec.status); // Kolom F: Status
      sheet.getRange(existingRow, 7).setValue(rec.keterangan || ""); // Kolom G: Keterangan
      sheet.getRange(existingRow, 8).setValue(waktuInput); // Kolom H: Waktu Update
      sheet.getRange(existingRow, 9).setValue(emailGuru); // Kolom I: Diupdate Oleh
    } else {
      // Buat ID Presensi unik baru
      const idPresensi = 'P' + Math.floor(100000 + Math.random() * 900000);
      sheet.appendRow([
        idPresensi,
        tanggal,
        rec.idSiswa,
        rec.namaSiswa,
        kelas,
        rec.status,
        rec.keterangan || "",
        waktuInput,
        emailGuru
      ]);
    }
  });
  
  return { success: true, message: "Berhasil menyimpan presensi kelas " + kelas + " tanggal " + tanggal };
}

/**
 * Mengambil histori presensi siswa tertentu untuk Orang Tua
 */
function getHistoriSiswa(idSiswa) {
  const ss = getDb();
  
  // Cari info siswa terlebih dahulu
  const sheetSiswa = ss.getSheetByName('Siswa');
  const dataSiswa = sheetSiswa.getDataRange().getValues();
  let infoSiswa = null;
  for (let i = 1; i < dataSiswa.length; i++) {
    if (dataSiswa[i][0] === idSiswa) {
      infoSiswa = {
        idSiswa: dataSiswa[i][0],
        namaSiswa: dataSiswa[i][1],
        kelas: dataSiswa[i][2]
      };
      break;
    }
  }
  
  if (!infoSiswa) return { success: false, message: "Siswa tidak ditemukan" };
  
  // Ambil data kehadiran siswa tersebut
  const sheetKehadiran = ss.getSheetByName('Kehadiran');
  const dataKehadiran = sheetKehadiran.getDataRange().getValues();
  const histori = [];
  
  for (let i = 1; i < dataKehadiran.length; i++) {
    if (dataKehadiran[i][2] === idSiswa) {
      let tgl = dataKehadiran[i][1];
      let tglStr = tgl instanceof Date ? tgl.toISOString().split('T')[0] : tgl.toString().split('T')[0];
      
      histori.push({
        idPresensi: dataKehadiran[i][0],
        tanggal: tglStr,
        status: dataKehadiran[i][5],
        keterangan: dataKehadiran[i][6],
        waktuInput: dataKehadiran[i][7]
      });
    }
  }
  
  // Urutkan histori berdasarkan tanggal terbaru
  histori.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
  
  return {
    success: true,
    siswa: infoSiswa,
    histori: histori
  };
}

/**
 * Mengambil seluruh histori kehadiran siswa untuk Guru (Rekap Presensi & Ekspor Excel)
 */
function getAllKehadiran() {
  const ss = getDb();
  const sheet = ss.getSheetByName('Kehadiran');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const list = [];
  
  for (let i = 1; i < data.length; i++) {
    let tgl = data[i][1];
    let tglStr = tgl instanceof Date ? tgl.toISOString().split('T')[0] : tgl.toString().split('T')[0];
    
    list.push({
      idPresensi: data[i][0],
      tanggal: tglStr,
      idSiswa: data[i][2],
      namaSiswa: data[i][3],
      kelas: data[i][4],
      status: data[i][5],
      keterangan: data[i][6],
      waktuInput: data[i][7],
      inputOleh: data[i][8]
    });
  }
  
  // Urutkan berdasarkan tanggal terbaru
  list.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
  return list;
}

/**
 * Menambahkan data siswa baru dan mendaftarkan akun orang tua terkait
 */
function tambahSiswaDanOrtu(idSiswa, namaSiswa, kelas, namaOrtu, emailOrtu) {
  try {
    const ss = getDb();
    
    idSiswa = idSiswa.trim().toUpperCase();
    namaSiswa = namaSiswa.trim();
    kelas = kelas.trim();
    namaOrtu = namaOrtu.trim();
    emailOrtu = emailOrtu.trim().toLowerCase();
    
    if (!idSiswa || !namaSiswa || !kelas || !namaOrtu || !emailOrtu) {
      return { success: false, message: "Semua kolom input wajib diisi." };
    }
    
    // 1. Periksa duplikasi ID Siswa di Sheet Siswa
    const sheetSiswa = ss.getSheetByName('Siswa');
    const dataSiswa = sheetSiswa.getDataRange().getValues();
    for (let i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0].toString().toUpperCase() === idSiswa) {
        return { success: false, message: "ID Siswa '" + idSiswa + "' sudah terdaftar sebelumnya!" };
      }
    }
    
    // 2. Periksa duplikasi Email di Sheet Akun
    const sheetAkun = ss.getSheetByName('Akun');
    const dataAkun = sheetAkun.getDataRange().getValues();
    for (let i = 1; i < dataAkun.length; i++) {
      if (dataAkun[i][0].toString().toLowerCase() === emailOrtu) {
        return { success: false, message: "Email Orang Tua '" + emailOrtu + "' sudah terdaftar untuk pengguna lain!" };
      }
    }
    
    // 3. Tambahkan ke Sheet Siswa
    sheetSiswa.appendRow([idSiswa, namaSiswa, kelas, emailOrtu]);
    
    // 4. Tambahkan ke Sheet Akun
    sheetAkun.appendRow([emailOrtu, namaOrtu, 'Orang Tua', idSiswa]);
    
    return { success: true, message: "Berhasil mendaftarkan siswa " + namaSiswa + " dan akun Orang Tua " + namaOrtu };
    
  } catch (error) {
    return { success: false, message: "Kesalahan sistem: " + error.message };
  }
}
