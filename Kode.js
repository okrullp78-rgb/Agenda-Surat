const SPREADSHEET_ID = '1qzvmSS5FcqdAaFBGm_GOBFYcGJresUDWYoMhKdBgcVE';

const NAMA_SEKOLAH = 'SMP Negeri 13 Tasikmalaya';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Agenda Surat SMPN 13 Tasikmalaya')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboardData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  return {
    namaSekolah: NAMA_SEKOLAH,
    suratMasuk: hitungData_(ss, 'Surat_Masuk'),
    suratKeluar: hitungData_(ss, 'Surat_Keluar'),
    suratTugas: hitungData_(ss, 'Surat_Tugas')
  };
}

function hitungData_(ss, namaSheet) {
  const sheet = ss.getSheetByName(namaSheet);

  if (!sheet) {
    throw new Error('Sheet "' + namaSheet + '" tidak ditemukan.');
  }

  if (sheet.getLastRow() < 2) {
    return 0;
  }

  const data = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 1
  ).getValues();

  return data.filter(function(row) {
    return row[0] !== '';
  }).length;
}function simpanSuratMasuk(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Masuk');

  if (!sheet) {
    throw new Error('Sheet Surat_Masuk tidak ditemukan.');
  }

  if (!form.tanggalSurat || !form.tanggalTerima ||
      !form.asalSurat || !form.perihal) {
    throw new Error('Tanggal surat, tanggal terima, asal surat, dan perihal wajib diisi.');
  }

  const id = 'SM-' + Utilities.getUuid().substring(0, 8).toUpperCase();

  sheet.appendRow([
    id,
    form.nomorSurat || '',
    form.tanggalSurat,
    form.tanggalTerima,
    form.asalSurat,
    form.perihal,
    form.kategori || '',
    form.disposisi || '',
    form.status || 'Belum Ditindaklanjuti',
    form.linkDokumen || '',
    form.keterangan || ''
  ]);

  return {
    sukses: true,
    pesan: 'Surat masuk berhasil disimpan.',
    id: id
  };
}

function ambilSuratMasuk() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Masuk');

  if (!sheet) {
    throw new Error('Sheet Surat_Masuk tidak ditemukan.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();

  return data.reverse().map(function(row) {
    return {
      id: row[0],
      nomorSurat: row[1],
      tanggalSurat: row[2],
      tanggalTerima: row[3],
      asalSurat: row[4],
      perihal: row[5],
      kategori: row[6],
      disposisi: row[7],
      status: row[8],
      linkDokumen: row[9],
      keterangan: row[10]
    };
  });
}function siapkanDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const struktur = {
    'Surat_Masuk': [
      'ID_Surat',
      'Nomor_Surat',
      'Tanggal_Surat',
      'Tanggal_Terima',
      'Asal_Surat',
      'Perihal',
      'Kategori',
      'Disposisi',
      'Status',
      'Link_Dokumen',
      'Keterangan'
    ],

    'Surat_Keluar': [
      'ID_Surat',
      'Nomor_Surat',
      'Tanggal_Surat',
      'Tujuan_Surat',
      'Perihal',
      'Kategori',
      'Penandatangan',
      'Status',
      'Link_Dokumen',
      'Keterangan'
    ],

    'Surat_Tugas': [
      'ID_Tugas',
      'Nomor_Surat',
      'Tanggal_Surat',
      'Nama_Pegawai',
      'NIP',
      'Jabatan',
      'Tujuan_Tugas',
      'Tanggal_Mulai',
      'Tanggal_Selesai',
      'Keperluan',
      'Penandatangan',
      'Link_Dokumen',
      'Status'
    ],

    'Pengaturan': [
      'Parameter',
      'Nilai',
      'Keterangan'
    ]
  };

  Object.keys(struktur).forEach(function(namaSheet) {
    let sheet = ss.getSheetByName(namaSheet);

    if (!sheet) {
      sheet = ss.insertSheet(namaSheet);
    }

    const header = struktur[namaSheet];
    const barisPertama = sheet.getRange(
      1, 1, 1, header.length
    ).getDisplayValues()[0];

    const kosong = barisPertama.every(function(nilai) {
      return nilai === '';
    });

    if (kosong) {
      sheet.getRange(1, 1, 1, header.length)
        .setValues([header]);
    }

    sheet.setFrozenRows(1);

    sheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#0f766e')
      .setFontColor('#ffffff');

    sheet.autoResizeColumns(1, header.length);
  });

  const pengaturan = ss.getSheetByName('Pengaturan');

  const dataPengaturan = [
    ['Nama_Sekolah', 'SMP Negeri 13 Tasikmalaya', 'Nama resmi sekolah'],
    ['Tahun_Aktif', new Date().getFullYear(), 'Tahun administrasi aktif'],
    ['Format_Nomor_Surat', 'Nomor urut/tahun', 'Format nomor surat'],
    ['Zona_Waktu', 'Asia/Jakarta', 'Zona waktu aplikasi']
  ];

  const parameterAda = pengaturan.getLastRow() > 1
    ? pengaturan.getRange(
        2, 1, pengaturan.getLastRow() - 1, 1
      ).getDisplayValues().flat()
    : [];

  dataPengaturan.forEach(function(baris) {
    if (parameterAda.indexOf(baris[0]) === -1) {
      pengaturan.appendRow(baris);
    }
  });

  SpreadsheetApp.flush();

  Logger.log('Database berhasil disiapkan: ' + ss.getName());

  return 'Database berhasil disiapkan.';
}// ==========================================
// FITUR SURAT KELUAR
// ==========================================

function simpanSuratKeluar(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Keluar');

  if (!sheet) {
    throw new Error('Sheet Surat_Keluar tidak ditemukan.');
  }

  // Validasi data wajib
  if (
    !form.nomorSurat ||
    !form.tanggalSurat ||
    !form.tujuanSurat ||
    !form.perihal
  ) {
    throw new Error(
      'Nomor surat, tanggal surat, tujuan surat, dan perihal wajib diisi.'
    );
  }

  // Membuat ID surat secara otomatis
  const id = 'SK-' +
    Utilities.getUuid().substring(0, 8).toUpperCase();

  // Menyimpan data ke Google Sheets
  sheet.appendRow([
    id,
    form.nomorSurat,
    form.tanggalSurat,
    form.tujuanSurat,
    form.perihal,
    form.kategori || '',
    form.penandatangan || '',
    form.status || 'Draft',
    form.linkDokumen || '',
    form.keterangan || ''
  ]);

  return {
    sukses: true,
    pesan: 'Surat keluar berhasil disimpan.',
    id: id
  };
}


// ==========================================
// MENGAMBIL DAFTAR SURAT KELUAR
// ==========================================

function ambilSuratKeluar() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Keluar');

  if (!sheet) {
    throw new Error('Sheet Surat_Keluar tidak ditemukan.');
  }

  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir < 2) {
    return [];
  }

  const data = sheet
    .getRange(2, 1, barisTerakhir - 1, 10)
    .getDisplayValues();

  return data.reverse().map(function(row) {
    return {
      id: row[0],
      nomorSurat: row[1],
      tanggalSurat: row[2],
      tujuanSurat: row[3],
      perihal: row[4],
      kategori: row[5],
      penandatangan: row[6],
      status: row[7],
      linkDokumen: row[8],
      keterangan: row[9]
    };
  });
}