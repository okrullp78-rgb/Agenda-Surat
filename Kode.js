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
}

function cariBarisSuratMasukById_(sheet, id) {
  const idCari = String(id || '').trim();

  if (!idCari) {
    return -1;
  }

  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir < 2) {
    return -1;
  }

  const ids = sheet.getRange(2, 1, barisTerakhir - 1, 1).getDisplayValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === idCari) {
      return i + 2;
    }
  }

  return -1;
}

function ambilSuratMasukById(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Masuk');

  if (!sheet) {
    throw new Error('Sheet Surat_Masuk tidak ditemukan.');
  }

  const nomorBaris = cariBarisSuratMasukById_(sheet, id);

  if (nomorBaris < 0) {
    throw new Error('Surat masuk dengan ID "' + id + '" tidak ditemukan.');
  }

  const row = sheet.getRange(nomorBaris, 1, 1, 11).getDisplayValues()[0];

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
}

function simpanSuratMasuk(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Masuk');

  if (!sheet) {
    throw new Error('Sheet Surat_Masuk tidak ditemukan.');
  }

  if (!form.tanggalSurat || !form.tanggalTerima ||
      !form.asalSurat || !form.perihal) {
    throw new Error('Tanggal surat, tanggal terima, asal surat, dan perihal wajib diisi.');
  }

  // Jika ada berkas dokumen yang diunggah dari komputer, simpan ke Google Drive
  if (form.berkasBase64 && form.berkasNama) {
    try {
      const linkFileDrive = simpanBerkasKeDrive_(form.berkasNama, form.berkasType, form.berkasBase64);
      if (linkFileDrive) {
        form.linkDokumen = linkFileDrive;
      }
    } catch (errDrive) {
      Logger.log('Gagal menyimpan berkas ke Google Drive: ' + errDrive.message);
      throw new Error('Gagal mengunggah berkas ke Drive: ' + errDrive.message);
    }
  }

  const idAda = String(form.idSurat || form.id || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    let nomorBaris = -1;
    let id = idAda;

    if (idAda) {
      nomorBaris = cariBarisSuratMasukById_(sheet, idAda);
      if (nomorBaris < 0) {
        throw new Error('Data surat masuk tidak ditemukan. Pembaruan dibatalkan.');
      }
    } else {
      id = 'SM-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const nilaiBaris = [
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
    ];

    if (nomorBaris > 0) {
      sheet.getRange(nomorBaris, 1, 1, 11).setValues([nilaiBaris]);
    } else {
      sheet.appendRow(nilaiBaris);
    }

    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: nomorBaris > 0
        ? 'Surat masuk berhasil diperbarui.'
        : 'Surat masuk berhasil disimpan.',
      id: id
    };
  } finally {
    lock.releaseLock();
  }
}

function ubahSuratMasuk(form) {
  return simpanSuratMasuk(form);
}

function hapusSuratMasuk(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Masuk');

  if (!sheet) {
    throw new Error('Sheet Surat_Masuk tidak ditemukan.');
  }

  const idCari = String(id || '').trim();
  if (!idCari) {
    throw new Error('ID surat masuk wajib disertakan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const nomorBaris = cariBarisSuratMasukById_(sheet, idCari);
    if (nomorBaris < 0) {
      throw new Error('Surat masuk dengan ID "' + idCari + '" tidak ditemukan atau sudah dihapus.');
    }

    sheet.deleteRow(nomorBaris);
    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: 'Surat masuk berhasil dihapus.',
      id: idCari
    };
  } finally {
    lock.releaseLock();
  }
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
}

function simpanBerkasKeDrive_(namaFile, mimeType, base64Data) {
  const namaFolder = 'Arsip_Surat_Masuk';
  const folderIter = DriveApp.getFoldersByName(namaFolder);
  let folder;
  if (folderIter.hasNext()) {
    folder = folderIter.next();
  } else {
    folder = DriveApp.createFolder(namaFolder);
  }

  let cleanBase64 = String(base64Data || '');
  if (cleanBase64.indexOf('base64,') > -1) {
    cleanBase64 = cleanBase64.split('base64,')[1];
  }

  const bytes = Utilities.base64Decode(cleanBase64);
  const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', namaFile || 'Berkas_Surat_Masuk.pdf');
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (eShare) {
    Logger.log('Set sharing error: ' + eShare.message);
  }

  return file.getUrl();
}

function unggahBerkasSuratMasuk(namaFile, mimeType, base64Data) {
  const url = simpanBerkasKeDrive_(namaFile, mimeType, base64Data);
  return {
    sukses: true,
    url: url
  };
}

function siapkanDatabase() {
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

    // Dokumen terkait Surat Tugas (format mengacu workbook sekolah).
    // Semua ditambahkan di akhir agar struktur sheet lama tidak berubah.
    'Surat_Tugas_Siswa': [
      'ID_Tugas',
      'Nomor_Surat',
      'Tanggal_Surat',
      'Nama_Siswa',
      'No_Induk',
      'Kelas',
      'Tujuan_Tugas',
      'Tanggal_Mulai',
      'Tanggal_Selesai',
      'Keperluan',
      'Penandatangan',
      'Link_Dokumen',
      'Status'
    ],

    'SPPD': [
      'ID_SPPD',
      'Nomor_Surat',
      'Pengguna_Anggaran',
      'Nama_Pegawai',
      'NIP',
      'Pangkat_Golongan',
      'Jabatan',
      'Tingkat_Biaya',
      'Maksud',
      'Alat_Angkut',
      'Tempat_Berangkat',
      'Tempat_Tujuan',
      'Lama_Perjalanan',
      'Tanggal_Berangkat',
      'Tanggal_Kembali',
      'Instansi',
      'Mata_Anggaran',
      'Keterangan',
      'Status'
    ],

    'SPPD_Lembar2': [
      'ID_Lembar',
      'Nomor_Surat',
      'Nama_Pegawai',
      'Urutan',
      'Tiba_Di',
      'Berangkat_Dari',
      'Ke',
      'Tanggal',
      'Status'
    ],

    'Nota_Dinas': [
      'ID_Nota',
      'Nomor_Surat',
      'Kepada_Yth',
      'Dari',
      'Perihal',
      'Tempat_Kegiatan',
      'Tanggal_Pelaksanaan',
      'Dasar_Penugasan',
      'Hasil',
      'Usul',
      'Status'
    ],

    'Daftar_Penyerahan': [
      'ID_Penyerahan',
      'Nomor_Surat',
      'Nama_Pegawai',
      'NIP',
      'Golongan',
      'Volume_Hari',
      'Tarif',
      'Jumlah',
      'Status'
    ],

    'Daftar_Penyerahan_Siswa': [
      'ID_Penyerahan',
      'Nomor_Surat',
      'Nama_Siswa',
      'Kelas',
      'Volume_Hari',
      'Tarif',
      'Jumlah',
      'Status'
    ],

    'Pengaturan': [
      'Parameter',
      'Nilai',
      'Keterangan'
    ],

    'Pejabat': [
      'ID_Pejabat',
      'Kategori',
      'Nama',
      'NIP',
      'Jabatan',
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
    ['Zona_Waktu', 'Asia/Jakarta', 'Zona waktu aplikasi'],
    ['Alamat_Sekolah', '', 'Alamat resmi sekolah (isi jika sudah tersedia)'],
    ['Telepon_Sekolah', '', 'Nomor telepon sekolah (isi jika sudah tersedia)'],
    ['Kota_Surat', 'Tasikmalaya', 'Kota pada tanggal surat'],
    ['Nama_Kepala_Sekolah', '', 'Nama kepala sekolah (isi jika sudah tersedia)'],
    ['Jabatan_Kepala_Sekolah', 'Kepala Sekolah', 'Jabatan pejabat pada kop/tanda tangan'],
    ['NIP_Kepala_Sekolah', '', 'NIP kepala sekolah (isi jika sudah tersedia)'],
    ['Nama_Bendahara', '', 'Nama bendahara sekolah'],
    ['Jabatan_Bendahara', 'Bendahara', 'Jabatan bendahara sekolah'],
    ['NIP_Bendahara', '', 'NIP bendahara sekolah'],
    ['Nama_Petugas', '', 'Nama petugas agenda / pengadministrasi surat'],
    ['Jabatan_Petugas', 'Pengadministrasi Persuratan', 'Jabatan petugas pengelola surat'],
    ['NIP_Petugas', '', 'NIP / NUPTK petugas agenda']
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

  const sheetPejabat = ss.getSheetByName('Pejabat');
  if (sheetPejabat && sheetPejabat.getLastRow() < 2) {
    sheetPejabat.appendRow(['PEJ-01', 'Kepala Sekolah', '', '', 'Kepala Sekolah', 'Pejabat Penandatangan Utama']);
    sheetPejabat.appendRow(['PEJ-02', 'Bendahara', '', '', 'Bendahara', 'Pejabat Keuangan']);
    sheetPejabat.appendRow(['PEJ-03', 'Petugas', '', '', 'Pengadministrasi Persuratan', 'Petugas Agenda & Arsip']);
  }

  const sheetKeluar = ss.getSheetByName('Surat_Keluar');
  if (sheetKeluar) {
    pastikanKolomSuratKeluar_(sheetKeluar);
  }

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

  const peta = pastikanKolomSuratKeluar_(sheet);

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

  const idAda = String(form.idSurat || form.id || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    let nomorBaris = -1;
    let id = idAda;

    if (idAda) {
      nomorBaris = cariBarisSuratKeluarById_(sheet, idAda);
      if (nomorBaris < 0) {
        throw new Error('Surat keluar tidak ditemukan. Pembaruan dibatalkan.');
      }
    } else {
      // Pembuatan surat keluar baru hanya melalui formulir Buat Surat
      // (simpanDraftSuratKeluar) agar tidak ada surat tercatat dua kali.
      throw new Error(
        'Pembuatan surat keluar baru hanya dapat dilakukan melalui menu Buat Surat.'
      );
    }

    const jumlahKolom = peta.jumlahKolom;
    let baris;

    if (nomorBaris > 0) {
      baris = sheet.getRange(nomorBaris, 1, 1, jumlahKolom).getValues()[0];
    } else {
      baris = [];
      for (let i = 0; i < jumlahKolom; i++) {
        baris.push('');
      }
    }

    baris[0] = id;
    baris[1] = form.nomorSurat;
    baris[2] = form.tanggalSurat;
    baris[3] = form.tujuanSurat;
    baris[4] = form.perihal;
    baris[5] = form.kategori || baris[5] || '';
    baris[6] = form.penandatangan || '';
    
    if (form.status) {
      baris[7] = form.status;
    } else if (nomorBaris > 0) {
      baris[7] = baris[7] || 'Draft';
    } else {
      baris[7] = 'Draft';
    }

    baris[8] = form.linkDokumen !== undefined ? (form.linkDokumen || '') : (baris[8] || '');
    baris[9] = form.keterangan !== undefined ? (form.keterangan || '') : (baris[9] || '');

    if (nomorBaris > 0) {
      sheet.getRange(nomorBaris, 1, 1, jumlahKolom).setValues([baris]);
    } else {
      sheet.appendRow(baris);
    }

    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: nomorBaris > 0
        ? 'Surat keluar berhasil diperbarui.'
        : 'Surat keluar berhasil disimpan.',
      id: id
    };
  } finally {
    lock.releaseLock();
  }
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

  const peta = petaKolomSuratKeluar_(sheet);
  const data = sheet
    .getRange(2, 1, barisTerakhir - 1, peta.jumlahKolom)
    .getDisplayValues();

  return data.reverse().map(function(row) {
    return objekSuratKeluarDariBaris_(row, peta);
  });
}

var KOLOM_SURAT_KELUAR_TAMBAHAN = [
  'Jenis_Surat',
  'Lampiran',
  'Jabatan_Penerima',
  'Isi_Surat',
  'Penutup',
  'Jabatan_Penandatangan',
  'Tembusan'
];

function pastikanKolomSuratKeluar_(sheet) {
  KOLOM_SURAT_KELUAR_TAMBAHAN.forEach(function(nama) {
    const peta = petaKolomSuratKeluar_(sheet);
    if (peta.indeks[nama]) {
      return;
    }

    const kolom = Math.max(sheet.getLastColumn(), 10) + 1;
    sheet.getRange(1, kolom)
      .setValue(nama)
      .setFontWeight('bold')
      .setBackground('#0f766e')
      .setFontColor('#ffffff');
  });

  return petaKolomSuratKeluar_(sheet);
}

function petaKolomSuratKeluar_(sheet) {
  const jumlahKolom = Math.max(sheet.getLastColumn(), 10);
  const headers = sheet.getRange(1, 1, 1, jumlahKolom).getDisplayValues()[0];
  const indeks = {};

  headers.forEach(function(nama, i) {
    if (nama) {
      indeks[nama] = i + 1;
    }
  });

  return {
    indeks: indeks,
    jumlahKolom: jumlahKolom,
    headers: headers
  };
}

function nilaiKolomBaris_(row, peta, nama, indeksLama) {
  if (peta.indeks[nama]) {
    return row[peta.indeks[nama] - 1] || '';
  }

  if (indeksLama >= 0 && indeksLama < row.length) {
    return row[indeksLama] || '';
  }

  return '';
}

function objekSuratKeluarDariBaris_(row, peta) {
  return {
    id: row[0] || '',
    nomorSurat: row[1] || '',
    tanggalSurat: row[2] || '',
    tujuanSurat: row[3] || '',
    perihal: row[4] || '',
    kategori: row[5] || '',
    penandatangan: row[6] || '',
    status: row[7] || '',
    linkDokumen: row[8] || '',
    keterangan: row[9] || '',
    jenisSurat: nilaiKolomBaris_(row, peta, 'Jenis_Surat', -1),
    lampiran: nilaiKolomBaris_(row, peta, 'Lampiran', -1),
    jabatanPenerima: nilaiKolomBaris_(row, peta, 'Jabatan_Penerima', -1),
    isiSurat: nilaiKolomBaris_(row, peta, 'Isi_Surat', -1),
    penutup: nilaiKolomBaris_(row, peta, 'Penutup', -1),
    jabatanPenandatangan: nilaiKolomBaris_(row, peta, 'Jabatan_Penandatangan', -1),
    tembusan: nilaiKolomBaris_(row, peta, 'Tembusan', -1)
  };
}

function kategoriDariJenis_(jenis) {
  const peta = {
    'Surat Undangan': 'Undangan',
    'Surat Pemberitahuan': 'Pemberitahuan',
    'Surat Permohonan': 'Permohonan',
    'Surat Keterangan': 'Keterangan',
    'Surat Tugas': 'Tugas',
    'Surat Pengantar': 'Pengantar',
    'Surat Balasan': 'Balasan',
    'Surat Lainnya': 'Lainnya'
  };

  return peta[jenis] || '';
}

function cariBarisSuratKeluarById_(sheet, id) {
  const idCari = String(id || '').trim();

  if (!idCari) {
    return -1;
  }

  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir < 2) {
    return -1;
  }

  const ids = sheet.getRange(2, 1, barisTerakhir - 1, 1).getDisplayValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === idCari) {
      return i + 2;
    }
  }

  return -1;
}

function tulisNilaiKolom_(baris, peta, nama, nilai) {
  if (!peta.indeks[nama]) {
    return;
  }

  baris[peta.indeks[nama] - 1] = nilai;
}

function ambilPengaturanKop() {
  const hasil = {
    namaSekolah: NAMA_SEKOLAH,
    alamat: '',
    telepon: '',
    kota: 'Tasikmalaya',
    namaKepala: '',
    jabatanKepala: 'Kepala Sekolah',
    nipKepala: '',
    namaBendahara: '',
    jabatanBendahara: 'Bendahara',
    nipBendahara: '',
    namaPetugas: '',
    jabatanPetugas: 'Pengadministrasi Persuratan',
    nipPetugas: ''
  };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Pengaturan');

  if (!sheet || sheet.getLastRow() < 2) {
    return hasil;
  }

  const data = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 2
  ).getDisplayValues();

  const peta = {};
  data.forEach(function(baris) {
    if (baris[0]) {
      peta[baris[0]] = String(baris[1] || '').trim();
    }
  });

  if (peta.Nama_Sekolah) hasil.namaSekolah = peta.Nama_Sekolah;
  if (peta.Alamat_Sekolah) hasil.alamat = peta.Alamat_Sekolah;
  if (peta.Telepon_Sekolah) hasil.telepon = peta.Telepon_Sekolah;
  if (peta.Kota_Surat) hasil.kota = peta.Kota_Surat;
  if (peta.Nama_Kepala_Sekolah) hasil.namaKepala = peta.Nama_Kepala_Sekolah;
  if (peta.Jabatan_Kepala_Sekolah) hasil.jabatanKepala = peta.Jabatan_Kepala_Sekolah;
  if (peta.NIP_Kepala_Sekolah) hasil.nipKepala = peta.NIP_Kepala_Sekolah;
  if (peta.Nama_Bendahara) hasil.namaBendahara = peta.Nama_Bendahara;
  if (peta.Jabatan_Bendahara) hasil.jabatanBendahara = peta.Jabatan_Bendahara;
  if (peta.NIP_Bendahara) hasil.nipBendahara = peta.NIP_Bendahara;
  if (peta.Nama_Petugas) hasil.namaPetugas = peta.Nama_Petugas;
  if (peta.Jabatan_Petugas) hasil.jabatanPetugas = peta.Jabatan_Petugas;
  if (peta.NIP_Petugas) hasil.nipPetugas = peta.NIP_Petugas;

  return hasil;
}

function pastikanSheetPejabat_(ss) {
  let sheet = ss.getSheetByName('Pejabat');
  if (!sheet) {
    sheet = ss.insertSheet('Pejabat');
  }
  const header = ['ID_Pejabat', 'Kategori', 'Nama', 'NIP', 'Jabatan', 'Keterangan'];
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#0f766e')
      .setFontColor('#ffffff');
    sheet.autoResizeColumns(1, header.length);
  }
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(['PEJ-01', 'Kepala Sekolah', '', '', 'Kepala Sekolah', 'Pejabat Penandatangan Utama']);
    sheet.appendRow(['PEJ-02', 'Bendahara', '', '', 'Bendahara', 'Pejabat Keuangan']);
    sheet.appendRow(['PEJ-03', 'Petugas', '', '', 'Pengadministrasi Persuratan', 'Petugas Agenda & Arsip']);
  }
  return sheet;
}

function ambilMasterPejabat() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  pastikanSheetPejabat_(ss);

  const hasil = {
    kop: {
      namaSekolah: NAMA_SEKOLAH,
      alamat: '',
      telepon: '',
      kota: 'Tasikmalaya'
    },
    kepalaSekolah: { nama: '', nip: '', jabatan: 'Kepala Sekolah' },
    bendahara: { nama: '', nip: '', jabatan: 'Bendahara' },
    petugas: { nama: '', nip: '', jabatan: 'Pengadministrasi Persuratan' }
  };

  const sheetPengaturan = ss.getSheetByName('Pengaturan');
  if (sheetPengaturan && sheetPengaturan.getLastRow() > 1) {
    const data = sheetPengaturan.getRange(
      2, 1, sheetPengaturan.getLastRow() - 1, 2
    ).getDisplayValues();

    const peta = {};
    data.forEach(function(r) {
      if (r[0]) peta[r[0]] = String(r[1] || '').trim();
    });

    if (peta.Nama_Sekolah) hasil.kop.namaSekolah = peta.Nama_Sekolah;
    if (peta.Alamat_Sekolah) hasil.kop.alamat = peta.Alamat_Sekolah;
    if (peta.Telepon_Sekolah) hasil.kop.telepon = peta.Telepon_Sekolah;
    if (peta.Kota_Surat) hasil.kop.kota = peta.Kota_Surat;

    if (peta.Nama_Kepala_Sekolah) hasil.kepalaSekolah.nama = peta.Nama_Kepala_Sekolah;
    if (peta.NIP_Kepala_Sekolah) hasil.kepalaSekolah.nip = peta.NIP_Kepala_Sekolah;
    if (peta.Jabatan_Kepala_Sekolah) hasil.kepalaSekolah.jabatan = peta.Jabatan_Kepala_Sekolah;

    if (peta.Nama_Bendahara) hasil.bendahara.nama = peta.Nama_Bendahara;
    if (peta.NIP_Bendahara) hasil.bendahara.nip = peta.NIP_Bendahara;
    if (peta.Jabatan_Bendahara) hasil.bendahara.jabatan = peta.Jabatan_Bendahara;

    if (peta.Nama_Petugas) hasil.petugas.nama = peta.Nama_Petugas;
    if (peta.NIP_Petugas) hasil.petugas.nip = peta.NIP_Petugas;
    if (peta.Jabatan_Petugas) hasil.petugas.jabatan = peta.Jabatan_Petugas;
  }

  const sheetPejabat = ss.getSheetByName('Pejabat');
  if (sheetPejabat && sheetPejabat.getLastRow() > 1) {
    const dataP = sheetPejabat.getRange(
      2, 1, sheetPejabat.getLastRow() - 1, 5
    ).getDisplayValues();

    dataP.forEach(function(row) {
      const kat = String(row[1] || '').toUpperCase();
      if (kat.indexOf('KEPALA') > -1) {
        if (row[2]) hasil.kepalaSekolah.nama = row[2];
        if (row[3]) hasil.kepalaSekolah.nip = row[3];
        if (row[4]) hasil.kepalaSekolah.jabatan = row[4];
      } else if (kat.indexOf('BENDAHARA') > -1) {
        if (row[2]) hasil.bendahara.nama = row[2];
        if (row[3]) hasil.bendahara.nip = row[3];
        if (row[4]) hasil.bendahara.jabatan = row[4];
      } else if (kat.indexOf('PETUGAS') > -1) {
        if (row[2]) hasil.petugas.nama = row[2];
        if (row[3]) hasil.petugas.nip = row[3];
        if (row[4]) hasil.petugas.jabatan = row[4];
      }
    });
  }

  return hasil;
}

function simpanMasterPejabat(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  pastikanSheetPejabat_(ss);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheetPengaturan = ss.getSheetByName('Pengaturan');
    if (sheetPengaturan) {
      const nilaiSimpan = {
        'Nama_Sekolah': (data.kop && data.kop.namaSekolah) || NAMA_SEKOLAH,
        'Alamat_Sekolah': (data.kop && data.kop.alamat) || '',
        'Telepon_Sekolah': (data.kop && data.kop.telepon) || '',
        'Kota_Surat': (data.kop && data.kop.kota) || 'Tasikmalaya',
        'Nama_Kepala_Sekolah': (data.kepalaSekolah && data.kepalaSekolah.nama) || '',
        'NIP_Kepala_Sekolah': (data.kepalaSekolah && data.kepalaSekolah.nip) || '',
        'Jabatan_Kepala_Sekolah': (data.kepalaSekolah && data.kepalaSekolah.jabatan) || 'Kepala Sekolah',
        'Nama_Bendahara': (data.bendahara && data.bendahara.nama) || '',
        'NIP_Bendahara': (data.bendahara && data.bendahara.nip) || '',
        'Jabatan_Bendahara': (data.bendahara && data.bendahara.jabatan) || 'Bendahara',
        'Nama_Petugas': (data.petugas && data.petugas.nama) || '',
        'NIP_Petugas': (data.petugas && data.petugas.nip) || '',
        'Jabatan_Petugas': (data.petugas && data.petugas.jabatan) || 'Pengadministrasi Persuratan'
      };

      const lastRow = sheetPengaturan.getLastRow();
      const existingParams = lastRow > 1
        ? sheetPengaturan.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(function(r) { return r[0]; })
        : [];

      Object.keys(nilaiSimpan).forEach(function(param) {
        const idx = existingParams.indexOf(param);
        if (idx > -1) {
          sheetPengaturan.getRange(idx + 2, 2).setValue(nilaiSimpan[param]);
        } else {
          sheetPengaturan.appendRow([param, nilaiSimpan[param], 'Pengaturan Master']);
          existingParams.push(param);
        }
      });
    }

    const sheetPejabat = ss.getSheetByName('Pejabat');
    if (sheetPejabat) {
      const rows = [
        ['PEJ-01', 'Kepala Sekolah', (data.kepalaSekolah && data.kepalaSekolah.nama) || '', (data.kepalaSekolah && data.kepalaSekolah.nip) || '', (data.kepalaSekolah && data.kepalaSekolah.jabatan) || 'Kepala Sekolah', 'Pejabat Penandatangan Utama'],
        ['PEJ-02', 'Bendahara', (data.bendahara && data.bendahara.nama) || '', (data.bendahara && data.bendahara.nip) || '', (data.bendahara && data.bendahara.jabatan) || 'Bendahara', 'Pejabat Keuangan'],
        ['PEJ-03', 'Petugas', (data.petugas && data.petugas.nama) || '', (data.petugas && data.petugas.nip) || '', (data.petugas && data.petugas.jabatan) || 'Pengadministrasi Persuratan', 'Petugas Agenda & Arsip']
      ];

      const lastRowP = sheetPejabat.getLastRow();
      if (lastRowP > 1) {
        sheetPejabat.getRange(2, 1, lastRowP - 1, 6).clearContent();
      }
      sheetPejabat.getRange(2, 1, rows.length, 6).setValues(rows);
    }

    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: 'Data Master Pejabat dan Kop Surat berhasil disimpan permanen ke Spreadsheet.'
    };
  } finally {
    lock.releaseLock();
  }
}

function ambilSuratKeluarById(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Keluar');

  if (!sheet) {
    throw new Error('Sheet Surat_Keluar tidak ditemukan.');
  }

  const nomorBaris = cariBarisSuratKeluarById_(sheet, id);

  if (nomorBaris < 0) {
    throw new Error('Surat keluar tidak ditemukan.');
  }

  const peta = petaKolomSuratKeluar_(sheet);
  const row = sheet.getRange(
    nomorBaris, 1, 1, peta.jumlahKolom
  ).getDisplayValues()[0];

  return objekSuratKeluarDariBaris_(row, peta);
}

function simpanDraftSuratKeluar(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Keluar');

  if (!sheet) {
    throw new Error('Sheet Surat_Keluar tidak ditemukan.');
  }

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

  const peta = pastikanKolomSuratKeluar_(sheet);
  const idAda = String(form.idSurat || form.id || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    let nomorBaris = -1;
    let id = idAda;

    if (idAda) {
      nomorBaris = cariBarisSuratKeluarById_(sheet, idAda);

      if (nomorBaris < 0) {
        throw new Error(
          'Draft tidak ditemukan. Penyimpanan dibatalkan agar data tidak terduplikasi.'
        );
      }
    } else {
      id = 'SK-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const jumlahKolom = peta.jumlahKolom;
    let baris;

    if (nomorBaris > 0) {
      baris = sheet.getRange(nomorBaris, 1, 1, jumlahKolom).getValues()[0];
    } else {
      baris = [];
      for (let i = 0; i < jumlahKolom; i++) {
        baris.push('');
      }
    }

    baris[0] = id;
    baris[1] = form.nomorSurat;
    baris[2] = form.tanggalSurat;
    baris[3] = form.tujuanSurat;
    baris[4] = form.perihal;
    baris[6] = form.penandatangan || '';

    // Pertahankan status yang ada jika sudah diproses, jangan otomatis reset ke Draft
    if (form.status) {
      baris[7] = form.status;
    } else if (nomorBaris > 0) {
      baris[7] = baris[7] || 'Draft';
    } else {
      baris[7] = 'Draft';
    }

    if (form.linkDokumen !== undefined) {
      baris[8] = form.linkDokumen || '';
    } else if (nomorBaris < 0) {
      baris[8] = '';
    }

    if (form.keterangan !== undefined) {
      baris[9] = form.keterangan || '';
    } else if (nomorBaris < 0) {
      baris[9] = '';
    }

    if (form.jenisSurat) {
      baris[5] = kategoriDariJenis_(form.jenisSurat) || baris[5] || '';
    } else if (nomorBaris < 0) {
      baris[5] = form.kategori || '';
    }

    var jenisLama = peta.indeks.Jenis_Surat
      ? (baris[peta.indeks.Jenis_Surat - 1] || '')
      : '';
    tulisNilaiKolom_(baris, peta, 'Jenis_Surat', form.jenisSurat || jenisLama);
    tulisNilaiKolom_(baris, peta, 'Lampiran', form.lampiran || '');
    tulisNilaiKolom_(baris, peta, 'Jabatan_Penerima', form.jabatanPenerima || '');
    tulisNilaiKolom_(baris, peta, 'Isi_Surat', form.isiSurat || '');
    tulisNilaiKolom_(baris, peta, 'Penutup', form.penutup || '');
    tulisNilaiKolom_(baris, peta, 'Jabatan_Penandatangan', form.jabatanPenandatangan || '');
    tulisNilaiKolom_(baris, peta, 'Tembusan', form.tembusan || '');

    if (nomorBaris > 0) {
      sheet.getRange(nomorBaris, 1, 1, jumlahKolom).setValues([baris]);
    } else {
      sheet.appendRow(baris);
    }

    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: nomorBaris > 0
        ? 'Draft surat keluar berhasil diperbarui.'
        : 'Draft surat keluar berhasil disimpan.',
      id: id
    };
  } finally {
    lock.releaseLock();
  }
}

function hapusSuratKeluar(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Keluar');

  if (!sheet) {
    throw new Error('Sheet Surat_Keluar tidak ditemukan.');
  }

  const idCari = String(id || '').trim();
  if (!idCari) {
    throw new Error('ID surat keluar wajib disertakan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const nomorBaris = cariBarisSuratKeluarById_(sheet, idCari);
    if (nomorBaris < 0) {
      throw new Error('Surat keluar dengan ID "' + idCari + '" tidak ditemukan atau sudah dihapus.');
    }

    sheet.deleteRow(nomorBaris);
    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: 'Surat keluar berhasil dihapus.',
      id: idCari
    };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// SURAT TUGAS: BACA DATA (UNTUK AGENDA SURAT)
// ==========================================

// Membaca data Surat_Tugas berdasarkan nama header sheet, tanpa mengubah
// struktur atau isi sheet. Struktur acuan mengikuti siapkanDatabase().
function ambilSuratTugas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Tugas');

  if (!sheet) {
    throw new Error('Sheet Surat_Tugas tidak ditemukan.');
  }

  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir < 2) {
    return [];
  }

  const jumlahKolom = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet
    .getRange(1, 1, 1, jumlahKolom)
    .getDisplayValues()[0];

  const indeks = {};
  headers.forEach(function(nama, i) {
    const kunci = String(nama || '').trim();
    if (kunci && indeks[kunci] === undefined) {
      indeks[kunci] = i;
    }
  });

  const data = sheet
    .getRange(2, 1, barisTerakhir - 1, jumlahKolom)
    .getDisplayValues();

  function nilai_(baris, nama, indeksCadangan) {
    if (indeks[nama] !== undefined) {
      return String(baris[indeks[nama]] || '').trim();
    }

    if (indeksCadangan >= 0 && indeksCadangan < baris.length) {
      return String(baris[indeksCadangan] || '').trim();
    }

    return '';
  }

  return data
    .filter(function(baris) {
      return String(baris[0] || '').trim() !== '';
    })
    .reverse()
    .map(function(baris) {
      return {
        id: nilai_(baris, 'ID_Tugas', 0),
        nomorSurat: nilai_(baris, 'Nomor_Surat', 1),
        tanggalSurat: nilai_(baris, 'Tanggal_Surat', 2),
        namaPegawai: nilai_(baris, 'Nama_Pegawai', 3),
        nip: nilai_(baris, 'NIP', 4),
        jabatan: nilai_(baris, 'Jabatan', 5),
        tujuanTugas: nilai_(baris, 'Tujuan_Tugas', 6),
        tanggalMulai: nilai_(baris, 'Tanggal_Mulai', 7),
        tanggalSelesai: nilai_(baris, 'Tanggal_Selesai', 8),
        keperluan: nilai_(baris, 'Keperluan', 9),
        penandatangan: nilai_(baris, 'Penandatangan', 10),
        linkDokumen: nilai_(baris, 'Link_Dokumen', 11),
        status: nilai_(baris, 'Status', 12)
      };
    });
}

// ==========================================
// SURAT TUGAS: KOLOM TAMBAHAN & PEMETAAN
// ==========================================

// Kolom tambahan (ditambahkan di AKHIR sheet, tanpa mengubah urutan/nama
// 13 kolom dasar) untuk mendukung format resmi Surat Tugas sekolah.
// Mengikuti pola yang sama dengan KOLOM_SURAT_KELUAR_TAMBAHAN.
var KOLOM_SURAT_TUGAS_TAMBAHAN = [
  'Dasar_Surat',
  'Hari',
  'Waktu',
  'Tempat',
  'Jabatan_Penandatangan',
  'Tembusan'
];

function pastikanKolomSuratTugas_(sheet) {
  KOLOM_SURAT_TUGAS_TAMBAHAN.forEach(function(nama) {
    const peta = petaKolomSuratTugas_(sheet);
    if (peta.indeks[nama]) {
      return;
    }

    const kolom = Math.max(sheet.getLastColumn(), 13) + 1;
    sheet.getRange(1, kolom)
      .setValue(nama)
      .setFontWeight('bold')
      .setBackground('#0f766e')
      .setFontColor('#ffffff');
  });

  return petaKolomSuratTugas_(sheet);
}

// Memetakan nama kolom ke indeks 1-based berdasarkan header sheet.
function petaKolomSuratTugas_(sheet) {
  const jumlahKolom = Math.max(sheet.getLastColumn(), 13);
  const headers = sheet
    .getRange(1, 1, 1, jumlahKolom)
    .getDisplayValues()[0];

  const indeks = {};
  headers.forEach(function(nama, i) {
    if (nama) {
      indeks[nama] = i + 1;
    }
  });

  return {
    indeks: indeks,
    jumlahKolom: jumlahKolom,
    headers: headers
  };
}

function tulisKolomTugas_(baris, peta, nama, nilai) {
  if (!peta.indeks[nama]) {
    return;
  }

  baris[peta.indeks[nama] - 1] = nilai;
}

function cariBarisSuratTugasById_(sheet, id) {
  const idCari = String(id || '').trim();
  if (!idCari) {
    return -1;
  }

  const barisTerakhir = sheet.getLastRow();
  if (barisTerakhir < 2) {
    return -1;
  }

  const ids = sheet.getRange(2, 1, barisTerakhir - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === idCari) {
      return i + 2;
    }
  }

  return -1;
}

// ==========================================
// SURAT TUGAS: AMBIL SATU SURAT (KELOMPOK
// PER NOMOR SURAT) UNTUK DIEDIT
// ==========================================

function ambilSuratTugasById(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Tugas');

  if (!sheet) {
    throw new Error('Sheet Surat_Tugas tidak ditemukan.');
  }

  const nomorBaris = cariBarisSuratTugasById_(sheet, id);
  if (nomorBaris < 0) {
    throw new Error('Surat tugas tidak ditemukan.');
  }

  const peta = petaKolomSuratTugas_(sheet);
  const jumlahKolom = peta.jumlahKolom;
  const barisSumber = sheet
    .getRange(nomorBaris, 1, 1, jumlahKolom)
    .getDisplayValues()[0];

  const idxNomor = peta.indeks['Nomor_Surat'];
  const nomorSurat = String(barisSumber[idxNomor - 1] || '').trim();

  function nilai(baris, nama) {
    if (!peta.indeks[nama]) {
      return '';
    }
    return String(baris[peta.indeks[nama] - 1] || '').trim();
  }

  const barisTerakhir = sheet.getLastRow();
  const semuaBaris = barisTerakhir >= 2
    ? sheet.getRange(2, 1, barisTerakhir - 1, jumlahKolom).getDisplayValues()
    : [];

  // Kumpulkan semua pegawai yang berbagi nomor surat yang sama.
  const pegawai = [];
  semuaBaris.forEach(function(baris) {
    if (String(baris[0] || '').trim() === '') {
      return;
    }
    if (String(baris[idxNomor - 1] || '').trim() !== nomorSurat) {
      return;
    }
    pegawai.push({
      id: String(baris[0] || '').trim(),
      namaPegawai: nilai(baris, 'Nama_Pegawai'),
      nip: nilai(baris, 'NIP'),
      jabatan: nilai(baris, 'Jabatan')
    });
  });

  return {
    id: String(barisSumber[0] || '').trim(),
    nomorSurat: nomorSurat,
    nomorSuratAwal: nomorSurat,
    tanggalSurat: nilai(barisSumber, 'Tanggal_Surat'),
    dasarSurat: nilai(barisSumber, 'Dasar_Surat'),
    tujuanTugas: nilai(barisSumber, 'Tujuan_Tugas'),
    hari: nilai(barisSumber, 'Hari'),
    tanggalMulai: nilai(barisSumber, 'Tanggal_Mulai'),
    tanggalSelesai: nilai(barisSumber, 'Tanggal_Selesai'),
    waktu: nilai(barisSumber, 'Waktu'),
    tempat: nilai(barisSumber, 'Tempat'),
    keperluan: nilai(barisSumber, 'Keperluan'),
    penandatangan: nilai(barisSumber, 'Penandatangan'),
    jabatanPenandatangan: nilai(barisSumber, 'Jabatan_Penandatangan'),
    tembusan: nilai(barisSumber, 'Tembusan'),
    linkDokumen: nilai(barisSumber, 'Link_Dokumen'),
    status: nilai(barisSumber, 'Status'),
    pegawai: pegawai
  };
}

// ==========================================
// SURAT TUGAS: SIMPAN (BUAT / EDIT)
// Satu baris per pegawai, dikelompokkan oleh
// Nomor_Surat. Baris kelompok lama diganti
// saat edit agar tidak ada duplikasi.
// ==========================================

function simpanSuratTugas(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Tugas');

  if (!sheet) {
    throw new Error('Sheet Surat_Tugas tidak ditemukan.');
  }

  const peta = pastikanKolomSuratTugas_(sheet);

  if (!form.nomorSurat || !form.tanggalSurat) {
    throw new Error('Nomor surat dan tanggal surat wajib diisi.');
  }

  const daftarPegawai = (form.pegawai || []).filter(function(p) {
    return String(p.namaPegawai || '').trim() !== '';
  });

  if (!daftarPegawai.length) {
    throw new Error('Minimal satu pegawai yang ditugasi wajib diisi.');
  }

  // Nomor acuan untuk mencari kelompok lama (pakai nomor awal saat edit).
  const nomorCari = String(form.nomorSuratAwal || form.nomorSurat || '').trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const barisTerakhir = sheet.getLastRow();
    const nomorAcuan = peta.indeks['Nomor_Surat'];
    const barisUntukHapus = [];

    if (barisTerakhir >= 2) {
      const kolomNomor = sheet
        .getRange(2, nomorAcuan, barisTerakhir - 1, 1)
        .getDisplayValues();

      for (let i = 0; i < kolomNomor.length; i++) {
        if (String(kolomNomor[i][0] || '').trim() === nomorCari) {
          barisUntukHapus.push(i + 2);
        }
      }
    }

    // Hapus baris kelompok lama dari bawah agar indeks tidak bergeser.
    for (let i = barisUntukHapus.length - 1; i >= 0; i--) {
      sheet.deleteRow(barisUntukHapus[i]);
    }

    const jumlahKolom = peta.jumlahKolom;
    const dataBaru = [];

    daftarPegawai.forEach(function(p) {
      const baris = [];
      for (let i = 0; i < jumlahKolom; i++) {
        baris.push('');
      }

      baris[0] = 'ST-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      baris[1] = form.nomorSurat;
      baris[2] = form.tanggalSurat;
      baris[3] = String(p.namaPegawai || '').trim();
      baris[4] = String(p.nip || '').trim();
      baris[5] = String(p.jabatan || '').trim();
      baris[6] = form.tujuanTugas || '';
      baris[7] = form.tanggalMulai || '';
      baris[8] = form.tanggalSelesai || '';
      baris[9] = form.keperluan || '';
      baris[10] = form.penandatangan || '';
      baris[11] = form.linkDokumen || '';
      baris[12] = form.status || 'Draft';

      tulisKolomTugas_(baris, peta, 'Dasar_Surat', form.dasarSurat || '');
      tulisKolomTugas_(baris, peta, 'Hari', form.hari || '');
      tulisKolomTugas_(baris, peta, 'Waktu', form.waktu || '');
      tulisKolomTugas_(baris, peta, 'Tempat', form.tempat || '');
      tulisKolomTugas_(baris, peta, 'Jabatan_Penandatangan', form.jabatanPenandatangan || '');
      tulisKolomTugas_(baris, peta, 'Tembusan', form.tembusan || '');

      dataBaru.push(baris);
    });

    const barisAwal = Math.max(sheet.getLastRow() + 1, 2);
    sheet
      .getRange(barisAwal, 1, dataBaru.length, jumlahKolom)
      .setValues(dataBaru);

    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: 'Surat tugas berhasil disimpan (' + dataBaru.length + ' pegawai ditugaskan).',
      id: dataBaru[0][0],
      nomorSurat: form.nomorSurat,
      jumlahPegawai: dataBaru.length
    };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// SURAT TUGAS: HAPUS SATU BARIS PEGAWAI
// ==========================================

function hapusSuratTugas(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Tugas');

  if (!sheet) {
    throw new Error('Sheet Surat_Tugas tidak ditemukan.');
  }

  const idCari = String(id || '').trim();
  if (!idCari) {
    throw new Error('ID surat tugas wajib disertakan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const nomorBaris = cariBarisSuratTugasById_(sheet, idCari);
    if (nomorBaris < 0) {
      throw new Error('Surat tugas tidak ditemukan atau sudah dihapus.');
    }

    sheet.deleteRow(nomorBaris);
    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: 'Surat tugas berhasil dihapus.',
      id: idCari
    };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// HELPER UMUM UNTUK SHEET DOKUMEN TERKAIT
// ==========================================

// Mencari baris berdasarkan ID (kolom A). Dipakai lintas sheet terkait.
function cariBarisTugasById_(sheet, id) {
  const idCari = String(id || '').trim();
  if (!idCari) return -1;

  const barisTerakhir = sheet.getLastRow();
  if (barisTerakhir < 2) return -1;

  const ids = sheet.getRange(2, 1, barisTerakhir - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === idCari) {
      return i + 2;
    }
  }
  return -1;
}

// Menulis nilai ke kolom berdasarkan nama header (1-based).
function tulisKolom_(baris, peta, nama, nilai) {
  if (!peta.indeks[nama]) return;
  baris[peta.indeks[nama] - 1] = nilai;
}

// Membaca nilai kolom berdasarkan nama header, dengan indeks cadangan.
function bacaKolom_(row, peta, nama, cadangan) {
  if (peta.indeks[nama]) {
    return String(row[peta.indeks[nama] - 1] || '').trim();
  }
  if (cadangan >= 0 && cadangan < row.length) {
    return String(row[cadangan] || '').trim();
  }
  return '';
}

// Menambahkan kolom tambahan di akhir sheet jika belum ada (pola aman).
function pastikanKolomTambahan_(sheet, daftarKolom, minimalKolom) {
  function peta() {
    const jumlahKolom = Math.max(sheet.getLastColumn(), minimalKolom);
    const headers = sheet.getRange(1, 1, 1, jumlahKolom).getDisplayValues()[0];
    const indeks = {};
    headers.forEach(function(nama, i) {
      if (nama) indeks[nama] = i + 1;
    });
    return { indeks: indeks, jumlahKolom: jumlahKolom };
  }

  daftarKolom.forEach(function(nama) {
    const p = peta();
    if (p.indeks[nama]) return;
    const kolom = Math.max(sheet.getLastColumn(), minimalKolom) + 1;
    sheet.getRange(1, kolom)
      .setValue(nama)
      .setFontWeight('bold')
      .setBackground('#0f766e')
      .setFontColor('#ffffff');
  });
}

// ==========================================
// SURAT TUGAS SISWA (mirror Surat Tugas untuk
// siswa: Nama/No_Induk/Kelas). Satu baris per
// siswa, dikelompokkan oleh Nomor_Surat.
// ==========================================

var KOLOM_TUGAS_SISWA_TAMBAHAN = [
  'Dasar_Surat', 'Nomor_Dasar', 'Tanggal_Dasar', 'Perihal_Dasar',
  'Hari', 'Waktu', 'Tempat'
];

function petaKolomTugasSiswa_(sheet) {
  const jumlahKolom = Math.max(sheet.getLastColumn(), 13);
  const headers = sheet.getRange(1, 1, 1, jumlahKolom).getDisplayValues()[0];
  const indeks = {};
  headers.forEach(function(nama, i) {
    if (nama) indeks[nama] = i + 1;
  });
  return { indeks: indeks, jumlahKolom: jumlahKolom, headers: headers };
}

function pastikanKolomTugasSiswa_(sheet) {
  pastikanKolomTambahan_(sheet, KOLOM_TUGAS_SISWA_TAMBAHAN, 13);
  return petaKolomTugasSiswa_(sheet);
}

function ambilSuratTugasSiswa() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Tugas_Siswa');

  if (!sheet) {
    throw new Error('Sheet Surat_Tugas_Siswa tidak ditemukan.');
  }

  const barisTerakhir = sheet.getLastRow();
  if (barisTerakhir < 2) return [];

  const peta = petaKolomTugasSiswa_(sheet);
  const jumlahKolom = peta.jumlahKolom;
  const data = sheet.getRange(2, 1, barisTerakhir - 1, jumlahKolom).getDisplayValues();

  return data.filter(function(row) {
    return String(row[0] || '').trim() !== '';
  }).map(function(row) {
    return {
      id: row[0] || '',
      nomorSurat: row[1] || '',
      tanggalSurat: row[2] || '',
      namaSiswa: row[3] || '',
      noInduk: row[4] || '',
      kelas: row[5] || '',
      tujuanTugas: row[6] || '',
      tanggalMulai: row[7] || '',
      tanggalSelesai: row[8] || '',
      keperluan: row[9] || '',
      penandatangan: row[10] || '',
      linkDokumen: row[11] || '',
      status: row[12] || '',
      dasarSurat: bacaKolom_(row, peta, 'Dasar_Surat', -1),
      hari: bacaKolom_(row, peta, 'Hari', -1),
      waktu: bacaKolom_(row, peta, 'Waktu', -1),
      tempat: bacaKolom_(row, peta, 'Tempat', -1)
    };
  });
}

function ambilSuratTugasSiswaById(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Tugas_Siswa');

  if (!sheet) {
    throw new Error('Sheet Surat_Tugas_Siswa tidak ditemukan.');
  }

  const nomorBaris = cariBarisTugasById_(sheet, id);
  if (nomorBaris < 0) {
    throw new Error('Surat tugas siswa tidak ditemukan.');
  }

  const peta = petaKolomTugasSiswa_(sheet);
  const jumlahKolom = peta.jumlahKolom;
  const barisSumber = sheet.getRange(nomorBaris, 1, 1, jumlahKolom).getDisplayValues()[0];
  const idxNomor = peta.indeks['Nomor_Surat'];
  const nomorSurat = String(barisSumber[idxNomor - 1] || '').trim();

  function nilai(baris, nama) {
    if (!peta.indeks[nama]) return '';
    return String(baris[peta.indeks[nama] - 1] || '').trim();
  }

  const barisTerakhir = sheet.getLastRow();
  const semuaBaris = barisTerakhir >= 2
    ? sheet.getRange(2, 1, barisTerakhir - 1, jumlahKolom).getDisplayValues()
    : [];

  const siswa = [];
  semuaBaris.forEach(function(baris) {
    if (String(baris[0] || '').trim() === '') return;
    if (String(baris[idxNomor - 1] || '').trim() !== nomorSurat) return;
    siswa.push({
      id: String(baris[0] || '').trim(),
      namaSiswa: nilai(baris, 'Nama_Siswa'),
      noInduk: nilai(baris, 'No_Induk'),
      kelas: nilai(baris, 'Kelas')
    });
  });

  return {
    id: String(barisSumber[0] || '').trim(),
    nomorSurat: nomorSurat,
    nomorSuratAwal: nomorSurat,
    tanggalSurat: nilai(barisSumber, 'Tanggal_Surat'),
    dasarSurat: nilai(barisSumber, 'Dasar_Surat'),
    nomorDasar: nilai(barisSumber, 'Nomor_Dasar'),
    tanggalDasar: nilai(barisSumber, 'Tanggal_Dasar'),
    perihalDasar: nilai(barisSumber, 'Perihal_Dasar'),
    tujuanTugas: nilai(barisSumber, 'Tujuan_Tugas'),
    hari: nilai(barisSumber, 'Hari'),
    tanggalMulai: nilai(barisSumber, 'Tanggal_Mulai'),
    tanggalSelesai: nilai(barisSumber, 'Tanggal_Selesai'),
    waktu: nilai(barisSumber, 'Waktu'),
    tempat: nilai(barisSumber, 'Tempat'),
    keperluan: nilai(barisSumber, 'Keperluan'),
    penandatangan: nilai(barisSumber, 'Penandatangan'),
    linkDokumen: nilai(barisSumber, 'Link_Dokumen'),
    status: nilai(barisSumber, 'Status'),
    siswa: siswa
  };
}

function simpanSuratTugasSiswa(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Tugas_Siswa');

  if (!sheet) {
    throw new Error('Sheet Surat_Tugas_Siswa tidak ditemukan.');
  }

  if (!form.nomorSurat || !form.tanggalSurat) {
    throw new Error('Nomor surat dan tanggal surat wajib diisi.');
  }

  const daftarSiswa = (form.siswa || []).filter(function(s) {
    return String(s.namaSiswa || '').trim() !== '';
  });

  if (!daftarSiswa.length) {
    throw new Error('Minimal satu siswa yang ditugasi wajib diisi.');
  }

  const nomorCari = String(form.nomorSuratAwal || form.nomorSurat || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const peta = pastikanKolomTugasSiswa_(sheet);
    const nomorAcuan = peta.indeks['Nomor_Surat'];
    const barisTerakhir = sheet.getLastRow();
    const barisUntukHapus = [];

    if (barisTerakhir >= 2) {
      const kolomNomor = sheet.getRange(2, nomorAcuan, barisTerakhir - 1, 1).getDisplayValues();
      for (let i = 0; i < kolomNomor.length; i++) {
        if (String(kolomNomor[i][0] || '').trim() === nomorCari) {
          barisUntukHapus.push(i + 2);
        }
      }
    }

    for (let i = barisUntukHapus.length - 1; i >= 0; i--) {
      sheet.deleteRow(barisUntukHapus[i]);
    }

    const jumlahKolom = peta.jumlahKolom;
    const dataBaru = [];

    daftarSiswa.forEach(function(s) {
      const baris = [];
      for (let i = 0; i < jumlahKolom; i++) baris.push('');

      baris[0] = 'STS-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      baris[1] = form.nomorSurat;
      baris[2] = form.tanggalSurat;
      baris[3] = String(s.namaSiswa || '').trim();
      baris[4] = String(s.noInduk || '').trim();
      baris[5] = String(s.kelas || '').trim();
      baris[6] = form.tujuanTugas || '';
      baris[7] = form.tanggalMulai || '';
      baris[8] = form.tanggalSelesai || '';
      baris[9] = form.keperluan || '';
      baris[10] = form.penandatangan || '';
      baris[11] = form.linkDokumen || '';
      baris[12] = form.status || 'Draft';

      tulisKolom_(baris, peta, 'Dasar_Surat', form.dasarSurat || '');
      tulisKolom_(baris, peta, 'Nomor_Dasar', form.nomorDasar || '');
      tulisKolom_(baris, peta, 'Tanggal_Dasar', form.tanggalDasar || '');
      tulisKolom_(baris, peta, 'Perihal_Dasar', form.perihalDasar || '');
      tulisKolom_(baris, peta, 'Hari', form.hari || '');
      tulisKolom_(baris, peta, 'Waktu', form.waktu || '');
      tulisKolom_(baris, peta, 'Tempat', form.tempat || '');

      dataBaru.push(baris);
    });

    const barisAwal = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(barisAwal, 1, dataBaru.length, jumlahKolom).setValues(dataBaru);
    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: 'Surat tugas siswa berhasil disimpan (' + dataBaru.length + ' siswa).',
      id: dataBaru[0][0],
      nomorSurat: form.nomorSurat,
      jumlahSiswa: dataBaru.length
    };
  } finally {
    lock.releaseLock();
  }
}

function hapusSuratTugasSiswa(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Surat_Tugas_Siswa');

  if (!sheet) {
    throw new Error('Sheet Surat_Tugas_Siswa tidak ditemukan.');
  }

  const idCari = String(id || '').trim();
  if (!idCari) {
    throw new Error('ID surat tugas siswa wajib disertakan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const nomorBaris = cariBarisTugasById_(sheet, idCari);
    if (nomorBaris < 0) {
      throw new Error('Surat tugas siswa tidak ditemukan atau sudah dihapus.');
    }
    sheet.deleteRow(nomorBaris);
    SpreadsheetApp.flush();
    return { sukses: true, pesan: 'Surat tugas siswa berhasil dihapus.', id: idCari };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MESIN CRUD BARIS UMUM
// Dipakai SPPD, SPPD Lembar 2, Nota Dinas, dan
// Daftar Penyerahan (pegawai & siswa) agar tidak
// menduplikasi fungsi yang mirip.
// Setiap sheet: 1 baris = 1 record, kolom ID di A.
// ==========================================

function ambilBarisDokumen_(namaSheet, kolom) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(namaSheet);

  if (!sheet) {
    throw new Error('Sheet ' + namaSheet + ' tidak ditemukan.');
  }

  const barisTerakhir = sheet.getLastRow();
  if (barisTerakhir < 2) return [];

  const jumlahKolom = Math.max(sheet.getLastColumn(), kolom.length);
  const data = sheet.getRange(2, 1, barisTerakhir - 1, jumlahKolom).getDisplayValues();

  return data.reverse().filter(function(row) {
    return String(row[0] || '').trim() !== '';
  }).map(function(row) {
    const obj = { id: row[0] || '' };
    for (let i = 1; i < kolom.length; i++) {
      obj[kolom[i]] = row[i] || '';
    }
    return obj;
  });
}

function ambilBarisDokumenById_(namaSheet, kolom, id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(namaSheet);

  if (!sheet) {
    throw new Error('Sheet ' + namaSheet + ' tidak ditemukan.');
  }

  const nomorBaris = cariBarisTugasById_(sheet, id);
  if (nomorBaris < 0) {
    throw new Error('Data pada ' + namaSheet + ' tidak ditemukan.');
  }

  const jumlahKolom = Math.max(sheet.getLastColumn(), kolom.length);
  const row = sheet.getRange(nomorBaris, 1, 1, jumlahKolom).getDisplayValues()[0];

  const obj = { id: row[0] || '' };
  for (let i = 1; i < kolom.length; i++) {
    obj[kolom[i]] = row[i] || '';
  }
  return obj;
}

function simpanBarisDokumen_(namaSheet, kolom, data, prefixId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(namaSheet);

  if (!sheet) {
    throw new Error('Sheet ' + namaSheet + ' tidak ditemukan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const jumlahKolom = Math.max(sheet.getLastColumn(), kolom.length);
    const idAda = String(data.id || '').trim();
    let nomorBaris = -1;

    if (idAda) {
      nomorBaris = cariBarisTugasById_(sheet, idAda);
    }

    let baris;
    if (nomorBaris > 0) {
      baris = sheet.getRange(nomorBaris, 1, 1, jumlahKolom).getValues()[0];
    } else {
      baris = [];
      for (let i = 0; i < jumlahKolom; i++) baris.push('');
      baris[0] = prefixId + '-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    for (let i = 1; i < kolom.length; i++) {
      if (data[kolom[i]] !== undefined) {
        baris[i] = data[kolom[i]];
      }
    }

    if (nomorBaris > 0) {
      sheet.getRange(nomorBaris, 1, 1, jumlahKolom).setValues([baris]);
    } else {
      sheet.appendRow(baris);
    }

    SpreadsheetApp.flush();

    return {
      sukses: true,
      pesan: nomorBaris > 0
        ? 'Data ' + namaSheet + ' berhasil diperbarui.'
        : 'Data ' + namaSheet + ' berhasil disimpan.',
      id: baris[0]
    };
  } finally {
    lock.releaseLock();
  }
}

function hapusBarisDokumen_(namaSheet, id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(namaSheet);

  if (!sheet) {
    throw new Error('Sheet ' + namaSheet + ' tidak ditemukan.');
  }

  const idCari = String(id || '').trim();
  if (!idCari) {
    throw new Error('ID data wajib disertakan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const nomorBaris = cariBarisTugasById_(sheet, idCari);
    if (nomorBaris < 0) {
      throw new Error('Data pada ' + namaSheet + ' tidak ditemukan atau sudah dihapus.');
    }
    sheet.deleteRow(nomorBaris);
    SpreadsheetApp.flush();
    return { sukses: true, pesan: 'Data ' + namaSheet + ' berhasil dihapus.', id: idCari };
  } finally {
    lock.releaseLock();
  }
}

// ---------- SPPD (Surat Perintah Perjalanan Dinas) ----------
var KOLOM_SPPD = [
  'ID_SPPD', 'Nomor_Surat', 'Pengguna_Anggaran', 'Nama_Pegawai', 'NIP',
  'Pangkat_Golongan', 'Jabatan', 'Tingkat_Biaya', 'Maksud', 'Alat_Angkut',
  'Tempat_Berangkat', 'Tempat_Tujuan', 'Lama_Perjalanan', 'Tanggal_Berangkat',
  'Tanggal_Kembali', 'Instansi', 'Mata_Anggaran', 'Keterangan', 'Status'
];
function ambilSPPD() { return ambilBarisDokumen_('SPPD', KOLOM_SPPD); }
function ambilSPPDById(id) { return ambilBarisDokumenById_('SPPD', KOLOM_SPPD, id); }
function simpanSPPD(form) { return simpanBarisDokumen_('SPPD', KOLOM_SPPD, form, 'SPPD'); }
function hapusSPPD(id) { return hapusBarisDokumen_('SPPD', id); }

// ---------- SPPD Lembar ke 2 (catatan perjalanan) ----------
var KOLOM_SPPD_LEMBAR2 = [
  'ID_Lembar', 'Nomor_Surat', 'Nama_Pegawai', 'Urutan',
  'Tiba_Di', 'Berangkat_Dari', 'Ke', 'Tanggal', 'Status'
];
function ambilSPPDLembar2() { return ambilBarisDokumen_('SPPD_Lembar2', KOLOM_SPPD_LEMBAR2); }
function ambilSPPDLembar2ById(id) { return ambilBarisDokumenById_('SPPD_Lembar2', KOLOM_SPPD_LEMBAR2, id); }
function simpanSPPDLembar2(form) { return simpanBarisDokumen_('SPPD_Lembar2', KOLOM_SPPD_LEMBAR2, form, 'SL2'); }
function hapusSPPDLembar2(id) { return hapusBarisDokumen_('SPPD_Lembar2', id); }

// ---------- Nota Dinas (laporan hasil kegiatan) ----------
var KOLOM_NOTA = [
  'ID_Nota', 'Nomor_Surat', 'Kepada_Yth', 'Dari', 'Perihal',
  'Tempat_Kegiatan', 'Tanggal_Pelaksanaan', 'Dasar_Penugasan',
  'Hasil', 'Usul', 'Status'
];
function ambilNotaDinas() { return ambilBarisDokumen_('Nota_Dinas', KOLOM_NOTA); }
function ambilNotaDinasById(id) { return ambilBarisDokumenById_('Nota_Dinas', KOLOM_NOTA, id); }
function simpanNotaDinas(form) { return simpanBarisDokumen_('Nota_Dinas', KOLOM_NOTA, form, 'ND'); }
function hapusNotaDinas(id) { return hapusBarisDokumen_('Nota_Dinas', id); }

// ---------- Daftar Penyerahan Transport (pegawai) ----------
var KOLOM_DP = [
  'ID_Penyerahan', 'Nomor_Surat', 'Nama_Pegawai', 'NIP', 'Golongan',
  'Volume_Hari', 'Tarif', 'Jumlah', 'Status'
];
function ambilDaftarPenyerahan() { return ambilBarisDokumen_('Daftar_Penyerahan', KOLOM_DP); }
function ambilDaftarPenyerahanById(id) { return ambilBarisDokumenById_('Daftar_Penyerahan', KOLOM_DP, id); }
function simpanDaftarPenyerahan(form) { return simpanBarisDokumen_('Daftar_Penyerahan', KOLOM_DP, form, 'DP'); }
function hapusDaftarPenyerahan(id) { return hapusBarisDokumen_('Daftar_Penyerahan', id); }

// ---------- Daftar Penyerahan Transport Siswa ----------
var KOLOM_DPS = [
  'ID_Penyerahan', 'Nomor_Surat', 'Nama_Siswa', 'Kelas',
  'Volume_Hari', 'Tarif', 'Jumlah', 'Status'
];
function ambilDaftarPenyerahanSiswa() { return ambilBarisDokumen_('Daftar_Penyerahan_Siswa', KOLOM_DPS); }
function ambilDaftarPenyerahanSiswaById(id) { return ambilBarisDokumenById_('Daftar_Penyerahan_Siswa', KOLOM_DPS, id); }
function simpanDaftarPenyerahanSiswa(form) { return simpanBarisDokumen_('Daftar_Penyerahan_Siswa', KOLOM_DPS, form, 'DPS'); }
function hapusDaftarPenyerahanSiswa(id) { return hapusBarisDokumen_('Daftar_Penyerahan_Siswa', id); }