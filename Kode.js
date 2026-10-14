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
    ['Zona_Waktu', 'Asia/Jakarta', 'Zona waktu aplikasi'],
    ['Alamat_Sekolah', '', 'Alamat resmi sekolah (isi jika sudah tersedia)'],
    ['Telepon_Sekolah', '', 'Nomor telepon sekolah (isi jika sudah tersedia)'],
    ['Kota_Surat', '', 'Kota pada tanggal surat (isi jika sudah tersedia)'],
    ['Nama_Kepala_Sekolah', '', 'Nama kepala sekolah (isi jika sudah tersedia)'],
    ['Jabatan_Kepala_Sekolah', '', 'Jabatan pejabat pada kop/tanda tangan'],
    ['NIP_Kepala_Sekolah', '', 'NIP kepala sekolah (isi jika sudah tersedia)']
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
    kota: '',
    namaKepala: '',
    jabatanKepala: '',
    nipKepala: ''
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

  return hasil;
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