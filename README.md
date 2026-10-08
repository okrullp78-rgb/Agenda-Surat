# Agenda-Surat

Aplikasi web lokal untuk arsip surat tugas dan rekap agenda surat masuk/keluar SMP Negeri 13 Tasikmalaya.

## Menjalankan

Pasang dependensi dan jalankan server pengembangan:

```sh
npm install
npm run dev
```

Alternatifnya, buka `index.html` langsung di browser. Tidak perlu server untuk penggunaan lokal sederhana.

## Fitur

- Nomor surat tugas diisi manual agar nomor lama dapat diarsipkan; arsip surat keluar disusun berdasarkan tanggal surat.
- Arsip surat tugas terhubung otomatis ke agenda surat keluar.
- Template master guru/siswa dapat diunduh sebagai Excel (.xlsx), diisi dengan format tabel, lalu diunggah kembali; CSV tetap didukung. Pilihan nama dapat ditambah beberapa baris dan tarif transport diisi manual per orang.
- Pencatatan surat masuk, pencarian arsip, ekspor agenda ke CSV, serta cetak register arsip.
- Tujuh dokumen tersedia melalui cetak browser: Surat Tugas, Surat Tugas Siswa, SPPD, SPPD Lembar ke-2, Nota Dinas, Daftar Penyerahan Transport, dan Daftar Penyerahan Transport Siswa. Redaksi dan label mengikuti workbook sumber.
- Data tersimpan di penyimpanan lokal browser. Gunakan cadangan JSON di halaman Pengaturan & data untuk memindahkan atau mengamankan arsip.

Catatan: aplikasi ini belum menggunakan server atau basis data bersama. Data pada browser/perangkat yang berbeda tidak tersinkron otomatis.