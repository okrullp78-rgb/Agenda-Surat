# Analisis File Referensi: SURAT TUGAS BARU 2025.xlsm

**Tanggal analisis:** 14 Oktober 2026
**Metode:** Ekstraksi arsip OOXML (.xlsm = ZIP) ke folder sementara, lalu parsing XML
(`workbook.xml`, `sharedStrings.xml`, `worksheets/sheetN.xml`, relasi, defined names)
dengan skrip Node.js. Sel-sel kunci diverifikasi ulang dari XML mentah per baris.
**Status:** File asli TIDAK diubah. Analisis murni baca.

---

## 1. Identitas File

- Nama: `SURAT TUGAS BARU 2025.xlsm` (776.372 byte, macro-enabled)
- Asal path penyimpanan terakhir: `D:\Data Sekolah\Surat Tugas\`
- Berisi: 11 worksheet, macro VBA (`vbaProject.bin`), 12 gambar (logo sekolah),
  1 external link ke add-in `D:\Master\Fungsi Terbilang - Info Menarik\Terbilang_Tertanggal.xlam`

## 2. Daftar Sheet (11)

| # | Sheet | File | Peran |
|---|---|---|---|
| 1 | Isian | sheet1 | Formulir input utama + 2 database (pegawai & siswa) |
| 2 | SPPD | sheet2 | Kerja internal penyusunan SPPD |
| 3 | Arsip | sheet3 | Tabel arsip polos (hanya header), diisi macro |
| 4 | **Surat Tugas** | sheet4 | **Format akhir Surat Tugas pegawai (print area A1:I64)** |
| 5 | Surat Tugas Siswa | sheet5 | Format akhir Surat Tugas siswa (mirror, untuk siswa) |
| 6 | SPPD Print | sheet6 | Format cetak SPPD |
| 7 | SPPD Lembar ke 2 | sheet7 | SPPD lanjutan |
| 8 | Nota Dinas | sheet8 | Nota dinas laporan hasil kegiatan |
| 9 | Daftar Penyerahan | sheet9 | Daftar penyerahan transport pegawai |
| 10 | Daftar Penyerahan Siswa | sheet10 | Daftar penyerahan siswa |
| 11 | TMP_SPPD | sheet11 | Template SPPD (struktur identik dengan Surat Tugas) |

## 3. Alur Data Antar-Sheet

```
                 ┌──────────────────────────────────┐
                 │  ISIAN (satu-satunya sheet input) │
                 │  blok input pegawai + kegiatan    │
                 │  + database Dapeg & Daf_siswa     │
                 └───────────────┬──────────────────┘
   ┌─────────────┬───────────────┼───────────────┬─────────────────┐
   ▼             ▼               ▼               ▼                 ▼
Surat Tugas   SPPD Print   SPPD Lembar 2   Nota Dinas    Daftar Penyerahan
(format akhir) (baca Isian  (baca Isian)    (baca Isian   (baca Isian +
(86 ref)       + Surat      (15 ref)        + Surat        Surat Tugas)
               Tugas)                        Tugas)         (47+1 ref)
   │
   └─► diarsipkan ke sheet "Arsip" + diekspor PDF oleh macro VBA (GeneratePDF_Arsip)
```

Jumlah referensi antar-sheet (dihitung dari semua rumus):
- Surat Tugas → Isian x86
- SPPD → Isian x40; SPPD Print → Isian x7 + Surat Tugas x4
- SPPD Lembar ke 2 → Isian x15; Nota Dinas → Isian x6 + Surat Tugas x1
- Daftar Penyerahan → Isian x47 + Surat Tugas x1
- Arsip → tidak ada rumus (diisi oleh macro)
- Surat Tugas Siswa & Daftar Penyerahan Siswa → masing-masing >5.000 ref ke Isian

## 4. Database Internal (Defined Names)

| Nama | Rentang | Isi |
|---|---|---|
| `Dapeg` | `Isian!$Q$2:$X$66` | Database pegawai, **60 baris data** (baris 3–62) |
| `Daf_siswa` | `Isian!$Z$2:$AE$1036` | Database siswa, **1.034 baris data** (baris 3–1036) |
| `Dapeg` (duplikat) | `SPPD!$L$2:$S$66`, `Isian!$Z$2:$AE$1036` | Duplikasi nama untuk sheet lain |

**Struktur kolom Dapeg (Isian Q..X):**
Q=No (No DUK), R=Nama, S=NIP, T=NUPTK, U=Gol, V=Pangkat, W=Jabatan, X=Mata Pelajaran

**Struktur kolom Daf_siswa (Isian Z..AE):**
Z=No (No Induk), AA=Induk Sekolah, AB=Induk Nasional, AC=Nama, AD=L/P, AE=Kelas

## 5. Formulir Input Sheet "Isian" (Yang Diisi Operator)

**Blok kop surat & dasar:**
- B2 = No Surat Tugas (nomor urut, contoh: 437)
- B3 = Dasar Surat (contoh: "Pangkalan SMK Manangga Pratama Tasikmalaya")
- B4 = Nomor dasar surat
- B5 = Tanggal dasar surat (teks, contoh: "1 Mei 2025")
- B6 = Perihal dasar surat

**Blok daftar pegawai yang ditugasi (maks 15 orang):**
- B8:B22 = No DUK pegawai (input utama; nama tampil otomatis di C8:C22 via VLOOKUP)
- G8:G22 = Jumlah transport per orang (dipakai sheet Daftar Penyerahan)
- E8:E22 = nomor urut (rumus otomatis)

**Blok kegiatan:**
- B23 = Nama Kegiatan
- B25 = Tanggal kegiatan (tanggal Excel; serial 45829 = **21 Juni 2025**)
- D25 = teks tanggal tambahan (opsional, saat ini kosong)
- E25 = tanggal akhir (opsional / rentang "s.d.")
- B26 = Waktu (contoh: "07.00 s.d Selesai")
- B27 = Tempat
- H25 = Lama perjalanan / jumlah hari (dirujuk formula "terbilang"; **saat ini kosong di file**)

**Blok daftar siswa (baris 30 ke bawah):** B31 dst = No Duk siswa; nama tampil
otomatis via `Daf_siswa`.

**Rumus otomatis di Isian:**
- C8:C22 = `IF(ISBLANK(Bn),"",VLOOKUP(Bn,Dapeg,2,FALSE))` — nama pegawai
- C4 = `COUNT(B8:B22)` — jumlah orang
- B24/E24 = `IF(DAY(Bn),CHOOSE(MOD(Bn,7)+1,"Sabtu",...,"Jumat"),0)` — hari dari tanggal
- I25 = `"( "&[1]!terbilang(H25)&" ) Hari"` — jumlah hari terbilang
  (butuh add-in eksternal `Terbilang_Tertanggal.xlam`)
