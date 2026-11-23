# Mendeley Citation Automation 

[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Platform Windows](https://img.shields.io/badge/Platform-Windows-green.svg)]()
[![Mendeley Cite Compatible](https://img.shields.io/badge/Mendeley%20Cite-Ready-red.svg)](https://www.mendeley.com/reference-management/mendeley-cite)

Tool otomatisasi pintar untuk mengonversi **placeholder sitasi teks** (format apa pun: `NR007`, `S12`, `REF-1`, `[1]`, `[1-5]`, `{A3}`, dll.) di dokumen Word (`.docx`) menjadi **field sitasi resmi Mendeley Cite (`ADDIN CSL_CITATION`)**.

Hasil konversi berupa dokumen baru **`manuscript - Mendeley.docx`** yang langsung terhubung ke Mendeley Cite, serta buku kerja Excel **`laporan_sitasi.xlsx`** yang memuat audit sitasi secara lengkap.

Dokumen Word asli Anda **aman dan tidak pernah ditimpa**.

---

## Daftar Isi
- [Latar Belakang](#latar-belakang)
- [Fitur Utama](#fitur-utama)
- [Prasyarat Sistem](#prasyarat-sistem)
- [Instalasi](#instalasi)
- [Tutorial Lengkap Penggunaan](#tutorial-lengkap-penggunaan)
  - [Langkah 1: Siapkan File Word dan Excel](#langkah-1-siapkan-file-word-dan-excel)
  - [Langkah 2: Jalankan Otomatisasi](#langkah-2-jalankan-otomatisasi)
  - [Langkah 3: Periksa Laporan Excel](#langkah-3-periksa-laporan-excel)
  - [Langkah 4: Tindak Lanjuti Referensi yang Kurang (Jika Ada)](#langkah-4-tindak-lanjuti-referensi-yang-kurang-jika-ada)
  - [Langkah 5: Finalisasi di Microsoft Word](#langkah-5-finalisasi-di-microsoft-word)
- [Penjelasan Laporan Excel (`laporan_sitasi.xlsx`)](#penjelasan-laporan-excel-laporan_sitasixlsx)
- [Opsi Perintah CLI Lanjutan](#opsi-perintah-cli-lanjutan)
- [Struktur Repositori](#struktur-repositori)
- [Troubleshooting & Solusi Masalah](#troubleshooting--solusi-masalah)
- [Lisensi](#lisensi)

---

## Latar Belakang

Sering kali saat menulis naskah ilmiah, artikel jurnal, atau skripsi secara kolaboratif atau melalui proses *systematic literature review* (SLR), sitasi sementara ditulis dalam bentuk kode sederhana seperti `(NR007)`, `[S12]`, atau `[1, 2, 5]`, sementara data pustakanya dicatat di spreadsheet Excel.

Sayangnya, add-in **Mendeley Cite** di Word tidak dapat langsung mengubah teks biasa tersebut menjadi sitasi aktif. Menghubungkannya satu per satu secara manual ke ratusan sitasi di Word membutuhkan waktu berhari-hari dan rawan salah.

**Tool ini menyelesaikan masalah tersebut dalam hitungan detik:**
Secara langsung pada level OpenXML (`word/document.xml`), tool ini menyisipkan struktur CSL field Mendeley yang valid dengan UUID dari database Mendeley Anda.

---

## Fitur Utama

- **Universal Placeholder Support**: Mendeteksi pola apa pun: huruf+angka (`NR001`, `S1`, `REF-01`), angka murni (`[1]`, `(1)`), tanda kurung bulat `()`, siku `[]`, maupun kurawal `{}`.
- **Auto Range Expansion**: Mengurai otomatis format rentang sitasi, seperti `[1-5]` menjadi `1, 2, 3, 4, 5` atau `(NR001-NR004)` menjadi `NR001, NR002, NR003, NR004`.
- **Deteksi Cerdas Kolom Excel**: Otomatis mengenali sheet dan kolom ID, Judul, DOI, Penulis, Tahun, dan Jurnal tanpa perlu mengubah format Excel Anda.
- **Pencocokan Bertingkat ke Mendeley (Multi-Tier Matching)**:
  1. Kecocokan persis nomor **DOI**.
  2. Kecocokan **Judul Persis** (dinormalisasi).
  3. Kecocokan **Fuzzy Judul** ($\ge 90\%$) dengan `difflib`.
- **Pengayaan Metadata via Crossref**: Mengambil data nomor volume, nomor terbitan/issue, dan halaman secara otomatis melalui Crossref API.
- **Audit Komprehensif di Excel**: Menghasilkan `laporan_sitasi.xlsx` 4-sheet dengan indikator warna status jelas (Berhasil, Gagal Tidak Ada di Mendeley, Gagal Tidak Ada di Excel).
- **Generator RIS Otomatis**: Jika ada referensi yang belum dimasukkan ke Mendeley, script otomatis membuatkan `referensi_belum_ada.ris` yang siap di-*drag & drop* ke Mendeley.
- **Aman**: Membaca database Mendeley melalui salinan temporer (tidak bentrok jika aplikasi Mendeley sedang terbuka).

---

## Prasyarat Sistem

1. **Sistem Operasi**: Windows 10 / 11.
2. **Python**: Versi 3.8 ke atas.
3. **Microsoft Word**: Dengan add-in **Mendeley Cite** terpasang.
4. **Mendeley Reference Manager**: Terpasang dan sudah login di komputer yang sama.

---

## Instalasi

1. Clone repositori ini ke komputer Anda:
   ```bash
   git clone https://github.com/Nashhansyah/MendeleycAutomation.git
   cd MendeleycAutomation
   ```

2. Pasang pustaka Python yang dibutuhkan:
   ```bash
   pip install -r requirements.txt
   ```

---

## Tutorial Lengkap Penggunaan

### Langkah 1: Siapkan File Word dan Excel
1. Pastikan Anda memiliki:
   - **1 file naskah Word (`.docx`)** yang berisi placeholder sitasi (misal: `...metode ini terbukti efektif (NR007; NR012).` atau `...menurut [1-3]...`).
   - **1 file Excel (`.xlsx`)** yang memuat daftar referensi tersebut (memiliki kolom ID/Kode, Judul, DOI/Penulis/Tahun).
2. Salin kedua file tersebut ke dalam folder repositori ini.

> [!NOTE]
> Nama file Word dan Excel bebas. Script otomatis memilih file terbaru yang ada di folder.

### Langkah 2: Jalankan Otomatisasi

**Cara 1 (Paling Mudah — Tanpa Terminal):**
- Cukup **double-click berkas `Run.bat`**.

**Cara 2 (Melalui Terminal / Command Prompt):**
```bash
python mendeley_cite.py
```

Tunggu beberapa detik hingga proses selesai:
```text
========================================================================
  WORD  ->  MENDELEY CITATION AUTOMATION (UNIVERSAL EDITION)
========================================================================
  Folder Kerja : C:\MENDELEY AUTOMATE
  Naskah Word  : naskah_penelitian.docx
  Data Excel   : data_referensi.xlsx

[1/6] Membaca naskah dokumen Word ...
[2/6] Menganalisis file Excel & mencocokkan kolom ...
[3/6] Mendeteksi pola placeholder sitasi di Word ...
[4/6] Menyusun metadata referensi ...
[5/6] Mencocokkan dengan library Mendeley Reference Manager ...
[6/6] Memodifikasi XML Word & menyisipkan field Mendeley ...

========================================================================
  HASIL EKSEKUSI:
  * 42 kelompok sitasi berhasil diubah menjadi field Mendeley
  * 30 ID unik sitasi sukses terhubung ke Mendeley
  * 2 ID belum ada di Mendeley : NR015, NR028

  Berkas Hasil Word   : manuscript - Mendeley.docx
  Berkas Laporan Excel: laporan_sitasi.xlsx
========================================================================
```

### Langkah 3: Periksa Laporan Excel
Buka berkas **`laporan_sitasi.xlsx`** yang baru dihasilkan.
- Periksa sheet **Ringkasan** untuk melihat persentase keberhasilan.
- Periksa sheet **Perlu Tindakan** jika ada sitasi yang berstatus Gagal.

### Langkah 4: Tindak Lanjuti Referensi yang Kurang (Jika Ada)
- **Jika ada referensi yang berstatus `GAGAL: TIDAK ADA DI MENDELEY`**:
  Tool otomatis membuat berkas **`referensi_belum_ada.ris`**.
  1. Buka aplikasi **Mendeley Reference Manager**.
  2. Tarik (*drag & drop*) file `referensi_belum_ada.ris` langsung ke jendela Mendeley.
  3. Jalankan kembali `Run.bat`. Sitasi tersebut sekarang akan berstatus `BERHASIL`.
- **Jika ada referensi yang berstatus `GAGAL: TIDAK ADA DI EXCEL`**:
  Kode tersebut ada di Word tetapi belum terdaftar di Excel. Buka file Excel Anda dan tambahkan baris data untuk kode tersebut.

### Langkah 5: Finalisasi di Microsoft Word
1. Buka file hasil: **`manuscript - Mendeley.docx`**.
2. Di Microsoft Word, buka tab **References** (atau **Insert**) lalu buka panel samping **Mendeley Cite**.
3. Di panel Mendeley Cite:
   - Pilih Citation Style yang diinginkan (misal: *APA 7th, IEEE, Nature, Harvard*, dll.).
   - Klik tombol **Update Citations and Bibliography** (atau ikon titik tiga `...` -> *Update from Library*).
   - Seluruh sitasi sementara akan langsung terformat rapi sesuai gaya sitasi yang dipilih.
4. Letakkan kursor di bagian akhir naskah (bab *References / Daftar Pustaka*), klik tombol titik tiga `...` di Mendeley Cite, lalu pilih **Insert Bibliography**.
5. Hapus daftar pustaka lama yang sebelumnya diketik manual agar tidak ganda.

---

## Penjelasan Laporan Excel (`laporan_sitasi.xlsx`)

File Excel laporan dirancang untuk mempermudah audit referensi:

| Sheet | Penjelasan |
|---|---|
| **`Ringkasan`** | Dashboard metrik: Total sitasi di Word, jumlah & persentase berhasil, jumlah gagal (tidak ada di Mendeley vs tidak ada di Excel), berkas yang diproses, dan database Mendeley yang digunakan. |
| **`Detail Sitasi Word`** | Tabel audit lengkap per kode sitasi: ID, Status (`BERHASIL` [Hijau], `GAGAL: TIDAK ADA DI MENDELEY` [Oranye], `GAGAL: TIDAK ADA DI EXCEL` [Kuning]), alasan, jumlah kemunculan, metode pencocokan, UUID Mendeley, metadata artikel, dan cuplikan kalimat paragraf di Word. |
| **`Perlu Tindakan`** | Filter khusus sitasi yang gagal beserta petunjuk penyelesaian langkah-demi-langkah. |
| **`Excel Belum Digunakan`** | Daftar referensi yang ada di file Excel tetapi tidak pernah dikutip di naskah Word. |

---

## Opsi Perintah CLI Lanjutan

Anda dapat menyesuaikan parameter eksekusi melalui argumen baris perintah:

```bash
python mendeley_cite.py [opsi]
```

| Opsi | Contoh Argumen | Keterangan |
|---|---|---|
| `--folder` | `--folder "D:\Riset"` | Menentukan folder tempat file Word & Excel berada |
| `--docx` | `--docx "draft.docx"` | Menentukan file Word naskah tertentu |
| `--workbook` | `--workbook "refs.xlsx"` | Menentukan file Excel referensi tertentu |
| `--out` | `--out "final_mendeley.docx"` | Menentukan nama file Word hasil |
| `--style` | `--style numeric` | Gaya tampilan sementara (`author-year` atau `numeric`) |
| `--id-regex` | `--id-regex "NR\d{3}"` | Memaksa pola regex ID tertentu jika ada pola unik |
| `--id-column` | `--id-column "Study_ID"` | Memaksa nama kolom ID di Excel tertentu |
| `--no-enrich` | *(flag)* | Menonaktifkan pengambilan volume/halaman dari Crossref API |
| `--dry-run` | *(flag)* | Hanya menganalisis dan mencocokkan tanpa menulis file Word baru |

---

## Struktur Repositori

```text
MendeleycAutomation/
├── mendeley_cite.py      # Script Python utama pengolah sitasi & pembuat laporan
├── Run.bat               # Launcher Windows 1-klik
├── config.json           # File konfigurasi default
├── requirements.txt      # Daftar dependensi Python (lxml, pandas, openpyxl)
├── WordMacro.bas         # Modul Macro VBA Word (opsional, untuk eksekusi via Word)
├── crossref_cache.json   # Cache lokal respons Crossref API
├── .gitignore            # Filter berkas sementara & output pengguna
└── README.md             # Dokumentasi & panduan lengkap
```

---

## Troubleshooting & Solusi Masalah

| Kendala | Kemungkinan Penyebab | Solusi |
|---|---|---|
| **Tidak ada file .docx / .xlsx di folder** | File belum ditaruh di folder yang sama | Letakkan 1 file `.docx` dan 1 file `.xlsx` di folder repo ini |
| **0 sitasi terkonversi** | Pola kurung tidak standar atau ID berbeda | Coba jalankan dengan opsi `--id-regex "POLA"` atau periksa kesamaan kode di Word dan kolom Excel |
| **Banyak ID berstatus TIDAK ADA DI MENDELEY** | PDF/entri artikel belum diimpor ke aplikasi Mendeley | Drag file `referensi_belum_ada.ris` langsung ke Mendeley Reference Manager |
| **Sitasi di Word tidak berubah setelah klik Update** | Add-in Mendeley Cite belum login / beda akun | Buka tab Mendeley Cite di Word, pastikan sudah login dengan akun Mendeley yang sama |
| **Database is locked** | Aplikasi Mendeley sedang melakukan sinkronisasi berat | Tutup sementara Mendeley Reference Manager selama 10 detik, lalu jalankan script |

---

## Lisensi

Didistribusikan di bawah Lisensi MIT. Bebas digunakan, dimodifikasi, dan dikembangkan untuk keperluan akademik maupun komersial.
