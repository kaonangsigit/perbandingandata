# 📊 Aplikasi Multi-Tool BPOM

<div align="center">

![Python](https://img.shields.io/badge/Python-3.11+-blue?logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.53-red?logo=streamlit&logoColor=white)
![License](https://img.shields.io/badge/License-Internal-gray)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac%20%7C%20Linux-lightgray)

**Aplikasi web berbasis Streamlit untuk membantu pekerjaan sehari-hari staf BPOM.**
Mulai dari perbandingan data impor, pengecekan HS Code, absensi, notulen rapat, hingga edit PDF — semua dalam satu aplikasi.

[🚀 Cara Instalasi](#-instalasi-lokal) · [✨ Fitur](#-fitur-lengkap) · [📋 Persyaratan](#-persyaratan)

</div>

---

## ✨ Fitur Lengkap

### 📋 1. Perbandingan Data Realisasi Impor
> Bandingkan dua file Excel data impor secara otomatis

- Upload **File Tarikan** (sistem) dan **File Data Anda**
- Pilih kolom pembanding secara fleksibel
- Hasil ditampilkan dengan **highlight warna** — cocok, tidak cocok, hanya ada di satu file
- Export hasil perbandingan ke Excel
- Mendukung multiple file tarikan sekaligus

---

### 💊 2. Cek HS Code Obat (INSW Otomatis)
> Scraping otomatis ke website INSW untuk cek regulasi per HS Code

- Upload file BPS Excel — **auto-deteksi semua prefix chapter**
- Filter chapter yang ingin dicek (28, 29, 30, 31, dst.)
- Deteksi regulasi secara otomatis:
  - ✅ **Lartas Border** & **Tata Niaga Post Border**
  - ✅ **Lartas Ekspor**
  - ✅ **Komoditi Obat, Narkotika, Psikotropika** (BPOM)
- Kolom **Jenis** terpisah: `IMPOR` / `EKSPOR` / `IMPOR & EKSPOR`
- **Color-coded**: Hijau (Obat), Pink (Impor & Ekspor), Biru (Impor), Kuning (Ekspor)
- Export ke Excel dengan sheet terpisah: Regulasi Impor, Ekspor, Terkait Obat
- Dual format pencarian HS Code (8 digit & format titik)

---

### 📈 3. Analisis Data
> Visualisasi data impor dalam bentuk grafik

- Grafik batang & pie chart interaktif
- Analisis distribusi komoditas
- Filter dan drill-down data

---

### 👤 4. Cek Petugas Loket S2
> Auto-fill nama petugas dan skor kepuasan dari Form Konsultasi

- Upload file **Loket S2** (format pivot) + **Form Konsultasi**
- Auto-deteksi header kolom (tanggal, nama, email)
- **Auto-fill Petugas** dari email + fallback nama pendek (misal: "Verda" → "Verda Dereviana Praningtyas")
- **Auto-fill Skor**: Sangat Puas=2, Puas=1, Tidak Puas=0
- Color-coded status: 🟢 Cocok · 🔵 Otomatis Terisi · 🔴 Tidak Cocok · 🟡 Kosong · ⚪ Tidak Ada di Form
- Export Excel per kategori status

---

### 📋 5. Cek Kehadiran
> Bandingkan daftar hadir dengan roster pegawai

- Auto-deteksi kolom Nama, Jabatan, Kehadiran, Waktu
- **Smart name matching** — strip gelar/titel untuk akurasi tinggi
- Tampilkan siapa yang **Hadir** dan **Tidak Hadir**
- Export Excel dengan sheet "Tidak Hadir" terpisah

---

### 🏢 6. Analisis Importir (AI)
> Klasifikasi bisnis importir secara otomatis menggunakan AI

- Upload Excel importir → pilih kolom nama/produk
- AI mengklasifikasikan: **CEK** (obat/kosmetik/OT/pangan) atau **NOM** (non-BPOM)
- Batch processing 25 importir per request + retry otomatis
- Color-coded: 🟢 NOM · 🟡 CEK
- Export Excel dengan ringkasan statistik
- *Membutuhkan API key OpenAI atau Groq*

---

### 🔗 7. Gabung Data Excel
> Merge dua file Excel tanpa menimpa data yang sudah ada

- Buka **File Utama** dengan format asli terjaga (filter, warna, font, lebar kolom)
- Isi sel kosong dari **File Pelengkap** — data lama **tidak pernah tertimpa**
- Pilih kolom tertentu atau merge semua sekaligus
- Atur range baris dan mode overwrite opsional
- Progress bar real-time per kolom

---

### 📝 8. Notulen Rapat
> Generate dokumen notulen rapat resmi (.docx) otomatis

- Format notulen pemerintah resmi (Arial 12pt, A4, justified)
- Bagian lengkap: Informasi Rapat, Pendahuluan, Pembahasan, Dokumentasi Foto, Kesimpulan, Notulis
- Upload foto dokumentasi hingga **35 foto** (2 per baris)
- Pembahasan dinamis: tambah/hapus topik speaker
- **🔥 Muat Rapat 7 Mei 2026** — isi otomatis 12 topik, 10 kesimpulan, 12 tindak lanjut (tanpa API!)
- Upload ringkasan **tldv.io** (.txt) untuk rapat multi-sesi
- Generate dengan AI (opsional: Groq / OpenAI)
- Download langsung sebagai `.docx`

---

### 🎓 9. Laporan Magang BPOM
> Generate laporan magang resmi BPOM (.docx)

- Template resmi sesuai format BPOM
- Isi data diri, unit penempatan, kegiatan magang
- Generate otomatis dokumen siap cetak
- Download sebagai `.docx`

---

### 📄 10. Editor PDF
> Edit PDF langsung di browser — tidak perlu software tambahan!

| Fitur | Keterangan |
|-------|-----------|
| 📎 **Gabung PDF** | Satukan beberapa PDF menjadi satu file |
| ✂️ **Pisah Halaman** | Ambil halaman tertentu (misal: 1-3, 5, 7-10) |
| 🗑️ **Hapus Halaman** | Hapus halaman tertentu, sisanya disimpan |
| 🔄 **Putar Halaman** | Putar 90°, 180°, atau 270° |
| 💧 **Watermark Teks** | Tambah watermark dengan warna, ukuran, sudut, transparansi kustom |
| 🔢 **Nomor Halaman** | Tambah nomor halaman dengan posisi & format bebas |
| 🗜️ **Kompres PDF** | Optimalkan ukuran file PDF |

---

## 🛠 Persyaratan

- **Python 3.11** atau lebih baru
- pip (sudah termasuk dalam instalasi Python)
- Koneksi internet (untuk fitur Cek HS Code & AI)

---

## 🚀 Instalasi Lokal

### Langkah 1 — Clone repository
```bash
git clone https://github.com/kaonangsigit/perbandingandata.git
cd perbandingandata
```

### Langkah 2 — Buat virtual environment
```bash
python -m venv venv
```

Aktifkan:
```bash
# Windows
venv\Scripts\activate

# Mac / Linux
source venv/bin/activate
```

### Langkah 3 — Install semua library
```bash
pip install -r requirements.txt
```

### Langkah 4 — Install Chromium (untuk fitur Cek HS Code)
```bash
playwright install chromium
```

### Langkah 5 — Jalankan aplikasi
```bash
streamlit run app.py --server.port 5000
```

Buka browser ke **http://localhost:5000** ✅

---

## 📦 Library yang Digunakan

| Library | Kegunaan |
|---------|---------|
| `streamlit` | Framework web aplikasi |
| `pandas` | Baca & proses file Excel |
| `openpyxl` | Baca/tulis Excel dengan format terjaga |
| `pdfplumber` | Baca konten PDF |
| `pypdf` | Manipulasi PDF (gabung, pisah, putar, dll.) |
| `reportlab` | Generate elemen PDF (watermark, nomor halaman) |
| `playwright` | Scraping INSW via headless browser |
| `python-docx` | Generate file Word (.docx) |
| `openai` | Integrasi AI untuk klasifikasi & notulen |
| `tenacity` | Retry otomatis saat rate limit API |
| `matplotlib` | Grafik dan visualisasi data |

---

## 📝 Catatan Penting

- **Fitur AI** (Analisis Importir & Notulen dengan AI) membutuhkan API key — masukkan langsung di tab masing-masing
- **Cek HS Code** membutuhkan Chromium dan koneksi internet ke insw.go.id
- **Notulen Rapat** bisa digunakan penuh **tanpa API key** menggunakan data pre-loaded
- Semua file yang diupload diproses di memori lokal — **tidak ada data dikirim ke server luar** (kecuali saat menggunakan fitur AI)
- Untuk update terbaru: `git pull origin main`

---

<div align="center">
  <sub>Dibuat untuk Direktorat Pengawasan KMEI · BPOM RI</sub>
</div>
<div align="center">
  <sub>BY Kaonang</sub>
</div>
