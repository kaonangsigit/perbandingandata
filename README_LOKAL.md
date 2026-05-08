# Panduan Menjalankan Aplikasi Secara Lokal

## Persyaratan
- Python 3.11 atau lebih baru
- pip

## Langkah-langkah

### 1. Clone repository
```bash
git clone https://github.com/kaonangsigit/perbandingandata.git
cd perbandingandata
```

### 2. Buat virtual environment (disarankan)
```bash
python -m venv venv

# Windows:
venv\Scripts\activate

# Mac/Linux:
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Install Chromium untuk fitur Cek HS Code (Playwright)
```bash
playwright install chromium
```

### 5. Jalankan aplikasi
```bash
streamlit run app.py --server.port 5000
```

Aplikasi akan terbuka otomatis di browser: http://localhost:5000

---

## Fitur yang Tersedia

| Tab | Fitur |
|-----|-------|
| 📋 Perbandingan Data | Bandingkan data impor dua file Excel |
| 💊 Cek HS Code Obat | Cek regulasi HS Code via INSW otomatis |
| 📈 Analisis Data | Grafik dan analisis data impor |
| 👤 Cek Petugas Loket S2 | Auto-fill petugas dari Form Konsultasi |
| 📋 Cek Kehadiran | Bandingkan daftar hadir vs roster pegawai |
| 🏢 Analisis Importir | Klasifikasi importir dengan AI (butuh API key) |
| 🔗 Gabung Data Excel | Merge dua file Excel tanpa menimpa data lama |
| 📝 Notulen Rapat | Generate dokumen notulen rapat (.docx) |
| 🎓 Laporan Magang BPOM | Generate laporan magang (.docx) |
| 📄 Edit PDF | Gabung, pisah, putar, watermark, nomor halaman, kompres PDF |

---

## Catatan

- Fitur **Analisis Importir** membutuhkan API key OpenAI atau Groq — masukkan di tab tersebut.
- Fitur **Cek HS Code Obat** membutuhkan koneksi internet dan Chromium (sudah diinstall di langkah 4).
- Fitur **Notulen Rapat** bisa digunakan tanpa API key (gunakan tombol "Muat Rapat 7 Mei 2026").
- Semua file yang diupload diproses di memori lokal — tidak ada data yang dikirim ke server luar (kecuali fitur AI).
