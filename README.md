# Perbandingan Data Realisasi Impor

Aplikasi Streamlit untuk membandingkan data realisasi impor antara file tarikan sistem dengan file data Anda berdasarkan NO. PIB, serta mengecek NO. INVOICE di file Bahan Tambahan Obat dan Bahan Kimia.

## Fitur

- Membandingkan NO. PIB antara File Tarikan dan File Data Anda
- Mengidentifikasi data tarikan yang belum ada di file Anda
- Mengecek NO. INVOICE di file Bahan Tambahan Obat (terpisah)
- Mengecek NO. INVOICE di file Bahan Kimia (terpisah)
- Mendukung NO. INVOICE dengan separator koma (,) dan titik koma (;)
- Download hasil perbandingan dalam format Excel

## Cara Menjalankan di Lokal

### Prasyarat
- Python 3.8 atau lebih baru
- pip (Python package manager)

### Langkah-langkah

1. Clone repository ini:
   ```bash
   git clone https://github.com/USERNAME/NAMA-REPO.git
   cd NAMA-REPO
   ```

2. Install dependencies:
   ```bash
   pip install streamlit pandas openpyxl
   ```

3. Jalankan aplikasi:
   ```bash
   streamlit run app.py
   ```

4. Buka browser dan akses `http://localhost:8501`

## Cara Penggunaan

1. Upload **File Tarikan** (data hasil tarikan dari sistem)
2. Upload **File Data Anda** (data yang ingin dibandingkan berdasarkan NO. PIB)
3. Upload **File Invoice Bahan Tambahan Obat** (opsional)
4. Upload **File Invoice Bahan Kimia** (opsional)
5. Klik tombol **Bandingkan Data**
6. Lihat hasil perbandingan dan download jika diperlukan

## Format File

- File harus berformat Excel (.xlsx atau .xls)
- File Tarikan dan File Data Anda harus memiliki kolom NO. PIB
- File Invoice harus memiliki kolom NO. INVOICE
