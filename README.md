# 📊 Perbandingan Data Realisasi Impor

Aplikasi Streamlit untuk membandingkan dan menganalisis data realisasi impor dengan tampilan yang menarik dan mudah digunakan.

![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=Streamlit&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)

## ✨ Fitur Utama

### 📋 Tab Perbandingan Data
- **Perbandingan Fleksibel**: Pilih kolom yang ingin dibandingkan (NO. PIB, NO. INVOICE, atau kolom lainnya)
- **Deteksi Otomatis**: Sistem mendeteksi kolom yang sama di kedua file
- **Pembersihan Data**: Opsi untuk membersihkan data numerik (menghapus karakter ', ", dll)
- **Cek Invoice Terpisah**: 
  - Cek di file Bahan Tambahan Obat
  - Cek di file Bahan Kimia
- **Support Multi-Separator**: Mendukung NO. INVOICE dengan separator koma (,) dan titik koma (;)
- **Download Hasil**: Export hasil perbandingan dalam format Excel

### 📈 Tab Analisis Data
- **Analisis Kolom**: Pilih kolom yang ingin dianalisis (Negara, Jenis Obat, dll)
- **Top N Data**: Tampilkan data terbanyak sesuai kebutuhan (5-50)
- **Visualisasi**:
  - Tabel data dengan persentase
  - Grafik Bar interaktif
  - Grafik Pie dengan legenda
- **Statistik Lengkap**: Total data, nilai unik, data terbanyak
- **Download**:
  - Data analisis dalam format Excel
  - Grafik dalam format PNG untuk laporan

## 🚀 Cara Menjalankan di Lokal

### Prasyarat
- Python 3.8 atau lebih baru
- pip (Python package manager)

### Langkah-langkah

1. **Clone repository ini:**
   ```bash
   git clone https://github.com/USERNAME/NAMA-REPO.git
   cd NAMA-REPO
   ```

2. **Install dependencies:**
   ```bash
   pip install streamlit pandas openpyxl matplotlib
   ```

3. **Jalankan aplikasi:**
   ```bash
   streamlit run app.py
   ```

4. **Buka browser** dan akses `http://localhost:8501`

## 📖 Cara Penggunaan

### Tab Perbandingan Data
1. Upload **File Tarikan** (data hasil tarikan dari sistem)
2. Upload **File Data Anda** (data yang ingin dibandingkan)
3. Pilih **kolom** yang ingin digunakan untuk perbandingan
4. Upload **File Invoice** (opsional):
   - Bahan Tambahan Obat
   - Bahan Kimia
5. Klik tombol **Bandingkan Data**
6. Lihat hasil dan download jika diperlukan

### Tab Analisis Data
1. Upload file Excel yang ingin dianalisis
2. Pilih kolom yang ingin dianalisis
3. Atur jumlah Top N data
4. Klik tombol **Analisis Data**
5. Lihat grafik dan statistik
6. Download hasil dalam format Excel atau PNG

## 📁 Format File

- File harus berformat Excel (`.xlsx` atau `.xls`)
- File Tarikan dan File Data Anda harus memiliki kolom yang sama untuk dibandingkan
- File Invoice harus memiliki kolom NO. INVOICE

## 🛠️ Teknologi

- **Streamlit** - Framework web app
- **Pandas** - Pengolahan data
- **Matplotlib** - Visualisasi grafik
- **OpenPyXL** - Baca/tulis file Excel

## 📝 Lisensi

MIT License

---

<p align="center">
  Dibuat oleh Kaonang S.P
</p>
