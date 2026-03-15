# Perbandingan Data Impor - Import Data Comparison Tool

## Overview

This is a Streamlit-based web application for comparing import realization data between two Excel files, and checking HS Code regulations via INSW (Indonesia National Single Window). The tool helps users identify discrepancies between system-pulled data ("File Tarikan") and user-provided data, and check import/export regulations for HS codes from BPS data files.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture
- **Framework**: Streamlit (Python-based web framework)
- **Layout**: Wide layout configuration for better data visualization
- **UI Pattern**: Two-column layout for side-by-side file comparison
- **Language**: Indonesian (Bahasa Indonesia) for all user-facing text

### Data Processing
- **Data Handling**: Pandas for Excel file parsing and data manipulation
- **File Format Support**: Excel files (.xlsx, .xls)
- **Comparison Logic**: Compares data between File Tarikan and File Data Anda based on user-selected columns, with symbol cleaning

### Features
1. **Perbandingan Data** (Tab 1): Compare import realization data between system files and user data
2. **Cek INSW Otomatis** (Tab 2): Primary feature - uses Playwright headless browser to scrape INSW website for each HS code:
   - **Flexible chapter filter**: Auto-detects all chapter prefixes from BPS file, user can select any combination (e.g., 28, 29, 30, 31, or any others)
   - **Import regulation detection**: Lartas Border, Tata Niaga Post Border
   - **Export regulation detection**: Lartas Ekspor
   - **Pharmaceutical classification**: Komoditi INSW (Obat, Narkotika, Psikotropika, etc.)
   - **BPOM regulation** presence
   - **Clear import/export differentiation**: Separate "Jenis" column (IMPOR/EKSPOR/IMPOR & EKSPOR), separate Keterangan Impor/Ekspor columns
   - **Color-coded results**: Green (Obat), Pink (Impor & Ekspor), Blue (Impor), Yellow (Ekspor)
   - **Optimized scraping**: Page reuse with back-navigation, element-based waits instead of fixed delays (~3s per HS code)
   - **Excel download**: Separate sheets for all results, Regulasi Impor, Regulasi Ekspor, Terkait Obat
3. **Analisis Data** (Tab 3): Data analysis features with bar/pie charts

### Application Entry Points
- `app.py`: Main Streamlit application (run with `streamlit run app.py`)
- `main.py`: Placeholder Python entry point (not actively used)

### Design Decisions
1. **Streamlit over Flask/Django**: Chosen for rapid prototyping of data-centric applications with minimal frontend code
2. **Pandas for Data Processing**: Standard choice for Excel/tabular data manipulation in Python
3. **Wide Layout**: Enables side-by-side comparison of datasets which is core to the application's purpose
4. **INSW-only approach**: Removed keyword/AI classification in favor of direct INSW website scraping for accuracy
5. **Background threading for INSW scraping**: Scraping runs in a daemon thread with file-based IPC (`/tmp/insw_{sid}.json`), while main Streamlit thread polls every 3s for progress updates. This prevents WebSocket timeout/disconnection during long scraping sessions. File-based state survives module reimports and process restarts (partial results preserved).
6. **Dual format HS Code search**: Tries plain 8-digit format first, then dotted format (XXXX.XX.XX) if INSW doesn't find it
7. **Heartbeat-based staleness detection**: Thread writes heartbeat timestamp on each update; polling detects stale heartbeats (>60s) to gracefully handle dead threads and show partial results

## External Dependencies

### Python Libraries
- **streamlit**: Web application framework for data apps
- **pandas**: Data manipulation and Excel file reading
- **openpyxl** (implicit): Required by pandas for .xlsx file support
- **playwright**: Headless browser automation for INSW website scraping (requires chromium + mesa-libgbm system dep)

### System Dependencies
- **mesa-libgbm**: Required by Chromium/Playwright (`/nix/store/24w3s75aa2lrvvxsybficn8y3zxd27kp-mesa-libgbm-25.1.0/lib`)
- **Chromium**: Installed via `PLAYWRIGHT_BROWSERS_PATH=/home/runner/.cache/ms-playwright python3 -m playwright install chromium`

### File I/O
- Excel file uploads handled via Streamlit's file_uploader component
- In-memory processing using pandas read_excel
- BPS files read with dtype=str to avoid mixed-type column issues

### External Services
- **INSW Website** (insw.go.id/intr/detail-komoditas): Scraped via Playwright for HS code regulation data
- No database connections
- No authentication systems

### Features
1. **Perbandingan Data** (Tab 1): Compare import realization data between system files and user data
2. **Cek HS Code Obat** (Tab 2): Automated INSW website scraping for HS code regulation checks
3. **Analisis Data** (Tab 3): Data analysis with bar/pie charts
4. **Cek Petugas Loket S2** (Tab 4 in app.py): Auto-fill officer names from Form Konsultasi into Loket S2 data
   - Parses Loket S2 pivot-table Excel format (date headers, name/email pairs, satisfaction levels)
   - Auto-fills empty Petugas from Form Konsultasi by matching email + date, with fallback to name matching
   - Auto-fills Skor from satisfaction level (Sangat Puas=2, Puas=1, Tidak Puas=0)
   - Short name matching (e.g., "Verda" matches "Verda Dereviana Praningtyas")
   - Color-coded statuses: Green (Cocok), Blue (Otomatis Terisi), Red (Tidak Cocok), Yellow (Kosong), Gray (Tidak Ada di Form)
   - Excel export with formatted sheets per status category

5. **Cek Kehadiran** (Tab 5 in app.py): Compare employee roster vs attendance list
   - Auto-detects column headers (Nama, Jabatan, Kehadiran, Waktu)
   - Smart name matching: strips titles/degrees for accurate comparison
   - Color-coded: Green (Hadir), Red (Tidak Hadir)
   - Excel export with separate "Tidak Hadir" sheet

6. **Analisis Importir** (Tab 6 in app.py): AI-powered importer business classification
   - Upload Excel → auto-detect headers → select columns
   - AI classifies each importer as CEK (obat/kosmetik/OT/food) or NOM (non-BPOM)
   - Uses OpenAI AI Integrations (gpt-5-mini) with batch processing (25 per batch)
   - Retry logic with exponential backoff for rate limits
   - Color-coded results: Green (NOM), Yellow (CEK)
   - Excel export with formatted data + summary sheet
   - Session state keyed to file identity to prevent stale results

7. **Gabung Data Excel** (Tab 7 in app.py): Merge data from two Excel files
   - Upload File Utama (with empty cells) and File Pelengkap (with filled data)
   - Auto-detect headers, select specific columns or merge all
   - Option to set row range and overwrite mode
   - Color-coded Excel output with merge statistics report sheet
   - Session state keyed to file identity

## Recent Changes
- 2026-03-15: Added Gabung Data Excel tab for merging two Excel files
- 2026-03-13: Added Analisis Importir tab with AI-powered business classification
- 2026-03-09: Added Cek Kehadiran tab for attendance checking
- 2026-03-09: Added Cek Petugas Loket S2 tab with auto-fill petugas & skor functionality
- 2026-02-09: Refactored INSW scraping to background threading with thread-safe shared state to prevent WebSocket disconnection
- 2026-02-09: Added dual format HS Code search (plain + dotted) for INSW compatibility
- 2026-02-09: Increased timeouts (60s page load, 20s element waits) for production stability
- 2026-02-09: Added thread liveness checking and dead thread recovery
- 2026-02-09: Removed keyword/AI classification section, focused on INSW Otomatis as primary feature
- 2026-02-09: Added flexible chapter filter (auto-detect all prefixes, multiselect any combination)
- 2026-02-09: Enhanced import/export differentiation (separate columns, Jenis field, color coding)
- 2026-02-09: Added separate Excel sheets for Regulasi Impor, Ekspor, and Terkait Obat
