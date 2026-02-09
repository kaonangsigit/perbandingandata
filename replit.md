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
5. **Page reuse optimization**: Navigate to INSW once, use back-navigation between checks instead of reloading

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

## Recent Changes
- 2026-02-09: Removed keyword/AI classification section, focused on INSW Otomatis as primary feature
- 2026-02-09: Added flexible chapter filter (auto-detect all prefixes, multiselect any combination)
- 2026-02-09: Optimized INSW scraping speed (page reuse, element waits, back-navigation)
- 2026-02-09: Enhanced import/export differentiation (separate columns, Jenis field, color coding)
- 2026-02-09: Added separate Excel sheets for Regulasi Impor, Ekspor, and Terkait Obat
