# Perbandingan Data Impor - Import Data Comparison Tool

## Overview

This is a Streamlit-based web application for comparing import realization data between two Excel files. The tool helps users identify discrepancies between system-pulled data ("File Tarikan") and user-provided data, specifically checking for records that exist in the system but are missing from the user's dataset, and whether those records have SKI (likely a document/certification identifier).

The application serves as a data reconciliation utility for import data management workflows.

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
2. **Cek HS Code Obat** (Tab 2): HS Code pharmaceutical classification with two methods:
   - **Keyword (Offline)**: Fast, no cost, uses pre-defined keyword lists
   - **AI / ChatGPT (Online)**: More accurate, uses OpenAI (gpt-5-nano) to classify each HS Code via Replit AI Integrations
   - **Cek INSW Otomatis**: Uses Playwright headless browser to scrape INSW (insw.go.id/intr/detail-komoditas) for each HS code, extracting:
     - Whether import regulations exist (Regulasi Impor / Lartas Border / Tata Niaga Post Border)
     - Whether export regulations exist (Regulasi Ekspor / Lartas Ekspor)
     - Komoditi classification from INSW (e.g., Obat, Obat Bahan Alam, Bahan Suplemen Kesehatan, Narkotika, etc.)
     - Whether the HS code is related to medicine/pharmaceuticals based on INSW data
     - BPOM regulation presence
3. **Analisis Data** (Tab 3): Data analysis features

### Application Entry Points
- `app.py`: Main Streamlit application (run with `streamlit run app.py`)
- `main.py`: Placeholder Python entry point (not actively used)

### Design Decisions
1. **Streamlit over Flask/Django**: Chosen for rapid prototyping of data-centric applications with minimal frontend code
2. **Pandas for Data Processing**: Standard choice for Excel/tabular data manipulation in Python
3. **Wide Layout**: Enables side-by-side comparison of datasets which is core to the application's purpose
4. **AI Classification**: Uses gpt-5-nano for cost efficiency; batch processing (30 items/batch) with fallback to keyword if AI fails

## External Dependencies

### Python Libraries
- **streamlit**: Web application framework for data apps
- **pandas**: Data manipulation and Excel file reading
- **openpyxl** (implicit): Required by pandas for .xlsx file support
- **openai**: OpenAI API client for AI-powered HS Code classification
- **playwright**: Headless browser automation for INSW website scraping (requires chromium + mesa-libgbm system dep)

### File I/O
- Excel file uploads handled via Streamlit's file_uploader component
- In-memory processing using pandas read_excel
- BPS files read with dtype=str to avoid mixed-type column issues

### External Services
- **OpenAI via Replit AI Integrations**: Used for HS Code pharmaceutical classification (env vars: AI_INTEGRATIONS_OPENAI_API_KEY, AI_INTEGRATIONS_OPENAI_BASE_URL)
- No database connections
- No authentication systems