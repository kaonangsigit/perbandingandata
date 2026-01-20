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
- **Comparison Logic**: Currently shows preview of both datasets; full comparison logic to be implemented

### Application Entry Points
- `app.py`: Main Streamlit application (run with `streamlit run app.py`)
- `main.py`: Placeholder Python entry point (not actively used)

### Design Decisions
1. **Streamlit over Flask/Django**: Chosen for rapid prototyping of data-centric applications with minimal frontend code
2. **Pandas for Data Processing**: Standard choice for Excel/tabular data manipulation in Python
3. **Wide Layout**: Enables side-by-side comparison of datasets which is core to the application's purpose

## External Dependencies

### Python Libraries
- **streamlit**: Web application framework for data apps
- **pandas**: Data manipulation and Excel file reading
- **openpyxl** (implicit): Required by pandas for .xlsx file support

### File I/O
- Excel file uploads handled via Streamlit's file_uploader component
- In-memory processing using pandas read_excel

### No External Services
- No database connections
- No external APIs
- No authentication systems
- Fully client-side data processing (files are processed in memory)