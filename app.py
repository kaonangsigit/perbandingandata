import streamlit as st
import pandas as pd
import io
import re
import os
import json
from openai import OpenAI
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(
    page_title="Perbandingan Data Impor", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        background: linear-gradient(90deg, #1e3a8a, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        padding: 1rem 0;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #64748b;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding: 10px 20px;
        background-color: #f1f5f9;
        border-radius: 10px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #3b82f6;
        color: white;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
    }
    .upload-section {
        background-color: #f8fafc;
        padding: 1.5rem;
        border-radius: 15px;
        border: 2px dashed #cbd5e1;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #dcfce7;
        border-left: 4px solid #22c55e;
        padding: 1rem;
        border-radius: 5px;
    }
    .warning-box {
        background-color: #fef3c7;
        border-left: 4px solid #f59e0b;
        padding: 1rem;
        border-radius: 5px;
    }
    .info-box {
        background-color: #e0f2fe;
        border-left: 4px solid #0ea5e9;
        padding: 1rem;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">📊 Perbandingan Data Realisasi Impor</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Aplikasi untuk membandingkan dan menganalisis data impor dengan mudah</p>', unsafe_allow_html=True)

tab_main, tab_hs, tab_analysis = st.tabs(["📋 Perbandingan Data", "💊 Cek HS Code Obat", "📈 Analisis Data"])

def clean_value(value):
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    val_str = val_str.replace(";", "").replace(",", "")
    
    date_patterns = [
        r'\s*/\s*\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+\d{4}',
        r'\s*/\s*\d{1,2}[-/]\d{1,2}[-/]\d{2,4}',
        r'\s*/\s*\d{4}[-/]\d{1,2}[-/]\d{1,2}',
        r'\s*-\s*\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+\d{4}',
    ]
    
    for pattern in date_patterns:
        val_str = re.sub(pattern, '', val_str, flags=re.IGNORECASE)
    
    val_str = re.sub(r'\s+', ' ', val_str)
    
    return val_str.strip()

def clean_number(value):
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    val_str = re.sub(r'[^\d]', '', val_str)
    return val_str

def get_invoice_list(value):
    if pd.isna(value):
        return []
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    
    if ';' in val_str or ',' in val_str:
        val_str = val_str.replace(';', ',')
        invoices = [inv.strip().strip(';').strip(',').strip() for inv in val_str.split(',')]
        invoices = [inv for inv in invoices if inv]
        return invoices
    
    val_str = val_str.strip(';').strip(',').strip()
    return [val_str] if val_str else []

def find_invoice_column(df):
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if 'invoice' in col_lower and 'no' in col_lower:
            return col
        if col_lower == 'no. invoice' or col_lower == 'no.invoice' or col_lower == 'noinvoice':
            return col
    for col in df.columns:
        if 'invoice' in str(col).lower():
            return col
    return None

def load_invoice_set(file_invoice, label):
    invoice_set = set()
    if file_invoice:
        df_invoice = pd.read_excel(file_invoice)
        invoice_col = find_invoice_column(df_invoice)
        if invoice_col:
            for inv_value in df_invoice[invoice_col].dropna():
                inv_list = get_invoice_list(inv_value)
                invoice_set.update(inv_list)
            st.success(f"✅ **{label}**: {len(invoice_set)} NO. INVOICE unik ditemukan")
        else:
            st.warning(f"⚠️ Kolom NO. INVOICE tidak ditemukan di {label}")
    return invoice_set

def is_numeric_column(col_name):
    col_lower = str(col_name).lower()
    numeric_keywords = ['pib', 'pengajuan']
    return any(keyword in col_lower for keyword in numeric_keywords)

with tab_main:
    st.markdown("### 📁 Upload File")
    
    with st.expander("📖 Petunjuk Penggunaan", expanded=False):
        st.markdown("""
        1. Upload **File Tarikan** (bisa multiple file, akan digabung otomatis)
        2. Upload **File Data Anda** (data yang ingin dibandingkan)
        3. **Pilih kolom** yang ingin digunakan untuk perbandingan
        4. Upload **File Invoice** (opsional) untuk cek NO. INVOICE
        5. Klik **Bandingkan Data**
        6. Download hasil: Data SAMA = **Kuning**, Data berbeda = **Putih**
        """)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 📥 File Tarikan (bisa multiple)")
        files_tarikan = st.file_uploader("Data hasil tarikan sistem", type=['xlsx', 'xls'], key="tarikan", help="Upload file Excel dari sistem (bisa pilih banyak file)", accept_multiple_files=True)

    with col2:
        st.markdown("#### 📤 File Data Anda")
        file_upload = st.file_uploader("Data Anda untuk dibandingkan", type=['xlsx', 'xls'], key="upload", help="Upload file Excel Anda")

    st.markdown("---")
    st.markdown("### 📑 File Invoice (Opsional)")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 💊 Bahan Tambahan Obat")
        file_invoice_obat = st.file_uploader("File Invoice Bahan Tambahan Obat", type=['xlsx', 'xls'], key="invoice_obat")

    with col2:
        st.markdown("#### 🧪 Bahan Kimia")
        file_invoice_kimia = st.file_uploader("File Invoice Bahan Kimia", type=['xlsx', 'xls'], key="invoice_kimia")

    if files_tarikan and file_upload:
        try:
            tarikan_files_data = []
            dfs_all = []
            for f in files_tarikan:
                df_temp = pd.read_excel(f)
                file_name = f.name.replace('.xlsx', '').replace('.xls', '')[:31]
                tarikan_files_data.append({'name': file_name, 'df': df_temp})
                dfs_all.append(df_temp)
            
            df_tarikan = pd.concat(dfs_all, ignore_index=True)
            st.success(f"✅ {len(files_tarikan)} file tarikan dimuat: {len(df_tarikan)} baris total")
            
            df_upload = pd.read_excel(file_upload)
            
            st.markdown("---")
            st.markdown("### ⚙️ Konfigurasi Perbandingan")
            
            col_tarikan_list = df_tarikan.columns.tolist()
            col_upload_list = df_upload.columns.tolist()
            
            common_cols = [col for col in col_tarikan_list if col in col_upload_list]
            
            col1, col2 = st.columns(2)
            
            with col1:
                selected_col_tarikan = st.selectbox(
                    "📌 Kolom File Tarikan",
                    options=col_tarikan_list,
                    index=0,
                    key="col_tarikan"
                )
            
            with col2:
                default_index = col_upload_list.index(selected_col_tarikan) if selected_col_tarikan in col_upload_list else 0
                selected_col_upload = st.selectbox(
                    "📌 Kolom File Data Anda",
                    options=col_upload_list,
                    index=default_index,
                    key="col_upload"
                )
            
            use_numeric_cleaning = st.checkbox(
                "🔢 Bersihkan numerik saja (HANYA untuk kolom angka murni seperti NO. PIB)",
                value=is_numeric_column(selected_col_tarikan),
                help="⚠️ JANGAN centang jika data mengandung huruf seperti ST.03.04.35.352A..."
            )
            
            st.markdown("---")
            st.markdown("### 📋 Pilihan Jenis Output Download")
            output_option = st.radio(
                "Pilih jenis output yang diinginkan:",
                options=[
                    "❌ Download HANYA data yang TIDAK ADA di file lain (Output Lama)",
                    "📊 Download SEMUA data dengan highlight kuning untuk yang SAMA (Output Baru)"
                ],
                index=0,
                help="Pilih jenis output: Output Lama = hanya data tidak cocok, Output Baru = semua data dengan warna"
            )
            
            if common_cols:
                st.info(f"💡 Kolom yang sama di kedua file: **{', '.join(common_cols)}**")
            
            invoice_col_tarikan = find_invoice_column(df_tarikan)
            
            invoice_set_obat = set()
            invoice_set_kimia = set()
            
            st.markdown("---")
            st.markdown("### 📋 Status File Invoice")
            
            col1, col2 = st.columns(2)
            with col1:
                if file_invoice_obat:
                    invoice_set_obat = load_invoice_set(file_invoice_obat, "Bahan Tambahan Obat")
                else:
                    st.info("📭 File Invoice Bahan Tambahan Obat belum diupload")
            
            with col2:
                if file_invoice_kimia:
                    invoice_set_kimia = load_invoice_set(file_invoice_kimia, "Bahan Kimia")
                else:
                    st.info("📭 File Invoice Bahan Kimia belum diupload")
            
            st.markdown("---")
            st.markdown("### 👀 Preview Data")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 📊 Data Tarikan")
                st.caption(f"📝 {len(df_tarikan)} baris | Kolom: **{selected_col_tarikan}**")
                st.dataframe(df_tarikan.head(5), use_container_width=True, height=200)
            
            with col2:
                st.markdown("#### 📊 Data Anda")
                st.caption(f"📝 {len(df_upload)} baris | Kolom: **{selected_col_upload}**")
                st.dataframe(df_upload.head(5), use_container_width=True, height=200)
            
            st.markdown("---")
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                compare_btn = st.button("🔍 Bandingkan Data", type="primary", use_container_width=True)
            
            if compare_btn:
                st.markdown("---")
                st.markdown(f"### 📊 Hasil Perbandingan: {selected_col_tarikan}")
                
                if use_numeric_cleaning:
                    df_tarikan['_clean_key'] = df_tarikan[selected_col_tarikan].apply(clean_number)
                    df_upload['_clean_key'] = df_upload[selected_col_upload].apply(clean_number)
                else:
                    df_tarikan['_clean_key'] = df_tarikan[selected_col_tarikan].apply(clean_value)
                    df_upload['_clean_key'] = df_upload[selected_col_upload].apply(clean_value)
                
                with st.expander("🔎 Preview Hasil Pembersihan Data (klik untuk lihat)", expanded=False):
                    st.markdown("**File Tarikan - Sample Data Sebelum & Sesudah Pembersihan:**")
                    preview_tarikan = df_tarikan[[selected_col_tarikan, '_clean_key']].head(5).copy()
                    preview_tarikan.columns = ['Data Asli', 'Setelah Dibersihkan']
                    st.dataframe(preview_tarikan, use_container_width=True)
                    
                    st.markdown("**File Anda - Sample Data Sebelum & Sesudah Pembersihan:**")
                    preview_upload = df_upload[[selected_col_upload, '_clean_key']].head(5).copy()
                    preview_upload.columns = ['Data Asli', 'Setelah Dibersihkan']
                    st.dataframe(preview_upload, use_container_width=True)
                
                tarikan_keys = set(df_tarikan['_clean_key'].dropna())
                tarikan_keys = {k for k in tarikan_keys if k != ''}
                
                upload_keys = set(df_upload['_clean_key'].dropna())
                upload_keys = {k for k in upload_keys if k != ''}
                
                matching_keys = tarikan_keys & upload_keys
                missing_in_upload = tarikan_keys - upload_keys
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("📥 Data Tarikan", len(tarikan_keys), help="Jumlah data unik di file tarikan")
                with col2:
                    st.metric("📤 Data Anda", len(upload_keys), help="Jumlah data unik di file Anda")
                with col3:
                    st.metric("✅ Data SAMA", len(matching_keys), delta=f"+{len(matching_keys)}" if matching_keys else None, delta_color="normal", help="Data yang ada di KEDUA file")
                with col4:
                    st.metric("❌ Tidak Ada", len(missing_in_upload), delta=f"-{len(missing_in_upload)}" if missing_in_upload else None, delta_color="inverse", help="Data tarikan yang tidak ada di file Anda")
                
                df_tarikan_display = df_tarikan.copy()
                df_tarikan_display['Status'] = df_tarikan_display['_clean_key'].apply(
                    lambda x: '✅ Sama' if x in matching_keys else '❌ Tidak Sama'
                )
                df_tarikan_display = df_tarikan_display.drop(columns=['_clean_key'])
                
                jumlah_sama = len(df_tarikan_display[df_tarikan_display['Status'] == '✅ Sama'])
                jumlah_tidak_sama = len(df_tarikan_display[df_tarikan_display['Status'] == '❌ Tidak Sama'])
                
                if missing_in_upload:
                    st.markdown(f"### 🔴 Data Tarikan yang Tidak Ada di File Anda")
                    st.warning(f"Ditemukan **{len(missing_in_upload)}** data unik dari tarikan yang tidak ada di file Anda.")
                    
                    df_missing = df_tarikan[df_tarikan['_clean_key'].isin(missing_in_upload)].copy()
                    df_missing = df_missing.drop(columns=['_clean_key'])
                    
                    st.dataframe(df_missing, use_container_width=True, height=300)
                
                st.markdown("---")
                
                show_only_missing = "HANYA" in output_option
                
                if show_only_missing:
                    st.markdown("### 📥 Download Data yang TIDAK ADA (Output Lama)")
                    st.markdown("File Excel berisi **hanya data yang tidak ada** di file lain")
                    
                    if missing_in_upload:
                        df_missing = df_tarikan[df_tarikan['_clean_key'].isin(missing_in_upload)].copy()
                        df_missing = df_missing.drop(columns=['_clean_key'])
                        
                        output_missing = io.BytesIO()
                        with pd.ExcelWriter(output_missing, engine='openpyxl') as writer:
                            header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                            header_font = Font(bold=True, color='FFFFFF')
                            thin_border = Border(
                                left=Side(style='thin'),
                                right=Side(style='thin'),
                                top=Side(style='thin'),
                                bottom=Side(style='thin')
                            )
                            
                            for file_data in tarikan_files_data:
                                df_file = file_data['df'].copy()
                                if use_numeric_cleaning:
                                    df_file['_clean_key'] = df_file[selected_col_tarikan].apply(clean_number)
                                else:
                                    df_file['_clean_key'] = df_file[selected_col_tarikan].apply(clean_value)
                                
                                df_file_missing = df_file[df_file['_clean_key'].isin(missing_in_upload)].copy()
                                df_file_missing = df_file_missing.drop(columns=['_clean_key'])
                                
                                if len(df_file_missing) > 0:
                                    sheet_name = file_data['name'][:31]
                                    df_file_missing.to_excel(writer, index=False, sheet_name=sheet_name)
                                    
                                    worksheet = writer.sheets[sheet_name]
                                    for col_idx, col in enumerate(df_file_missing.columns, 1):
                                        cell = worksheet.cell(row=1, column=col_idx)
                                        cell.fill = header_fill
                                        cell.font = header_font
                                        cell.alignment = Alignment(horizontal='center')
                                        cell.border = thin_border
                                    
                                    for row_idx in range(2, len(df_file_missing) + 2):
                                        for col_idx in range(1, len(df_file_missing.columns) + 1):
                                            cell = worksheet.cell(row=row_idx, column=col_idx)
                                            cell.border = thin_border
                                    
                                    for col_idx, col in enumerate(df_file_missing.columns, 1):
                                        max_len = max(df_file_missing[col].astype(str).apply(len).max(), len(str(col))) + 2
                                        worksheet.column_dimensions[worksheet.cell(row=1, column=col_idx).column_letter].width = min(max_len, 50)
                        
                        output_missing.seek(0)
                        
                        st.metric("❌ Total Data Tidak Ada", len(missing_in_upload))
                        
                        st.download_button(
                            label="📥 Download Data yang Tidak Ada",
                            data=output_missing,
                            file_name="data_tidak_ada.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.success("✅ Semua data tarikan sudah ada di file Anda!")
                
                else:
                    st.markdown("### 📊 Download Data Lengkap dengan Warna (Output Baru)")
                    st.markdown(f"File Excel akan memiliki **{len(tarikan_files_data) + 1} sheet/laman**:")
                    for i, file_data in enumerate(tarikan_files_data, 1):
                        st.markdown(f"- 📥 **Sheet {i}**: {file_data['name']}")
                    st.markdown(f"- 📤 **Sheet Terakhir**: Data Anda")
                    st.markdown("- 🟡 **Warna Kuning**: Data yang **SAMA** di kedua file")
                    st.markdown("- ⬜ **Putih**: Data yang **TIDAK ADA** di file lain")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("🟡 Data Kuning (Sama)", jumlah_sama)
                    with col2:
                        st.metric("⬜ Data Putih (Tidak Sama)", jumlah_tidak_sama)
                    
                    df_upload_display = df_upload.copy()
                    df_upload_display['Status'] = df_upload_display['_clean_key'].apply(
                        lambda x: '✅ Sama' if x in matching_keys else '❌ Tidak Sama'
                    )
                    df_upload_display = df_upload_display.drop(columns=['_clean_key'])
                    
                    output_colored = io.BytesIO()
                    with pd.ExcelWriter(output_colored, engine='openpyxl') as writer:
                        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                        header_font = Font(bold=True, color='FFFFFF')
                        thin_border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
                        
                        for file_data in tarikan_files_data:
                            df_file = file_data['df'].copy()
                            if use_numeric_cleaning:
                                df_file['_clean_key'] = df_file[selected_col_tarikan].apply(clean_number)
                            else:
                                df_file['_clean_key'] = df_file[selected_col_tarikan].apply(clean_value)
                            df_file['Status'] = df_file['_clean_key'].apply(
                                lambda x: '✅ Sama' if x in matching_keys else '❌ Tidak Sama'
                            )
                            df_file = df_file.drop(columns=['_clean_key'])
                            
                            sheet_name = file_data['name'][:31]
                            df_file.to_excel(writer, index=False, sheet_name=sheet_name)
                            
                            worksheet = writer.sheets[sheet_name]
                            for col_idx, col in enumerate(df_file.columns, 1):
                                cell = worksheet.cell(row=1, column=col_idx)
                                cell.fill = header_fill
                                cell.font = header_font
                                cell.alignment = Alignment(horizontal='center')
                                cell.border = thin_border
                            
                            status_col_idx = df_file.columns.get_loc('Status') + 1
                            for row_idx in range(2, len(df_file) + 2):
                                status_cell = worksheet.cell(row=row_idx, column=status_col_idx)
                                if '✅' in str(status_cell.value):
                                    for col_idx in range(1, len(df_file.columns) + 1):
                                        cell = worksheet.cell(row=row_idx, column=col_idx)
                                        cell.fill = yellow_fill
                                        cell.border = thin_border
                                else:
                                    for col_idx in range(1, len(df_file.columns) + 1):
                                        cell = worksheet.cell(row=row_idx, column=col_idx)
                                        cell.border = thin_border
                            
                            for col_idx, col in enumerate(df_file.columns, 1):
                                max_len = max(df_file[col].astype(str).apply(len).max(), len(str(col))) + 2
                                worksheet.column_dimensions[worksheet.cell(row=1, column=col_idx).column_letter].width = min(max_len, 50)
                        
                        df_upload_display.to_excel(writer, index=False, sheet_name='Data Anda')
                        worksheet_upload = writer.sheets['Data Anda']
                        
                        for col_idx, col in enumerate(df_upload_display.columns, 1):
                            cell = worksheet_upload.cell(row=1, column=col_idx)
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center')
                            cell.border = thin_border
                        
                        status_col_idx = df_upload_display.columns.get_loc('Status') + 1
                        for row_idx in range(2, len(df_upload_display) + 2):
                            status_cell = worksheet_upload.cell(row=row_idx, column=status_col_idx)
                            if '✅' in str(status_cell.value):
                                for col_idx in range(1, len(df_upload_display.columns) + 1):
                                    cell = worksheet_upload.cell(row=row_idx, column=col_idx)
                                    cell.fill = yellow_fill
                                    cell.border = thin_border
                            else:
                                for col_idx in range(1, len(df_upload_display.columns) + 1):
                                    cell = worksheet_upload.cell(row=row_idx, column=col_idx)
                                    cell.border = thin_border
                        
                        for col_idx, col in enumerate(df_upload_display.columns, 1):
                            max_len = max(df_upload_display[col].astype(str).apply(len).max(), len(str(col))) + 2
                            worksheet_upload.column_dimensions[worksheet_upload.cell(row=1, column=col_idx).column_letter].width = min(max_len, 50)
                        
                    output_colored.seek(0)
                    
                    st.download_button(
                        label="📥 Download Excel dengan Warna",
                        data=output_colored,
                        file_name="hasil_perbandingan_berwarna.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                if missing_in_upload:
                    
                    if invoice_col_tarikan and (invoice_set_obat or invoice_set_kimia):
                        st.markdown("---")
                        st.markdown("### 📋 Cek NO. INVOICE")
                        
                        def check_invoice_obat(inv_value):
                            if not invoice_set_obat:
                                return '-'
                            inv_list = get_invoice_list(inv_value)
                            if not inv_list:
                                return '❌ Tidak Ada'
                            found = sum(1 for inv in inv_list if inv in invoice_set_obat)
                            if found == len(inv_list):
                                return '✅ Ada'
                            elif found > 0:
                                return f'⚠️ Sebagian ({found}/{len(inv_list)})'
                            else:
                                return '❌ Tidak Ada'
                        
                        def check_invoice_kimia(inv_value):
                            if not invoice_set_kimia:
                                return '-'
                            inv_list = get_invoice_list(inv_value)
                            if not inv_list:
                                return '❌ Tidak Ada'
                            found = sum(1 for inv in inv_list if inv in invoice_set_kimia)
                            if found == len(inv_list):
                                return '✅ Ada'
                            elif found > 0:
                                return f'⚠️ Sebagian ({found}/{len(inv_list)})'
                            else:
                                return '❌ Tidak Ada'
                        
                        df_invoice_check = df_missing.copy()
                        
                        if invoice_set_obat:
                            st.markdown("#### 💊 Cek di Bahan Tambahan Obat")
                            df_invoice_check['Cek Bahan Obat'] = df_invoice_check[invoice_col_tarikan].apply(check_invoice_obat)
                            
                            ada_obat = df_invoice_check[df_invoice_check['Cek Bahan Obat'] == '✅ Ada']
                            sebagian_obat = df_invoice_check[df_invoice_check['Cek Bahan Obat'].str.contains('Sebagian', na=False)]
                            tidak_obat = df_invoice_check[df_invoice_check['Cek Bahan Obat'] == '❌ Tidak Ada']
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("✅ Ada", len(ada_obat))
                            with col2:
                                st.metric("⚠️ Sebagian", len(sebagian_obat))
                            with col3:
                                st.metric("❌ Tidak Ada", len(tidak_obat))
                        
                        if invoice_set_kimia:
                            st.markdown("#### 🧪 Cek di Bahan Kimia")
                            df_invoice_check['Cek Bahan Kimia'] = df_invoice_check[invoice_col_tarikan].apply(check_invoice_kimia)
                            
                            ada_kimia = df_invoice_check[df_invoice_check['Cek Bahan Kimia'] == '✅ Ada']
                            sebagian_kimia = df_invoice_check[df_invoice_check['Cek Bahan Kimia'].str.contains('Sebagian', na=False)]
                            tidak_kimia = df_invoice_check[df_invoice_check['Cek Bahan Kimia'] == '❌ Tidak Ada']
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("✅ Ada", len(ada_kimia))
                            with col2:
                                st.metric("⚠️ Sebagian", len(sebagian_kimia))
                            with col3:
                                st.metric("❌ Tidak Ada", len(tidak_kimia))
                        
                        st.markdown("#### 📊 Data Lengkap dengan Status Invoice")
                        st.dataframe(df_invoice_check, use_container_width=True, height=300)
                        
                        if invoice_set_obat:
                            st.markdown("---")
                            st.markdown("##### 💊 Filter Bahan Tambahan Obat")
                            tab1, tab2, tab3 = st.tabs(["✅ Ada", "⚠️ Sebagian", "❌ Tidak Ada"])
                            
                            with tab1:
                                if len(ada_obat) > 0:
                                    st.dataframe(ada_obat, use_container_width=True)
                                else:
                                    st.info("Tidak ada data")
                            
                            with tab2:
                                if len(sebagian_obat) > 0:
                                    st.dataframe(sebagian_obat, use_container_width=True)
                                else:
                                    st.info("Tidak ada data")
                            
                            with tab3:
                                if len(tidak_obat) > 0:
                                    st.dataframe(tidak_obat, use_container_width=True)
                                else:
                                    st.info("Tidak ada data")
                        
                        if invoice_set_kimia:
                            st.markdown("---")
                            st.markdown("##### 🧪 Filter Bahan Kimia")
                            tab1, tab2, tab3 = st.tabs(["✅ Ada ", "⚠️ Sebagian ", "❌ Tidak Ada "])
                            
                            with tab1:
                                if len(ada_kimia) > 0:
                                    st.dataframe(ada_kimia, use_container_width=True)
                                else:
                                    st.info("Tidak ada data")
                            
                            with tab2:
                                if len(sebagian_kimia) > 0:
                                    st.dataframe(sebagian_kimia, use_container_width=True)
                                else:
                                    st.info("Tidak ada data")
                            
                            with tab3:
                                if len(tidak_kimia) > 0:
                                    st.dataframe(tidak_kimia, use_container_width=True)
                                else:
                                    st.info("Tidak ada data")
                        
                        output_invoice = io.BytesIO()
                        with pd.ExcelWriter(output_invoice, engine='openpyxl') as writer:
                            df_invoice_check.to_excel(writer, index=False, sheet_name='Hasil Cek Invoice')
                        output_invoice.seek(0)
                        
                        st.download_button(
                            label="📥 Download Hasil Cek Invoice",
                            data=output_invoice,
                            file_name="hasil_cek_invoice.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                else:
                    st.success("✅ Semua data dari tarikan sudah tersedia di file Anda!")
                
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan: {str(e)}")
            st.info("💡 Pastikan file Excel dalam format yang benar (.xlsx atau .xls)")

    else:
        st.info("👆 Silakan upload **File Tarikan** dan **File Data Anda** untuk memulai perbandingan.")

with tab_hs:
    st.markdown("### 💊 Pengecekan HS Code Obat/Bahan Baku Obat")
    st.markdown("Upload file data ekspor/impor dari BPS untuk filter HS Code yang termasuk obat/bahan baku obat.")
    
    with st.expander("📖 Petunjuk Penggunaan", expanded=False):
        st.markdown("""
        1. Upload file data dari **BPS** (format .xlsx/.xls)
        2. Sistem akan **otomatis filter** HS Code awalan 28, 29, 30, 31
        3. Klasifikasi otomatis mana yang **benar-benar masuk obat/bahan baku obat**
        4. Kolom **HS Code** dan **Nama Bahan Baku Obat/Obat** akan terisi otomatis
        5. Download hasil dalam format Excel
        
        **Kategori HS Code:**
        - **28**: Bahan kimia anorganik (sebagian bahan baku obat)
        - **29**: Bahan kimia organik (banyak bahan baku obat)
        - **30**: Produk farmasi (semua termasuk obat)
        - **31**: Pupuk (umumnya bukan obat)
        """)
    
    PHARMA_KEYWORDS = [
        'medic', 'pharma', 'drug', 'vaccine', 'antibiotic', 'vitamin',
        'hormone', 'insulin', 'steroid', 'alkaloid', 'glycoside',
        'analges', 'antisep', 'anaesthe', 'antipyr', 'antimal',
        'anthelm', 'contracepti', 'prophylac', 'therapeutic',
        'surgical', 'first-aid', 'dressing', 'bandage', 'catgut',
        'quinine', 'sulphonamide', 'saccharin', 'sucralose',
        'lysine', 'glutam', 'amino acid', 'nucleic acid',
        'provitamin', 'penicillin', 'streptomycin', 'erythromycin',
        'tetracycline', 'chloramphenicol', 'aspirin', 'paracetamol',
        'acetylsalicylic', 'ibuprofen', 'caffeine', 'ephedrine',
        'pseudoephedrine', 'codeine', 'morphine', 'herbal medic',
        'immunolog', 'serum', 'toxin', 'antitoxin',
        'cancer', 'tumour', 'tumor', 'intractable',
        'sodium chloride', 'glucose', 'infusion',
        'isoniazid', 'chlorpheniramine', 'mebendazole', 'parbendazole',
        'hydantoin', 'lactam', 'imidazole',
    ]
    
    PHARMA_RAW_KEYWORDS_28 = [
        'zinc oxide', 'aluminium hydroxide', 'hydrogen peroxide',
        'sodium hydroxide', 'calcium', 'phosphat',
        'ammonia', 'oxygen', 'carbon dioxide', 'silicon dioxide',
        'sodium sulphite', 'potassium', 'magnesium',
        'iodine', 'iodide', 'bromide', 'fluoride',
        'ferrous', 'ferric', 'iron oxide',
        'manganese', 'copper sulphate', 'boric acid',
        'sodium bicarbonate', 'sodium carbonate',
        'calcium carbonate', 'magnesium oxide',
        'magnesium hydroxide', 'titanium dioxide',
        'sulphuric acid', 'hydrochloric acid', 'nitric acid',
    ]
    
    PHARMA_RAW_KEYWORDS_29 = [
        'methanol', 'ethanol', 'alcohol', 'glycerol', 'glycol',
        'mannitol', 'sorbitol', 'phenol', 'vanillin',
        'citric acid', 'acetic acid', 'benzoic acid', 'salicylic',
        'stearic', 'palmitic', 'oleic', 'lauric',
        'formaldehyde', 'paraformaldehyde', 'acetone',
        'ether', 'ester', 'lactone', 'coumarin',
        'amine', 'amide', 'amino', 'urea',
        'menthol', 'camphor', 'thymol', 'eucalyptol',
        'benzyl alcohol', 'isopropyl alcohol',
        'propylene glycol', 'ethylene glycol',
        'acrylic acid', 'methacrylic',
        'parathion', 'organo-phosphor',
        'azodicarbonamide',
        'butanol', 'propanol', 'octanol',
    ]
    
    def classify_hs_pharma(hs_code, description):
        desc_lower = description.lower()
        code_prefix = hs_code[:2]
        
        if code_prefix == '30':
            return True, 'Produk Farmasi'
        
        for kw in PHARMA_KEYWORDS:
            if kw.lower() in desc_lower:
                if code_prefix == '29':
                    return True, 'Bahan Baku Obat (Kimia Organik)'
                elif code_prefix == '28':
                    return True, 'Bahan Baku Obat (Kimia Anorganik)'
                elif code_prefix == '31':
                    return True, 'Bahan Terkait Farmasi'
                return True, 'Terkait Farmasi'
        
        if code_prefix == '29':
            for kw in PHARMA_RAW_KEYWORDS_29:
                if kw.lower() in desc_lower:
                    return True, 'Bahan Baku Obat (Kimia Organik)'
        
        if code_prefix == '28':
            for kw in PHARMA_RAW_KEYWORDS_28:
                if kw.lower() in desc_lower:
                    return True, 'Bahan Baku Obat (Kimia Anorganik)'
        
        if code_prefix == '31':
            return False, 'Pupuk (Bukan Obat)'
        
        return False, 'Bukan Obat/Bahan Baku Obat'
    
    def classify_hs_with_ai(hs_items_batch):
        client = OpenAI(
            api_key=os.environ.get("AI_INTEGRATIONS_OPENAI_API_KEY"),
            base_url=os.environ.get("AI_INTEGRATIONS_OPENAI_BASE_URL"),
        )
        
        items_text = ""
        for i, h in enumerate(hs_items_batch):
            items_text += f"{i+1}. HS Code: {h['hs_code']} - {h['description']}\n"
        
        prompt = f"""You are a pharmaceutical and drug substance classification expert. 
Analyze each HS Code below and determine if it is a pharmaceutical product, drug raw material (bahan baku obat), or related to medicine/healthcare.

Consider these criteria:
- Active Pharmaceutical Ingredients (API) / Bahan aktif obat
- Excipients used in drug formulation / Bahan tambahan obat
- Finished pharmaceutical products / Produk farmasi jadi
- Medical devices and surgical supplies / Alat kesehatan
- Vaccine, serum, blood products / Vaksin dan produk darah
- Traditional/herbal medicine ingredients / Bahan obat tradisional
- Chemical compounds commonly used in pharmaceutical manufacturing
- Substances listed in pharmacopoeia

For each item, respond with a JSON array. Each element must have:
- "index": the item number (1-based)
- "is_pharma": true or false
- "kategori": one of: "Produk Farmasi", "Bahan Baku Obat (Kimia Organik)", "Bahan Baku Obat (Kimia Anorganik)", "Bahan Terkait Farmasi", "Pupuk (Bukan Obat)", "Bukan Obat/Bahan Baku Obat"
- "alasan": brief reason in Indonesian (max 15 words)

Items to classify:
{items_text}

Respond ONLY with the JSON array, no other text."""

        try:
            response = client.chat.completions.create(
                model="gpt-5-nano",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
            )
            
            content = response.choices[0].message.content.strip()
            if content.startswith("```"):
                content = content.split("\n", 1)[1]
                content = content.rsplit("```", 1)[0]
            
            ai_results = json.loads(content)
            return ai_results
        except Exception as e:
            st.error(f"Error dari AI: {str(e)}")
            return None
    
    file_hs = st.file_uploader("📁 Upload file data BPS", type=['xlsx', 'xls'], key="hs_check")
    
    if file_hs:
        try:
            df_hs_raw = pd.read_excel(file_hs, header=None, dtype=str)
            for col in df_hs_raw.columns:
                df_hs_raw[col] = df_hs_raw[col].astype(str).replace('nan', '')
            
            header_row = None
            for i in range(min(10, len(df_hs_raw))):
                val = str(df_hs_raw.iloc[i, 0]).strip().lower()
                if 'kode hs' in val or 'hs code' in val:
                    header_row = i
                    break
            
            if header_row is None:
                header_row = 3
            
            data_start = header_row + 1
            
            hs_items = []
            for idx in range(data_start, len(df_hs_raw)):
                val = str(df_hs_raw.iloc[idx, 0]).strip()
                match = re.match(r'\[(\d+)\]\s*(.*)', val)
                if match:
                    code = match.group(1)
                    desc = match.group(2).strip()
                    hs_items.append({
                        'row_idx': idx,
                        'raw_value': val,
                        'hs_code': code,
                        'description': desc,
                        'prefix': code[:2]
                    })
            
            st.success(f"Total **{len(hs_items)}** HS Code ditemukan dalam file")
            
            hs_filtered = [h for h in hs_items if h['prefix'] in ['28', '29', '30', '31']]
            
            st.markdown("---")
            st.markdown("### 📊 Hasil Filter HS Code 28, 29, 30, 31")
            
            col1, col2, col3, col4, col5 = st.columns(5)
            count_28 = len([h for h in hs_filtered if h['prefix'] == '28'])
            count_29 = len([h for h in hs_filtered if h['prefix'] == '29'])
            count_30 = len([h for h in hs_filtered if h['prefix'] == '30'])
            count_31 = len([h for h in hs_filtered if h['prefix'] == '31'])
            
            with col1:
                st.metric("Total Filter", len(hs_filtered))
            with col2:
                st.metric("HS 28 (Anorganik)", count_28)
            with col3:
                st.metric("HS 29 (Organik)", count_29)
            with col4:
                st.metric("HS 30 (Farmasi)", count_30)
            with col5:
                st.metric("HS 31 (Pupuk)", count_31)
            
            st.markdown("---")
            st.markdown("### Metode Klasifikasi")
            
            metode = st.radio(
                "Pilih metode klasifikasi:",
                ["Keyword (Offline)", "AI / ChatGPT (Online)"],
                horizontal=True,
                help="Keyword: cepat, tanpa biaya, berdasarkan daftar kata kunci. AI: lebih akurat, menggunakan kecerdasan buatan untuk menganalisis setiap HS Code."
            )
            
            use_ai = metode == "AI / ChatGPT (Online)"
            
            results = []
            
            if use_ai:
                st.info("Menggunakan AI untuk klasifikasi. Proses ini memerlukan beberapa saat...")
                
                batch_size = 30
                progress_bar = st.progress(0)
                total_batches = max(1, (len(hs_filtered) + batch_size - 1) // batch_size)
                
                for batch_idx in range(0, len(hs_filtered), batch_size):
                    batch = hs_filtered[batch_idx:batch_idx + batch_size]
                    current_batch = batch_idx // batch_size + 1
                    progress_bar.progress(current_batch / total_batches, text=f"Memproses batch {current_batch}/{total_batches}...")
                    
                    ai_results = classify_hs_with_ai(batch)
                    
                    ai_map = {}
                    if ai_results:
                        for ai_item in ai_results:
                            idx = ai_item.get('index', 0) - 1
                            if 0 <= idx < len(batch):
                                ai_map[idx] = ai_item
                    
                    for i, h in enumerate(batch):
                        if i in ai_map:
                            ai_item = ai_map[i]
                            is_pharma = ai_item.get('is_pharma', False)
                            kategori = ai_item.get('kategori', 'Tidak Diketahui')
                            alasan = ai_item.get('alasan', '')
                        else:
                            is_pharma, kategori = classify_hs_pharma(h['hs_code'], h['description'])
                            alasan = '(Fallback ke keyword)'
                        
                        results.append({
                            'HS Code': h['hs_code'],
                            'Deskripsi (English)': h['description'],
                            'Kategori': kategori,
                            'Masuk Obat/Bahan Obat': 'YA' if is_pharma else 'TIDAK',
                            'Alasan AI': alasan,
                            'Chapter': f"HS {h['prefix']}",
                            '_row_idx': h['row_idx']
                        })
                
                progress_bar.progress(1.0, text="Selesai!")
            else:
                for h in hs_filtered:
                    is_pharma, kategori = classify_hs_pharma(h['hs_code'], h['description'])
                    results.append({
                        'HS Code': h['hs_code'],
                        'Deskripsi (English)': h['description'],
                        'Kategori': kategori,
                        'Masuk Obat/Bahan Obat': 'YA' if is_pharma else 'TIDAK',
                        'Chapter': f"HS {h['prefix']}",
                        '_row_idx': h['row_idx']
                    })
            
            df_results = pd.DataFrame(results)
            
            results_lookup = {}
            for r in results:
                results_lookup[r['_row_idx']] = r
            
            df_display = df_results.drop(columns=['_row_idx'], errors='ignore')
            
            pharma_count = len(df_display[df_display['Masuk Obat/Bahan Obat'] == 'YA'])
            non_pharma_count = len(df_display[df_display['Masuk Obat/Bahan Obat'] == 'TIDAK'])
            
            st.markdown("---")
            st.markdown("### 💊 Hasil Klasifikasi Otomatis")
            
            if use_ai:
                st.caption("Klasifikasi menggunakan AI - kolom 'Alasan AI' menunjukkan alasan klasifikasi")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("💊 Masuk Obat/Bahan Obat", pharma_count)
            with col2:
                st.metric("❌ Bukan Obat", non_pharma_count)
            
            tab_all, tab_pharma, tab_non = st.tabs(["📋 Semua Data", "💊 Obat/Bahan Obat", "❌ Bukan Obat"])
            
            with tab_all:
                def highlight_pharma(row):
                    if row['Masuk Obat/Bahan Obat'] == 'YA':
                        return ['background-color: #dcfce7'] * len(row)
                    else:
                        return ['background-color: #fef2f2'] * len(row)
                
                styled_df = df_display.style.apply(highlight_pharma, axis=1)
                st.dataframe(styled_df, use_container_width=True, height=400)
            
            with tab_pharma:
                df_pharma = df_display[df_display['Masuk Obat/Bahan Obat'] == 'YA']
                st.dataframe(df_pharma, use_container_width=True, height=400)
            
            with tab_non:
                df_non_pharma = df_display[df_display['Masuk Obat/Bahan Obat'] == 'TIDAK']
                st.dataframe(df_non_pharma, use_container_width=True, height=400)
            
            st.markdown("---")
            st.markdown("### 📥 Download Hasil")
            
            hs_code_col = None
            nama_obat_col = None
            for j in range(min(10, len(df_hs_raw.columns))):
                header_val = str(df_hs_raw.iloc[header_row, j]).strip().lower() if header_row < len(df_hs_raw) else ''
                if header_val == 'hs code' or header_val == 'kode hs':
                    if hs_code_col is None and j > 0:
                        hs_code_col = j
                if 'nama' in header_val and ('obat' in header_val or 'bahan' in header_val):
                    nama_obat_col = j
            
            if hs_code_col is None:
                hs_code_col = 1
            if nama_obat_col is None:
                nama_obat_col = 2
            
            df_output = df_hs_raw.copy()
            
            for h in hs_filtered:
                r = results_lookup.get(h['row_idx'])
                if r and r['Masuk Obat/Bahan Obat'] == 'YA':
                    df_output.iloc[h['row_idx'], hs_code_col] = h['hs_code']
                    df_output.iloc[h['row_idx'], nama_obat_col] = r['Kategori'] + ': ' + h['description']
            
            df_export = df_display.copy()
            
            output_hs = io.BytesIO()
            with pd.ExcelWriter(output_hs, engine='openpyxl') as writer:
                green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                header_font = Font(bold=True, color='FFFFFF')
                thin_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                
                df_output.to_excel(writer, index=False, header=False, sheet_name='Data Asli Terisi')
                ws_asli = writer.sheets['Data Asli Terisi']
                
                pharma_rows = set()
                for h in hs_filtered:
                    r = results_lookup.get(h['row_idx'])
                    if r and r['Masuk Obat/Bahan Obat'] == 'YA':
                        pharma_rows.add(h['row_idx'] + 1)
                
                for row_idx in range(1, len(df_output) + 1):
                    for col_idx in range(1, min(len(df_output.columns) + 1, 10)):
                        cell = ws_asli.cell(row=row_idx, column=col_idx)
                        cell.border = thin_border
                        if row_idx in pharma_rows:
                            cell.fill = green_fill
                
                df_export.to_excel(writer, index=False, sheet_name='Klasifikasi HS Code')
                ws_klasifikasi = writer.sheets['Klasifikasi HS Code']
                
                status_col_idx = list(df_export.columns).index('Masuk Obat/Bahan Obat') + 1 if 'Masuk Obat/Bahan Obat' in df_export.columns else 4
                
                for col_idx in range(1, len(df_export.columns) + 1):
                    cell = ws_klasifikasi.cell(row=1, column=col_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center')
                    cell.border = thin_border
                
                for row_idx in range(2, len(df_export) + 2):
                    status_cell = ws_klasifikasi.cell(row=row_idx, column=status_col_idx)
                    for col_idx in range(1, len(df_export.columns) + 1):
                        cell = ws_klasifikasi.cell(row=row_idx, column=col_idx)
                        cell.border = thin_border
                        if str(status_cell.value) == 'YA':
                            cell.fill = green_fill
                
                for col_idx, col in enumerate(df_export.columns, 1):
                    max_len = max(df_export[col].astype(str).apply(len).max(), len(str(col))) + 2
                    ws_klasifikasi.column_dimensions[ws_klasifikasi.cell(row=1, column=col_idx).column_letter].width = min(max_len, 60)
                
                df_pharma_only = df_export[df_export['Masuk Obat/Bahan Obat'] == 'YA'].copy()
                df_pharma_only.to_excel(writer, index=False, sheet_name='Obat & Bahan Baku Obat')
                ws_obat = writer.sheets['Obat & Bahan Baku Obat']
                
                for col_idx in range(1, len(df_pharma_only.columns) + 1):
                    cell = ws_obat.cell(row=1, column=col_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center')
                    cell.border = thin_border
                
                for row_idx in range(2, len(df_pharma_only) + 2):
                    for col_idx in range(1, len(df_pharma_only.columns) + 1):
                        cell = ws_obat.cell(row=row_idx, column=col_idx)
                        cell.border = thin_border
                        cell.fill = green_fill
                
                for col_idx, col in enumerate(df_pharma_only.columns, 1):
                    if len(df_pharma_only) > 0:
                        max_len = max(df_pharma_only[col].astype(str).apply(len).max(), len(str(col))) + 2
                    else:
                        max_len = len(str(col)) + 2
                    ws_obat.column_dimensions[ws_obat.cell(row=1, column=col_idx).column_letter].width = min(max_len, 60)
            
            output_hs.seek(0)
            
            st.download_button(
                label="📥 Download Hasil Klasifikasi (Excel)",
                data=output_hs,
                file_name="klasifikasi_hs_code_obat.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.markdown("---")
            st.markdown("### 🌐 Cek INSW Otomatis (Indonesia National Single Window)")
            st.markdown("Cek otomatis setiap HS Code di website INSW untuk mengetahui apakah barang tersebut memiliki **regulasi impor**, **regulasi ekspor**, atau terkait **bahan baku obat**.")
            
            all_hs_codes_insw = list(dict.fromkeys([h['hs_code'] for h in hs_items]))
            all_hs_desc_map = {h['hs_code']: h['description'] for h in hs_items}
            
            playwright_available = False
            try:
                from playwright.sync_api import sync_playwright as _pw_check
                known_gbm = "/nix/store/24w3s75aa2lrvvxsybficn8y3zxd27kp-mesa-libgbm-25.1.0/lib"
                chromium_path = os.path.expanduser("~/.cache/ms-playwright")
                if os.path.exists(chromium_path) and (os.path.exists(known_gbm + "/libgbm.so.1") or 'libgbm' in os.environ.get("LD_LIBRARY_PATH", "")):
                    playwright_available = True
            except ImportError:
                pass
            
            btn_insw = False
            if not playwright_available:
                st.warning("Browser otomatis (Playwright/Chromium) tidak tersedia. Gunakan link berikut untuk cek manual.")
                st.markdown(f"[Buka INSW Detail Komoditas](https://insw.go.id/intr/detail-komoditas)")
                insw_manual_data = [{'No': i+1, 'HS Code': h['hs_code'], 'Deskripsi': h['description']} for i, h in enumerate(hs_items[:200])]
                st.dataframe(pd.DataFrame(insw_manual_data), use_container_width=True, height=300)
                if len(hs_items) > 200:
                    st.caption(f"Menampilkan 200 dari {len(hs_items)} HS Code")
            else:
                st.info(f"Total **{len(all_hs_codes_insw)}** HS Code unik ditemukan dalam file.")
                
                insw_scope = st.radio(
                    "Pilih HS Code yang akan dicek:",
                    ["Semua HS Code", "Hanya Chapter 28-31 (Kimia & Farmasi)", "Pilih range sendiri"],
                    horizontal=True,
                    key="insw_scope"
                )
                
                if insw_scope == "Semua HS Code":
                    selected_hs_insw = all_hs_codes_insw
                elif insw_scope == "Hanya Chapter 28-31 (Kimia & Farmasi)":
                    selected_hs_insw = [h for h in all_hs_codes_insw if h[:2] in ('28', '29', '30', '31')]
                else:
                    col_range1, col_range2 = st.columns(2)
                    with col_range1:
                        start_idx = st.number_input("Dari urutan ke-", min_value=1, max_value=len(all_hs_codes_insw), value=1, key="insw_start", help="Urutan HS Code dalam file (bukan nomor HS)")
                    with col_range2:
                        end_idx = st.number_input("Sampai urutan ke-", min_value=1, max_value=len(all_hs_codes_insw), value=min(50, len(all_hs_codes_insw)), key="insw_end")
                    selected_hs_insw = all_hs_codes_insw[start_idx-1:end_idx]
                
                max_batch = st.slider("Batas maksimal HS Code per sesi", min_value=10, max_value=500, value=min(100, len(selected_hs_insw)), step=10, key="insw_batch_limit")
                
                codes_to_check = selected_hs_insw[:max_batch]
                est_seconds = len(codes_to_check) * 8
                est_minutes = max(1, est_seconds // 60)
                
                st.markdown(f"Akan mengecek **{len(codes_to_check)}** HS Code. Estimasi waktu: **~{est_minutes} menit**.")
                
                if len(codes_to_check) > 0:
                    preview_data = [{'No': i+1, 'HS Code': c, 'Deskripsi': all_hs_desc_map.get(c, '')} for i, c in enumerate(codes_to_check[:10])]
                    st.dataframe(pd.DataFrame(preview_data), use_container_width=True, height=200)
                    if len(codes_to_check) > 10:
                        st.caption(f"... dan {len(codes_to_check) - 10} HS Code lainnya")
                
                if len(codes_to_check) == 0:
                    st.warning("Tidak ada HS Code yang dipilih untuk dicek.")
                else:
                    col_insw1, col_insw2, col_insw3 = st.columns([1, 2, 1])
                    with col_insw2:
                        btn_insw = st.button("🔍 Mulai Cek INSW Otomatis", type="primary", use_container_width=True, key="btn_insw")
            
            if playwright_available and btn_insw:
                insw_results = []
                progress_insw = st.progress(0)
                status_text = st.empty()
                results_container = st.empty()
                
                try:
                    if 'libgbm' not in os.environ.get("LD_LIBRARY_PATH", ""):
                        known_gbm = "/nix/store/24w3s75aa2lrvvxsybficn8y3zxd27kp-mesa-libgbm-25.1.0/lib"
                        if os.path.exists(known_gbm + "/libgbm.so.1"):
                            os.environ["LD_LIBRARY_PATH"] = known_gbm + ":" + os.environ.get("LD_LIBRARY_PATH", "")
                        else:
                            import subprocess as sp_find
                            gbm_r = sp_find.run(["find", "/nix/store", "-name", "libgbm.so.1", "-maxdepth", "3"], capture_output=True, text=True, timeout=10)
                            gbm_paths = [p.rsplit("/", 1)[0] for p in gbm_r.stdout.strip().split("\n") if p]
                            if gbm_paths:
                                os.environ["LD_LIBRARY_PATH"] = gbm_paths[0] + ":" + os.environ.get("LD_LIBRARY_PATH", "")
                    
                    from playwright.sync_api import sync_playwright
                    import time as time_mod
                    
                    INSW_URL = "https://insw.go.id/intr/detail-komoditas"
                    OBAT_KEYWORDS = ['obat', 'farmasi', 'pharmaceutical', 'medicine', 'drug',
                                    'suplemen kesehatan', 'bahan baku obat', 'kosmetik',
                                    'vaksin', 'vitamin', 'narkotik', 'psikotropik',
                                    'kuasi', 'prekursor']
                    
                    def extract_insw_detail(pw_page, hs_code, desc_text=''):
                        entry = {
                            'HS Code': hs_code,
                            'Deskripsi': desc_text,
                            'Ada Regulasi Impor': 'Tidak',
                            'Lartas Border': 'Tidak',
                            'Tata Niaga Post Border': 'Tidak',
                            'Ada Regulasi Ekspor': 'Tidak',
                            'Komoditi INSW': '-',
                            'Terkait Obat (INSW)': 'Tidak',
                            'Ada BPOM': 'Tidak',
                            'Keterangan': '-'
                        }
                        
                        pw_page.goto(INSW_URL, timeout=30000)
                        pw_page.wait_for_selector("input[placeholder='Cari kode HS / Uraian HS']", timeout=15000)
                        
                        search_input = pw_page.query_selector("input[placeholder='Cari kode HS / Uraian HS']")
                        search_input.fill(hs_code)
                        search_input.press("Enter")
                        
                        time_mod.sleep(3)
                        pw_page.wait_for_timeout(2000)
                        
                        detail_btns = pw_page.query_selector_all("button:has-text('Detail')")
                        if not detail_btns:
                            entry['Keterangan'] = 'Tidak ditemukan di INSW'
                            return entry
                        
                        detail_btns[0].click()
                        time_mod.sleep(3)
                        pw_page.wait_for_timeout(2000)
                        
                        for _ in range(3):
                            pw_page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                            time_mod.sleep(0.5)
                        
                        body = pw_page.inner_text("body")
                        
                        has_lartas_border = "Regulasi Impor (Lartas Border)" in body
                        has_tata_niaga = "Regulasi Impor (Tata Niaga Post Border)" in body
                        has_import = has_lartas_border or has_tata_niaga or "Regulasi Impor" in body
                        has_export = "Regulasi Ekspor" in body
                        
                        entry['Ada Regulasi Impor'] = 'YA' if has_import else 'Tidak'
                        entry['Lartas Border'] = 'YA' if has_lartas_border else 'Tidak'
                        entry['Tata Niaga Post Border'] = 'YA' if has_tata_niaga else 'Tidak'
                        entry['Ada Regulasi Ekspor'] = 'YA' if has_export else 'Tidak'
                        
                        komoditi_list = []
                        is_obat = False
                        keterangan_parts = []
                        
                        lines = body.split('\n')
                        for li, line in enumerate(lines):
                            stripped = line.strip()
                            if stripped == 'Komoditi':
                                for offset in range(1, 6):
                                    if li + offset < len(lines):
                                        next_line = lines[li + offset].strip()
                                        if next_line.startswith('[') and next_line.endswith(']'):
                                            komoditi_val = next_line[1:-1]
                                            if komoditi_val and komoditi_val not in komoditi_list:
                                                komoditi_list.append(komoditi_val)
                                            break
                                        elif next_line == ':':
                                            continue
                                        elif next_line and next_line not in ('Regulasi', 'Deskripsi', ''):
                                            break
                        
                        if komoditi_list:
                            entry['Komoditi INSW'] = '; '.join(komoditi_list)
                            for k_val in komoditi_list:
                                k_lower = k_val.lower()
                                for ok in OBAT_KEYWORDS:
                                    if ok in k_lower:
                                        is_obat = True
                                        break
                        
                        body_lower = body.lower()
                        if 'bahan obat' in body_lower or 'bahan baku obat' in body_lower:
                            is_obat = True
                            keterangan_parts.append('Terkait Bahan Obat')
                        
                        has_bpom = 'BPOM' in body
                        entry['Ada BPOM'] = 'YA' if has_bpom else 'Tidak'
                        if has_bpom:
                            keterangan_parts.append('Ada regulasi BPOM')
                        
                        if has_import and has_export:
                            keterangan_parts.insert(0, 'Impor & Ekspor')
                        elif has_import:
                            keterangan_parts.insert(0, 'Hanya Impor')
                        elif has_export:
                            keterangan_parts.insert(0, 'Hanya Ekspor')
                        else:
                            keterangan_parts.insert(0, 'Tidak ada lartas')
                        
                        entry['Terkait Obat (INSW)'] = 'YA' if is_obat else 'Tidak'
                        entry['Keterangan'] = '; '.join(keterangan_parts) if keterangan_parts else '-'
                        
                        return entry
                    
                    with sync_playwright() as pw:
                        pw_browser = pw.chromium.launch(headless=True)
                        pw_page = pw_browser.new_page()
                        pw_page.set_default_timeout(20000)
                        
                        for idx_hs, hs_code in enumerate(codes_to_check):
                            progress_val = (idx_hs + 1) / len(codes_to_check)
                            progress_insw.progress(progress_val, text=f"Mengecek HS Code {hs_code} ({idx_hs+1}/{len(codes_to_check)})...")
                            status_text.text(f"Sedang memproses: {hs_code} - {all_hs_desc_map.get(hs_code, '')[:60]}")
                            
                            try:
                                result_entry = extract_insw_detail(pw_page, hs_code, all_hs_desc_map.get(hs_code, ''))
                            except Exception as e_hs:
                                result_entry = {
                                    'HS Code': hs_code,
                                    'Deskripsi': all_hs_desc_map.get(hs_code, ''),
                                    'Ada Regulasi Impor': '-',
                                    'Lartas Border': '-',
                                    'Tata Niaga Post Border': '-',
                                    'Ada Regulasi Ekspor': '-',
                                    'Komoditi INSW': '-',
                                    'Terkait Obat (INSW)': '-',
                                    'Ada BPOM': '-',
                                    'Keterangan': f'Error: {str(e_hs)[:80]}'
                                }
                            
                            insw_results.append(result_entry)
                        
                        pw_browser.close()
                    
                    progress_insw.progress(1.0, text="Selesai!")
                    status_text.text("Pengecekan INSW selesai!")
                    
                except Exception as e_insw:
                    st.error(f"Error saat mengakses INSW: {str(e_insw)}")
                    status_text.text("Terjadi error saat mengakses INSW")
                
                if insw_results:
                    df_insw_results = pd.DataFrame(insw_results)
                    
                    insw_impor_count = len(df_insw_results[df_insw_results['Ada Regulasi Impor'] == 'YA'])
                    insw_ekspor_count = len(df_insw_results[df_insw_results['Ada Regulasi Ekspor'] == 'YA'])
                    insw_obat_count = len(df_insw_results[df_insw_results['Terkait Obat (INSW)'] == 'YA'])
                    insw_bpom_count = len(df_insw_results[df_insw_results['Ada BPOM'] == 'YA'])
                    insw_no_lartas = len(df_insw_results[
                        (df_insw_results['Ada Regulasi Impor'] == 'Tidak') & 
                        (df_insw_results['Ada Regulasi Ekspor'] == 'Tidak')
                    ])
                    
                    st.markdown("#### Ringkasan Hasil INSW")
                    col_r1, col_r2, col_r3, col_r4, col_r5 = st.columns(5)
                    with col_r1:
                        st.metric("Total Dicek", len(insw_results))
                    with col_r2:
                        st.metric("Ada Regulasi Impor", insw_impor_count)
                    with col_r3:
                        st.metric("Ada Regulasi Ekspor", insw_ekspor_count)
                    with col_r4:
                        st.metric("Terkait Obat", insw_obat_count)
                    with col_r5:
                        st.metric("Tidak Ada Lartas", insw_no_lartas)
                    
                    st.markdown("#### Detail Hasil Pengecekan INSW")
                    
                    filter_hs_prefix = st.multiselect(
                        "🔎 Filter Hasil: Tampilkan HS Code dengan awalan",
                        options=["28", "29", "30", "31"],
                        default=[],
                        key="insw_filter_prefix",
                        help="Pilih satu atau lebih awalan HS Code untuk memfilter hasil. Kosongkan untuk menampilkan semua."
                    )
                    
                    if filter_hs_prefix:
                        df_insw_display = df_insw_results[df_insw_results['HS Code'].astype(str).str[:2].isin(filter_hs_prefix)]
                        st.caption(f"Menampilkan {len(df_insw_display)} dari {len(df_insw_results)} HS Code (filter: awalan {', '.join(filter_hs_prefix)})")
                    else:
                        df_insw_display = df_insw_results
                    
                    def highlight_insw(row):
                        if row.get('Terkait Obat (INSW)') == 'YA':
                            return ['background-color: #dcfce7'] * len(row)
                        elif row.get('Ada Regulasi Ekspor') == 'YA':
                            return ['background-color: #fef3c7'] * len(row)
                        elif row.get('Ada Regulasi Impor') == 'YA':
                            return ['background-color: #dbeafe'] * len(row)
                        elif 'Error' in str(row.get('Keterangan', '')):
                            return ['background-color: #fef2f2'] * len(row)
                        return [''] * len(row)
                    
                    if len(df_insw_display) > 0:
                        styled_insw = df_insw_display.style.apply(highlight_insw, axis=1)
                        st.dataframe(styled_insw, use_container_width=True, height=400)
                    else:
                        st.info("Tidak ada HS Code yang cocok dengan filter yang dipilih.")
                    
                    tab_insw_all, tab_insw_impor, tab_insw_ekspor, tab_insw_obat = st.tabs(
                        ["Semua", "Ada Regulasi Impor", "Ada Regulasi Ekspor", "Terkait Obat"]
                    )
                    
                    with tab_insw_all:
                        st.dataframe(df_insw_display, use_container_width=True, height=300)
                    
                    with tab_insw_impor:
                        df_insw_imp = df_insw_display[df_insw_display['Ada Regulasi Impor'] == 'YA']
                        if len(df_insw_imp) > 0:
                            st.dataframe(df_insw_imp, use_container_width=True, height=300)
                        else:
                            st.info("Tidak ada HS Code yang memiliki regulasi impor")
                    
                    with tab_insw_ekspor:
                        df_insw_eks = df_insw_display[df_insw_display['Ada Regulasi Ekspor'] == 'YA']
                        if len(df_insw_eks) > 0:
                            st.dataframe(df_insw_eks, use_container_width=True, height=300)
                        else:
                            st.info("Tidak ada HS Code yang memiliki regulasi ekspor")
                    
                    with tab_insw_obat:
                        df_insw_obat = df_insw_display[df_insw_display['Terkait Obat (INSW)'] == 'YA']
                        if len(df_insw_obat) > 0:
                            st.dataframe(df_insw_obat, use_container_width=True, height=300)
                        else:
                            st.info("Tidak ada HS Code yang terkait obat menurut INSW")
                    
                    output_insw = io.BytesIO()
                    with pd.ExcelWriter(output_insw, engine='openpyxl') as writer:
                        header_fill_insw = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                        header_font_insw = Font(bold=True, color='FFFFFF')
                        green_fill_insw = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                        yellow_fill_insw = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                        blue_fill_insw = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
                        thin_border_insw = Border(
                            left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin')
                        )
                        
                        df_insw_results.to_excel(writer, index=False, sheet_name='Hasil Cek INSW')
                        ws_insw = writer.sheets['Hasil Cek INSW']
                        
                        for col_idx in range(1, len(df_insw_results.columns) + 1):
                            cell = ws_insw.cell(row=1, column=col_idx)
                            cell.fill = header_fill_insw
                            cell.font = header_font_insw
                            cell.alignment = Alignment(horizontal='center')
                            cell.border = thin_border_insw
                        
                        obat_col = list(df_insw_results.columns).index('Terkait Obat (INSW)') + 1
                        ekspor_col = list(df_insw_results.columns).index('Ada Regulasi Ekspor') + 1
                        impor_col = list(df_insw_results.columns).index('Ada Regulasi Impor') + 1
                        
                        for row_idx in range(2, len(df_insw_results) + 2):
                            obat_val = ws_insw.cell(row=row_idx, column=obat_col).value
                            ekspor_val = ws_insw.cell(row=row_idx, column=ekspor_col).value
                            impor_val = ws_insw.cell(row=row_idx, column=impor_col).value
                            for col_idx in range(1, len(df_insw_results.columns) + 1):
                                cell = ws_insw.cell(row=row_idx, column=col_idx)
                                cell.border = thin_border_insw
                                if obat_val == 'YA':
                                    cell.fill = green_fill_insw
                                elif ekspor_val == 'YA':
                                    cell.fill = yellow_fill_insw
                                elif impor_val == 'YA':
                                    cell.fill = blue_fill_insw
                        
                        for col_idx, col in enumerate(df_insw_results.columns, 1):
                            max_len = max(df_insw_results[col].astype(str).apply(len).max(), len(str(col))) + 2
                            ws_insw.column_dimensions[ws_insw.cell(row=1, column=col_idx).column_letter].width = min(max_len, 60)
                    
                    output_insw.seek(0)
                    st.download_button(
                        label="Download Hasil Cek INSW (Excel)",
                        data=output_insw,
                        file_name="hasil_cek_insw.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            st.markdown("---")
            st.info("**Catatan:** Klasifikasi otomatis berdasarkan deskripsi HS Code. Untuk verifikasi lebih lanjut, cek di [INSW INTR](https://insw.go.id/intr/detail-komoditas)")
            
        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")
            st.info("Pastikan file dalam format yang benar (.xlsx atau .xls)")
    else:
        st.info("Silakan upload file data BPS untuk memulai pengecekan HS Code.")

with tab_analysis:
    st.markdown("### 📈 Analisis Data")
    st.markdown("Upload file Excel untuk menganalisis dan memvisualisasikan data Anda.")
    
    file_analysis = st.file_uploader("📁 Upload file untuk analisis", type=['xlsx', 'xls'], key="analysis")
    
    if file_analysis:
        try:
            df_analysis = pd.read_excel(file_analysis)
            
            st.markdown("---")
            st.markdown("### 📋 Preview Data")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📝 Jumlah Baris", len(df_analysis))
            with col2:
                st.metric("📊 Jumlah Kolom", len(df_analysis.columns))
            with col3:
                st.metric("📁 Ukuran Data", f"{df_analysis.memory_usage(deep=True).sum() / 1024:.1f} KB")
            
            st.dataframe(df_analysis.head(10), use_container_width=True, height=250)
            
            st.markdown("---")
            st.markdown("### ⚙️ Konfigurasi Analisis")
            
            col_list = df_analysis.columns.tolist()
            
            col1, col2 = st.columns(2)
            
            with col1:
                selected_analysis_col = st.selectbox(
                    "📌 Pilih kolom untuk dianalisis",
                    options=col_list,
                    key="analysis_col",
                    help="Pilih kolom yang ingin Anda analisis (misalnya: Negara, Jenis Obat, dll)"
                )
            
            with col2:
                top_n = st.slider("🔢 Tampilkan Top N data", min_value=5, max_value=50, value=10, key="top_n")
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                analysis_btn = st.button("🔍 Analisis Data", type="primary", use_container_width=True, key="btn_analysis")
            
            if analysis_btn:
                st.markdown("---")
                
                value_counts = df_analysis[selected_analysis_col].value_counts().head(top_n)
                
                st.markdown(f"### 📊 Top {top_n} {selected_analysis_col}")
                
                total_data = len(df_analysis)
                unique_values = df_analysis[selected_analysis_col].nunique()
                top_value = value_counts.index[0] if len(value_counts) > 0 else '-'
                top_count = value_counts.values[0] if len(value_counts) > 0 else 0
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("📝 Total Data", total_data)
                with col2:
                    st.metric("🔢 Nilai Unik", unique_values)
                with col3:
                    st.metric("🏆 Terbanyak", str(top_value)[:20])
                with col4:
                    st.metric("📊 Jumlah", top_count)
                
                st.markdown("---")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("#### 📋 Tabel Data")
                    df_counts = value_counts.reset_index()
                    df_counts.columns = [selected_analysis_col, 'Jumlah']
                    df_counts['Persentase'] = (df_counts['Jumlah'] / df_counts['Jumlah'].sum() * 100).round(2).astype(str) + '%'
                    df_counts.index = range(1, len(df_counts) + 1)
                    st.dataframe(df_counts, use_container_width=True, height=400)
                
                with col2:
                    st.markdown("#### 📊 Grafik Bar")
                    st.bar_chart(value_counts, use_container_width=True, height=400)
                
                st.markdown("---")
                st.markdown("#### 🥧 Grafik Pie")
                
                import matplotlib.pyplot as plt
                
                fig, ax = plt.subplots(figsize=(12, 8))
                colors = plt.cm.Set3(range(len(value_counts)))
                
                wedges, texts, autotexts = ax.pie(
                    value_counts.values, 
                    labels=None,
                    autopct='%1.1f%%',
                    colors=colors,
                    startangle=90,
                    explode=[0.02] * len(value_counts)
                )
                
                for autotext in autotexts:
                    autotext.set_fontsize(9)
                    autotext.set_fontweight('bold')
                
                ax.legend(
                    wedges, 
                    [f"{str(label)[:30]} ({count:,})" for label, count in zip(value_counts.index, value_counts.values)],
                    title=selected_analysis_col,
                    loc="center left",
                    bbox_to_anchor=(1, 0, 0.5, 1),
                    fontsize=9
                )
                
                ax.set_title(f"Distribusi {selected_analysis_col}", fontsize=14, fontweight='bold', pad=20)
                plt.tight_layout()
                
                st.pyplot(fig)
                
                st.markdown("---")
                st.markdown("### 📥 Download Hasil Analisis")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    output_analysis = io.BytesIO()
                    with pd.ExcelWriter(output_analysis, engine='openpyxl') as writer:
                        df_counts.to_excel(writer, index=True, sheet_name='Hasil Analisis')
                    output_analysis.seek(0)
                    
                    st.download_button(
                        label="📥 Download Data (Excel)",
                        data=output_analysis,
                        file_name=f"analisis_{selected_analysis_col}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    img_buffer = io.BytesIO()
                    fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight', facecolor='white')
                    img_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Grafik (PNG)",
                        data=img_buffer,
                        file_name=f"grafik_{selected_analysis_col}.png",
                        mime="image/png",
                        use_container_width=True
                    )
                
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan: {str(e)}")
            st.info("💡 Pastikan file Excel dalam format yang benar (.xlsx atau .xls)")
    else:
        st.info("👆 Silakan upload file Excel untuk memulai analisis data.")

st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #64748b; padding: 1rem;">
    <p>📊 Aplikasi Perbandingan Data Realisasi Impor</p>
    <p style="font-size: 0.8rem;">Dibuat dengan ❤️ menggunakan Streamlit</p>
</div>
""", unsafe_allow_html=True)
