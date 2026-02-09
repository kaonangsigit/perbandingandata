import streamlit as st
import pandas as pd
import io
import re
import os
import subprocess
import logging
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def _setup_playwright_env():
    if os.environ.get('_PLAYWRIGHT_SETUP_DONE'):
        return

    if 'libgbm' not in os.environ.get("LD_LIBRARY_PATH", ""):
        gbm_lib_dir = None
        try:
            r = subprocess.run(["pkg-config", "--libs-only-L", "gbm"],
                             capture_output=True, text=True, timeout=5)
            if r.returncode == 0 and r.stdout.strip():
                gbm_lib_dir = r.stdout.strip().replace("-L", "")
        except Exception:
            pass

        if not gbm_lib_dir:
            try:
                r = subprocess.run(["nix-build", "<nixpkgs>", "-A", "libgbm", "--no-out-link"],
                                 capture_output=True, text=True, timeout=30)
                p = r.stdout.strip()
                if p and os.path.exists(p + "/lib/libgbm.so.1"):
                    gbm_lib_dir = p + "/lib"
            except Exception:
                pass

        if gbm_lib_dir and os.path.isdir(gbm_lib_dir):
            os.environ["LD_LIBRARY_PATH"] = gbm_lib_dir + ":" + os.environ.get("LD_LIBRARY_PATH", "")

    for bpath in [os.path.expanduser("~/.cache/ms-playwright"),
                  os.path.join(os.getcwd(), ".cache/ms-playwright"),
                  "/home/runner/workspace/.cache/ms-playwright"]:
        if os.path.exists(bpath) and os.listdir(bpath):
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = bpath
            break
    else:
        try:
            bp = os.path.expanduser("~/.cache/ms-playwright")
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = bp
            subprocess.run(["python3", "-m", "playwright", "install", "chromium"],
                         capture_output=True, timeout=120, env={**os.environ, "PLAYWRIGHT_BROWSERS_PATH": bp})
        except Exception:
            pass

    os.environ['_PLAYWRIGHT_SETUP_DONE'] = '1'

_setup_playwright_env()

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
    st.markdown("### 🌐 Cek INSW Otomatis (Indonesia National Single Window)")
    st.markdown("Upload file data dari BPS, pilih chapter yang ingin dicek, lalu sistem akan mengecek otomatis di website INSW untuk mengetahui **regulasi impor**, **regulasi ekspor**, dan klasifikasi **obat/farmasi**.")

    with st.expander("📖 Petunjuk Penggunaan", expanded=False):
        st.markdown("""
        1. Upload file data dari **BPS** (format .xlsx/.xls)
        2. Sistem mendeteksi semua **chapter** HS Code dalam file
        3. **Pilih chapter** yang ingin dicek (bebas pilih berapa pun)
        4. Klik **Mulai Cek INSW Otomatis**
        5. Lihat hasil: regulasi **impor**, **ekspor**, **BPOM**, dan klasifikasi **obat**
        6. Download hasil dalam format Excel
        """)

    file_hs = st.file_uploader("📁 Upload file data BPS", type=['xlsx', 'xls'], key="hs_check")

    if file_hs:
        try:
            xls = pd.ExcelFile(file_hs)
            all_sheet_names = xls.sheet_names

            selected_sheet = all_sheet_names[0]
            if len(all_sheet_names) > 1:
                selected_sheet = st.selectbox(
                    "📄 Pilih Sheet:",
                    options=all_sheet_names,
                    key="sheet_select"
                )

            df_hs_raw = pd.read_excel(xls, sheet_name=selected_sheet, header=None, dtype=str)
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

            st.success(f"Sheet **{selected_sheet}**: Total **{len(hs_items)}** HS Code ditemukan")

            all_prefixes = sorted(list(set(h['prefix'] for h in hs_items)))
            prefix_counts = {}
            for p in all_prefixes:
                prefix_counts[p] = len([h for h in hs_items if h['prefix'] == p])

            st.markdown("---")
            st.markdown("### 📊 Pilih Chapter untuk Dicek")

            chapter_labels = []
            for p in all_prefixes:
                chapter_labels.append(f"{p} ({prefix_counts[p]} HS Code)")

            default_chapters = [lbl for lbl in chapter_labels if lbl.startswith(('28 ', '29 ', '30 ', '31 '))]

            selected_chapter_labels = st.multiselect(
                "🔎 Pilih Chapter (awalan HS Code):",
                options=chapter_labels,
                default=default_chapters,
                key="chapter_select",
                help="Pilih chapter yang ingin dicek di INSW. Bisa pilih berapa pun, bebas kombinasi."
            )

            selected_prefixes = [lbl.split(' ')[0] for lbl in selected_chapter_labels]

            hs_filtered = [h for h in hs_items if h['prefix'] in selected_prefixes]

            if selected_prefixes:
                n_cols = min(len(selected_prefixes) + 1, 6)
                cols = st.columns(n_cols)
                with cols[0]:
                    st.metric("Total Terpilih", len(hs_filtered))
                for i, p in enumerate(selected_prefixes[:n_cols-1]):
                    with cols[i + 1]:
                        st.metric(f"Chapter {p}", prefix_counts.get(p, 0))

            all_hs_desc_map = {h['hs_code']: h['description'] for h in hs_items}
            codes_to_check = list(dict.fromkeys([h['hs_code'] for h in hs_filtered]))

            st.markdown("---")

            if 'playwright_available' not in st.session_state:
                try:
                    from playwright.sync_api import sync_playwright as _pw_check
                    with _pw_check() as _pw_test:
                        _test_browser = _pw_test.chromium.launch(
                            headless=True,
                            args=['--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--single-process']
                        )
                        _test_browser.close()
                        st.session_state['playwright_available'] = True
                except Exception as _pw_err:
                    st.session_state['playwright_available'] = False
                    st.session_state['playwright_error'] = str(_pw_err)
            playwright_available = st.session_state['playwright_available']

            btn_insw = False
            if not playwright_available:
                st.warning("Browser otomatis (Playwright/Chromium) tidak tersedia. Gunakan link berikut untuk cek manual.")
                st.markdown(f"[Buka INSW Detail Komoditas](https://insw.go.id/intr/detail-komoditas)")
                insw_manual_data = [{'No': i+1, 'HS Code': h['hs_code'], 'Deskripsi': h['description']} for i, h in enumerate(hs_filtered[:200])]
                if insw_manual_data:
                    st.dataframe(pd.DataFrame(insw_manual_data), use_container_width=True, height=300)
                if len(hs_filtered) > 200:
                    st.caption(f"Menampilkan 200 dari {len(hs_filtered)} HS Code")
            else:
                if len(codes_to_check) == 0:
                    st.warning("Tidak ada HS Code yang dipilih. Pilih minimal 1 chapter di atas.")
                else:
                    est_seconds = len(codes_to_check) * 5
                    est_minutes = max(1, est_seconds // 60)

                    st.info(f"Akan mengecek **{len(codes_to_check)}** HS Code unik dari chapter **{', '.join(selected_prefixes)}**. Estimasi waktu: **~{est_minutes} menit**.")

                    if len(codes_to_check) > 0:
                        preview_data = [{'No': i+1, 'HS Code': c, 'Deskripsi': all_hs_desc_map.get(c, '')} for i, c in enumerate(codes_to_check[:10])]
                        st.dataframe(pd.DataFrame(preview_data), use_container_width=True, height=200)
                        if len(codes_to_check) > 10:
                            st.caption(f"... dan {len(codes_to_check) - 10} HS Code lainnya")

                    col_insw1, col_insw2, col_insw3 = st.columns([1, 2, 1])
                    with col_insw2:
                        btn_insw = st.button("🔍 Mulai Cek INSW Otomatis", type="primary", use_container_width=True, key="btn_insw")

            if playwright_available and btn_insw and len(codes_to_check) > 0 and not st.session_state.get('insw_running', False):
                st.session_state['insw_running'] = True
                st.session_state['insw_complete'] = False
                st.session_state.pop('insw_error', None)
                insw_temp_results = []
                st.session_state['insw_checked_prefixes'] = selected_prefixes
                progress_insw = st.progress(0)
                status_text = st.empty()
                error_container = st.empty()

                from playwright.sync_api import sync_playwright

                INSW_URL = "https://insw.go.id/intr/detail-komoditas"
                OBAT_KEYWORDS = ['obat', 'farmasi', 'pharmaceutical', 'medicine', 'drug',
                                'suplemen kesehatan', 'bahan baku obat', 'kosmetik',
                                'vaksin', 'vitamin', 'narkotik', 'psikotropik',
                                'kuasi', 'prekursor', 'narkotika', 'psikotropika']

                def format_hs_dotted(code):
                    if len(code) == 8:
                        return f"{code[:4]}.{code[4:6]}.{code[6:8]}"
                    return code

                def search_and_click_detail(pw_page, hs_code):
                    search_queries = [hs_code, format_hs_dotted(hs_code)]
                    for attempt, query in enumerate(search_queries):
                        try:
                            logger.info(f"[INSW] Searching {hs_code} with query '{query}' (attempt {attempt+1})")
                            pw_page.goto(INSW_URL, timeout=60000, wait_until='domcontentloaded')
                            pw_page.wait_for_timeout(2000)
                            search_input = pw_page.wait_for_selector("input[placeholder='Cari kode HS / Uraian HS']", timeout=20000)
                            search_input.fill(query)
                            search_input.press("Enter")

                            try:
                                pw_page.wait_for_selector("button:has-text('Detail')", timeout=20000)
                            except Exception:
                                logger.info(f"[INSW] No Detail button found for query '{query}'")
                                continue

                            pw_page.wait_for_timeout(1500)
                            body_text = pw_page.inner_text("body")
                            if hs_code not in body_text:
                                logger.info(f"[INSW] HS code {hs_code} not in search results for query '{query}'")
                                continue

                            rows = pw_page.query_selector_all("tr")
                            for row in rows:
                                row_text = row.inner_text()
                                if hs_code in row_text:
                                    detail_btn = row.query_selector("button:has-text('Detail')")
                                    if detail_btn:
                                        detail_btn.click()
                                        pw_page.wait_for_timeout(3000)
                                        logger.info(f"[INSW] Clicked Detail for {hs_code}")
                                        return True

                            detail_btns = pw_page.query_selector_all("button:has-text('Detail')")
                            if detail_btns:
                                detail_btns[0].click()
                                pw_page.wait_for_timeout(3000)
                                logger.info(f"[INSW] Clicked first Detail button for {hs_code}")
                                return True
                        except Exception as e:
                            logger.error(f"[INSW] Error searching {hs_code} with query '{query}': {str(e)[:100]}")
                            continue

                    logger.warning(f"[INSW] {hs_code} not found with any format")
                    return False

                def extract_insw_detail(pw_page, hs_code, desc_text=''):
                    entry = {
                        'HS Code': hs_code,
                        'Deskripsi': desc_text,
                        'Jenis': '-',
                        'Ada Regulasi Impor': 'Tidak',
                        'Lartas Border': 'Tidak',
                        'Tata Niaga Post Border': 'Tidak',
                        'Ada Regulasi Ekspor': 'Tidak',
                        'Lartas Ekspor': 'Tidak',
                        'Komoditi INSW': '-',
                        'Terkait Obat (INSW)': 'Tidak',
                        'Ada BPOM': 'Tidak',
                        'Keterangan Impor': '-',
                        'Keterangan Ekspor': '-',
                    }

                    found = search_and_click_detail(pw_page, hs_code)
                    if not found:
                        entry['Jenis'] = 'Tidak ditemukan'
                        entry['Keterangan Impor'] = 'Tidak ditemukan di INSW'
                        entry['Keterangan Ekspor'] = 'Tidak ditemukan di INSW'
                        return entry

                    pw_page.evaluate("window.scrollTo(0, document.body.scrollHeight)")

                    body = pw_page.inner_text("body")

                    has_lartas_border = "Regulasi Impor (Lartas Border)" in body
                    has_tata_niaga = "Regulasi Impor (Tata Niaga Post Border)" in body
                    has_import = has_lartas_border or has_tata_niaga or "Regulasi Impor" in body
                    has_lartas_ekspor = "Regulasi Ekspor (Lartas Ekspor)" in body or "Lartas Ekspor" in body
                    has_export = has_lartas_ekspor or "Regulasi Ekspor" in body

                    entry['Ada Regulasi Impor'] = 'YA' if has_import else 'Tidak'
                    entry['Lartas Border'] = 'YA' if has_lartas_border else 'Tidak'
                    entry['Tata Niaga Post Border'] = 'YA' if has_tata_niaga else 'Tidak'
                    entry['Ada Regulasi Ekspor'] = 'YA' if has_export else 'Tidak'
                    entry['Lartas Ekspor'] = 'YA' if has_lartas_ekspor else 'Tidak'

                    komoditi_list = []
                    is_obat = False
                    ket_impor_parts = []
                    ket_ekspor_parts = []

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

                    has_bpom = 'BPOM' in body
                    entry['Ada BPOM'] = 'YA' if has_bpom else 'Tidak'

                    if has_lartas_border:
                        ket_impor_parts.append('Lartas Border')
                    if has_tata_niaga:
                        ket_impor_parts.append('Tata Niaga Post Border')
                    if has_bpom:
                        ket_impor_parts.append('BPOM')
                    if is_obat:
                        ket_impor_parts.append('Terkait Obat/Farmasi')

                    if has_lartas_ekspor:
                        ket_ekspor_parts.append('Lartas Ekspor')

                    entry['Keterangan Impor'] = '; '.join(ket_impor_parts) if ket_impor_parts else 'Tidak ada regulasi impor'
                    entry['Keterangan Ekspor'] = '; '.join(ket_ekspor_parts) if ket_ekspor_parts else 'Tidak ada regulasi ekspor'

                    if has_import and has_export:
                        entry['Jenis'] = 'IMPOR & EKSPOR'
                    elif has_import:
                        entry['Jenis'] = 'IMPOR'
                    elif has_export:
                        entry['Jenis'] = 'EKSPOR'
                    else:
                        entry['Jenis'] = 'Tidak ada lartas'

                    entry['Terkait Obat (INSW)'] = 'YA' if is_obat else 'Tidak'

                    return entry

                pw_browser = None
                error_count = 0
                max_retries = 3

                BROWSER_ARGS = [
                    '--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu',
                    '--single-process', '--disable-extensions',
                    '--disable-background-networking',
                    '--disable-software-rasterizer',
                    '--disable-translate',
                    '--no-first-run',
                    '--no-zygote',
                ]

                try:
                    with sync_playwright() as pw:
                        logger.info("[INSW] Launching Chromium browser...")
                        pw_browser = pw.chromium.launch(
                            headless=True,
                            args=BROWSER_ARGS
                        )
                        pw_page = pw_browser.new_page()
                        pw_page.set_default_timeout(60000)
                        logger.info("[INSW] Browser launched successfully")

                        for idx_hs, hs_code in enumerate(codes_to_check):
                            progress_val = (idx_hs + 1) / len(codes_to_check)
                            progress_insw.progress(progress_val, text=f"Mengecek HS Code {hs_code} ({idx_hs+1}/{len(codes_to_check)})...")
                            status_text.text(f"Sedang memproses: {hs_code} - {all_hs_desc_map.get(hs_code, '')[:60]}")

                            last_error_msg = ''
                            result_entry = None
                            for retry in range(max_retries + 1):
                                try:
                                    result_entry = extract_insw_detail(pw_page, hs_code, all_hs_desc_map.get(hs_code, ''))
                                    break
                                except Exception as e_hs:
                                    last_error_msg = str(e_hs)[:120]
                                    logger.error(f"[INSW] Error on {hs_code} retry {retry}: {last_error_msg}")
                                    if retry < max_retries:
                                        try:
                                            pw_page.close()
                                        except Exception:
                                            pass
                                        try:
                                            pw_browser.close()
                                        except Exception:
                                            pass
                                        try:
                                            pw_page.wait_for_timeout(2000)
                                        except Exception:
                                            import time; time.sleep(2)
                                        try:
                                            pw_browser = pw.chromium.launch(
                                                headless=True,
                                                args=BROWSER_ARGS
                                            )
                                            pw_page = pw_browser.new_page()
                                            pw_page.set_default_timeout(60000)
                                            logger.info(f"[INSW] Browser restarted for retry {retry+1}")
                                        except Exception as e_launch:
                                            last_error_msg = f'Browser restart error: {str(e_launch)[:80]}'
                                            logger.error(f"[INSW] {last_error_msg}")
                                            break

                            if result_entry is None:
                                error_count += 1
                                result_entry = {
                                    'HS Code': hs_code,
                                    'Deskripsi': all_hs_desc_map.get(hs_code, ''),
                                    'Jenis': 'Error',
                                    'Ada Regulasi Impor': '-', 'Lartas Border': '-',
                                    'Tata Niaga Post Border': '-', 'Ada Regulasi Ekspor': '-',
                                    'Lartas Ekspor': '-', 'Komoditi INSW': '-',
                                    'Terkait Obat (INSW)': '-', 'Ada BPOM': '-',
                                    'Keterangan Impor': f'Error: {last_error_msg}',
                                    'Keterangan Ekspor': '-',
                                }

                            insw_temp_results.append(result_entry)
                            st.session_state['insw_results'] = list(insw_temp_results)

                        try:
                            pw_browser.close()
                        except Exception:
                            pass

                    st.session_state['insw_results'] = insw_temp_results
                    st.session_state['insw_running'] = False
                    st.session_state['insw_complete'] = True
                    progress_insw.progress(1.0, text="Selesai!")
                    total_checked = len(insw_temp_results)
                    total_expected = len(codes_to_check)
                    if error_count > 0:
                        status_text.text(f"Selesai! {total_checked}/{total_expected} HS Code dicek ({error_count} error)")
                    else:
                        status_text.text(f"Pengecekan INSW selesai! {total_checked}/{total_expected} HS Code dicek")
                    st.rerun()

                except Exception as e_insw:
                    error_detail = str(e_insw)
                    logger.error(f"[INSW] Fatal error: {error_detail}")
                    if insw_temp_results:
                        st.session_state['insw_results'] = insw_temp_results
                        st.session_state['insw_complete'] = True
                        st.session_state['insw_error'] = f"Proses terhenti setelah {len(insw_temp_results)} HS Code. Error: {error_detail[:150]}"
                    else:
                        if 'Executable doesn' in error_detail or 'browser' in error_detail.lower():
                            st.session_state['insw_error'] = f"Browser Chromium tidak dapat dijalankan. Silakan coba lagi atau hubungi admin. Detail: {error_detail[:120]}"
                        elif 'timeout' in error_detail.lower() or 'Timeout' in error_detail:
                            st.session_state['insw_error'] = f"Koneksi ke INSW timeout. Website INSW mungkin sedang lambat. Silakan coba lagi. Detail: {error_detail[:120]}"
                        else:
                            st.session_state['insw_error'] = f"Error saat mengakses INSW: {error_detail[:150]}"
                    st.session_state['insw_running'] = False
                    try:
                        if pw_browser:
                            pw_browser.close()
                    except Exception:
                        pass
                    st.rerun()

            if st.session_state.get('insw_error'):
                err_msg = st.session_state.pop('insw_error')
                if st.session_state.get('insw_results'):
                    st.warning(err_msg + " Hasil parsial ditampilkan di bawah.")
                else:
                    st.error(err_msg)

            insw_results_stored = st.session_state.get('insw_results', [])
            if insw_results_stored:
                df_insw_results = pd.DataFrame(insw_results_stored)

                st.markdown("---")
                st.markdown("### 📊 Hasil Pengecekan INSW")

                insw_impor_count = len(df_insw_results[df_insw_results['Ada Regulasi Impor'] == 'YA'])
                insw_ekspor_count = len(df_insw_results[df_insw_results['Ada Regulasi Ekspor'] == 'YA'])
                insw_obat_count = len(df_insw_results[df_insw_results['Terkait Obat (INSW)'] == 'YA'])
                insw_bpom_count = len(df_insw_results[df_insw_results['Ada BPOM'] == 'YA'])
                insw_both_count = len(df_insw_results[
                    (df_insw_results['Ada Regulasi Impor'] == 'YA') &
                    (df_insw_results['Ada Regulasi Ekspor'] == 'YA')
                ])
                insw_no_lartas = len(df_insw_results[
                    (df_insw_results['Ada Regulasi Impor'] == 'Tidak') &
                    (df_insw_results['Ada Regulasi Ekspor'] == 'Tidak')
                ])

                col_r1, col_r2, col_r3, col_r4, col_r5, col_r6 = st.columns(6)
                with col_r1:
                    st.metric("Total Dicek", len(insw_results_stored))
                with col_r2:
                    st.metric("Regulasi Impor", insw_impor_count)
                with col_r3:
                    st.metric("Regulasi Ekspor", insw_ekspor_count)
                with col_r4:
                    st.metric("Terkait Obat", insw_obat_count)
                with col_r5:
                    st.metric("Ada BPOM", insw_bpom_count)
                with col_r6:
                    st.metric("Tidak Ada Lartas", insw_no_lartas)

                st.markdown("---")

                result_prefixes = sorted(list(set(str(r.get('HS Code', ''))[:2] for r in insw_results_stored if r.get('HS Code', ''))))

                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    filter_hs_prefix = st.multiselect(
                        "🔎 Filter per Chapter:",
                        options=result_prefixes,
                        default=[],
                        key="insw_filter_prefix",
                        help="Filter berdasarkan chapter. Kosongkan untuk menampilkan semua."
                    )
                with col_f2:
                    filter_insw_type = st.multiselect(
                        "🔎 Filter berdasarkan hasil:",
                        options=["Ada Regulasi Impor", "Ada Regulasi Ekspor", "Impor & Ekspor", "Terkait Obat", "Ada BPOM", "Tidak Ada Lartas"],
                        default=[],
                        key="insw_filter_type",
                        help="Filter berdasarkan jenis regulasi. Kosongkan untuk menampilkan semua."
                    )

                df_insw_display = df_insw_results

                if filter_hs_prefix:
                    df_insw_display = df_insw_display[df_insw_display['HS Code'].astype(str).str[:2].isin(filter_hs_prefix)]

                if filter_insw_type:
                    mask = pd.Series([False] * len(df_insw_display), index=df_insw_display.index)
                    if "Ada Regulasi Impor" in filter_insw_type:
                        mask = mask | (df_insw_display['Ada Regulasi Impor'] == 'YA')
                    if "Ada Regulasi Ekspor" in filter_insw_type:
                        mask = mask | (df_insw_display['Ada Regulasi Ekspor'] == 'YA')
                    if "Impor & Ekspor" in filter_insw_type:
                        mask = mask | ((df_insw_display['Ada Regulasi Impor'] == 'YA') & (df_insw_display['Ada Regulasi Ekspor'] == 'YA'))
                    if "Terkait Obat" in filter_insw_type:
                        mask = mask | (df_insw_display['Terkait Obat (INSW)'] == 'YA')
                    if "Ada BPOM" in filter_insw_type:
                        mask = mask | (df_insw_display['Ada BPOM'] == 'YA')
                    if "Tidak Ada Lartas" in filter_insw_type:
                        mask = mask | ((df_insw_display['Ada Regulasi Impor'] == 'Tidak') & (df_insw_display['Ada Regulasi Ekspor'] == 'Tidak'))
                    df_insw_display = df_insw_display[mask]

                if filter_hs_prefix or filter_insw_type:
                    st.caption(f"Menampilkan {len(df_insw_display)} dari {len(df_insw_results)} HS Code")

                tab_insw_all, tab_insw_impor, tab_insw_ekspor, tab_insw_obat, tab_insw_bpom = st.tabs(
                    ["📋 Semua", "📦 Regulasi Impor", "🚢 Regulasi Ekspor", "💊 Terkait Obat", "🏥 Ada BPOM"]
                )

                def highlight_insw(row):
                    jenis = str(row.get('Jenis', ''))
                    if row.get('Terkait Obat (INSW)') == 'YA':
                        return ['background-color: #dcfce7'] * len(row)
                    elif jenis == 'IMPOR & EKSPOR':
                        return ['background-color: #fce7f3'] * len(row)
                    elif row.get('Ada Regulasi Ekspor') == 'YA':
                        return ['background-color: #fef3c7'] * len(row)
                    elif row.get('Ada Regulasi Impor') == 'YA':
                        return ['background-color: #dbeafe'] * len(row)
                    elif 'Error' in str(row.get('Keterangan Impor', '')):
                        return ['background-color: #fef2f2'] * len(row)
                    return [''] * len(row)

                with tab_insw_all:
                    if len(df_insw_display) > 0:
                        st.markdown("**Legenda warna:** 🟢 Terkait Obat | 🩷 Impor & Ekspor | 🔵 Impor | 🟡 Ekspor | ⬜ Tidak ada lartas")
                        styled_insw = df_insw_display.style.apply(highlight_insw, axis=1)
                        st.dataframe(styled_insw, use_container_width=True, height=400)
                    else:
                        st.info("Tidak ada HS Code yang cocok dengan filter yang dipilih.")

                with tab_insw_impor:
                    df_insw_imp = df_insw_display[df_insw_display['Ada Regulasi Impor'] == 'YA']
                    if len(df_insw_imp) > 0:
                        st.success(f"**{len(df_insw_imp)}** HS Code memiliki regulasi impor")
                        impor_cols = ['HS Code', 'Deskripsi', 'Jenis', 'Lartas Border', 'Tata Niaga Post Border', 'Komoditi INSW', 'Ada BPOM', 'Terkait Obat (INSW)', 'Keterangan Impor']
                        display_cols = [c for c in impor_cols if c in df_insw_imp.columns]
                        st.dataframe(df_insw_imp[display_cols], use_container_width=True, height=400)
                    else:
                        st.info("Tidak ada HS Code yang memiliki regulasi impor")

                with tab_insw_ekspor:
                    df_insw_eks = df_insw_display[df_insw_display['Ada Regulasi Ekspor'] == 'YA']
                    if len(df_insw_eks) > 0:
                        st.success(f"**{len(df_insw_eks)}** HS Code memiliki regulasi ekspor")
                        ekspor_cols = ['HS Code', 'Deskripsi', 'Jenis', 'Lartas Ekspor', 'Komoditi INSW', 'Keterangan Ekspor']
                        display_cols = [c for c in ekspor_cols if c in df_insw_eks.columns]
                        st.dataframe(df_insw_eks[display_cols], use_container_width=True, height=400)
                    else:
                        st.info("Tidak ada HS Code yang memiliki regulasi ekspor")

                with tab_insw_obat:
                    df_insw_obat_data = df_insw_display[df_insw_display['Terkait Obat (INSW)'] == 'YA']
                    if len(df_insw_obat_data) > 0:
                        st.success(f"**{len(df_insw_obat_data)}** HS Code terkait obat/farmasi")
                        obat_cols = ['HS Code', 'Deskripsi', 'Jenis', 'Komoditi INSW', 'Ada BPOM', 'Keterangan Impor', 'Keterangan Ekspor']
                        display_cols = [c for c in obat_cols if c in df_insw_obat_data.columns]
                        st.dataframe(df_insw_obat_data[display_cols], use_container_width=True, height=400)
                    else:
                        st.info("Tidak ada HS Code yang terkait obat menurut INSW")

                with tab_insw_bpom:
                    df_insw_bpom_data = df_insw_display[df_insw_display['Ada BPOM'] == 'YA']
                    if len(df_insw_bpom_data) > 0:
                        st.success(f"**{len(df_insw_bpom_data)}** HS Code memiliki regulasi BPOM")
                        bpom_cols = ['HS Code', 'Deskripsi', 'Jenis', 'Komoditi INSW', 'Ada BPOM', 'Keterangan Impor']
                        display_cols = [c for c in bpom_cols if c in df_insw_bpom_data.columns]
                        st.dataframe(df_insw_bpom_data[display_cols], use_container_width=True, height=400)
                    else:
                        st.info("Tidak ada HS Code yang memiliki regulasi BPOM")

                st.markdown("---")
                st.markdown("### 📥 Download Hasil INSW")

                output_insw = io.BytesIO()
                with pd.ExcelWriter(output_insw, engine='openpyxl') as writer:
                    header_fill_insw = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                    header_font_insw = Font(bold=True, color='FFFFFF')
                    green_fill_insw = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                    yellow_fill_insw = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                    blue_fill_insw = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
                    pink_fill_insw = PatternFill(start_color='FCE4EC', end_color='FCE4EC', fill_type='solid')
                    thin_border_insw = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )

                    df_insw_results.to_excel(writer, index=False, sheet_name='Semua Hasil INSW')
                    ws_insw = writer.sheets['Semua Hasil INSW']

                    for col_idx in range(1, len(df_insw_results.columns) + 1):
                        cell = ws_insw.cell(row=1, column=col_idx)
                        cell.fill = header_fill_insw
                        cell.font = header_font_insw
                        cell.alignment = Alignment(horizontal='center')
                        cell.border = thin_border_insw

                    obat_col = list(df_insw_results.columns).index('Terkait Obat (INSW)') + 1
                    jenis_col = list(df_insw_results.columns).index('Jenis') + 1
                    ekspor_col = list(df_insw_results.columns).index('Ada Regulasi Ekspor') + 1
                    impor_col = list(df_insw_results.columns).index('Ada Regulasi Impor') + 1

                    for row_idx in range(2, len(df_insw_results) + 2):
                        obat_val = ws_insw.cell(row=row_idx, column=obat_col).value
                        jenis_val = ws_insw.cell(row=row_idx, column=jenis_col).value
                        ekspor_val = ws_insw.cell(row=row_idx, column=ekspor_col).value
                        impor_val = ws_insw.cell(row=row_idx, column=impor_col).value
                        for col_idx in range(1, len(df_insw_results.columns) + 1):
                            cell = ws_insw.cell(row=row_idx, column=col_idx)
                            cell.border = thin_border_insw
                            if obat_val == 'YA':
                                cell.fill = green_fill_insw
                            elif jenis_val == 'IMPOR & EKSPOR':
                                cell.fill = pink_fill_insw
                            elif ekspor_val == 'YA':
                                cell.fill = yellow_fill_insw
                            elif impor_val == 'YA':
                                cell.fill = blue_fill_insw

                    for col_idx, col in enumerate(df_insw_results.columns, 1):
                        max_len = max(df_insw_results[col].astype(str).apply(len).max(), len(str(col))) + 2
                        ws_insw.column_dimensions[ws_insw.cell(row=1, column=col_idx).column_letter].width = min(max_len, 60)

                    df_impor_only = df_insw_results[df_insw_results['Ada Regulasi Impor'] == 'YA'].copy()
                    if len(df_impor_only) > 0:
                        df_impor_only.to_excel(writer, index=False, sheet_name='Regulasi Impor')
                        ws_imp = writer.sheets['Regulasi Impor']
                        for col_idx in range(1, len(df_impor_only.columns) + 1):
                            cell = ws_imp.cell(row=1, column=col_idx)
                            cell.fill = header_fill_insw
                            cell.font = header_font_insw
                            cell.alignment = Alignment(horizontal='center')
                            cell.border = thin_border_insw
                        for row_idx in range(2, len(df_impor_only) + 2):
                            for col_idx in range(1, len(df_impor_only.columns) + 1):
                                ws_imp.cell(row=row_idx, column=col_idx).border = thin_border_insw
                                ws_imp.cell(row=row_idx, column=col_idx).fill = blue_fill_insw

                    df_ekspor_only = df_insw_results[df_insw_results['Ada Regulasi Ekspor'] == 'YA'].copy()
                    if len(df_ekspor_only) > 0:
                        df_ekspor_only.to_excel(writer, index=False, sheet_name='Regulasi Ekspor')
                        ws_eks = writer.sheets['Regulasi Ekspor']
                        for col_idx in range(1, len(df_ekspor_only.columns) + 1):
                            cell = ws_eks.cell(row=1, column=col_idx)
                            cell.fill = header_fill_insw
                            cell.font = header_font_insw
                            cell.alignment = Alignment(horizontal='center')
                            cell.border = thin_border_insw
                        for row_idx in range(2, len(df_ekspor_only) + 2):
                            for col_idx in range(1, len(df_ekspor_only.columns) + 1):
                                ws_eks.cell(row=row_idx, column=col_idx).border = thin_border_insw
                                ws_eks.cell(row=row_idx, column=col_idx).fill = yellow_fill_insw

                    df_obat_only = df_insw_results[df_insw_results['Terkait Obat (INSW)'] == 'YA'].copy()
                    if len(df_obat_only) > 0:
                        df_obat_only.to_excel(writer, index=False, sheet_name='Terkait Obat')
                        ws_obat = writer.sheets['Terkait Obat']
                        for col_idx in range(1, len(df_obat_only.columns) + 1):
                            cell = ws_obat.cell(row=1, column=col_idx)
                            cell.fill = header_fill_insw
                            cell.font = header_font_insw
                            cell.alignment = Alignment(horizontal='center')
                            cell.border = thin_border_insw
                        for row_idx in range(2, len(df_obat_only) + 2):
                            for col_idx in range(1, len(df_obat_only.columns) + 1):
                                ws_obat.cell(row=row_idx, column=col_idx).border = thin_border_insw
                                ws_obat.cell(row=row_idx, column=col_idx).fill = green_fill_insw

                output_insw.seek(0)
                st.download_button(
                    label="📥 Download Hasil Cek INSW (Excel)",
                    data=output_insw,
                    file_name="hasil_cek_insw.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")
            st.info("Pastikan file dalam format yang benar (.xlsx atau .xls)")
    else:
        st.info("Silakan upload file data BPS untuk memulai pengecekan INSW.")

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
