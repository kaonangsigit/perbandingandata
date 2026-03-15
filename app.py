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

import threading as _threading
import json as _json
_insw_threads = {}

def _insw_state_path(sid):
    return f"/tmp/insw_{sid}.json"

def _write_insw_state(sid, state):
    path = _insw_state_path(sid)
    tmp_path = path + ".tmp"
    try:
        with open(tmp_path, 'w') as f:
            _json.dump(state, f)
        os.rename(tmp_path, path)
    except Exception as e:
        logger.error(f"[INSW] Failed to write state file: {e}")

def _read_insw_state(sid):
    path = _insw_state_path(sid)
    try:
        with open(path, 'r') as f:
            return _json.load(f)
    except (FileNotFoundError, _json.JSONDecodeError, OSError):
        return {}

def _cleanup_insw_state(sid):
    path = _insw_state_path(sid)
    try:
        os.remove(path)
    except (FileNotFoundError, OSError):
        pass

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

if os.environ.get('ENABLE_PLAYWRIGHT_SETUP') == '1':
    _setup_playwright_env()

st.set_page_config(
    page_title="Perbandingan Data Impor", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    :root {
        --bpom-green: #0f766e;
        --bpom-green-2: #16a34a;
        --bpom-bg: #f6fbf9;
        --bpom-card: #ffffff;
        --bpom-text: #0f172a;
        --bpom-muted: #475569;
        --bpom-border: #e2e8f0;
        --bpom-shadow: 0 6px 24px rgba(2, 44, 34, 0.08);
        --bpom-radius: 14px;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }

    .bpom-hero {
        background: linear-gradient(90deg, rgba(15,118,110,0.12), rgba(22,163,74,0.10));
        border: 1px solid var(--bpom-border);
        border-radius: var(--bpom-radius);
        padding: 18px 18px;
        box-shadow: var(--bpom-shadow);
        margin-bottom: 14px;
    }
    .bpom-hero-title {
        font-size: 2.05rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        color: var(--bpom-text);
        margin: 0;
        line-height: 1.1;
    }
    .bpom-hero-sub {
        font-size: 1rem;
        color: var(--bpom-muted);
        margin: 6px 0 0 0;
    }
    .bpom-badge {
        display: inline-block;
        font-size: 0.85rem;
        font-weight: 700;
        color: var(--bpom-green);
        background: rgba(15,118,110,0.10);
        border: 1px solid rgba(15,118,110,0.20);
        padding: 4px 10px;
        border-radius: 999px;
        margin-bottom: 10px;
    }

    .main-header {
        display: none;
    }
    .sub-header {
        display: none;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding: 10px 20px;
        background-color: rgba(15,118,110,0.06);
        border: 1px solid rgba(15,118,110,0.15);
        border-radius: 999px;
        font-weight: 700;
        color: var(--bpom-green);
    }
    .stTabs [aria-selected="true"] {
        background-color: var(--bpom-green);
        color: white;
    }
    .upload-section {
        background-color: var(--bpom-card);
        padding: 1.5rem;
        border-radius: var(--bpom-radius);
        border: 1px dashed rgba(15,118,110,0.35);
        margin-bottom: 1rem;
        box-shadow: var(--bpom-shadow);
    }
    .success-box {
        background-color: #dcfce7;
        border-left: 4px solid #22c55e;
        padding: 1rem;
        border-radius: 10px;
    }
    .warning-box {
        background-color: #fef3c7;
        border-left: 4px solid #f59e0b;
        padding: 1rem;
        border-radius: 10px;
    }
    .info-box {
        background-color: #e0f2fe;
        border-left: 4px solid #0ea5e9;
        padding: 1rem;
        border-radius: 10px;
    }

    div[data-testid="stMetric"] {
        background: var(--bpom-card);
        border: 1px solid var(--bpom-border);
        border-radius: var(--bpom-radius);
        padding: 12px 12px;
        box-shadow: var(--bpom-shadow);
    }

    div[data-testid="stExpander"] > details {
        background: var(--bpom-card);
        border: 1px solid var(--bpom-border);
        border-radius: var(--bpom-radius);
        box-shadow: var(--bpom-shadow);
    }

    .stButton>button {
        border-radius: 999px;
        font-weight: 700;
        border: 1px solid rgba(15,118,110,0.35);
    }
    .stButton>button[kind="primary"] {
        background: var(--bpom-green);
        border: 1px solid var(--bpom-green);
    }
    .stButton>button[kind="primary"]:hover {
        background: #0b5f58;
        border-color: #0b5f58;
    }
</style>
""", unsafe_allow_html=True)

st.markdown(
    """
    <div class="bpom-hero">
        <div class="bpom-badge">BPOM • Analisis & Perbandingan Data</div>
        <p class="bpom-hero-title">Perbandingan Data Realisasi Impor</p>
        <p class="bpom-hero-sub">Aplikasi untuk membandingkan dan menganalisis data impor secara cepat, rapi, dan mudah di-review</p>
    </div>
    """,
    unsafe_allow_html=True
)

tab_main, tab_hs, tab_analysis, tab_petugas, tab_absen, tab_merge = st.tabs(["📋 Perbandingan Data", "💊 Cek HS Code Obat", "📈 Analisis Data", "👤 Cek Petugas Loket S2", "📋 Cek Kehadiran", "🔗 Gabung Data Excel"])

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
                st.session_state['insw_results'] = []
                st.session_state['insw_progress_current'] = 0
                st.session_state['insw_progress_total'] = len(codes_to_check)
                st.session_state['insw_progress_hs'] = ''
                st.session_state['insw_progress_desc'] = ''
                st.session_state['insw_checked_prefixes'] = selected_prefixes
                st.session_state['insw_codes_to_check'] = list(codes_to_check)
                st.session_state['insw_desc_map'] = dict(all_hs_desc_map)

                import threading

                def _run_insw_scraping(codes, desc_map, session_id):
                    import time as _time

                    _file_state = {
                        'results': [],
                        'complete': False,
                        'current': 0,
                        'total': len(codes),
                        'current_hs': '',
                        'current_desc': '',
                        'error_count': 0,
                        'error_msg': '',
                        'heartbeat': _time.time(),
                        'status': 'running',
                    }
                    _write_insw_state(session_id, _file_state)

                    def _update_shared(key, value):
                        _file_state[key] = value
                        _file_state['heartbeat'] = _time.time()
                        _write_insw_state(session_id, _file_state)

                    def _update_shared_multi(updates):
                        _file_state.update(updates)
                        _file_state['heartbeat'] = _time.time()
                        _write_insw_state(session_id, _file_state)

                    INSW_URL = "https://insw.go.id/intr/detail-komoditas"
                    OBAT_KEYWORDS = ['obat', 'farmasi', 'pharmaceutical', 'medicine', 'drug',
                                    'suplemen kesehatan', 'bahan baku obat', 'kosmetik',
                                    'vaksin', 'vitamin', 'narkotik', 'psikotropik',
                                    'kuasi', 'prekursor', 'narkotika', 'psikotropika']
                    BROWSER_ARGS = [
                        '--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu',
                        '--single-process', '--disable-extensions',
                        '--disable-background-networking',
                        '--disable-software-rasterizer',
                        '--disable-translate',
                        '--no-first-run',
                        '--no-zygote',
                    ]
                    max_retries = 3

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
                                    return True
                            except Exception as e:
                                logger.error(f"[INSW] Error searching {hs_code} with query '{query}': {str(e)[:100]}")
                                continue
                        return False

                    def extract_insw_detail(pw_page, hs_code, desc_text=''):
                        entry = {
                            'HS Code': hs_code, 'Deskripsi': desc_text, 'Jenis': '-',
                            'Ada Regulasi Impor': 'Tidak', 'Lartas Border': 'Tidak',
                            'Tata Niaga Post Border': 'Tidak', 'Ada Regulasi Ekspor': 'Tidak',
                            'Lartas Ekspor': 'Tidak', 'Komoditi INSW': '-',
                            'Terkait Obat (INSW)': 'Tidak', 'Ada BPOM': 'Tidak',
                            'Keterangan Impor': '-', 'Keterangan Ekspor': '-',
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

                    results = []
                    error_count = 0
                    pw_browser = None

                    try:
                        from playwright.sync_api import sync_playwright
                        with sync_playwright() as pw:
                            logger.info("[INSW-Thread] Launching Chromium browser...")
                            pw_browser = pw.chromium.launch(headless=True, args=BROWSER_ARGS)
                            pw_page = pw_browser.new_page()
                            pw_page.set_default_timeout(60000)
                            logger.info("[INSW-Thread] Browser launched successfully")

                            for idx_hs, hs_code in enumerate(codes):
                                _update_shared_multi({
                                    'current': idx_hs + 1,
                                    'current_hs': hs_code,
                                    'current_desc': desc_map.get(hs_code, '')[:60],
                                })

                                last_error_msg = ''
                                result_entry = None
                                for retry in range(max_retries + 1):
                                    try:
                                        result_entry = extract_insw_detail(pw_page, hs_code, desc_map.get(hs_code, ''))
                                        break
                                    except Exception as e_hs:
                                        last_error_msg = str(e_hs)[:120]
                                        logger.error(f"[INSW-Thread] Error on {hs_code} retry {retry}: {last_error_msg}")
                                        if retry < max_retries:
                                            try:
                                                pw_page.close()
                                            except Exception:
                                                pass
                                            try:
                                                pw_browser.close()
                                            except Exception:
                                                pass
                                            _time.sleep(2)
                                            try:
                                                pw_browser = pw.chromium.launch(headless=True, args=BROWSER_ARGS)
                                                pw_page = pw_browser.new_page()
                                                pw_page.set_default_timeout(60000)
                                                logger.info(f"[INSW-Thread] Browser restarted for retry {retry+1}")
                                            except Exception as e_launch:
                                                last_error_msg = f'Browser restart error: {str(e_launch)[:80]}'
                                                logger.error(f"[INSW-Thread] {last_error_msg}")
                                                break

                                if result_entry is None:
                                    error_count += 1
                                    result_entry = {
                                        'HS Code': hs_code, 'Deskripsi': desc_map.get(hs_code, ''),
                                        'Jenis': 'Error',
                                        'Ada Regulasi Impor': '-', 'Lartas Border': '-',
                                        'Tata Niaga Post Border': '-', 'Ada Regulasi Ekspor': '-',
                                        'Lartas Ekspor': '-', 'Komoditi INSW': '-',
                                        'Terkait Obat (INSW)': '-', 'Ada BPOM': '-',
                                        'Keterangan Impor': f'Error: {last_error_msg}',
                                        'Keterangan Ekspor': '-',
                                    }

                                results.append(result_entry)
                                _update_shared('results', list(results))

                            try:
                                pw_browser.close()
                            except Exception:
                                pass

                        _update_shared_multi({
                            'results': results,
                            'complete': True,
                            'error_count': error_count,
                            'status': 'completed',
                        })
                        logger.info(f"[INSW-Thread] Completed. {len(results)}/{len(codes)} checked, {error_count} errors")

                    except Exception as e_insw:
                        error_detail = str(e_insw)
                        logger.error(f"[INSW-Thread] Fatal error: {error_detail}")
                        _update_shared_multi({
                            'results': results,
                            'complete': True,
                            'error_msg': error_detail[:200],
                            'error_count': error_count,
                            'status': 'error',
                        })
                        try:
                            if pw_browser:
                                pw_browser.close()
                        except Exception:
                            pass

                import uuid
                sid = str(uuid.uuid4())[:8]
                st.session_state['insw_session_id'] = sid

                st.session_state['insw_thread_started'] = True

                t = threading.Thread(
                    target=_run_insw_scraping,
                    args=(list(codes_to_check), dict(all_hs_desc_map), sid),
                    daemon=True
                )
                t.start()
                _insw_threads[sid] = t
                st.rerun()

            if st.session_state.get('insw_running', False) and st.session_state.get('insw_thread_started', False):
                import time as _time
                sid = st.session_state.get('insw_session_id', '')

                shared = _read_insw_state(sid)

                thread = _insw_threads.get(sid)
                thread_alive = thread is not None and thread.is_alive()

                file_status = shared.get('status', '')
                heartbeat = shared.get('heartbeat', 0)
                heartbeat_stale = (heartbeat > 0 and (_time.time() - heartbeat) > 60)

                if not shared and not thread_alive:
                    st.session_state['insw_running'] = False
                    st.session_state.pop('insw_thread_started', None)
                    st.session_state['insw_error'] = "Proses terganggu (koneksi terputus). Silakan klik tombol 'Mulai Cek INSW Otomatis' lagi."
                    st.rerun()

                total = shared.get('total', st.session_state.get('insw_progress_total', 0))
                current = shared.get('current', 0)
                current_hs = shared.get('current_hs', '')
                current_desc = shared.get('current_desc', '')
                is_complete = shared.get('complete', False) or file_status in ('completed', 'error')
                partial_results = shared.get('results', [])
                error_msg = shared.get('error_msg', '')
                error_count = shared.get('error_count', 0)

                if not is_complete and heartbeat_stale:
                    is_complete = True
                    partial_results = shared.get('results', [])
                    error_msg = error_msg or "Proses scraping berhenti (tidak ada update selama 60 detik)"
                    logger.warning(f"[INSW] Heartbeat stale for {sid}, marking complete with {len(partial_results)} partial results")

                if not is_complete:
                    progress_val = current / total if total > 0 else 0
                    st.progress(progress_val, text=f"Mengecek HS Code {current_hs} ({current}/{total})...")
                    st.info(f"Sedang memproses: **{current_hs}** - {current_desc}")

                    if partial_results:
                        st.caption(f"{len(partial_results)} HS Code sudah dicek...")

                    _time.sleep(3)
                    st.rerun()
                else:
                    st.session_state['insw_results'] = partial_results
                    st.session_state['insw_running'] = False
                    st.session_state['insw_complete'] = True
                    st.session_state.pop('insw_thread_started', None)

                    if error_msg:
                        st.session_state['insw_error'] = f"Proses selesai dengan error. {len(partial_results)}/{total} HS Code dicek. Error: {error_msg}"
                    elif error_count > 0:
                        st.session_state['insw_error'] = f"Selesai! {len(partial_results)}/{total} HS Code dicek ({error_count} error)"

                    _cleanup_insw_state(sid)
                    _insw_threads.pop(sid, None)

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

with tab_petugas:
    st.markdown("### 👤 Cek & Lengkapi Nama Petugas Loket S2")
    st.markdown("Otomatis melengkapi **Nama Petugas** yang kosong di data Loket S2 berdasarkan data **Form Konsultasi**, serta mengisi **Skor** berdasarkan tingkat kepuasan.")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.markdown("**📁 File Loket S2**")
        file_loket = st.file_uploader("Upload file Loket S2 (.xlsx)", type=["xlsx", "xls"], key="loket_s2_file")
    with col_up2:
        st.markdown("**📁 File Form Konsultasi (bisa lebih dari 1)**")
        files_form = st.file_uploader("Upload file Form Konsultasi (.xlsx)", type=["xlsx", "xls"], key="form_konsul_files", accept_multiple_files=True)

    if 'petugas_result' not in st.session_state:
        st.session_state.petugas_result = None
    if 'petugas_excel' not in st.session_state:
        st.session_state.petugas_excel = None

    if file_loket and files_form:
        if st.button("🔍 Proses & Lengkapi Data Petugas", key="btn_cek_petugas"):
            with st.spinner("Memproses data..."):
                try:
                    df_loket_raw = pd.read_excel(file_loket)

                    skor_map = {'Sangat Puas': 2, 'Puas': 1, 'Tidak Puas': 0}

                    loket_records = []
                    current_date = None
                    current_satisfaction = None
                    idx_loket = 0
                    while idx_loket < len(df_loket_raw):
                        val0 = str(df_loket_raw.iloc[idx_loket, 0]).strip() if pd.notna(df_loket_raw.iloc[idx_loket, 0]) else ''
                        val1 = str(df_loket_raw.iloc[idx_loket, 1]).strip() if pd.notna(df_loket_raw.iloc[idx_loket, 1]) else ''

                        if val0 in ['', 'Row Labels', 'Grand Total'] or 'Nama petugas' in val1:
                            idx_loket += 1
                            continue

                        if 'Sangat Puas' in val0:
                            current_satisfaction = 'Sangat Puas'
                            idx_loket += 1
                            continue
                        elif 'Tidak Puas' in val0:
                            current_satisfaction = 'Tidak Puas'
                            idx_loket += 1
                            continue
                        elif 'Puas' in val0 and 'Sangat' not in val0 and 'Tidak' not in val0:
                            current_satisfaction = 'Puas'
                            idx_loket += 1
                            continue

                        try:
                            date_val = pd.to_datetime(val0)
                            if date_val.year >= 2025:
                                current_date = date_val
                                idx_loket += 1
                                continue
                        except (ValueError, TypeError):
                            pass

                        if '@' in val0:
                            idx_loket += 1
                            continue

                        if idx_loket + 1 < len(df_loket_raw):
                            next_val = str(df_loket_raw.iloc[idx_loket + 1, 0]).strip() if pd.notna(df_loket_raw.iloc[idx_loket + 1, 0]) else ''
                            if '@' in next_val:
                                nama = val0
                                petugas_loket = val1 if val1 else ''
                                email = next_val.lower()
                                skor_otomatis = skor_map.get(current_satisfaction, '')
                                loket_records.append({
                                    'Tanggal': current_date,
                                    'Nama': nama,
                                    'Email': email,
                                    'Petugas_Loket': petugas_loket,
                                    'Kepuasan': current_satisfaction if current_satisfaction else '',
                                    'Skor': skor_otomatis
                                })
                                idx_loket += 2
                                continue

                        idx_loket += 1

                    df_loket = pd.DataFrame(loket_records)

                    form_records = []
                    for ff in files_form:
                        df_f = pd.read_excel(ff)
                        has_loket_col = 'Pilihan Loket Layanan' in df_f.columns
                        if has_loket_col:
                            df_f = df_f[df_f['Pilihan Loket Layanan'].astype(str).str.contains('S2', case=False, na=False)]

                        col_nama = 'Nama' if 'Nama' in df_f.columns else None
                        col_email = 'Email Address' if 'Email Address' in df_f.columns else None
                        col_tanggal = 'Tanggal Konsultasi' if 'Tanggal Konsultasi' in df_f.columns else None
                        col_petugas = 'Nama Petugas' if 'Nama Petugas' in df_f.columns else None

                        if not all([col_nama, col_email, col_petugas]):
                            continue

                        for _, row in df_f.iterrows():
                            f_nama = str(row[col_nama]).strip() if pd.notna(row[col_nama]) else ''
                            f_email = str(row[col_email]).strip().lower() if pd.notna(row[col_email]) else ''
                            f_tanggal = None
                            if col_tanggal and pd.notna(row[col_tanggal]):
                                try:
                                    f_tanggal = pd.to_datetime(row[col_tanggal])
                                except (ValueError, TypeError):
                                    pass
                            f_petugas = str(row[col_petugas]).strip() if pd.notna(row[col_petugas]) else ''

                            if f_nama and f_email:
                                form_records.append({
                                    'Nama_Form': f_nama,
                                    'Email_Form': f_email,
                                    'Tanggal_Form': f_tanggal,
                                    'Petugas_Form': f_petugas,
                                    'Sumber': ff.name
                                })

                    df_forms = pd.DataFrame(form_records)

                    if df_loket.empty:
                        st.session_state.petugas_result = None
                        st.error("❌ Tidak ada data yang berhasil diparsing dari file Loket S2.")
                    elif df_forms.empty:
                        st.session_state.petugas_result = None
                        st.error("❌ Tidak ada data Form Konsultasi yang ditemukan.")
                    else:
                        def normalize_name(name):
                            if not name:
                                return ''
                            return re.sub(r'\s+', ' ', str(name).strip().lower())

                        def find_form_petugas(email, tanggal, nama):
                            candidates = df_forms[df_forms['Email_Form'] == email]
                            if not candidates.empty and tanggal is not None:
                                for _, fr in candidates.iterrows():
                                    if fr['Tanggal_Form'] is not None:
                                        try:
                                            if pd.to_datetime(tanggal).date() == pd.to_datetime(fr['Tanggal_Form']).date():
                                                return fr['Petugas_Form'], fr['Sumber']
                                        except (ValueError, TypeError):
                                            pass
                            if not candidates.empty:
                                first = candidates.iloc[0]
                                if first['Petugas_Form']:
                                    return first['Petugas_Form'], first['Sumber']
                            norm_nama = normalize_name(nama)
                            if norm_nama:
                                for _, fr in df_forms.iterrows():
                                    if normalize_name(fr['Nama_Form']) == norm_nama:
                                        if tanggal is not None and fr['Tanggal_Form'] is not None:
                                            try:
                                                if pd.to_datetime(tanggal).date() == pd.to_datetime(fr['Tanggal_Form']).date():
                                                    return fr['Petugas_Form'], fr['Sumber']
                                            except (ValueError, TypeError):
                                                pass
                            return '', ''

                        def match_short_to_full(short_name, full_name):
                            sn = normalize_name(short_name)
                            fn = normalize_name(full_name)
                            if not sn or not fn:
                                return False
                            if sn == fn:
                                return True
                            if sn in fn or fn in sn:
                                return True
                            sn_parts = sn.split()
                            fn_parts = fn.split()
                            for p in sn_parts:
                                if len(p) > 2 and p in fn_parts:
                                    return True
                            return False

                        results = []
                        for _, lr in df_loket.iterrows():
                            petugas_loket = lr['Petugas_Loket']
                            form_petugas, sumber = find_form_petugas(lr['Email'], lr['Tanggal'], lr['Nama'])

                            if petugas_loket and form_petugas:
                                if match_short_to_full(petugas_loket, form_petugas):
                                    status = 'Cocok'
                                    petugas_final = form_petugas
                                else:
                                    status = 'Tidak Cocok'
                                    petugas_final = petugas_loket
                            elif petugas_loket and not form_petugas:
                                status = 'Tidak Ada di Form'
                                petugas_final = petugas_loket
                            elif not petugas_loket and form_petugas:
                                status = 'Otomatis Terisi'
                                petugas_final = form_petugas
                            else:
                                status = 'Kosong'
                                petugas_final = ''

                            tanggal_str = ''
                            if lr['Tanggal'] is not None:
                                try:
                                    tanggal_str = pd.to_datetime(lr['Tanggal']).strftime('%d-%m-%Y')
                                except Exception:
                                    tanggal_str = str(lr['Tanggal'])

                            results.append({
                                'Tanggal': tanggal_str,
                                'Nama': lr['Nama'],
                                'Email': lr['Email'],
                                'Kepuasan': lr['Kepuasan'],
                                'Skor': lr['Skor'],
                                'Petugas (Loket S2)': petugas_loket if petugas_loket else '-',
                                'Petugas (Form)': form_petugas if form_petugas else '-',
                                'Petugas Final': petugas_final if petugas_final else '-',
                                'Status': status,
                                'Sumber File': sumber if sumber else '-'
                            })

                        df_result = pd.DataFrame(results)
                        st.session_state.petugas_result = df_result

                        out_buf = io.BytesIO()
                        with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
                            df_result.to_excel(writer, index=False, sheet_name='Hasil Lengkap')

                            wb = writer.book
                            ws = wb['Hasil Lengkap']

                            green_fill = PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid')
                            blue_fill = PatternFill(start_color='CCE5FF', end_color='CCE5FF', fill_type='solid')
                            red_fill = PatternFill(start_color='F8D7DA', end_color='F8D7DA', fill_type='solid')
                            yellow_fill = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
                            gray_fill = PatternFill(start_color='E2E3E5', end_color='E2E3E5', fill_type='solid')
                            hdr_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                            hdr_font = Font(bold=True, color='FFFFFF')
                            border_thin = Border(
                                left=Side(style='thin'), right=Side(style='thin'),
                                top=Side(style='thin'), bottom=Side(style='thin')
                            )

                            for cell in ws[1]:
                                cell.fill = hdr_fill
                                cell.font = hdr_font
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                cell.border = border_thin

                            status_ci = list(df_result.columns).index('Status') + 1
                            fill_map = {
                                'Cocok': green_fill, 'Otomatis Terisi': blue_fill,
                                'Tidak Cocok': red_fill, 'Kosong': yellow_fill,
                                'Tidak Ada di Form': gray_fill
                            }
                            for ri in range(2, ws.max_row + 1):
                                sv = ws.cell(row=ri, column=status_ci).value or ''
                                fl = fill_map.get(sv)
                                for ci_x in range(1, ws.max_column + 1):
                                    c = ws.cell(row=ri, column=ci_x)
                                    c.border = border_thin
                                    if fl:
                                        c.fill = fl

                            for ci_x in range(1, ws.max_column + 1):
                                ml = 0
                                for ri in range(1, ws.max_row + 1):
                                    cv = ws.cell(row=ri, column=ci_x).value
                                    if cv:
                                        ml = max(ml, len(str(cv)))
                                ws.column_dimensions[ws.cell(row=1, column=ci_x).column_letter].width = min(ml + 3, 40)

                            for sheet_status, sheet_name in [('Cocok', 'Cocok'), ('Otomatis Terisi', 'Otomatis Terisi'), ('Tidak Cocok', 'Tidak Cocok'), ('Kosong', 'Petugas Kosong'), ('Tidak Ada di Form', 'Tidak Ada di Form')]:
                                df_sheet = df_result[df_result['Status'] == sheet_status]
                                if not df_sheet.empty:
                                    df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)

                        out_buf.seek(0)
                        st.session_state.petugas_excel = out_buf.getvalue()
                        st.rerun()

                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

    if st.session_state.petugas_result is not None:
        df_result = st.session_state.petugas_result

        total = len(df_result)
        cocok = len(df_result[df_result['Status'] == 'Cocok'])
        otomatis = len(df_result[df_result['Status'] == 'Otomatis Terisi'])
        tidak_cocok = len(df_result[df_result['Status'] == 'Tidak Cocok'])
        kosong = len(df_result[df_result['Status'] == 'Kosong'])
        no_form = len(df_result[df_result['Status'] == 'Tidak Ada di Form'])

        st.success(f"✅ Berhasil memproses {total} data Loket S2")

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Total", total)
        c2.metric("Cocok", cocok)
        c3.metric("Otomatis Terisi", otomatis)
        c4.metric("Tidak Cocok", tidak_cocok)
        c5.metric("Kosong", kosong)
        c6.metric("Tidak Ada di Form", no_form)

        filter_st = st.selectbox("Filter Status:", ["Semua", "Cocok", "Otomatis Terisi", "Tidak Cocok", "Kosong", "Tidak Ada di Form"], key="filter_petugas")
        df_show = df_result if filter_st == "Semua" else df_result[df_result['Status'] == filter_st]

        def color_row(row):
            s = row['Status']
            if s == 'Cocok':
                return ['background-color: #d4edda'] * len(row)
            elif s == 'Otomatis Terisi':
                return ['background-color: #cce5ff'] * len(row)
            elif s == 'Tidak Cocok':
                return ['background-color: #f8d7da'] * len(row)
            elif s == 'Kosong':
                return ['background-color: #fff3cd'] * len(row)
            elif s == 'Tidak Ada di Form':
                return ['background-color: #e2e3e5'] * len(row)
            return [''] * len(row)

        st.dataframe(df_show.style.apply(color_row, axis=1), height=500)

        if st.session_state.petugas_excel is not None:
            st.download_button(
                label="📥 Download Hasil Pengecekan Petugas (Excel)",
                data=st.session_state.petugas_excel,
                file_name="hasil_cek_petugas_loket_s2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif not file_loket or not files_form:
        st.info("👆 Silakan upload file **Loket S2** dan **Form Konsultasi** untuk memulai pengecekan petugas.")

with tab_absen:
    st.markdown("### 📋 Cek Kehadiran Pegawai")
    st.markdown("Bandingkan data **Pegawai** dengan **Daftar Hadir** untuk mengetahui siapa saja yang **tidak hadir**.")

    col_ab1, col_ab2 = st.columns(2)
    with col_ab1:
        st.markdown("**📁 File Data Pegawai**")
        file_pegawai = st.file_uploader("Upload file Pegawai (.xlsx)", type=["xlsx", "xls"], key="file_pegawai")
    with col_ab2:
        st.markdown("**📁 File Daftar Hadir**")
        file_hadir = st.file_uploader("Upload file Daftar Hadir (.xlsx)", type=["xlsx", "xls"], key="file_hadir")

    if 'absen_result' not in st.session_state:
        st.session_state.absen_result = None
    if 'absen_excel' not in st.session_state:
        st.session_state.absen_excel = None

    if file_pegawai and file_hadir:
        if st.button("🔍 Cek Kehadiran", key="btn_cek_absen"):
            with st.spinner("Memproses data..."):
                try:
                    df_peg_raw = pd.read_excel(file_pegawai)
                    df_hadir_raw = pd.read_excel(file_hadir)

                    pegawai_list = []
                    nama_col_peg = None
                    jabatan_col_peg = None
                    for ci in range(min(10, df_peg_raw.shape[1])):
                        for ri in range(min(10, len(df_peg_raw))):
                            val = str(df_peg_raw.iloc[ri, ci]).strip().upper() if pd.notna(df_peg_raw.iloc[ri, ci]) else ''
                            if val == 'NAMA':
                                nama_col_peg = ci
                                start_row_peg = ri + 1
                                if ci + 1 < df_peg_raw.shape[1]:
                                    jabatan_col_peg = ci + 1
                                break
                        if nama_col_peg is not None:
                            break

                    if nama_col_peg is None:
                        for ci in range(min(10, df_peg_raw.shape[1])):
                            for ri in range(min(10, len(df_peg_raw))):
                                val = str(df_peg_raw.iloc[ri, ci]).strip().lower() if pd.notna(df_peg_raw.iloc[ri, ci]) else ''
                                if 'nama' in val:
                                    nama_col_peg = ci
                                    start_row_peg = ri + 1
                                    break
                            if nama_col_peg is not None:
                                break

                    if nama_col_peg is None:
                        nama_col_peg = 2
                        jabatan_col_peg = 3
                        start_row_peg = 4

                    for ri in range(start_row_peg, len(df_peg_raw)):
                        nama = str(df_peg_raw.iloc[ri, nama_col_peg]).strip() if pd.notna(df_peg_raw.iloc[ri, nama_col_peg]) else ''
                        jabatan = ''
                        if jabatan_col_peg is not None and jabatan_col_peg < df_peg_raw.shape[1]:
                            jabatan = str(df_peg_raw.iloc[ri, jabatan_col_peg]).strip() if pd.notna(df_peg_raw.iloc[ri, jabatan_col_peg]) else ''
                        if nama and nama != 'nan':
                            pegawai_list.append({'Nama': nama, 'Jabatan': jabatan if jabatan != 'nan' else ''})

                    hadir_list = []
                    nama_col_h = None
                    kehadiran_col_h = None
                    waktu_col_h = None
                    for ci in range(min(10, df_hadir_raw.shape[1])):
                        for ri in range(min(10, len(df_hadir_raw))):
                            val = str(df_hadir_raw.iloc[ri, ci]).strip().lower() if pd.notna(df_hadir_raw.iloc[ri, ci]) else ''
                            if 'nama' in val:
                                nama_col_h = ci
                                start_row_h = ri + 1
                                break
                        if nama_col_h is not None:
                            break

                    if nama_col_h is None:
                        nama_col_h = 0
                        start_row_h = 6

                    for ci in range(min(10, df_hadir_raw.shape[1])):
                        for ri in range(min(10, len(df_hadir_raw))):
                            val = str(df_hadir_raw.iloc[ri, ci]).strip().lower() if pd.notna(df_hadir_raw.iloc[ri, ci]) else ''
                            if 'kehadiran' in val:
                                kehadiran_col_h = ci
                            if 'waktu' in val:
                                waktu_col_h = ci

                    for ri in range(start_row_h, len(df_hadir_raw)):
                        nama = str(df_hadir_raw.iloc[ri, nama_col_h]).strip() if pd.notna(df_hadir_raw.iloc[ri, nama_col_h]) else ''
                        kehadiran = ''
                        waktu = ''
                        if kehadiran_col_h is not None:
                            kehadiran = str(df_hadir_raw.iloc[ri, kehadiran_col_h]).strip() if pd.notna(df_hadir_raw.iloc[ri, kehadiran_col_h]) else ''
                        if waktu_col_h is not None:
                            waktu = str(df_hadir_raw.iloc[ri, waktu_col_h]).strip() if pd.notna(df_hadir_raw.iloc[ri, waktu_col_h]) else ''
                        if nama and nama != 'nan':
                            hadir_list.append({'Nama_Hadir': nama, 'Kehadiran': kehadiran if kehadiran != 'nan' else '', 'Waktu': waktu if waktu != 'nan' else ''})

                    def clean_name_absen(name):
                        cleaned = re.sub(r',?\s*(S\.Si|S\.Farm|S\.E|S\.Kom|S\.IP|S\.Ak|S\.Sos|S\.K\.M|SKM|A\.Md|Apt|apt|M\.Si|M\.S|M\.Sc|M\.Farm|M\.Med\.Sc|M\.Epid|M\.K\.M|MKM|M\.T|Dra\.|Drs\.|Dr\.|drg\.|Rr\.)\.*', '', name, flags=re.IGNORECASE)
                        cleaned = re.sub(r'[,.]', ' ', cleaned)
                        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
                        return cleaned.lower()

                    hadir_clean_map = {}
                    for h in hadir_list:
                        hadir_clean_map[clean_name_absen(h['Nama_Hadir'])] = h

                    results_absen = []
                    for p in pegawai_list:
                        p_clean = clean_name_absen(p['Nama'])
                        matched_hadir = None
                        for h_clean, h_data in hadir_clean_map.items():
                            if p_clean == h_clean:
                                matched_hadir = h_data
                                break
                            if p_clean in h_clean or h_clean in p_clean:
                                matched_hadir = h_data
                                break
                            p_parts = p_clean.split()
                            h_parts = h_clean.split()
                            if len(p_parts) >= 2 and len(h_parts) >= 2:
                                if p_parts[0] == h_parts[0] and p_parts[-1] == h_parts[-1]:
                                    matched_hadir = h_data
                                    break

                        if matched_hadir:
                            results_absen.append({
                                'Nama Pegawai': p['Nama'],
                                'Jabatan': p['Jabatan'],
                                'Status': 'Hadir',
                                'Kehadiran': matched_hadir.get('Kehadiran', ''),
                                'Waktu': matched_hadir.get('Waktu', '')
                            })
                        else:
                            results_absen.append({
                                'Nama Pegawai': p['Nama'],
                                'Jabatan': p['Jabatan'],
                                'Status': 'Tidak Hadir',
                                'Kehadiran': '-',
                                'Waktu': '-'
                            })

                    df_absen = pd.DataFrame(results_absen)
                    st.session_state.absen_result = df_absen

                    out_absen = io.BytesIO()
                    with pd.ExcelWriter(out_absen, engine='openpyxl') as writer:
                        df_absen.to_excel(writer, index=False, sheet_name='Semua Pegawai')

                        wb = writer.book
                        ws = wb['Semua Pegawai']

                        green_f = PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid')
                        red_f = PatternFill(start_color='F8D7DA', end_color='F8D7DA', fill_type='solid')
                        hdr_f = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                        hdr_fn = Font(bold=True, color='FFFFFF')
                        bdr = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                        for cell in ws[1]:
                            cell.fill = hdr_f
                            cell.font = hdr_fn
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = bdr

                        sc = list(df_absen.columns).index('Status') + 1
                        for ri in range(2, ws.max_row + 1):
                            sv = ws.cell(row=ri, column=sc).value or ''
                            fl = green_f if sv == 'Hadir' else red_f if sv == 'Tidak Hadir' else None
                            for ci_x in range(1, ws.max_column + 1):
                                cell = ws.cell(row=ri, column=ci_x)
                                cell.border = bdr
                                if fl:
                                    cell.fill = fl

                        for ci_x in range(1, ws.max_column + 1):
                            ml = 0
                            for ri in range(1, ws.max_row + 1):
                                cv = ws.cell(row=ri, column=ci_x).value
                                if cv:
                                    ml = max(ml, len(str(cv)))
                            ws.column_dimensions[ws.cell(row=1, column=ci_x).column_letter].width = min(ml + 3, 50)

                        df_tidak = df_absen[df_absen['Status'] == 'Tidak Hadir']
                        if not df_tidak.empty:
                            df_tidak.to_excel(writer, index=False, sheet_name='Tidak Hadir')

                    out_absen.seek(0)
                    st.session_state.absen_excel = out_absen.getvalue()
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

    if st.session_state.absen_result is not None:
        df_absen = st.session_state.absen_result

        total_peg = len(df_absen)
        total_hadir = len(df_absen[df_absen['Status'] == 'Hadir'])
        total_tidak = len(df_absen[df_absen['Status'] == 'Tidak Hadir'])

        st.success(f"✅ Total Pegawai: {total_peg} | Hadir: {total_hadir} | Tidak Hadir: {total_tidak}")

        c_a1, c_a2, c_a3 = st.columns(3)
        c_a1.metric("Total Pegawai", total_peg)
        c_a2.metric("Hadir", total_hadir)
        c_a3.metric("Tidak Hadir", total_tidak)

        filter_absen = st.selectbox("Filter:", ["Semua", "Hadir", "Tidak Hadir"], key="filter_absen")
        df_show_absen = df_absen if filter_absen == "Semua" else df_absen[df_absen['Status'] == filter_absen]

        def color_absen(row):
            if row['Status'] == 'Hadir':
                return ['background-color: #d4edda'] * len(row)
            elif row['Status'] == 'Tidak Hadir':
                return ['background-color: #f8d7da'] * len(row)
            return [''] * len(row)

        st.dataframe(df_show_absen.style.apply(color_absen, axis=1), height=500)

        if st.session_state.absen_excel is not None:
            st.download_button(
                label="📥 Download Hasil Cek Kehadiran (Excel)",
                data=st.session_state.absen_excel,
                file_name="hasil_cek_kehadiran.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif not file_pegawai or not file_hadir:
        st.info("👆 Silakan upload file **Data Pegawai** dan **Daftar Hadir** untuk memulai pengecekan.")

if False:
    st.markdown("### 🏢 Analisis Bidang Usaha Importir")
    st.markdown("Upload file Excel, pilih kolom **Nama Importir**, lalu sistem akan **otomatis menganalisis bidang usaha** setiap importir menggunakan AI dan menentukan apakah termasuk **CEK** (obat/kosmetik/OT/food) atau **NOM** (bukan komoditas BPOM).")

    file_importir = st.file_uploader("Upload file Excel Importir (.xlsx/.xls)", type=["xlsx", "xls"], key="file_importir_upload")

    if 'importir_df_raw' not in st.session_state:
        st.session_state.importir_df_raw = None
    if 'importir_headers' not in st.session_state:
        st.session_state.importir_headers = None
    if 'importir_header_row' not in st.session_state:
        st.session_state.importir_header_row = None
    if 'importir_result' not in st.session_state:
        st.session_state.importir_result = None
    if 'importir_excel' not in st.session_state:
        st.session_state.importir_excel = None
    if 'importir_progress' not in st.session_state:
        st.session_state.importir_progress = None
    if 'importir_file_id' not in st.session_state:
        st.session_state.importir_file_id = None

    if file_importir:
        current_file_id = f"{file_importir.name}_{file_importir.size}"
        if st.session_state.importir_file_id != current_file_id:
            st.session_state.importir_result = None
            st.session_state.importir_excel = None
            st.session_state.importir_file_id = current_file_id
    elif not file_importir and st.session_state.importir_file_id is not None:
        st.session_state.importir_result = None
        st.session_state.importir_excel = None
        st.session_state.importir_file_id = None

    if file_importir:
        try:
            df_raw_imp = pd.read_excel(file_importir, header=None)
            header_row_imp = 0
            for ri in range(min(10, len(df_raw_imp))):
                row_vals = [str(df_raw_imp.iloc[ri, ci]).strip().upper() if pd.notna(df_raw_imp.iloc[ri, ci]) else '' for ci in range(min(10, df_raw_imp.shape[1]))]
                if any('NAMA' in v for v in row_vals):
                    header_row_imp = ri
                    break

            df_with_header = pd.read_excel(file_importir, header=header_row_imp)
            st.session_state.importir_df_raw = df_with_header
            st.session_state.importir_headers = list(df_with_header.columns)
            st.session_state.importir_header_row = header_row_imp

            st.success(f"✅ File berhasil dibaca! **{len(df_with_header)}** baris data, **{len(df_with_header.columns)}** kolom.")

            with st.expander("📋 Preview Data (10 baris pertama)", expanded=True):
                st.dataframe(df_with_header.head(10), height=300)

            with st.expander("📊 Struktur Kolom"):
                col_info = []
                for i, col in enumerate(df_with_header.columns):
                    col_letter = chr(65 + i) if i < 26 else chr(65 + (i // 26 - 1)) + chr(65 + (i % 26))
                    non_null = df_with_header[col].notna().sum()
                    col_info.append({
                        'Kolom Excel': col_letter,
                        'Nama Kolom': str(col)[:50],
                        'Jumlah Data': non_null,
                        'Kosong': len(df_with_header) - non_null
                    })
                st.dataframe(pd.DataFrame(col_info), height=400)

            st.markdown("---")
            st.markdown("#### ⚙️ Pengaturan Analisis")

            col_names = [str(c) for c in df_with_header.columns]

            nama_col_default = 0
            for i, cn in enumerate(col_names):
                if 'NAMA_IMPORTIR' in cn.upper() or 'NAMA IMPORTIR' in cn.upper():
                    nama_col_default = i
                    break

            selected_nama_col = st.selectbox(
                "Pilih kolom **Nama Importir** yang akan dianalisis:",
                options=col_names,
                index=nama_col_default,
                key="sel_nama_importir"
            )

            keterangan_col_options = ["(Buat kolom baru)"] + col_names
            ket_default = 0
            for i, cn in enumerate(col_names):
                cn_up = cn.upper().strip()
                if 'PENJELASAN' in cn_up or 'KETERANGAN' in cn_up:
                    ket_default = i + 1
                    break
            selected_ket_col = st.selectbox(
                "Pilih kolom untuk **Keterangan (NOM/CEK)**:",
                options=keterangan_col_options,
                index=ket_default,
                key="sel_ket_col"
            )

            bidang_col_options = ["(Buat kolom baru)"] + col_names
            bid_default = 0
            for i, cn in enumerate(col_names):
                cn_up = cn.upper().strip()
                if 'HASIL' in cn_up and 'ANALISIS' in cn_up:
                    bid_default = i + 1
                    break
            selected_bidang_col = st.selectbox(
                "Pilih kolom untuk **Bidang Usaha / Hasil Analisis**:",
                options=bidang_col_options,
                index=bid_default,
                key="sel_bidang_col"
            )

            only_empty = st.checkbox("Hanya analisis baris yang kolom Keterangan-nya masih kosong", value=True, key="only_empty_importir")

            product_col_candidates = {}
            for cn_s in col_names:
                cn_up = cn_s.upper()
                if 'BRGURAI' in cn_up:
                    product_col_candidates['brgurai'] = cn_s
                elif 'NOHS' in cn_up or cn_up == 'NOHS':
                    product_col_candidates['nohs'] = cn_s
                elif 'URAIAN_HS' in cn_up or 'URAIAN HS' in cn_up:
                    product_col_candidates['uraian_hs'] = cn_s
                elif 'ALAMAT' in cn_up:
                    product_col_candidates['alamat'] = cn_s

            unique_importers = df_with_header[selected_nama_col].dropna().unique()
            if only_empty and selected_ket_col != "(Buat kolom baru)":
                mask_empty = df_with_header[selected_ket_col].isna() | (df_with_header[selected_ket_col].astype(str).str.strip() == '')
                unique_importers = df_with_header.loc[mask_empty, selected_nama_col].dropna().unique()

            unique_importers = [str(n).strip() for n in unique_importers if str(n).strip() and str(n).strip().lower() != 'nan']
            unique_importers = list(dict.fromkeys(unique_importers))

            importir_context = {}
            for imp_name in unique_importers:
                rows_imp = df_with_header[df_with_header[selected_nama_col].astype(str).str.strip() == imp_name]
                products = []
                alamat = ''
                for _, row_imp in rows_imp.head(5).iterrows():
                    prod_info = {}
                    if 'brgurai' in product_col_candidates:
                        v = str(row_imp.get(product_col_candidates['brgurai'], '')).strip()
                        if v and v != 'nan':
                            prod_info['barang'] = v[:80]
                    if 'nohs' in product_col_candidates:
                        v = str(row_imp.get(product_col_candidates['nohs'], '')).strip()
                        if v and v != 'nan':
                            prod_info['hs'] = v
                    if 'uraian_hs' in product_col_candidates:
                        v = str(row_imp.get(product_col_candidates['uraian_hs'], '')).strip()
                        if v and v != 'nan':
                            prod_info['uraian'] = v[:80]
                    if prod_info:
                        products.append(prod_info)
                    if not alamat and 'alamat' in product_col_candidates:
                        v = str(row_imp.get(product_col_candidates['alamat'], '')).strip()
                        if v and v != 'nan':
                            alamat = v[:100]
                importir_context[imp_name] = {'products': products, 'alamat': alamat}

            st.info(f"📊 Ditemukan **{len(unique_importers)}** importir unik yang perlu dianalisis.")

            analysis_mode = st.radio(
                "Mode analisis:",
                ["Tanpa API (Gratis)", "Dengan AI (Butuh API Key)"],
                index=0,
                key="importir_analysis_mode"
            )

            def _apply_importir_results(all_results, source_label):
                df_result = df_with_header.copy()

                if selected_ket_col == "(Buat kolom baru)":
                    ket_col_name = "Keterangan_AI"
                    df_result[ket_col_name] = ""
                else:
                    ket_col_name = selected_ket_col

                if selected_bidang_col == "(Buat kolom baru)":
                    bidang_col_name = "Bidang_Usaha_AI"
                    df_result[bidang_col_name] = ""
                else:
                    bidang_col_name = selected_bidang_col

                alasan_col_name = "Alasan_Analisis"
                df_result[alasan_col_name] = ""

                def normalize_imp_name(n):
                    return re.sub(r'[^A-Z0-9\s]', '', n).strip()

                norm_map = {normalize_imp_name(k): v for k, v in all_results.items()}

                filled_count = 0
                for idx in range(len(df_result)):
                    nama_val = str(df_result.at[idx, selected_nama_col]).strip().upper() if pd.notna(df_result.at[idx, selected_nama_col]) else ''
                    if not nama_val or nama_val == 'NAN':
                        continue

                    if only_empty and selected_ket_col != "(Buat kolom baru)":
                        existing = str(df_result.at[idx, ket_col_name]).strip() if pd.notna(df_result.at[idx, ket_col_name]) else ''
                        if existing:
                            continue

                    nama_norm = normalize_imp_name(nama_val)
                    matched = all_results.get(nama_val)
                    if not matched:
                        matched = norm_map.get(nama_norm)
                    if not matched:
                        nama_words = set(nama_val.split())
                        best_score = 0
                        for key, val in all_results.items():
                            key_words = set(key.split())
                            if len(nama_words) >= 2 and len(key_words) >= 2:
                                common = len(nama_words & key_words)
                                score = common / max(len(nama_words), len(key_words))
                                if score > best_score and score >= 0.6:
                                    best_score = score
                                    matched = val

                    if matched:
                        df_result.at[idx, ket_col_name] = matched['kelas']
                        df_result.at[idx, bidang_col_name] = matched['bidang']
                        df_result.at[idx, alasan_col_name] = matched.get('alasan', '')
                        filled_count += 1
                    else:
                        df_result.at[idx, ket_col_name] = 'CEK'
                        df_result.at[idx, bidang_col_name] = 'Perlu cek manual'
                        df_result.at[idx, alasan_col_name] = f"({source_label}) Data importir tidak dapat dianalisis secara otomatis, perlu pengecekan manual."
                        filled_count += 1

                st.session_state.importir_result = df_result
                st.session_state.importir_ket_col = ket_col_name
                st.session_state.importir_bidang_col = bidang_col_name
                st.session_state.importir_alasan_col = alasan_col_name

                out_imp = io.BytesIO()
                with pd.ExcelWriter(out_imp, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Data Lengkap')

                    wb = writer.book
                    ws = wb['Data Lengkap']

                    hdr_f_imp = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                    hdr_fn_imp = Font(bold=True, color='FFFFFF')
                    green_fi = PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid')
                    yellow_fi = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
                    bdr_i = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                    for cell in ws[1]:
                        cell.fill = hdr_f_imp
                        cell.font = hdr_fn_imp
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = bdr_i

                    col_names_result = list(df_result.columns)
                    ket_ci = col_names_result.index(ket_col_name) + 1 if ket_col_name in col_names_result else None
                    bid_ci = col_names_result.index(bidang_col_name) + 1 if bidang_col_name in col_names_result else None

                    if ket_ci:
                        for ri in range(2, ws.max_row + 1):
                            sv = str(ws.cell(row=ri, column=ket_ci).value or '').strip().upper()
                            if sv == 'NOM':
                                ws.cell(row=ri, column=ket_ci).fill = green_fi
                                if bid_ci:
                                    ws.cell(row=ri, column=bid_ci).fill = green_fi
                            elif sv == 'CEK':
                                ws.cell(row=ri, column=ket_ci).fill = yellow_fi
                                if bid_ci:
                                    ws.cell(row=ri, column=bid_ci).fill = yellow_fi
                            ws.cell(row=ri, column=ket_ci).border = bdr_i
                            if bid_ci:
                                ws.cell(row=ri, column=bid_ci).border = bdr_i

                    summary_data = []
                    for name_upper, info in all_results.items():
                        summary_data.append({
                            'Nama Importir': name_upper,
                            'Bidang Usaha': info['bidang'],
                            'Klasifikasi': info['kelas'],
                            'Alasan': info.get('alasan', '')
                        })
                    df_summary = pd.DataFrame(summary_data)
                    df_summary.to_excel(writer, index=False, sheet_name='Ringkasan Importir')

                    ws2 = wb['Ringkasan Importir']
                    for cell in ws2[1]:
                        cell.fill = hdr_f_imp
                        cell.font = hdr_fn_imp
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = bdr_i

                    kls_ci = 3
                    for ri in range(2, ws2.max_row + 1):
                        sv = str(ws2.cell(row=ri, column=kls_ci).value or '').strip().upper()
                        for ci_x in range(1, ws2.max_column + 1):
                            ws2.cell(row=ri, column=ci_x).border = bdr_i
                        if sv == 'NOM':
                            for ci_x in range(1, ws2.max_column + 1):
                                ws2.cell(row=ri, column=ci_x).fill = green_fi
                        elif sv == 'CEK':
                            for ci_x in range(1, ws2.max_column + 1):
                                ws2.cell(row=ri, column=ci_x).fill = yellow_fi

                    for ws_x in [ws, ws2]:
                        for ci_x in range(1, ws_x.max_column + 1):
                            ml = 0
                            for ri in range(1, min(ws_x.max_row + 1, 100)):
                                cv = ws_x.cell(row=ri, column=ci_x).value
                                if cv:
                                    ml = max(ml, len(str(cv)))
                            ws_x.column_dimensions[ws_x.cell(row=1, column=ci_x).column_letter].width = min(ml + 3, 50)

                out_imp.seek(0)
                st.session_state.importir_excel = out_imp.getvalue()
                st.rerun()

            if analysis_mode == "Tanpa API (Gratis)":
                if st.button("⚙️ Mulai Analisis Otomatis (Tanpa API)", key="btn_analisis_importir_no_api"):
                    cek_keywords = [
                        'obat', 'farmasi', 'pharmaceutical', 'medicine', 'drug', 'bpom',
                        'kosmetik', 'cosmetic', 'suplemen', 'supplement', 'vitamin', 'vaksin',
                        'pangan', 'makanan', 'minuman', 'food', 'beverage',
                        'herbal', 'tradisional', 'jamu'
                    ]
                    bidang_map = {
                        'farmasi': ['farmasi', 'pharmaceutical', 'obat', 'medicine', 'drug', 'bahan baku obat', 'bahan obat'],
                        'kosmetik': ['kosmetik', 'cosmetic', 'skincare', 'parfum', 'fragrance'],
                        'pangan': ['pangan', 'makanan', 'minuman', 'food', 'beverage', 'flavour', 'flavor', 'ingredient'],
                        'suplemen': ['suplemen', 'supplement', 'vitamin', 'vaksin'],
                    }

                    def _guess(importir_name, ctx):
                        alamat = (ctx.get('alamat') or '')
                        products = ctx.get('products') or []
                        prod_texts = []
                        for p in products[:5]:
                            if p.get('barang'):
                                prod_texts.append(str(p.get('barang')))
                            if p.get('uraian'):
                                prod_texts.append(str(p.get('uraian')))
                        combined = (importir_name + ' ' + alamat + ' ' + ' '.join(prod_texts)).lower()

                        is_cek = any(k in combined for k in cek_keywords)

                        bidang = 'Perlu cek manual'
                        for bname, keys in bidang_map.items():
                            if any(k in combined for k in keys):
                                if bname == 'farmasi':
                                    bidang = 'Farmasi'
                                elif bname == 'kosmetik':
                                    bidang = 'Kosmetik'
                                elif bname == 'pangan':
                                    bidang = 'Pangan'
                                elif bname == 'suplemen':
                                    bidang = 'Suplemen Kesehatan'
                                break

                        kelas = 'CEK' if is_cek else 'NOM'

                        contoh = ''
                        if products:
                            p0 = products[0]
                            if p0.get('barang'):
                                contoh = str(p0.get('barang'))
                            elif p0.get('uraian'):
                                contoh = str(p0.get('uraian'))
                        if not contoh:
                            contoh = '-'

                        if kelas == 'CEK':
                            alasan = f"(Tanpa API) Importir terindikasi terkait komoditas yang berpotensi diawasi BPOM (contoh: {contoh}). Disarankan klasifikasi CEK untuk verifikasi manual lebih lanjut berdasarkan uraian barang/HS dan profil perusahaan."
                        else:
                            alasan = f"(Tanpa API) Berdasarkan teks nama/alamat/uraian barang yang tersedia (contoh: {contoh}), tidak ditemukan indikasi kuat terkait komoditas BPOM. Diklasifikasikan NOM, namun tetap perlu verifikasi jika ada informasi tambahan."

                        return {'bidang': bidang, 'kelas': kelas, 'alasan': alasan}

                    all_results = {}
                    for imp_name in unique_importers:
                        all_results[imp_name.strip().upper()] = _guess(imp_name, importir_context.get(imp_name, {}))

                    _apply_importir_results(all_results, "Tanpa API")

            else:
                if st.button("🤖 Mulai Analisis Otomatis dengan AI", key="btn_analisis_importir"):
                    from openai import OpenAI
                    from concurrent.futures import ThreadPoolExecutor, as_completed
                    from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception
                    import time

                    ai_key = os.environ.get("AI_INTEGRATIONS_OPENAI_API_KEY")
                    ai_url = os.environ.get("AI_INTEGRATIONS_OPENAI_BASE_URL")

                    if not ai_key or not ai_url:
                        st.error("❌ AI Integration belum dikonfigurasi. Pastikan OpenAI AI Integration sudah terinstall.")
                    else:
                        client = OpenAI(api_key=ai_key, base_url=ai_url)

                    def is_rate_limit_error(exception):
                        error_msg = str(exception)
                        return ("429" in error_msg or "RATELIMIT" in error_msg.upper()
                                or "quota" in error_msg.lower() or "rate limit" in error_msg.lower()
                                or (hasattr(exception, "status_code") and exception.status_code == 429))

                    @retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=1, min=2, max=60), retry=retry_if_exception(is_rate_limit_error), reraise=True)
                    def classify_batch(names_batch, context_map):
                        entries = []
                        for i, n in enumerate(names_batch):
                            ctx = context_map.get(n, {})
                            entry = f"{i+1}. Importir: {n}"
                            if ctx.get('alamat'):
                                entry += f"\n   Alamat: {ctx['alamat']}"
                            if ctx.get('products'):
                                prods = ctx['products'][:3]
                                prod_strs = []
                                for p in prods:
                                    ps = ""
                                    if p.get('barang'):
                                        ps += p['barang']
                                    if p.get('hs'):
                                        ps += f" (HS: {p['hs']})"
                                    if p.get('uraian'):
                                        ps += f" - {p['uraian']}"
                                    if ps:
                                        prod_strs.append(ps)
                                if prod_strs:
                                    entry += "\n   Produk yang diimpor: " + "; ".join(prod_strs)
                            entries.append(entry)
                        names_text = "\n".join(entries)

                        prompt = f"""Kamu adalah analis perdagangan Indonesia yang sangat ahli dalam mengidentifikasi bidang usaha perusahaan importir dan menganalisis komoditas impor terkait regulasi BPOM.

Untuk setiap importir di bawah ini, berikan:
1. "bidang": Bidang usaha utama perusahaan (singkat, 2-5 kata)
2. "kelas": "CEK" jika bidang usahanya terkait obat/farmasi, kosmetik, obat tradisional/herbal, makanan/minuman/food/pangan, suplemen kesehatan, atau bahan baku untuk produk-produk tersebut. "NOM" jika BUKAN terkait hal-hal tersebut.
3. "alasan": Penjelasan detail (2-3 kalimat) yang mencakup:
   - Deskripsi produk yang diimpor (berdasarkan data barang/HS Code jika tersedia)
   - Kegunaan produk tersebut
   - Mengapa diklasifikasikan NOM atau CEK
   - Informasi tentang importir dan bidang usahanya

Contoh alasan yang baik:
"Thermal grease dengan CAS Number 63148-62-9 merupakan bahan berbasis Polydimethylsiloxane (silicone oil) yang digunakan sebagai pasta penghantar panas (heat transfer compound), dan tidak termasuk bahan obat maupun makanan. Importir produk ini adalah PT Jaya Refrigeration Equipment yaitu perusahaan yang bergerak di bidang perdagangan dan penyediaan peralatan sistem pendingin (refrigeration) serta komponen pendukungnya."

PENTING:
- Perusahaan bahan kimia industri/specialty chemicals yang TIDAK spesifik untuk farmasi/food → NOM
- Perusahaan yang jelas bergerak di farmasi/pharmaceutical → CEK
- Perusahaan food ingredients/flavor/fragrance → CEK
- Jika ragu, klasifikasikan sebagai CEK
- Alasan harus ditulis dalam Bahasa Indonesia yang formal dan jelas
- Gunakan informasi produk (nama barang, HS Code, uraian) untuk memberikan alasan yang spesifik

Jawab HANYA dalam format JSON object dengan key "results" berisi array, contoh:
{{"results": [{{"nama": "PT ABC", "bidang": "Farmasi", "kelas": "CEK", "alasan": "PT ABC merupakan perusahaan farmasi yang mengimpor bahan baku obat..."}}, {{"nama": "PT XYZ", "bidang": "Peralatan Pendingin", "kelas": "NOM", "alasan": "PT XYZ mengimpor komponen refrigerasi yang bukan merupakan komoditas BPOM..."}}]}}

Daftar importir:
{names_text}"""

                        response = client.chat.completions.create(
                            model="gpt-5-mini",
                            messages=[{"role": "user", "content": prompt}],
                            response_format={"type": "json_object"},
                            max_completion_tokens=8192
                        )
                        content = response.choices[0].message.content or "[]"
                        try:
                            parsed = _json.loads(content)
                            if isinstance(parsed, dict):
                                for key in parsed:
                                    if isinstance(parsed[key], list):
                                        return parsed[key]
                                return []
                            return parsed
                        except:
                            return []

                    progress_bar = st.progress(0, text="Memulai analisis...")
                    status_text = st.empty()

                    batch_size = 15
                    batches = [unique_importers[i:i+batch_size] for i in range(0, len(unique_importers), batch_size)]
                    all_results = {}
                    total_batches = len(batches)
                    errors_count = 0

                    for bi, batch in enumerate(batches):
                        progress = (bi + 1) / total_batches
                        progress_bar.progress(progress, text=f"Menganalisis batch {bi+1}/{total_batches} ({len(all_results)}/{len(unique_importers)} importir)...")
                        status_text.text(f"🔄 Sedang memproses: {batch[0][:30]}... s/d {batch[-1][:30]}...")

                        try:
                            batch_ctx = {n: importir_context.get(n, {}) for n in batch}
                            results = classify_batch(batch, batch_ctx)
                            matched_in_batch = set()
                            if isinstance(results, list):
                                for r in results:
                                    if isinstance(r, dict) and 'nama' in r:
                                        rname = r['nama'].strip().upper()
                                        all_results[rname] = {
                                            'bidang': r.get('bidang', ''),
                                            'kelas': r.get('kelas', 'CEK'),
                                            'alasan': r.get('alasan', '')
                                        }
                                        matched_in_batch.add(rname)
                            for name in batch:
                                if name.strip().upper() not in matched_in_batch:
                                    norm_name = re.sub(r'[^A-Z0-9\s]', '', name.strip().upper()).strip()
                                    found = False
                                    for mk in matched_in_batch:
                                        mk_norm = re.sub(r'[^A-Z0-9\s]', '', mk).strip()
                                        if norm_name == mk_norm or (len(norm_name) > 5 and (norm_name in mk_norm or mk_norm in norm_name)):
                                            found = True
                                            break
                                    if not found:
                                        all_results[name.strip().upper()] = {'bidang': 'Perlu cek manual', 'kelas': 'CEK', 'alasan': 'Data importir tidak dapat dianalisis secara otomatis, perlu pengecekan manual.'}
                        except Exception as e:
                            errors_count += 1
                            st.warning(f"⚠️ Error batch {bi+1}: {str(e)[:100]}")
                            for name in batch:
                                all_results[name.strip().upper()] = {'bidang': 'Error - perlu cek manual', 'kelas': 'CEK', 'alasan': 'Terjadi error saat analisis, perlu pengecekan manual.'}

                        if bi < total_batches - 1:
                            time.sleep(0.5)

                    progress_bar.progress(1.0, text="✅ Analisis selesai!")
                    status_text.text(f"✅ Selesai! {len(all_results)} importir dianalisis" + (f" ({errors_count} batch error)" if errors_count else ""))

                    _apply_importir_results(all_results, "AI")

        except Exception as e:
            st.error(f"❌ Gagal membaca file: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

    if st.session_state.importir_result is not None:
        df_res = st.session_state.importir_result
        ket_col_name = st.session_state.get('importir_ket_col', 'Keterangan_AI')
        bidang_col_name = st.session_state.get('importir_bidang_col', 'Bidang_Usaha_AI')
        alasan_col_name = st.session_state.get('importir_alasan_col', 'Alasan_Analisis')

        if ket_col_name in df_res.columns:
            total_nom = len(df_res[df_res[ket_col_name].astype(str).str.strip().str.upper() == 'NOM'])
            total_cek = len(df_res[df_res[ket_col_name].astype(str).str.strip().str.upper() == 'CEK'])
            total_kosong = len(df_res[df_res[ket_col_name].isna() | (df_res[ket_col_name].astype(str).str.strip() == '')])

            st.success(f"✅ Analisis selesai! NOM: {total_nom} | CEK: {total_cek} | Belum diisi: {total_kosong}")

            c_i1, c_i2, c_i3 = st.columns(3)
            c_i1.metric("NOM (Bukan BPOM)", total_nom)
            c_i2.metric("CEK (Perlu Dicek)", total_cek)
            c_i3.metric("Belum Diisi", total_kosong)

        filter_imp = st.selectbox("Filter Klasifikasi:", ["Semua", "CEK", "NOM", "Belum Diisi"], key="filter_importir")
        if filter_imp == "CEK":
            df_show_imp = df_res[df_res[ket_col_name].astype(str).str.strip().str.upper() == 'CEK']
        elif filter_imp == "NOM":
            df_show_imp = df_res[df_res[ket_col_name].astype(str).str.strip().str.upper() == 'NOM']
        elif filter_imp == "Belum Diisi":
            df_show_imp = df_res[df_res[ket_col_name].isna() | (df_res[ket_col_name].astype(str).str.strip() == '')]
        else:
            df_show_imp = df_res

        show_cols = []
        for cn_s in df_res.columns:
            cn_upper = str(cn_s).upper()
            if 'NAMA_IMPORTIR' in cn_upper or 'NAMA IMPORTIR' in cn_upper or cn_s == ket_col_name or cn_s == bidang_col_name or cn_s == alasan_col_name or 'BRGURAI' in cn_upper or 'NOHS' in cn_upper or 'STATUS' in cn_upper:
                show_cols.append(cn_s)

        if not show_cols:
            show_cols = list(df_res.columns)

        def color_imp_row(row):
            kls = str(row.get(ket_col_name, '')).strip().upper()
            if kls == 'NOM':
                return ['background-color: #d4edda'] * len(row)
            elif kls == 'CEK':
                return ['background-color: #fff3cd'] * len(row)
            return [''] * len(row)

        st.dataframe(df_show_imp[show_cols].style.apply(color_imp_row, axis=1), height=500)

        if st.session_state.importir_excel is not None:
            st.download_button(
                label="📥 Download Hasil Analisis Importir (Excel)",
                data=st.session_state.importir_excel,
                file_name="hasil_analisis_importir.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif not file_importir:
        st.info("👆 Silakan upload file **Excel** untuk memulai analisis importir.")

with tab_merge:
    st.markdown("### 🔗 Gabung Data dari 2 File Excel")
    st.markdown("""Upload **2 file Excel** dengan struktur yang sama. Sistem akan:
- **Mempertahankan format asli** File Utama (filter, warna, font, lebar kolom, dll)
- **Hanya mengisi sel kosong** di File Utama dengan data dari File Pelengkap
- **Tidak mengubah** data yang sudah ada""")

    col_mg1, col_mg2 = st.columns(2)
    with col_mg1:
        st.markdown("**📁 File Utama** (yang ada sel kosong)")
        file_merge_main = st.file_uploader("Upload File Utama (.xlsx)", type=["xlsx"], key="file_merge_main")
    with col_mg2:
        st.markdown("**📁 File Pelengkap** (yang datanya lebih lengkap)")
        file_merge_source = st.file_uploader("Upload File Pelengkap (.xlsx)", type=["xlsx"], key="file_merge_source")

    if 'merge_excel' not in st.session_state:
        st.session_state.merge_excel = None
    if 'merge_stats' not in st.session_state:
        st.session_state.merge_stats = None
    if 'merge_file_id' not in st.session_state:
        st.session_state.merge_file_id = None
    if 'merge_filename' not in st.session_state:
        st.session_state.merge_filename = None

    if file_merge_main and file_merge_source:
        current_merge_id = f"{file_merge_main.name}_{file_merge_main.size}_{file_merge_source.name}_{file_merge_source.size}"
        if st.session_state.merge_file_id != current_merge_id:
            st.session_state.merge_excel = None
            st.session_state.merge_stats = None
            st.session_state.merge_file_id = current_merge_id

        try:
            from openpyxl import load_workbook

            df_main_peek = pd.read_excel(file_merge_main, header=None, nrows=5)
            file_merge_main.seek(0)
            df_src_peek = pd.read_excel(file_merge_source, header=None, nrows=5)
            file_merge_source.seek(0)

            df_main_info = pd.read_excel(file_merge_main, header=None, nrows=0)
            file_merge_main.seek(0)
            df_src_info = pd.read_excel(file_merge_source, header=None, nrows=0)
            file_merge_source.seek(0)

            main_cols = df_main_peek.shape[1]
            src_cols = df_src_peek.shape[1]

            main_rows_est = None
            src_rows_est = None

            st.success(f"✅ File berhasil dibaca!")

            col_info_mg1, col_info_mg2 = st.columns(2)
            with col_info_mg1:
                st.info(f"**File Utama**: {main_cols} kolom")
            with col_info_mg2:
                st.info(f"**File Pelengkap**: {src_cols} kolom")

            col_headers = []
            for ci in range(main_cols):
                found_header = None
                for ri in range(min(5, len(df_main_peek))):
                    v = df_main_peek.iloc[ri, ci]
                    if pd.notna(v) and str(v).strip() and str(v).strip().upper() != 'NAN':
                        vs = str(v).strip().upper()
                        if any(kw in vs for kw in ['NAMA', 'NO', 'TANGGAL', 'KODE', 'STATUS', 'KANTOR', 'ALAMAT', 'NPWP', 'SERIAL', 'SATUAN', 'KEMASAN', 'NEGARA', 'PELABUHAN', 'PENJELASAN', 'HASIL', 'ESTIMASI', 'CENTANG', 'NOMOR', 'ALASAN']):
                            found_header = str(v).strip()
                            break
                if not found_header:
                    for ri in range(min(5, len(df_main_peek))):
                        v = df_main_peek.iloc[ri, ci]
                        if pd.notna(v) and str(v).strip() and str(v).strip().upper() != 'NAN' and len(str(v).strip()) > 2:
                            found_header = str(v).strip()
                            break
                if not found_header:
                    found_header = f'Kolom_{ci}'

                col_idx_1 = ci + 1
                col_letter = chr(64 + col_idx_1) if col_idx_1 <= 26 else chr(64 + ((col_idx_1 - 1) // 26)) + chr(65 + ((col_idx_1 - 1) % 26))
                col_headers.append(f"{col_letter}: {found_header[:40]}")

            st.markdown("---")
            st.markdown("#### ⚙️ Pengaturan Penggabungan")

            mode_merge = st.radio(
                "Mode penggabungan:",
                ["Isi semua sel kosong di File Utama dari File Pelengkap", "Pilih kolom tertentu saja"],
                key="mode_merge"
            )

            selected_cols_merge = list(range(1, main_cols + 1))
            if mode_merge == "Pilih kolom tertentu saja":
                selected_headers = st.multiselect(
                    "Pilih kolom yang ingin digabungkan:",
                    options=col_headers,
                    default=[],
                    key="sel_cols_merge"
                )
                selected_cols_merge = []
                for sh in selected_headers:
                    col_letter = sh.split(":")[0].strip()
                    if len(col_letter) == 1:
                        selected_cols_merge.append(ord(col_letter) - 64)
                    elif len(col_letter) == 2:
                        selected_cols_merge.append((ord(col_letter[0]) - 64) * 26 + (ord(col_letter[1]) - 64))

            start_row_mg = st.number_input("Mulai dari baris ke- (di Excel):", min_value=1, value=1, step=1, key="start_row_merge")
            end_row_mg = st.number_input("Sampai baris ke- (di Excel, 0 = sampai akhir):", min_value=0, value=0, step=1, key="end_row_merge")

            overwrite_mode = st.checkbox("Timpa data yang sudah ada (overwrite)", value=False, key="overwrite_merge")

            if st.button("🔄 Gabungkan Data", key="btn_merge"):
                progress_mg = st.progress(0, text="Membaca File Pelengkap (cepat via pandas)...")

                file_merge_source.seek(0)
                df_src_all = pd.read_excel(file_merge_source, header=None)
                src_total_rows = len(df_src_all)

                progress_mg.progress(15, text="Membuka File Utama (mempertahankan format)...")

                file_merge_main.seek(0)
                wb_main = load_workbook(file_merge_main)
                ws_main = wb_main.active

                main_total_rows = ws_main.max_row or 0
                max_r = min(main_total_rows, src_total_rows)
                max_c = min(ws_main.max_column or 0, df_src_all.shape[1])

                actual_start = max(start_row_mg, 1)
                actual_end = end_row_mg if end_row_mg > 0 else max_r

                if main_total_rows != src_total_rows:
                    st.warning(f"⚠️ Jumlah baris berbeda! File Utama: {main_total_rows}, File Pelengkap: {src_total_rows}. Diproses sampai baris terpendek.")

                progress_mg.progress(30, text="Mengisi sel kosong...")

                fill_stats = {}
                total_cols_to_process = len([c for c in selected_cols_merge if c <= max_c])
                processed_cols = 0

                for ci in selected_cols_merge:
                    if ci > max_c:
                        continue
                    col_label = col_headers[ci - 1] if ci - 1 < len(col_headers) else f'Kolom_{ci}'
                    count = 0
                    pandas_ci = ci - 1

                    for ri in range(actual_start, min(actual_end + 1, max_r + 1)):
                        v_main = ws_main.cell(row=ri, column=ci).value
                        main_empty = v_main is None or str(v_main).strip() == '' or str(v_main).strip().lower() == 'nan'

                        if main_empty or overwrite_mode:
                            src_ri = ri - 1
                            if src_ri < len(df_src_all):
                                v_src = df_src_all.iloc[src_ri, pandas_ci]
                                src_filled = pd.notna(v_src) and str(v_src).strip() != '' and str(v_src).strip().lower() != 'nan'
                                if src_filled:
                                    ws_main.cell(row=ri, column=ci).value = v_src if not isinstance(v_src, float) or not v_src != v_src else None
                                    count += 1

                    if count > 0:
                        fill_stats[col_label] = count

                    processed_cols += 1
                    pct = 30 + int(50 * processed_cols / max(total_cols_to_process, 1))
                    progress_mg.progress(pct, text=f"Mengisi kolom {col_label}... ({processed_cols}/{total_cols_to_process})")

                progress_mg.progress(85, text="Menyimpan file (mempertahankan format asli)...")

                out_merge = io.BytesIO()
                wb_main.save(out_merge)
                wb_main.close()
                out_merge.seek(0)

                st.session_state.merge_excel = out_merge.getvalue()
                st.session_state.merge_stats = fill_stats

                base_name = file_merge_main.name
                if base_name.endswith('.xlsx'):
                    base_name = base_name[:-5]
                st.session_state.merge_filename = f"{base_name}_LENGKAP.xlsx"

                progress_mg.progress(100, text="Selesai!")
                st.rerun()

        except Exception as e:
            st.error(f"❌ Gagal membaca file: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

    if st.session_state.merge_stats is not None:
        fill_stats = st.session_state.merge_stats
        total_filled = sum(fill_stats.values())

        st.success(f"✅ Penggabungan selesai! Total **{total_filled}** sel berhasil diisi.")
        st.markdown("📌 **Format asli file dipertahankan** (filter, warna, font, lebar kolom, dll.)")

        if fill_stats:
            st.markdown("**📊 Rincian per kolom:**")
            stats_df = pd.DataFrame([{'Kolom': k, 'Jumlah Sel Terisi': v} for k, v in fill_stats.items()])
            st.dataframe(stats_df, height=min(len(stats_df) * 40 + 50, 400))
        else:
            st.info("Tidak ada sel yang perlu diisi. Data di File Utama sudah lengkap atau tidak ada data pelengkap yang cocok.")

        if st.session_state.merge_excel is not None:
            dl_name = st.session_state.merge_filename or "file_gabungan.xlsx"
            st.download_button(
                label="📥 Download File Gabungan (Excel)",
                data=st.session_state.merge_excel,
                file_name=dl_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif not file_merge_main or not file_merge_source:
        st.info("👆 Silakan upload **File Utama** dan **File Pelengkap** untuk memulai penggabungan data.")

st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #64748b; padding: 1rem;">
    <p>📊 Aplikasi Perbandingan Data Realisasi Impor</p>
    <p style="font-size: 0.8rem;">Dibuat dengan ❤️ menggunakan Streamlit</p>
</div>
""", unsafe_allow_html=True)
