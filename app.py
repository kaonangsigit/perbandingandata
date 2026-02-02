import streamlit as st
import pandas as pd
import io
import re
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

tab_main, tab_analysis = st.tabs(["📋 Perbandingan Data", "📈 Analisis Data"])

def clean_value(value):
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    
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
        1. Upload **File Tarikan** (data hasil tarikan dari sistem)
        2. Upload **File Data Anda** (data yang ingin dibandingkan)
        3. **Pilih kolom** yang ingin digunakan untuk perbandingan
        4. Upload **File Invoice** (opsional) untuk cek NO. INVOICE
        5. Klik **Bandingkan Data**
        """)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 📥 File Tarikan")
        file_tarikan = st.file_uploader("Data hasil tarikan sistem", type=['xlsx', 'xls'], key="tarikan", help="Upload file Excel dari sistem")

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

    if file_tarikan and file_upload:
        try:
            df_tarikan = pd.read_excel(file_tarikan)
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
                st.markdown("### 📊 Download Data Lengkap dengan Warna")
                st.markdown("File Excel akan memiliki:")
                st.markdown("- 🟡 **Warna Kuning**: Data yang **SAMA** di kedua file")
                st.markdown("- ⬜ **Tanpa Warna (Putih)**: Data yang **TIDAK SAMA** / tidak ada di file lain")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("🟡 Data Kuning (Sama)", jumlah_sama)
                with col2:
                    st.metric("⬜ Data Putih (Tidak Sama)", jumlah_tidak_sama)
                
                output_colored = io.BytesIO()
                with pd.ExcelWriter(output_colored, engine='openpyxl') as writer:
                    df_tarikan_display.to_excel(writer, index=False, sheet_name='Hasil Perbandingan')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Hasil Perbandingan']
                    
                    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                    header_font = Font(bold=True, color='FFFFFF')
                    thin_border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    
                    for col_idx, col in enumerate(df_tarikan_display.columns, 1):
                        cell = worksheet.cell(row=1, column=col_idx)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal='center')
                        cell.border = thin_border
                    
                    status_col_idx = df_tarikan_display.columns.get_loc('Status') + 1
                    
                    for row_idx in range(2, len(df_tarikan_display) + 2):
                        status_cell = worksheet.cell(row=row_idx, column=status_col_idx)
                        if '✅' in str(status_cell.value):
                            for col_idx in range(1, len(df_tarikan_display.columns) + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.fill = yellow_fill
                                cell.border = thin_border
                        else:
                            for col_idx in range(1, len(df_tarikan_display.columns) + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.border = thin_border
                    
                    for col_idx, col in enumerate(df_tarikan_display.columns, 1):
                        max_length = max(
                            df_tarikan_display[col].astype(str).apply(len).max(),
                            len(str(col))
                        ) + 2
                        worksheet.column_dimensions[worksheet.cell(row=1, column=col_idx).column_letter].width = min(max_length, 50)
                    
                    summary_row = len(df_tarikan_display) + 4
                    worksheet.cell(row=summary_row, column=1, value='RINGKASAN:')
                    worksheet.cell(row=summary_row, column=1).font = Font(bold=True)
                    worksheet.cell(row=summary_row + 1, column=1, value='Data Kuning (Sama di kedua file):')
                    worksheet.cell(row=summary_row + 1, column=2, value=jumlah_sama)
                    worksheet.cell(row=summary_row + 1, column=1).fill = yellow_fill
                    worksheet.cell(row=summary_row + 1, column=2).fill = yellow_fill
                    worksheet.cell(row=summary_row + 2, column=1, value='Data Putih (Tidak sama / tidak ada):')
                    worksheet.cell(row=summary_row + 2, column=2, value=jumlah_tidak_sama)
                    worksheet.cell(row=summary_row + 3, column=1, value='Total Data:')
                    worksheet.cell(row=summary_row + 3, column=2, value=len(df_tarikan_display))
                    worksheet.cell(row=summary_row + 3, column=1).font = Font(bold=True)
                    
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
