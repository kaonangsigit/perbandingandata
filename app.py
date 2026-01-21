import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Perbandingan Data Impor", page_icon="📊", layout="wide")

st.title("📊 Perbandingan Data Realisasi Impor")
st.markdown("---")

REQUIRED_COLUMNS = ['NO', 'CUSDECID', 'NO. INVOICE', 'TGL. INVOICE', 'NO. PIB', 'TGL. PIB', 'TGL. SPPB', 'SERIAL', 'URAIAN BARANG', 'NO. HS', 'URAIAN HS', 'NEGARA ASAL', 'JML SATUAN', 'SATUAN', 'KEMASAN', 'TGL. REALISASI', 'NO. SKI', 'TGL. SKI', 'NPWP', 'NAMA IMPORTIR', 'ALAMAT', 'KPBC', 'PEL. MUAT', 'PEL. BONGKAR', 'STATUS', 'STATUS PERIKSA']

st.markdown("""
### Petunjuk Penggunaan:
1. Upload **File Tarikan** (data hasil tarikan dari sistem)
2. Upload **File Data Anda** (data yang ingin dibandingkan)
3. Upload **File Invoice** (untuk mengecek keberadaan NO. INVOICE)
4. Sistem akan:
   - Mengidentifikasi data tarikan yang belum ada di file Anda
   - Mengecek apakah NO. INVOICE dari hasil perbandingan ada di file Invoice
   - Mengecek status ijin (SKI)
""")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📁 File Tarikan")
    file_tarikan = st.file_uploader("Data hasil tarikan sistem", type=['xlsx', 'xls'], key="tarikan")

with col2:
    st.subheader("📁 File Data Anda")
    file_upload = st.file_uploader("Data Anda untuk dibandingkan", type=['xlsx', 'xls'], key="upload")

with col3:
    st.subheader("📁 File Invoice")
    file_invoice = st.file_uploader("File untuk cek NO. INVOICE", type=['xlsx', 'xls'], key="invoice")

def clean_number(value):
    """Membersihkan nilai dari tanda kutip ' dan " serta karakter non-numerik lainnya"""
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    val_str = re.sub(r'[^\d]', '', val_str)
    return val_str

def clean_text(value):
    """Membersihkan nilai dari tanda kutip ' dan " """
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    return val_str.strip()

def extract_columns(df, required_cols):
    """Extract only required columns from dataframe, handling column name variations"""
    available_cols = []
    col_mapping = {}
    
    for req_col in required_cols:
        for df_col in df.columns:
            if req_col.lower().strip() == str(df_col).lower().strip():
                available_cols.append(df_col)
                col_mapping[df_col] = req_col
                break
            elif req_col.lower().replace('.', '').replace(' ', '') == str(df_col).lower().replace('.', '').replace(' ', ''):
                available_cols.append(df_col)
                col_mapping[df_col] = req_col
                break
    
    if available_cols:
        df_filtered = df[available_cols].copy()
        df_filtered.columns = [col_mapping.get(col, col) for col in df_filtered.columns]
        return df_filtered, available_cols
    return df, list(df.columns)

if file_tarikan and file_upload:
    try:
        df_tarikan_raw = pd.read_excel(file_tarikan)
        df_upload_raw = pd.read_excel(file_upload)
        
        df_tarikan, tarikan_cols_found = extract_columns(df_tarikan_raw, REQUIRED_COLUMNS)
        df_upload, upload_cols_found = extract_columns(df_upload_raw, REQUIRED_COLUMNS)
        
        df_invoice = None
        invoice_numbers = set()
        if file_invoice:
            df_invoice_raw = pd.read_excel(file_invoice)
            df_invoice, _ = extract_columns(df_invoice_raw, REQUIRED_COLUMNS)
            if 'NO. INVOICE' in df_invoice.columns:
                invoice_numbers = set(df_invoice['NO. INVOICE'].apply(clean_text).dropna())
                invoice_numbers = {inv for inv in invoice_numbers if inv != ''}
        
        st.markdown("---")
        st.subheader("📋 Preview Data")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Data Tarikan (Sistem)**")
            st.write(f"Jumlah baris: {len(df_tarikan)}")
            st.dataframe(df_tarikan.head(5), use_container_width=True)
        
        with col2:
            st.markdown("**Data Upload Anda**")
            st.write(f"Jumlah baris: {len(df_upload)}")
            st.dataframe(df_upload.head(5), use_container_width=True)
        
        if file_invoice and df_invoice is not None:
            st.markdown("**File Invoice**")
            st.write(f"Jumlah baris: {len(df_invoice)}, Jumlah NO. INVOICE unik: {len(invoice_numbers)}")
            st.dataframe(df_invoice.head(5), use_container_width=True)
        
        st.markdown("---")
        st.subheader("⚙️ Konfigurasi Perbandingan")
        
        common_cols = list(set(df_tarikan.columns) & set(df_upload.columns))
        
        if common_cols:
            default_key = 'NO. PIB' if 'NO. PIB' in common_cols else common_cols[0]
            default_idx = common_cols.index(default_key) if default_key in common_cols else 0
            
            key_column = st.selectbox(
                "Pilih kolom kunci untuk perbandingan",
                options=common_cols,
                index=default_idx
            )
            
            if st.button("🔍 Bandingkan Data", type="primary"):
                st.markdown("---")
                st.subheader("📊 Hasil Perbandingan")
                
                df_tarikan['_clean_key'] = df_tarikan[key_column].apply(clean_number)
                df_upload['_clean_key'] = df_upload[key_column].apply(clean_number)
                
                tarikan_keys = set(df_tarikan['_clean_key'].dropna())
                tarikan_keys = {k for k in tarikan_keys if k != ''}
                
                upload_keys = set(df_upload['_clean_key'].dropna())
                upload_keys = {k for k in upload_keys if k != ''}
                
                missing_in_upload = tarikan_keys - upload_keys
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Data di Tarikan", len(tarikan_keys))
                with col2:
                    st.metric("Data di File Anda", len(upload_keys))
                with col3:
                    st.metric("Data Tarikan Tidak Ada di File Anda", len(missing_in_upload))
                
                if missing_in_upload:
                    st.markdown("### 🔴 Data Tarikan yang Tidak Ada di File Anda")
                    st.write(f"Ditemukan **{len(missing_in_upload)}** data dari tarikan yang tidak ada di file Anda.")
                    
                    df_missing = df_tarikan[df_tarikan['_clean_key'].isin(missing_in_upload)].copy()
                    df_missing = df_missing.drop(columns=['_clean_key'])
                    
                    if 'NO. INVOICE' in df_missing.columns:
                        df_missing['_clean_invoice'] = df_missing['NO. INVOICE'].apply(clean_text)
                        
                        if invoice_numbers:
                            df_missing['Invoice di File Invoice'] = df_missing['_clean_invoice'].apply(
                                lambda x: '✅ Ada' if x in invoice_numbers else '❌ Tidak Ada'
                            )
                        
                        df_missing = df_missing.drop(columns=['_clean_invoice'])
                    
                    if 'STATUS' in df_missing.columns:
                        df_missing['Status Ijin'] = df_missing['STATUS'].apply(
                            lambda x: '✅ Sudah Ada Ijin' if pd.notna(x) and 'sudah ada ijin' in str(x).lower().strip() else '❌ Belum Ada Ijin'
                        )
                        
                        ada_ijin = df_missing[df_missing['Status Ijin'] == '✅ Sudah Ada Ijin']
                        belum_ijin = df_missing[df_missing['Status Ijin'] == '❌ Belum Ada Ijin']
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Sudah Ada Ijin", len(ada_ijin))
                        with col2:
                            st.metric("Belum Ada Ijin", len(belum_ijin))
                    
                    if invoice_numbers and 'Invoice di File Invoice' in df_missing.columns:
                        ada_invoice = df_missing[df_missing['Invoice di File Invoice'] == '✅ Ada']
                        tidak_ada_invoice = df_missing[df_missing['Invoice di File Invoice'] == '❌ Tidak Ada']
                        
                        st.markdown("#### Status Invoice:")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Invoice Ada di File Invoice", len(ada_invoice))
                        with col2:
                            st.metric("Invoice Tidak Ada di File Invoice", len(tidak_ada_invoice))
                    
                    st.markdown("#### Data Lengkap:")
                    st.dataframe(df_missing, use_container_width=True)
                    
                    st.markdown("---")
                    
                    if 'Status Ijin' in df_missing.columns:
                        tab1, tab2 = st.tabs(["✅ Sudah Ada Ijin", "❌ Belum Ada Ijin"])
                        
                        with tab1:
                            if len(ada_ijin) > 0:
                                st.dataframe(ada_ijin, use_container_width=True)
                            else:
                                st.info("Tidak ada data yang sudah ada ijin")
                        
                        with tab2:
                            if len(belum_ijin) > 0:
                                st.dataframe(belum_ijin, use_container_width=True)
                            else:
                                st.info("Semua data sudah memiliki ijin")
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_missing.to_excel(writer, index=False, sheet_name='Data Tidak Ada di Upload')
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Download Hasil Perbandingan",
                        data=output,
                        file_name="hasil_perbandingan.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.success("✅ Semua data dari tarikan sudah tersedia di file Anda!")
                
                df_tarikan = df_tarikan.drop(columns=['_clean_key'], errors='ignore')
                df_upload = df_upload.drop(columns=['_clean_key'], errors='ignore')
                
        else:
            st.error("⚠️ Tidak ditemukan kolom yang sama antara kedua file.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")
        st.info("Pastikan file Excel dalam format yang benar (.xlsx atau .xls)")

else:
    st.info("👆 Silakan upload File Tarikan dan File Data Anda untuk memulai perbandingan. File Invoice bersifat opsional.")

st.markdown("---")
st.markdown("*Aplikasi Perbandingan Data Realisasi Impor*")
