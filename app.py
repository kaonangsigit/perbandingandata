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
2. Upload **File Data Anda** (data yang ingin dibandingkan berdasarkan NO. PIB)
3. Upload **File Invoice** (untuk sinkronisasi NO. INVOICE)
4. Sistem akan:
   - Mengidentifikasi data tarikan yang belum ada di file Anda (berdasarkan NO. PIB)
   - Sinkronisasi NO. INVOICE dari hasil perbandingan dengan file Invoice
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
    file_invoice = st.file_uploader("File untuk sinkronisasi Invoice", type=['xlsx', 'xls'], key="invoice")

def clean_number(value):
    """Membersihkan nilai dari tanda kutip ' dan " serta karakter non-numerik lainnya"""
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    val_str = re.sub(r'[^\d]', '', val_str)
    return val_str

def clean_invoice(value):
    """Membersihkan nilai invoice dari tanda kutip"""
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    return val_str.strip()

def get_invoice_list(value):
    """Memecah invoice yang mengandung ; menjadi list"""
    if pd.isna(value):
        return []
    val_str = clean_invoice(value)
    if ';' in val_str:
        invoices = [inv.strip() for inv in val_str.split(';') if inv.strip()]
        return invoices
    return [val_str] if val_str else []

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
        invoice_set = set()
        if file_invoice:
            df_invoice_raw = pd.read_excel(file_invoice)
            df_invoice, _ = extract_columns(df_invoice_raw, REQUIRED_COLUMNS)
            if 'NO. INVOICE' in df_invoice.columns:
                for inv_value in df_invoice['NO. INVOICE'].dropna():
                    inv_list = get_invoice_list(inv_value)
                    invoice_set.update(inv_list)
        
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
            st.write(f"Jumlah baris: {len(df_invoice)}, Jumlah NO. INVOICE unik (termasuk yg dipisah ;): {len(invoice_set)}")
            st.dataframe(df_invoice.head(5), use_container_width=True)
        
        st.markdown("---")
        
        if st.button("🔍 Bandingkan Data", type="primary"):
            st.markdown("---")
            st.subheader("📊 Hasil Perbandingan")
            
            df_tarikan['_clean_key'] = df_tarikan['NO. PIB'].apply(clean_number)
            df_upload['_clean_key'] = df_upload['NO. PIB'].apply(clean_number)
            
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
                
                if 'NO. INVOICE' in df_missing.columns and invoice_set:
                    def check_invoice_sync(inv_value):
                        inv_list = get_invoice_list(inv_value)
                        if not inv_list:
                            return '❌ Tidak Ada'
                        found = []
                        not_found = []
                        for inv in inv_list:
                            if inv in invoice_set:
                                found.append(inv)
                            else:
                                not_found.append(inv)
                        if len(found) == len(inv_list):
                            return '✅ Ada Semua'
                        elif found:
                            return f'⚠️ Sebagian ({len(found)}/{len(inv_list)})'
                        else:
                            return '❌ Tidak Ada'
                    
                    df_missing['Sinkron Invoice'] = df_missing['NO. INVOICE'].apply(check_invoice_sync)
                    
                    ada_semua = df_missing[df_missing['Sinkron Invoice'] == '✅ Ada Semua']
                    sebagian = df_missing[df_missing['Sinkron Invoice'].str.contains('Sebagian', na=False)]
                    tidak_ada = df_missing[df_missing['Sinkron Invoice'] == '❌ Tidak Ada']
                    
                    st.markdown("#### Status Sinkronisasi Invoice:")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Invoice Ada Semua", len(ada_semua))
                    with col2:
                        st.metric("Invoice Sebagian", len(sebagian))
                    with col3:
                        st.metric("Invoice Tidak Ada", len(tidak_ada))
                
                st.markdown("#### Data Lengkap:")
                st.dataframe(df_missing, use_container_width=True)
                
                if 'Sinkron Invoice' in df_missing.columns:
                    st.markdown("---")
                    tab1, tab2, tab3 = st.tabs(["✅ Invoice Ada Semua", "⚠️ Invoice Sebagian", "❌ Invoice Tidak Ada"])
                    
                    with tab1:
                        if len(ada_semua) > 0:
                            st.dataframe(ada_semua, use_container_width=True)
                        else:
                            st.info("Tidak ada data dengan invoice lengkap")
                    
                    with tab2:
                        if len(sebagian) > 0:
                            st.dataframe(sebagian, use_container_width=True)
                        else:
                            st.info("Tidak ada data dengan invoice sebagian")
                    
                    with tab3:
                        if len(tidak_ada) > 0:
                            st.dataframe(tidak_ada, use_container_width=True)
                        else:
                            st.info("Semua invoice sudah tersinkron")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_missing.to_excel(writer, index=False, sheet_name='Hasil Perbandingan')
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
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")
        st.info("Pastikan file Excel dalam format yang benar (.xlsx atau .xls)")

else:
    st.info("👆 Silakan upload File Tarikan dan File Data Anda untuk memulai perbandingan. File Invoice bersifat opsional untuk sinkronisasi.")

st.markdown("---")
st.markdown("*Aplikasi Perbandingan Data Realisasi Impor*")
