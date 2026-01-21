import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Perbandingan Data Impor", page_icon="📊", layout="wide")

st.title("📊 Perbandingan Data Realisasi Impor")
st.markdown("---")

st.markdown("""
### Petunjuk Penggunaan:
1. Upload **File Tarikan** (data hasil tarikan dari sistem)
2. Upload **File Data Anda** (data yang ingin dibandingkan berdasarkan NO. PIB)
3. Upload **File Invoice** (untuk sinkronisasi NO. INVOICE)
4. Sistem akan:
   - Mengidentifikasi data tarikan yang belum ada di file Anda (berdasarkan NO. PIB)
   - Sinkronisasi NO. INVOICE dari hasil perbandingan dengan file Invoice
   - Membersihkan tanda ; dari NO. INVOICE secara otomatis
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
    """Membersihkan nilai dari tanda kutip ' dan " serta karakter non-numerik lainnya untuk NO. PIB"""
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    val_str = re.sub(r'[^\d]', '', val_str)
    return val_str

def clean_invoice(value):
    """Membersihkan nilai invoice dari tanda kutip dan ; di awal/akhir"""
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    val_str = val_str.strip(';').strip()
    return val_str

def get_invoice_list(value):
    """Memecah invoice yang mengandung ; menjadi list dan membersihkan setiap item"""
    if pd.isna(value):
        return []
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    
    if ';' in val_str:
        invoices = [inv.strip().strip(';').strip() for inv in val_str.split(';')]
        invoices = [inv for inv in invoices if inv]
        return invoices
    
    val_str = val_str.strip(';').strip()
    return [val_str] if val_str else []

def find_invoice_column(df):
    """Mencari kolom NO. INVOICE dengan berbagai variasi nama"""
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

def find_pib_column(df):
    """Mencari kolom NO. PIB dengan berbagai variasi nama"""
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if 'pib' in col_lower and 'no' in col_lower:
            return col
        if col_lower == 'no. pib' or col_lower == 'no.pib' or col_lower == 'nopib':
            return col
    for col in df.columns:
        if 'pib' in str(col).lower():
            return col
    return None

if file_tarikan and file_upload:
    try:
        df_tarikan = pd.read_excel(file_tarikan)
        df_upload = pd.read_excel(file_upload)
        
        pib_col_tarikan = find_pib_column(df_tarikan)
        pib_col_upload = find_pib_column(df_upload)
        invoice_col_tarikan = find_invoice_column(df_tarikan)
        
        if not pib_col_tarikan:
            st.error("Kolom NO. PIB tidak ditemukan di File Tarikan")
            st.stop()
        if not pib_col_upload:
            st.error("Kolom NO. PIB tidak ditemukan di File Data Anda")
            st.stop()
        
        df_invoice = None
        invoice_set = set()
        invoice_col_invoice = None
        
        if file_invoice:
            df_invoice = pd.read_excel(file_invoice)
            invoice_col_invoice = find_invoice_column(df_invoice)
            
            if invoice_col_invoice:
                st.success(f"Kolom invoice ditemukan: **{invoice_col_invoice}**")
                for inv_value in df_invoice[invoice_col_invoice].dropna():
                    inv_list = get_invoice_list(inv_value)
                    invoice_set.update(inv_list)
                st.info(f"Total **{len(invoice_set)}** NO. INVOICE unik ditemukan (sudah dibersihkan dari tanda ;)")
            else:
                st.warning("Kolom NO. INVOICE tidak ditemukan di File Invoice")
        
        st.markdown("---")
        st.subheader("📋 Preview Data")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Data Tarikan (Sistem)**")
            st.write(f"Jumlah baris: {len(df_tarikan)}")
            st.write(f"Kolom NO. PIB: {pib_col_tarikan}")
            if invoice_col_tarikan:
                st.write(f"Kolom NO. INVOICE: {invoice_col_tarikan}")
            st.dataframe(df_tarikan.head(5), use_container_width=True)
        
        with col2:
            st.markdown("**Data Upload Anda**")
            st.write(f"Jumlah baris: {len(df_upload)}")
            st.write(f"Kolom NO. PIB: {pib_col_upload}")
            st.dataframe(df_upload.head(5), use_container_width=True)
        
        if file_invoice and df_invoice is not None:
            st.markdown("**File Invoice**")
            st.write(f"Jumlah baris: {len(df_invoice)}")
            if invoice_col_invoice:
                st.write(f"Kolom NO. INVOICE: {invoice_col_invoice}")
                sample_invoices = df_invoice[invoice_col_invoice].head(5).tolist()
                cleaned_samples = [get_invoice_list(inv) for inv in sample_invoices]
                st.write("Contoh invoice (setelah dibersihkan):")
                for orig, cleaned in zip(sample_invoices, cleaned_samples):
                    st.write(f"  `{orig}` → `{cleaned}`")
            st.dataframe(df_invoice.head(5), use_container_width=True)
        
        st.markdown("---")
        
        if st.button("🔍 Bandingkan Data", type="primary"):
            st.markdown("---")
            st.subheader("📊 Hasil Perbandingan")
            
            df_tarikan['_clean_pib'] = df_tarikan[pib_col_tarikan].apply(clean_number)
            df_upload['_clean_pib'] = df_upload[pib_col_upload].apply(clean_number)
            
            tarikan_keys = set(df_tarikan['_clean_pib'].dropna())
            tarikan_keys = {k for k in tarikan_keys if k != ''}
            
            upload_keys = set(df_upload['_clean_pib'].dropna())
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
                
                df_missing = df_tarikan[df_tarikan['_clean_pib'].isin(missing_in_upload)].copy()
                df_missing = df_missing.drop(columns=['_clean_pib'])
                
                if invoice_col_tarikan and invoice_set:
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
                            return '✅ Ada'
                        elif found:
                            return f'⚠️ Sebagian ({len(found)}/{len(inv_list)})'
                        else:
                            return '❌ Tidak Ada'
                    
                    df_missing['Sinkron Invoice'] = df_missing[invoice_col_tarikan].apply(check_invoice_sync)
                    
                    ada_semua = df_missing[df_missing['Sinkron Invoice'] == '✅ Ada']
                    sebagian = df_missing[df_missing['Sinkron Invoice'].str.contains('Sebagian', na=False)]
                    tidak_ada = df_missing[df_missing['Sinkron Invoice'] == '❌ Tidak Ada']
                    
                    st.markdown("#### Status Sinkronisasi Invoice:")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Invoice Ada", len(ada_semua))
                    with col2:
                        st.metric("Invoice Sebagian", len(sebagian))
                    with col3:
                        st.metric("Invoice Tidak Ada", len(tidak_ada))
                
                st.markdown("#### Data Lengkap:")
                st.dataframe(df_missing, use_container_width=True)
                
                if 'Sinkron Invoice' in df_missing.columns:
                    st.markdown("---")
                    tab1, tab2, tab3 = st.tabs(["✅ Invoice Ada", "⚠️ Invoice Sebagian", "❌ Invoice Tidak Ada"])
                    
                    with tab1:
                        if len(ada_semua) > 0:
                            st.dataframe(ada_semua, use_container_width=True)
                        else:
                            st.info("Tidak ada data dengan invoice yang tersinkron")
                    
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
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")
        st.info("Pastikan file Excel dalam format yang benar (.xlsx atau .xls)")

else:
    st.info("👆 Silakan upload File Tarikan dan File Data Anda untuk memulai perbandingan. File Invoice bersifat opsional untuk sinkronisasi.")

st.markdown("---")
st.markdown("*Aplikasi Perbandingan Data Realisasi Impor*")
