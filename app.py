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
3. Upload **File Invoice Bahan Tambahan Obat** (untuk cek NO. INVOICE)
4. Upload **File Invoice Bahan Kimia** (untuk cek NO. INVOICE)
5. Sistem akan:
   - Mengidentifikasi data tarikan yang belum ada di file Anda (berdasarkan NO. PIB)
   - Mengecek NO. INVOICE di file Bahan Tambahan Obat
   - Mengecek NO. INVOICE di file Bahan Kimia
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 File Tarikan")
    file_tarikan = st.file_uploader("Data hasil tarikan sistem", type=['xlsx', 'xls'], key="tarikan")

with col2:
    st.subheader("📁 File Data Anda")
    file_upload = st.file_uploader("Data Anda untuk dibandingkan", type=['xlsx', 'xls'], key="upload")

st.markdown("---")
st.subheader("📁 File Invoice")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Bahan Tambahan Obat**")
    file_invoice_obat = st.file_uploader("File Invoice Bahan Tambahan Obat", type=['xlsx', 'xls'], key="invoice_obat")

with col2:
    st.markdown("**Bahan Kimia**")
    file_invoice_kimia = st.file_uploader("File Invoice Bahan Kimia", type=['xlsx', 'xls'], key="invoice_kimia")

def clean_number(value):
    """Membersihkan nilai dari tanda kutip ' dan " serta karakter non-numerik lainnya untuk NO. PIB"""
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    val_str = re.sub(r'[^\d]', '', val_str)
    return val_str

def get_invoice_list(value):
    """Memecah invoice yang mengandung ; atau , menjadi list dan membersihkan setiap item"""
    if pd.isna(value):
        return []
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    
    # Cek apakah ada separator ; atau ,
    if ';' in val_str or ',' in val_str:
        # Ganti semua separator dengan satu jenis lalu split
        val_str = val_str.replace(';', ',')
        invoices = [inv.strip().strip(';').strip(',').strip() for inv in val_str.split(',')]
        invoices = [inv for inv in invoices if inv]
        return invoices
    
    val_str = val_str.strip(';').strip(',').strip()
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

def load_invoice_set(file_invoice, label):
    """Load invoice set from file"""
    invoice_set = set()
    if file_invoice:
        df_invoice = pd.read_excel(file_invoice)
        invoice_col = find_invoice_column(df_invoice)
        if invoice_col:
            for inv_value in df_invoice[invoice_col].dropna():
                inv_list = get_invoice_list(inv_value)
                invoice_set.update(inv_list)
            st.success(f"**{label}**: {len(invoice_set)} NO. INVOICE unik ditemukan")
        else:
            st.warning(f"Kolom NO. INVOICE tidak ditemukan di {label}")
    return invoice_set

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
        
        invoice_set_obat = set()
        invoice_set_kimia = set()
        
        st.markdown("---")
        st.subheader("📋 Status File Invoice")
        
        if file_invoice_obat:
            invoice_set_obat = load_invoice_set(file_invoice_obat, "Bahan Tambahan Obat")
        
        if file_invoice_kimia:
            invoice_set_kimia = load_invoice_set(file_invoice_kimia, "Bahan Kimia")
        
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
        
        st.markdown("---")
        
        if st.button("🔍 Bandingkan Data", type="primary"):
            st.markdown("---")
            st.subheader("📊 Hasil Perbandingan NO. PIB")
            
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
                
                st.markdown("#### Data Hasil Perbandingan:")
                st.dataframe(df_missing, use_container_width=True)
                
                output_compare = io.BytesIO()
                with pd.ExcelWriter(output_compare, engine='openpyxl') as writer:
                    df_missing.to_excel(writer, index=False, sheet_name='Hasil Perbandingan')
                output_compare.seek(0)
                
                st.download_button(
                    label="📥 Download Hasil Perbandingan",
                    data=output_compare,
                    file_name="hasil_perbandingan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                if invoice_col_tarikan and (invoice_set_obat or invoice_set_kimia):
                    st.markdown("---")
                    st.subheader("📋 Cek NO. INVOICE")
                    
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
                        st.markdown("#### Cek di Bahan Tambahan Obat:")
                        df_invoice_check['Cek Bahan Obat'] = df_invoice_check[invoice_col_tarikan].apply(check_invoice_obat)
                        
                        ada_obat = df_invoice_check[df_invoice_check['Cek Bahan Obat'] == '✅ Ada']
                        sebagian_obat = df_invoice_check[df_invoice_check['Cek Bahan Obat'].str.contains('Sebagian', na=False)]
                        tidak_obat = df_invoice_check[df_invoice_check['Cek Bahan Obat'] == '❌ Tidak Ada']
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Ada di Bahan Obat", len(ada_obat))
                        with col2:
                            st.metric("Sebagian di Bahan Obat", len(sebagian_obat))
                        with col3:
                            st.metric("Tidak Ada di Bahan Obat", len(tidak_obat))
                    
                    if invoice_set_kimia:
                        st.markdown("#### Cek di Bahan Kimia:")
                        df_invoice_check['Cek Bahan Kimia'] = df_invoice_check[invoice_col_tarikan].apply(check_invoice_kimia)
                        
                        ada_kimia = df_invoice_check[df_invoice_check['Cek Bahan Kimia'] == '✅ Ada']
                        sebagian_kimia = df_invoice_check[df_invoice_check['Cek Bahan Kimia'].str.contains('Sebagian', na=False)]
                        tidak_kimia = df_invoice_check[df_invoice_check['Cek Bahan Kimia'] == '❌ Tidak Ada']
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Ada di Bahan Kimia", len(ada_kimia))
                        with col2:
                            st.metric("Sebagian di Bahan Kimia", len(sebagian_kimia))
                        with col3:
                            st.metric("Tidak Ada di Bahan Kimia", len(tidak_kimia))
                    
                    st.markdown("#### Data Lengkap dengan Status Invoice:")
                    st.dataframe(df_invoice_check, use_container_width=True)
                    
                    if invoice_set_obat:
                        st.markdown("---")
                        st.markdown("##### Filter Bahan Tambahan Obat:")
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
                        st.markdown("##### Filter Bahan Kimia:")
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
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
            else:
                st.success("✅ Semua data dari tarikan sudah tersedia di file Anda!")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")
        st.info("Pastikan file Excel dalam format yang benar (.xlsx atau .xls)")

else:
    st.info("👆 Silakan upload File Tarikan dan File Data Anda untuk memulai perbandingan. File Invoice bersifat opsional.")

st.markdown("---")
st.markdown("*Aplikasi Perbandingan Data Realisasi Impor*")
