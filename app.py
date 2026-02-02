import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Perbandingan Data Impor", page_icon="📊", layout="wide")

st.title("📊 Perbandingan Data Realisasi Impor")

tab_main, tab_analysis = st.tabs(["📋 Perbandingan Data", "📈 Analisis Data"])

def clean_value(value):
    """Membersihkan nilai dari tanda kutip dan spasi"""
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    return val_str.strip()

def clean_number(value):
    """Membersihkan nilai dari tanda kutip ' dan " serta karakter non-numerik lainnya"""
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
    
    if ';' in val_str or ',' in val_str:
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

def is_numeric_column(col_name):
    """Cek apakah kolom berisi data numerik berdasarkan nama"""
    col_lower = str(col_name).lower()
    numeric_keywords = ['no', 'pib', 'invoice', 'nomor', 'kode', 'id', 'number']
    return any(keyword in col_lower for keyword in numeric_keywords)

with tab_main:
    st.markdown("---")
    st.markdown("""
    ### Petunjuk Penggunaan:
    1. Upload **File Tarikan** (data hasil tarikan dari sistem)
    2. Upload **File Data Anda** (data yang ingin dibandingkan)
    3. **Pilih kolom** yang ingin digunakan untuk perbandingan
    4. Upload **File Invoice** (opsional) untuk cek NO. INVOICE
    5. Klik **Bandingkan Data**
    """)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📁 File Tarikan")
        file_tarikan = st.file_uploader("Data hasil tarikan sistem", type=['xlsx', 'xls'], key="tarikan")

    with col2:
        st.subheader("📁 File Data Anda")
        file_upload = st.file_uploader("Data Anda untuk dibandingkan", type=['xlsx', 'xls'], key="upload")

    st.markdown("---")
    st.subheader("📁 File Invoice (Opsional)")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Bahan Tambahan Obat**")
        file_invoice_obat = st.file_uploader("File Invoice Bahan Tambahan Obat", type=['xlsx', 'xls'], key="invoice_obat")

    with col2:
        st.markdown("**Bahan Kimia**")
        file_invoice_kimia = st.file_uploader("File Invoice Bahan Kimia", type=['xlsx', 'xls'], key="invoice_kimia")

    if file_tarikan and file_upload:
        try:
            df_tarikan = pd.read_excel(file_tarikan)
            df_upload = pd.read_excel(file_upload)
            
            st.markdown("---")
            st.subheader("⚙️ Pilih Kolom untuk Perbandingan")
            
            col_tarikan_list = df_tarikan.columns.tolist()
            col_upload_list = df_upload.columns.tolist()
            
            common_cols = [col for col in col_tarikan_list if col in col_upload_list]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Kolom di File Tarikan:**")
                selected_col_tarikan = st.selectbox(
                    "Pilih kolom untuk perbandingan (File Tarikan)",
                    options=col_tarikan_list,
                    index=0,
                    key="col_tarikan"
                )
            
            with col2:
                st.markdown("**Kolom di File Data Anda:**")
                default_index = col_upload_list.index(selected_col_tarikan) if selected_col_tarikan in col_upload_list else 0
                selected_col_upload = st.selectbox(
                    "Pilih kolom untuk perbandingan (File Anda)",
                    options=col_upload_list,
                    index=default_index,
                    key="col_upload"
                )
            
            use_numeric_cleaning = st.checkbox(
                "Bersihkan data numerik (hapus karakter non-angka seperti ', \", dll)",
                value=is_numeric_column(selected_col_tarikan),
                help="Centang jika kolom berisi nomor seperti NO. PIB, NO. INVOICE, dll"
            )
            
            if common_cols:
                st.info(f"💡 Kolom yang sama di kedua file: {', '.join(common_cols)}")
            
            invoice_col_tarikan = find_invoice_column(df_tarikan)
            
            invoice_set_obat = set()
            invoice_set_kimia = set()
            
            st.markdown("---")
            st.subheader("📋 Status File Invoice")
            
            if file_invoice_obat:
                invoice_set_obat = load_invoice_set(file_invoice_obat, "Bahan Tambahan Obat")
            else:
                st.info("File Invoice Bahan Tambahan Obat belum diupload")
            
            if file_invoice_kimia:
                invoice_set_kimia = load_invoice_set(file_invoice_kimia, "Bahan Kimia")
            else:
                st.info("File Invoice Bahan Kimia belum diupload")
            
            st.markdown("---")
            st.subheader("📋 Preview Data")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Data Tarikan (Sistem)**")
                st.write(f"Jumlah baris: {len(df_tarikan)}")
                st.write(f"Kolom dipilih: **{selected_col_tarikan}**")
                st.dataframe(df_tarikan.head(5), use_container_width=True)
            
            with col2:
                st.markdown("**Data Upload Anda**")
                st.write(f"Jumlah baris: {len(df_upload)}")
                st.write(f"Kolom dipilih: **{selected_col_upload}**")
                st.dataframe(df_upload.head(5), use_container_width=True)
            
            st.markdown("---")
            
            if st.button("🔍 Bandingkan Data", type="primary"):
                st.markdown("---")
                st.subheader(f"📊 Hasil Perbandingan: {selected_col_tarikan}")
                
                if use_numeric_cleaning:
                    df_tarikan['_clean_key'] = df_tarikan[selected_col_tarikan].apply(clean_number)
                    df_upload['_clean_key'] = df_upload[selected_col_upload].apply(clean_number)
                else:
                    df_tarikan['_clean_key'] = df_tarikan[selected_col_tarikan].apply(clean_value)
                    df_upload['_clean_key'] = df_upload[selected_col_upload].apply(clean_value)
                
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
                    st.markdown(f"### 🔴 Data Tarikan yang Tidak Ada di File Anda (berdasarkan {selected_col_tarikan})")
                    st.write(f"Ditemukan **{len(missing_in_upload)}** data unik dari tarikan yang tidak ada di file Anda.")
                    
                    df_missing = df_tarikan[df_tarikan['_clean_key'].isin(missing_in_upload)].copy()
                    df_missing = df_missing.drop(columns=['_clean_key'])
                    
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

with tab_analysis:
    st.markdown("---")
    st.markdown("""
    ### Analisis Data
    Upload file Excel untuk menganalisis data dan membuat grafik.
    Anda bisa melihat data terbanyak berdasarkan kolom yang dipilih (misalnya: Negara, Jenis Obat, dll).
    """)
    
    file_analysis = st.file_uploader("Upload file untuk analisis", type=['xlsx', 'xls'], key="analysis")
    
    if file_analysis:
        try:
            df_analysis = pd.read_excel(file_analysis)
            
            st.markdown("---")
            st.subheader("📋 Preview Data")
            st.write(f"Jumlah baris: {len(df_analysis)}")
            st.write(f"Jumlah kolom: {len(df_analysis.columns)}")
            st.dataframe(df_analysis.head(10), use_container_width=True)
            
            st.markdown("---")
            st.subheader("📊 Analisis Berdasarkan Kolom")
            
            col_list = df_analysis.columns.tolist()
            
            selected_analysis_col = st.selectbox(
                "Pilih kolom untuk dianalisis",
                options=col_list,
                key="analysis_col"
            )
            
            top_n = st.slider("Tampilkan Top N data", min_value=5, max_value=50, value=10, key="top_n")
            
            if st.button("🔍 Analisis Data", type="primary", key="btn_analysis"):
                st.markdown("---")
                
                value_counts = df_analysis[selected_analysis_col].value_counts().head(top_n)
                
                st.subheader(f"📈 Top {top_n} {selected_analysis_col}")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("#### Tabel Data")
                    df_counts = value_counts.reset_index()
                    df_counts.columns = [selected_analysis_col, 'Jumlah']
                    df_counts['Persentase'] = (df_counts['Jumlah'] / df_counts['Jumlah'].sum() * 100).round(2).astype(str) + '%'
                    st.dataframe(df_counts, use_container_width=True)
                
                with col2:
                    st.markdown("#### Grafik Bar")
                    st.bar_chart(value_counts)
                
                st.markdown("---")
                st.markdown("#### Grafik Pie")
                
                import matplotlib.pyplot as plt
                
                fig, ax = plt.subplots(figsize=(10, 8))
                colors = plt.cm.Set3(range(len(value_counts)))
                
                wedges, texts, autotexts = ax.pie(
                    value_counts.values, 
                    labels=None,
                    autopct='%1.1f%%',
                    colors=colors,
                    startangle=90
                )
                
                ax.legend(
                    wedges, 
                    [f"{label} ({count})" for label, count in zip(value_counts.index, value_counts.values)],
                    title=selected_analysis_col,
                    loc="center left",
                    bbox_to_anchor=(1, 0, 0.5, 1)
                )
                
                ax.set_title(f"Distribusi {selected_analysis_col}")
                plt.tight_layout()
                
                st.pyplot(fig)
                
                st.markdown("---")
                st.subheader("📥 Download Hasil Analisis")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    output_analysis = io.BytesIO()
                    with pd.ExcelWriter(output_analysis, engine='openpyxl') as writer:
                        df_counts.to_excel(writer, index=False, sheet_name='Hasil Analisis')
                    output_analysis.seek(0)
                    
                    st.download_button(
                        label="📥 Download Data (Excel)",
                        data=output_analysis,
                        file_name=f"analisis_{selected_analysis_col}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    img_buffer = io.BytesIO()
                    fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
                    img_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Grafik (PNG)",
                        data=img_buffer,
                        file_name=f"grafik_{selected_analysis_col}.png",
                        mime="image/png"
                    )
                
                st.markdown("---")
                st.subheader("📊 Statistik Lengkap")
                
                total_data = len(df_analysis)
                unique_values = df_analysis[selected_analysis_col].nunique()
                top_value = value_counts.index[0] if len(value_counts) > 0 else '-'
                top_count = value_counts.values[0] if len(value_counts) > 0 else 0
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Data", total_data)
                with col2:
                    st.metric(f"Jumlah {selected_analysis_col} Unik", unique_values)
                with col3:
                    st.metric("Terbanyak", top_value)
                with col4:
                    st.metric("Jumlah Terbanyak", top_count)
                
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")
            st.info("Pastikan file Excel dalam format yang benar (.xlsx atau .xls)")
    else:
        st.info("👆 Silakan upload file Excel untuk memulai analisis data.")

st.markdown("---")
st.markdown("*Aplikasi Perbandingan Data Realisasi Impor*")
