import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Perbandingan Data Impor", page_icon="📊", layout="wide")

st.title("📊 Perbandingan Data Realisasi Impor")
st.markdown("---")

REQUIRED_COLUMNS = ['NO', 'CAR', 'NO. PIB', 'TGL. PIB', 'TGL. SPPB', 'NPWP', 'NAMA IMPORTIR', 'ALAMAT', 'STATUS', 'STATUS PERIKSA']

st.markdown("""
### Petunjuk Penggunaan:
1. Upload **File Tarikan** (data hasil tarikan dari sistem)
2. Upload **File Data Anda** (data yang ingin dibandingkan)
3. Sistem akan mengidentifikasi data tarikan yang belum tersedia di file Anda
4. Untuk data tersebut, akan dicek apakah sudah ada SKI-nya atau tidak
""")

st.info(f"**Kolom yang akan diambil:** {', '.join(REQUIRED_COLUMNS)}")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 File Tarikan (Data Sistem)")
    file_tarikan = st.file_uploader("Upload file hasil tarikan", type=['xlsx', 'xls'], key="tarikan")

with col2:
    st.subheader("📁 File Data Anda")
    file_upload = st.file_uploader("Upload file data Anda untuk dibandingkan", type=['xlsx', 'xls'], key="upload")

def clean_number(value):
    """Membersihkan nilai dari tanda kutip ' dan " serta karakter non-numerik lainnya"""
    if pd.isna(value):
        return ''
    val_str = str(value).strip()
    val_str = val_str.replace("'", "").replace('"', "").replace("'", "").replace("'", "")
    val_str = re.sub(r'[^\d]', '', val_str)
    return val_str

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
        
        st.markdown("---")
        st.subheader("📋 Preview Data (Kolom yang Diambil)")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Data Tarikan (Sistem)**")
            st.write(f"Jumlah baris: {len(df_tarikan)}")
            st.write(f"Kolom ditemukan: {len(tarikan_cols_found)} dari {len(REQUIRED_COLUMNS)}")
            missing_tarikan = set(REQUIRED_COLUMNS) - set(df_tarikan.columns)
            if missing_tarikan:
                st.warning(f"Kolom tidak ditemukan: {', '.join(missing_tarikan)}")
            st.dataframe(df_tarikan.head(10), use_container_width=True)
        
        with col2:
            st.markdown("**Data Upload Anda**")
            st.write(f"Jumlah baris: {len(df_upload)}")
            st.write(f"Kolom ditemukan: {len(upload_cols_found)} dari {len(REQUIRED_COLUMNS)}")
            missing_upload = set(REQUIRED_COLUMNS) - set(df_upload.columns)
            if missing_upload:
                st.warning(f"Kolom tidak ditemukan: {', '.join(missing_upload)}")
            st.dataframe(df_upload.head(10), use_container_width=True)
        
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
            
            ski_options = list(df_tarikan.columns)
            ski_column = st.selectbox(
                "Pilih kolom untuk cek status SKI",
                options=ski_options,
                index=ski_options.index('STATUS') if 'STATUS' in ski_options else 0
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
                    st.metric("Data Tarikan Belum Ada di File Anda", len(missing_in_upload))
                
                if missing_in_upload:
                    st.markdown("### 🔴 Data Tarikan yang Belum Tersedia di File Anda")
                    st.write(f"Ditemukan **{len(missing_in_upload)}** data dari tarikan yang belum ada di file Anda.")
                    
                    df_missing = df_tarikan[df_tarikan['_clean_key'].isin(missing_in_upload)].copy()
                    df_missing = df_missing.drop(columns=['_clean_key'])
                    
                    if ski_column in df_missing.columns:
                        df_missing['Status SKI'] = df_missing[ski_column].apply(
                            lambda x: '✅ Ada SKI' if pd.notna(x) and str(x).strip() != '' else '❌ Belum Ada SKI'
                        )
                        
                        ada_ski = df_missing[df_missing['Status SKI'] == '✅ Ada SKI']
                        belum_ski = df_missing[df_missing['Status SKI'] == '❌ Belum Ada SKI']
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Sudah Ada SKI", len(ada_ski), delta=None)
                        with col2:
                            st.metric("Belum Ada SKI", len(belum_ski), delta=None)
                        
                        st.markdown("#### Data Lengkap dengan Status SKI:")
                        st.dataframe(df_missing, use_container_width=True)
                        
                        st.markdown("---")
                        
                        tab1, tab2 = st.tabs(["✅ Sudah Ada SKI", "❌ Belum Ada SKI"])
                        
                        with tab1:
                            if len(ada_ski) > 0:
                                st.dataframe(ada_ski, use_container_width=True)
                            else:
                                st.info("Tidak ada data dengan SKI")
                        
                        with tab2:
                            if len(belum_ski) > 0:
                                st.dataframe(belum_ski, use_container_width=True)
                            else:
                                st.info("Semua data sudah memiliki SKI")
                    else:
                        st.dataframe(df_missing, use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_missing.to_excel(writer, index=False, sheet_name='Data Tidak Ada di Upload')
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Download Data yang Belum Ada",
                        data=output,
                        file_name="data_belum_tersedia.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.success("✅ Semua data dari tarikan sudah tersedia di file Anda!")
                
                df_tarikan = df_tarikan.drop(columns=['_clean_key'], errors='ignore')
                df_upload = df_upload.drop(columns=['_clean_key'], errors='ignore')
                
        else:
            st.error("⚠️ Tidak ditemukan kolom yang sama antara kedua file. Pastikan kedua file memiliki struktur kolom yang serupa.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")
        st.info("Pastikan file Excel yang diupload dalam format yang benar (.xlsx atau .xls)")

else:
    st.info("👆 Silakan upload kedua file untuk memulai perbandingan")

st.markdown("---")
st.markdown("*Aplikasi Perbandingan Data Realisasi Impor - TW3*")
