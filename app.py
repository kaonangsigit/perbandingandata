import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Perbandingan Data Impor", page_icon="📊", layout="wide")

st.title("📊 Perbandingan Data Realisasi Impor")
st.markdown("---")

st.markdown("""
### Petunjuk Penggunaan:
1. Upload **File Tarikan** (data hasil tarikan dari sistem)
2. Upload **File Data Anda** (data yang ingin dibandingkan)
3. Sistem akan mengidentifikasi data yang belum tersedia di file Anda
4. Untuk data tersebut, akan dicek apakah sudah ada SKI-nya atau tidak
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 File Tarikan (Data Sistem)")
    file_tarikan = st.file_uploader("Upload file hasil tarikan", type=['xlsx', 'xls'], key="tarikan")

with col2:
    st.subheader("📁 File Data Anda")
    file_upload = st.file_uploader("Upload file data Anda untuk dibandingkan", type=['xlsx', 'xls'], key="upload")

if file_tarikan and file_upload:
    try:
        df_tarikan = pd.read_excel(file_tarikan)
        df_upload = pd.read_excel(file_upload)
        
        st.markdown("---")
        st.subheader("📋 Preview Data")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Data Tarikan (Sistem)**")
            st.write(f"Jumlah baris: {len(df_tarikan)}")
            st.write(f"Kolom: {', '.join(df_tarikan.columns.tolist())}")
            st.dataframe(df_tarikan.head(10), use_container_width=True)
        
        with col2:
            st.markdown("**Data Upload Anda**")
            st.write(f"Jumlah baris: {len(df_upload)}")
            st.write(f"Kolom: {', '.join(df_upload.columns.tolist())}")
            st.dataframe(df_upload.head(10), use_container_width=True)
        
        st.markdown("---")
        st.subheader("⚙️ Konfigurasi Perbandingan")
        
        common_cols = list(set(df_tarikan.columns) & set(df_upload.columns))
        
        if common_cols:
            key_column = st.selectbox(
                "Pilih kolom kunci untuk perbandingan (misalnya: No. Pengajuan, ID, dll)",
                options=common_cols,
                index=0
            )
            
            ski_columns = [col for col in df_tarikan.columns if 'ski' in col.lower() or 'surat' in col.lower() or 'keterangan' in col.lower() or 'izin' in col.lower()]
            
            if ski_columns:
                ski_column = st.selectbox(
                    "Pilih kolom SKI (Surat Keterangan Impor)",
                    options=ski_columns + ['Lainnya...'],
                    index=0
                )
                if ski_column == 'Lainnya...':
                    ski_column = st.selectbox("Pilih kolom SKI dari semua kolom:", options=df_tarikan.columns.tolist())
            else:
                ski_column = st.selectbox(
                    "Pilih kolom SKI (Surat Keterangan Impor)",
                    options=df_tarikan.columns.tolist()
                )
            
            if st.button("🔍 Bandingkan Data", type="primary"):
                st.markdown("---")
                st.subheader("📊 Hasil Perbandingan")
                
                tarikan_keys = set(df_tarikan[key_column].dropna().astype(str))
                upload_keys = set(df_upload[key_column].dropna().astype(str))
                
                missing_in_upload = tarikan_keys - upload_keys
                missing_in_tarikan = upload_keys - tarikan_keys
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Data di Tarikan", len(tarikan_keys))
                with col2:
                    st.metric("Data di Upload Anda", len(upload_keys))
                with col3:
                    st.metric("Data Belum Ada di Upload Anda", len(missing_in_upload))
                
                if missing_in_upload:
                    st.markdown("### 🔴 Data yang Belum Tersedia di File Anda")
                    st.write(f"Ditemukan **{len(missing_in_upload)}** data dari sistem yang belum ada di file Anda.")
                    
                    df_missing = df_tarikan[df_tarikan[key_column].astype(str).isin(missing_in_upload)].copy()
                    
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
                    st.success("✅ Semua data dari sistem sudah tersedia di file Anda!")
                
                if missing_in_tarikan:
                    st.markdown("### 🟡 Data di File Anda yang Tidak Ada di Sistem")
                    st.write(f"Ditemukan **{len(missing_in_tarikan)}** data di file Anda yang tidak ada di sistem.")
                    
                    df_extra = df_upload[df_upload[key_column].astype(str).isin(missing_in_tarikan)]
                    st.dataframe(df_extra, use_container_width=True)
        else:
            st.error("⚠️ Tidak ditemukan kolom yang sama antara kedua file. Pastikan kedua file memiliki struktur kolom yang serupa.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")
        st.info("Pastikan file Excel yang diupload dalam format yang benar (.xlsx atau .xls)")

else:
    st.info("👆 Silakan upload kedua file untuk memulai perbandingan")

st.markdown("---")
st.markdown("*Aplikasi Perbandingan Data Realisasi Impor - TW3*")
