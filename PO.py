import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import warnings
import matplotlib.pyplot as plt
import pmdarima as pm
from pmdarima import auto_arima
from sklearn.linear_model import LinearRegression
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from fpdf import FPDF
import base64
warnings.filterwarnings('ignore')


st.markdown('<h1 class="main-header"> Dashboard Analisis Performa </h1>', unsafe_allow_html=True)
with st.sidebar:
    st.header("📁 Upload Data")
    uploaded_file = st.file_uploader("Upload file Excel", type=['xlsx', 'xls'])

    uploaded_file_2 = st.file_uploader("Upload File Excel BSM", type=['xlsx', 'xls'], key="file2")

    uploaded_con = st.file_uploader("Upload Master Consumable", type=['xlsx', 'xls'], key="master_con")
    uploaded_noncon = st.file_uploader("Upload Master Non-Consumable", type=['xlsx', 'xls'], key="master_noncon")
    
    st.markdown("---")
    st.header("📅 Filter Periode")
    
    # Default value: Hari ini
    today = datetime.now().date()
    # Default start: 1 tahun lalu (opsional, bisa disesuaikan)
    last_year = today.replace(year=today.year - 1)

    start_date = st.date_input("Mulai Tanggal", value=last_year)
    end_date = st.date_input("Sampai Tanggal", value=today)
    
    st.markdown("---")
    st.info("""
    **Instruksi:**
    1. Upload file Excel dengan 3 sheet (PO, BP, SJ)
    2. Pilih jenis analisis
    3. Gunakan filter untuk menyesuaikan data
    4. Ekspor hasil jika diperlukan
    """)




# Fungsi untuk memproses data
@st.cache_data
def process_excel_file(file):
    try:
        # Baca semua sheet
        po_sheet = pd.read_excel(file, sheet_name='PO',header=2)
        bp_sheet = pd.read_excel(file, sheet_name='BP',header=2)
        sj_sheet = pd.read_excel(file, sheet_name='SJ',header=3)
        
        # Standardisasi nama kolom
        po_sheet.columns = po_sheet.columns.str.strip().str.upper()
        bp_sheet.columns = bp_sheet.columns.str.strip().str.upper()
        sj_sheet.columns = sj_sheet.columns.str.strip().str.upper()

        po_sheet['PO_CREATED_ON'] = pd.to_datetime(po_sheet['PO_CREATED_ON'], dayfirst=True, errors='coerce')
        po_sheet['PO_APPROVED_ON'] = pd.to_datetime(po_sheet['PO_APPROVED_ON'], dayfirst=True, errors='coerce')
        bp_sheet['BP_CREATED_ON'] = pd.to_datetime(bp_sheet['BP_CREATED_ON'], dayfirst=True, errors='coerce')
        sj_sheet['SJ_CREATED_ON'] = pd.to_datetime(sj_sheet['SJ_CREATED_ON'], dayfirst=True, errors='coerce')
        sj_sheet['SJ_CLOSED_ON'] = pd.to_datetime(sj_sheet['SJ_CLOSED_ON'], dayfirst=True, errors='coerce')
        sj_sheet['TGLAPP'] = pd.to_datetime(sj_sheet['TGLAPP'], dayfirst=True, errors='coerce')   
        
        


        cols_to_drop_po = [
         "DOCKING","KODE_KATEGORI","KODE_BARANG","SATUAN","KODEAKUN_BARANG","NAMAKODEAKUN_BARANG","KODEAKUN_PO","NAMAKODEAKUN_PO",
        "PURCHASE_TYPE","JUMLAH","JML_DITERIMA","DELIVERY_TIME","VESSELID","VESSELNAME","KODEALATBERAT","KODE2","NAMAALATBERAT","HVE_OR_VESSEL","REF","REF2",
        "CATEGORY","STATUS","USER_ID","PAYMENT_TERM","DELIVERY_TERM","PO_APPROVED_BY","KODE_LOKASI","PO_KETERANGAN","PO_REMARK","PO_REMARK2",
        "BPS_NO","BPS_KETERANGAN"
            ]
        po_sheet = po_sheet.drop(columns=cols_to_drop_po, errors="ignore")

        cols_to_drop_bp = [
        'KODE_SUPPLIER',		'SATUAN',	'KODEAKUN_BARANG',	'NAMAKODEAKUN_BARANG',	'BP_KETERANGAN',
        'KODE_BARANG',	'DETAIL_KETERANGAN',	'TUJUAN',	'VESSELCODE',	'JML_DITERIMA',
            ]
        bp_sheet = bp_sheet.drop(columns=cols_to_drop_bp, errors="ignore")

        cols_to_drop_sj = [
        'KEPADA',	'VESSELID',	'NAMA_KAPAL','NAMA_KATEGORI',	'SATUAN',	'KODEAKUN_BARANG',	'NAMAKODEAKUN_BARANG',	'JUMLAH',
        'JMLDISETUJUI',	'CREATED_BY', 'KETERANGAN',	'JML_DITERIMA',	'RELEASE',	'BPK_NO',	'USERIDAPP',	
        'STATUSAPP',	'SJ_CLOSED_BY',	'NOSBT',	'LOKASIBUAT',	
            ]
        sj_sheet = sj_sheet.drop(columns=cols_to_drop_sj, errors="ignore")


        return {
            'PO': po_sheet,
            'BP': bp_sheet,
            'SJ': sj_sheet
        }
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
        return None
    



def process_excel_file2(file):
    try:
        # Baca semua sheet
        bsm_sheet = pd.read_excel(file)
        # bsm_sheet = pd.read_excel(file, sheet_name='Export Worksheet')  

        bsm_sheet.columns = bsm_sheet.columns.str.strip().str.upper()

        bsm_sheet['BSM_CREATED_ON'] = pd.to_datetime(bsm_sheet['BSM_CREATED_ON'], dayfirst=True, errors='coerce')  
        bsm_sheet['DATEAPP'] = pd.to_datetime(bsm_sheet['DATEAPP'], dayfirst=True, errors='coerce')  
        bsm_sheet['BSM_CLOSED_ON'] = pd.to_datetime(bsm_sheet['BSM_CLOSED_ON'], dayfirst=True, errors='coerce')  

        cols_to_drop = [
            'NOBSM',	'STATUS',
            'BPB_NO',	'NOSPB','JENISSTOK',	'NOLOKK',	'NOLOKB',	'NOLA',	'USERIDAPP',	
            'BSM_CLOSED_BY',	'LOKASIBUAT',	'KODEKAPAL',	'KODEALATBERAT',	'NAMAOBJEK',	
            'KODEAKUNBRG',	'KODEAKUNBSM',	'KETBPB',	'KETSPB',	'BAGIAN',
            ]
        bsm_sheet = bsm_sheet.drop(columns=cols_to_drop, errors="ignore")

        return {
            'BSM': bsm_sheet
        }
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
        return None


#FUNCTION ANALISIS
def tampilkan_analisis_po(po_raw, kolom_pilihan):
    st.title("Analisis Durasi Approval PO")
    st.subheader("Tabel Data")
    
    # Menggunakan dataframe agar bisa di-scroll dan di-sort
    st.dataframe(po_raw[kolom_pilihan], use_container_width=True) 

    st.subheader(f"Rata-rata waktu tunggu: {po_raw['durasi_proses'].mean()}")   
    # Jika ingin tabel statis (tanpa scroll) bisa gunakan ini:
    # st.table(po_raw[kolom_pilihan])
    
    # Menghitung total baris dan menampilkan metric
    total_baris = len(po_raw)
    st.metric("Total Transaksi", f"{total_baris} Row")


def tracking_po_tanpa_bp(po_raw, bp_raw):
    # 1. Menggabungkan data PO dan BP
    merged = pd.merge(po_raw, bp_raw, on="PO_NO", how="left")
    
    st.markdown("---")
    
    # Memfilter PO yang sudah approved tapi belum ada BP
    outstanding_bp = merged[
        (merged['PO_APPROVED_ON'].notnull()) & 
        (merged['BP_CREATED_ON'].isna())
    ].copy()

    today_ts = pd.Timestamp(datetime.now().date()) # Ambil tanggal hari ini
    
    # Hitung selisih hari
    outstanding_bp['durasi'] = (today_ts - outstanding_bp['PO_APPROVED_ON']).dt.days

    # 2. Tampilkan di Streamlit
    st.header("🔍 Tracking PO Belum ada BP")

    # Tampilkan metrik jumlah PO yang menggantung
    total_outstanding = len(outstanding_bp)
    
    if total_outstanding > 0:
        st.warning(f"Terdapat {total_outstanding} PO yang sudah Approved tapi belum ada BP.")
        
        # Urutkan dari yang terlama (durasi terbesar) agar prioritas penanganan terlihat
        outstanding_bp = outstanding_bp.sort_values(by='durasi', ascending=False)

        # Masukkan kolom 'durasi' ke dalam list kolom yang akan ditampilkan
        kolom_outstanding = ['PO_NO', 'PO_APPROVED_ON', 'durasi'] 
        
        # Tampilkan Dataframe
        st.dataframe(
            outstanding_bp[kolom_outstanding], 
            use_container_width=True,
            column_config={
                "durasi": st.column_config.NumberColumn(
                    "Durasi (Hari)",
                    help="Selisih hari dari PO Approved hingga Sekarang",
                    format="%d Hari" 
                )
            }
        )
    else:
        st.success("Kerja Bagus! Tidak ada PO yang menunggu BP (Outstanding).")

def analisis_durasi_po_ke_bp(merged, bp_raw):
    st.markdown("---") 
    
    # Filter hanya data yang sudah ada BP-nya
    df_clean = merged[merged['BP_CREATED_ON'].notnull()].copy()
    
    # Hitung durasi proses
    df_clean['durasi_prosesBP'] = df_clean['BP_CREATED_ON'] - df_clean['PO_APPROVED_ON']

    kolom_pilihanBP = ['PO_NO', 'BP_NO', 'PO_APPROVED_ON', 'BP_CREATED_ON', 'durasi_prosesBP']
    
    # Tampilkan di Streamlit
    st.title("Analisis Durasi Vendor dari PO hingga BP")
    st.subheader("Tabel Data")
    st.dataframe(df_clean[kolom_pilihanBP], use_container_width=True) 
    
    # Menghitung dan menampilkan rata-rata waktu tunggu
    rata_rata = df_clean['durasi_prosesBP'].mean()
    st.subheader(f"Rata-rata waktu tunggu: {rata_rata}")
    
    # Menampilkan total transaksi
    total_baris = len(df_clean)
    st.metric("Total Transaksi", f"{total_baris} Row")
    
    # Hitung kemunculan tiap PO_NO di tabel BP
    po_counts = bp_raw['PO_NO'].value_counts()
    
    # Mengembalikan hasil jika dibutuhkan di luar function
    return df_clean, po_counts

def analisis_performa_supplier_bar(df_clean):
    st.markdown("---")
    st.header("🏆 Performa Supplier (Durasi PO Approved ➔ BP Created)")

    # 1. Cari kolom yang mengandung kata 'SUPPLIER' atau 'VENDOR'
    possible_cols = [col for col in df_clean.columns if 'SUPPLIER' in str(col).upper() or 'VENDOR' in str(col).upper()]
    
    if not possible_cols:
        st.error("⚠️ Kolom nama supplier tidak ditemukan di data.")
        return

    # Gunakan kolom pertama yang cocok
    col_supplier = possible_cols[0]
    
    # 2. Konversi durasi (timedelta) menjadi angka hari
    df_clean['durasi_hari'] = df_clean['durasi_prosesBP'].dt.days

    # 3. Filter data yang valid (hindari durasi minus jika ada error input tanggal)
    df_valid = df_clean[df_clean['durasi_hari'] >= 0]

    if df_valid.empty:
        st.warning("Tidak ada data durasi yang valid (≥ 0 hari) untuk dianalisis.")
        return

    # 4. Hitung rata-rata durasi per supplier
    avg_durasi = df_valid.groupby(col_supplier)['durasi_hari'].mean().reset_index()
    avg_durasi['durasi_hari'] = avg_durasi['durasi_hari'].round(1)

    # 5. Ambil Top 5 Tercepat (Ascending) & Top 5 Terlama (Descending)
    top5_fastest = avg_durasi.sort_values(by='durasi_hari', ascending=True).head(5)
    top5_slowest = avg_durasi.sort_values(by='durasi_hari', ascending=False).head(5)

    col1, col2 = st.columns(2)

    # --- CHART 1: TOP 5 TERCEPAT (BAR CHART) ---
    with col1:
        st.subheader("🚀 Top 5 Tercepat")
        if not top5_fastest.empty:
            fig_fast = px.bar(
                top5_fastest,
                x='durasi_hari',
                y=col_supplier,
                orientation='h', # Bikin bar chart horizontal agar nama supplier terbaca jelas
                text='durasi_hari', # Tampilkan angka di ujung batang
                color_discrete_sequence=['#2CA02C'] # Warna Hijau
            )
            # Balik urutan Y-axis agar yang tercepat (paling atas)
            fig_fast.update_layout(
                yaxis={'categoryorder':'total descending'},
                xaxis_title="Rata-rata Durasi (Hari)",
                yaxis_title="Nama Supplier"
            )
            fig_fast.update_traces(textposition='outside')
            st.plotly_chart(fig_fast, use_container_width=True)
        else:
            st.info("Data tidak cukup.")

    # --- CHART 2: TOP 5 TERLAMA (BAR CHART) ---
    with col2:
        st.subheader("🐢 Top 5 Terlama")
        if not top5_slowest.empty:
            fig_slow = px.bar(
                top5_slowest,
                x='durasi_hari',
                y=col_supplier,
                orientation='h',
                text='durasi_hari',
                color_discrete_sequence=['#D62728'] # Warna Merah
            )
            # Urutan Y-axis dibalik agar yang terlama di posisi paling atas
            fig_slow.update_layout(
                yaxis={'categoryorder':'total ascending'},
                xaxis_title="Rata-rata Durasi (Hari)",
                yaxis_title="Nama Supplier"
            )
            fig_slow.update_traces(textposition='outside')
            st.plotly_chart(fig_slow, use_container_width=True)
        else:
            st.info("Data tidak cukup.")

def deteksi_po_multi_bp(bp_raw):
    st.subheader("⚠️ Deteksi PO dengan Multi BP (Pengiriman Bertahap)")

    # 1. Hitung Berapa Kali PO Muncul di Tabel BP
    po_counts = bp_raw['PO_NO'].value_counts()
    
    # 2. Ambil PO yang muncul > 1 kali (Ringkasan)
    duplicates = po_counts[po_counts > 1].reset_index()
    duplicates.columns = ['PO_NO', 'TOTAL_BP_COUNT'] # Kolom baru: Total BP dalam 1 PO

    if not duplicates.empty:
        st.warning(f"Ditemukan {len(duplicates)} PO yang memiliki pengiriman bertahap (Multi-BP).")

        # 3. Ambil Data Detail dari BP Raw
        # Filter bp_raw hanya untuk PO yang ada di daftar 'duplicates'
        list_po_multi = duplicates['PO_NO'].tolist()
        detail_data = bp_raw[bp_raw['PO_NO'].isin(list_po_multi)].copy()

        # 4. LAKUKAN PENGGABUNGAN (MERGE)
        # Menambahkan kolom 'TOTAL_BP_COUNT' ke setiap baris detail barang
        combined_df = pd.merge(detail_data, duplicates, on='PO_NO', how='left')

        # 5. Atur Urutan Kolom Agar Rapi
        # Kita pilih kolom mana saja yang mau ditampilkan
        target_cols = ['PO_NO', 'TOTAL_BP_COUNT', 'BP_NO', 'BP_CREATED_ON', 'NAMABRG']
        
        # Opsional: Jika kolom NAMA_SUPPLIER ada, kita ikutkan juga
        if 'NAMA_SUPPLIER' in combined_df.columns:
            target_cols.insert(1, 'NAMA_SUPPLIER') # Masukkan setelah PO_NO

        # Filter hanya kolom yang benar-benar ada di data (untuk mencegah error)
        valid_cols = [col for col in target_cols if col in combined_df.columns]
        
        # Urutkan Data: Berdasarkan PO, lalu Tanggal BP
        final_table = combined_df[valid_cols].sort_values(by=['PO_NO', 'BP_CREATED_ON'])

        # 6. Tampilkan Tabel Gabungan
        st.dataframe(
            final_table, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "TOTAL_BP_COUNT": st.column_config.NumberColumn(
                    "Jml BP",
                    help="Total berapa kali BP diterbitkan untuk PO ini",
                    format="%d x" # Format tampilan angka
                )
            }
        )

        # 7. Tombol Download Satu File Lengkap
        csv_combined = final_table.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="⬇️ Download Data Lengkap (Gabungan CSV)",
            data=csv_combined,
            file_name='monitoring_po_multi_bp_lengkap.csv',
            mime='text/csv',
            help="Berisi data detail barang ditambah kolom jumlah total BP per PO"
        )

        # 8. Drill Down (Opsional - Tetap ada untuk cek cepat)
        with st.expander("🔍 Filter Cepat per Nomor PO"):
            pilih_po = st.selectbox("Pilih No PO:", duplicates['PO_NO'])
            if pilih_po:
                st.dataframe(final_table[final_table['PO_NO'] == pilih_po], use_container_width=True)

    else:
        st.success("✅ Kerja Bagus! Tidak ada PO dengan pengiriman parsial (Semua 1x Kirim).")

def analisis_perilaku_multi_bp(bp_raw):
    st.markdown("---")
    st.header("🕵️ Analisis Perilaku Pengiriman Bertahap (Multi-BP)")
    st.info("Menganalisis barang apa saja yang sering dicicil pengirimannya dan supplier mana yang paling sering melakukan praktik ini.")

    # 1. Identifikasi PO yang dipecah (Multi-BP)
    po_counts = bp_raw['PO_NO'].value_counts()
    multi_po = po_counts[po_counts > 1].index.tolist()
    
    # Filter data hanya untuk transaksi yang dicicil
    df_multi = bp_raw[bp_raw['PO_NO'].isin(multi_po)].copy()
    
    if df_multi.empty:
        st.success("Aman! Tidak ada data pengiriman bertahap untuk dianalisis saat ini.")
        return

    col1, col2 = st.columns(2)

    # ==========================================
    # CHART 1: TOP 10 BARANG YANG SERING DICICIL
    # ==========================================
    with col1:
        st.subheader("📦 Top 10 Barang Sering Dicicil")
        if 'NAMABRG' in df_multi.columns:
            # Kita hitung berapa banyak PO unik yang memuat barang ini dan akhirnya dicicil
            item_freq = df_multi.groupby('NAMABRG')['PO_NO'].nunique().reset_index()
            item_freq.columns = ['Nama Barang', 'Frekuensi PO Dicicil']
            
            top_items = item_freq.sort_values(by='Frekuensi PO Dicicil', ascending=False).head(10)
            
            if not top_items.empty:
                fig_items = px.bar(
                    top_items,
                    x='Frekuensi PO Dicicil',
                    y='Nama Barang',
                    orientation='h',
                    text='Frekuensi PO Dicicil',
                    color='Frekuensi PO Dicicil',
                    color_continuous_scale='Blues',
                    title="Berdasarkan Jumlah PO yang Dipecah"
                )
                fig_items.update_layout(yaxis={'categoryorder':'total ascending'})
                fig_items.update_traces(textposition='outside')
                st.plotly_chart(fig_items, use_container_width=True)
            else:
                st.info("Data barang tidak cukup.")
        else:
            st.error("Kolom 'NAMABRG' tidak ditemukan.")

    # ==========================================
    # CHART 2: TOP 10 SUPPLIER YANG SERING MENCICIL
    # ==========================================
    with col2:
        st.subheader("🏢 Top 10 Supplier Sering Mencicil")
        
        # Cari kolom nama supplier secara dinamis
        possible_cols = [col for col in df_multi.columns if 'SUPPLIER' in str(col).upper() or 'VENDOR' in str(col).upper()]
        
        if possible_cols:
            col_supp = possible_cols[0]
            
            # Hitung berapa banyak PO unik dari supplier ini yang akhirnya dicicil
            supp_freq = df_multi.groupby(col_supp)['PO_NO'].nunique().reset_index()
            supp_freq.columns = ['Nama Supplier', 'Frekuensi PO Dicicil']
            
            top_supp = supp_freq.sort_values(by='Frekuensi PO Dicicil', ascending=False).head(10)
            
            if not top_supp.empty:
                fig_supp = px.bar(
                    top_supp,
                    x='Frekuensi PO Dicicil',
                    y='Nama Supplier',
                    orientation='h',
                    text='Frekuensi PO Dicicil',
                    color='Frekuensi PO Dicicil',
                    color_continuous_scale='Oranges',
                    title="Berdasarkan Jumlah PO yang Dipecah"
                )
                fig_supp.update_layout(yaxis={'categoryorder':'total ascending'})
                fig_supp.update_traces(textposition='outside')
                st.plotly_chart(fig_supp, use_container_width=True)
            else:
                st.info("Data supplier tidak cukup.")
        else:
            st.error("Kolom nama supplier/vendor tidak ditemukan di tabel BP.")

    # --- TABEL DETAIL UNTUK INVESTIGASI ---
    with st.expander("🔍 Klik untuk melihat Detail Transaksi Barang yang Dicicil (Pengecekan Logika)"):
        st.write("Gunakan tabel ini untuk mengecek apakah barang yang dicicil memang wajar (kuantitas masif) atau tidak wajar (kuantitas kecil tapi dicicil).")
        
        # Tampilkan kolom penting
        kolom_tampil = ['PO_NO', 'BP_NO', 'BP_CREATED_ON']
        if 'NAMABRG' in df_multi.columns: kolom_tampil.append('NAMABRG')
        if possible_cols: kolom_tampil.append(possible_cols[0])
        # Tambahkan JML_DITERIMA jika kamu tidak men-dropnya, agar user bisa melihat kuantitas cicilannya
        if 'JML_DITERIMA' in bp_raw.columns: kolom_tampil.append('JML_DITERIMA')
        
        # Filter kolom yang benar-benar ada
        kolom_valid = [col for col in kolom_tampil if col in df_multi.columns]
        
        # Urutkan berdasarkan PO dan Tanggal agar kelihatan historis cicilannya
        tabel_detail = df_multi[kolom_valid].sort_values(by=['PO_NO', 'BP_CREATED_ON'])
        st.dataframe(tabel_detail, use_container_width=True, hide_index=True)

def deteksi_po_ekstrem(bp_raw, batas_bp=49):
    st.markdown("---") 
    st.subheader(f"🔥 Alert: PO dengan Aktivitas BP Ekstrem (> {batas_bp + 1} BP)")

    # 1. Hitung Jumlah BP per PO
    po_counts = bp_raw['PO_NO'].value_counts()
    
    # 2. FILTER KHUSUS: Hanya ambil yang jumlahnya > batas_bp
    extreme_po = po_counts[po_counts > batas_bp].reset_index()
    extreme_po.columns = ['PO_NO', 'TOTAL_BP_COUNT']

    if not extreme_po.empty:
        st.error(f"Ditemukan {len(extreme_po)} PO yang memiliki lebih dari {batas_bp + 1} BP! (Indikasi cicilan pengiriman yang sangat banyak).")

        # 3. Ambil Detail Data dari BP Raw hanya untuk PO tersebut
        list_extreme_po = extreme_po['PO_NO'].tolist()
        detail_extreme = bp_raw[bp_raw['PO_NO'].isin(list_extreme_po)].copy()

        # 4. GABUNGKAN (MERGE) untuk mendapatkan kolom TOTAL_BP_COUNT
        combined_extreme = pd.merge(detail_extreme, extreme_po, on='PO_NO', how='left')

        # 5. Pilih Kolom
        target_cols = ['PO_NO', 'TOTAL_BP_COUNT', 'BP_NO', 'BP_CREATED_ON', 'NAMABRG']
        
        # Cek jika kolom NAMA_SUPPLIER ada, masukkan juga
        if 'NAMA_SUPPLIER' in combined_extreme.columns:
            target_cols.insert(1, 'NAMA_SUPPLIER') 

        # Filter kolom yang valid saja
        valid_cols = [col for col in target_cols if col in combined_extreme.columns]

        # 6. Sorting: Urutkan dari yang TOTAL_BP-nya paling BANYAK (Descending)
        final_extreme = combined_extreme[valid_cols].sort_values(
            by=['TOTAL_BP_COUNT', 'PO_NO', 'BP_CREATED_ON'], 
            ascending=[False, True, True]
        )

        # 7. Tampilkan Tabel
        st.dataframe(
            final_extreme, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "TOTAL_BP_COUNT": st.column_config.NumberColumn(
                    "Total BP",
                    help="Jumlah total BP untuk PO ini",
                    format="%d ⚠️" # Menambahkan icon warning di angka
                )
            }
        )

        # 8. Tombol Download
        csv_extreme = final_extreme.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=f"⬇️ Download Data PO Ekstrem (>{batas_bp + 1} BP)",
            data=csv_extreme,
            file_name='data_po_extreme.csv',
            mime='text/csv',
            key='download_extreme_po_alert' # Key unik agar tidak bentrok
        )

    else:
        st.success(f"✅ Aman! Tidak ada PO yang memiliki BP di atas {batas_bp + 1}.")


# TAB 2
def tracking_bp_open(bp_raw):
    st.markdown("---") 
    st.subheader("🔍 Tracking BP dengan Status OPEN")

    # Pastikan kolom yang dibutuhkan ada di dataframe
    if 'STATUS' in bp_raw.columns and 'NAMA_SUPPLIER' in bp_raw.columns:
        
        # 1. Filter hanya BP yang berstatus 'OPEN'
        # Menggunakan str.strip().str.upper() untuk menghindari error karena spasi, misal "OPEN " atau "open"
        bp_open = bp_raw[bp_raw['STATUS'].astype(str).str.strip().str.upper() == 'OPEN'].copy()

        # 2. Tentukan kolom yang ingin ditampilkan
        kolom_bp_open = ['BP_NO', 'BP_CREATED_ON', 'NAMA_SUPPLIER', 'STATUS']

        if not bp_open.empty:
            st.warning(f"Ditemukan {len(bp_open)} Bukti Penerimaan (BP) yang masih berstatus OPEN.")

            # Urutkan dari tanggal pembuatan paling lama ke terbaru
            bp_open = bp_open.sort_values(by='BP_CREATED_ON', ascending=True)

            # 3. Tampilkan Tabel
            st.dataframe(
                bp_open[kolom_bp_open], 
                use_container_width=True, 
                hide_index=True
            )

            # 4. Tombol Download CSV
            csv_bp_open = bp_open[kolom_bp_open].to_csv(index=False).encode('utf-8')
            st.download_button(
                label="⬇️ Download Data BP OPEN (CSV)",
                data=csv_bp_open,
                file_name='data_bp_status_open.csv',
                mime='text/csv',
                key='download_bp_open_status' # Key unik
            )
        else:
            st.success("✅ Aman! Tidak ada BP yang berstatus OPEN saat ini.")
            
    else:
        st.error("Gagal menampilkan tabel: Kolom 'STATUS' atau 'NAMA_SUPPLIER' tidak ditemukan. Pastikan Anda sudah menghapusnya dari daftar `cols_to_drop_bp` di fungsi `process_excel_file`.")

# TAB 3
def analisis_durasi_approval_sj(sj_raw):
    st.markdown("---")
    
    # 1. Menghitung durasi approval
    sj_raw['durasi_app_SJ'] = sj_raw['TGLAPP'] - sj_raw['SJ_CREATED_ON']
    
    # Menentukan kolom yang akan ditampilkan di tabel
    kolom_pilihanSJ = ['SJ_NO', 'SJ_CREATED_ON', 'TGLAPP', 'durasi_app_SJ']
    
    # 2. Tampilkan di Streamlit
    st.title("Analisis Durasi Approval SJ")
    st.subheader("Tabel Data")
    
    # Menampilkan dataframe
    st.dataframe(sj_raw[kolom_pilihanSJ], use_container_width=True) 
    
    # Menghitung dan menampilkan rata-rata waktu tunggu
    rata_rata = sj_raw['durasi_app_SJ'].mean()
    st.subheader(f"Rata-rata waktu tunggu: {rata_rata}")   

    # Jika ingin tabel statis (tanpa scroll) bisa gunakan ini:
    # st.table(sj_raw[kolom_pilihanSJ])
    
    # Menghitung total baris dan menampilkan metric
    total_baris = len(sj_raw)
    st.metric("Total Transaksi", f"{total_baris} Row")
    
    # Mengembalikan dataframe yang sudah ditambah kolom durasi jika dibutuhkan di luar
    return sj_raw

def deteksi_sj_telat(sj_raw, batas_hari=14):
    st.markdown("---")
    st.subheader(f"⚠️ Alert: Approval SJ Melebihi {batas_hari} Hari")

    # Keamanan: Pastikan kolom 'durasi_app_SJ' sudah ada dari proses sebelumnya
    if 'durasi_app_SJ' not in sj_raw.columns:
        st.error("Kolom 'durasi_app_SJ' tidak ditemukan. Pastikan Anda sudah menjalankan fungsi analisis durasi terlebih dahulu.")
        return

    # 1. Buat kolom bantuan untuk menghitung hari (integer)
    sj_raw['hari_tunggu_app'] = sj_raw['durasi_app_SJ'].dt.days

    # 2. Filter data yang > batas_hari
    late_sj = sj_raw[sj_raw['hari_tunggu_app'] > batas_hari].copy()

    # Urutkan dari yang paling lama menunggunya
    late_sj = late_sj.sort_values(by='hari_tunggu_app', ascending=False)

    # Kolom yang akan ditampilkan dan didownload
    cols_late = ['SJ_NO', 'SJ_CREATED_ON', 'TGLAPP', 'durasi_app_SJ']
    
    # Filter hanya kolom yang valid agar tidak error jika ada kolom yang hilang
    valid_cols = [col for col in cols_late if col in late_sj.columns]

    # 3. Tampilkan Logika Streamlit
    if not late_sj.empty:
        st.error(f"Ditemukan {len(late_sj)} Surat Jalan yang approval-nya lebih dari {batas_hari} hari!")
        
        # Tampilkan Tabel
        st.dataframe(late_sj[valid_cols], use_container_width=True)

        # Siapkan data CSV
        csv_late_sj = late_sj[valid_cols].to_csv(index=False).encode('utf-8')

        # Tombol Download
        st.download_button(
            label="⬇️ Download Data SJ Telat (CSV)",
            data=csv_late_sj,
            file_name=f'sj_approval_over_{batas_hari}_days.csv',
            mime='text/csv',
            key='download_late_sj'
        )
    else:
        st.success(f"✅ Aman! Tidak ada Approval SJ yang melebihi {batas_hari} hari.")
def analisis_durasi_proses_sj(sj_raw):
    st.markdown("---")

    # 1. Menghitung durasi proses (SJ Closed - SJ Approved)
    sj_raw['durasi_proses_SJ'] = sj_raw['SJ_CLOSED_ON'] - sj_raw['TGLAPP']
    
    # Menentukan kolom yang akan ditampilkan di tabel
    kolom_pilihanSJ = ['SJ_NO', 'TGLAPP', 'SJ_CLOSED_ON', 'durasi_proses_SJ']

    # 2. Tampilkan di Streamlit
    st.title("Analisis Durasi Proses SJ Approve --> SJ Close")
    st.subheader("Tabel Data")
    
    # Menampilkan dataframe
    st.dataframe(sj_raw[kolom_pilihanSJ], use_container_width=True) 
    
    # Menghitung dan menampilkan rata-rata waktu tunggu
    rata_rata = sj_raw['durasi_proses_SJ'].mean()
    st.subheader(f"Rata-rata waktu tunggu: {rata_rata}")   
    
    # Menghitung total baris dan menampilkan metric
    total_baris = len(sj_raw['durasi_proses_SJ'])
    st.metric("Total Transaksi", f"{total_baris} Row")
    
    return sj_raw


def tampilkan_kpi_dashboard(po_raw, merged, sj_raw):
    st.header("Ringkasan Performa")
    
    # Membuat 4 kolom sejajar untuk menempatkan Card KPI
    col1, col2, col3, col4 = st.columns(4)
    
    if 'durasi_proses' in po_raw.columns:
        avg_app_po = po_raw['durasi_proses'].mean() / pd.Timedelta(days=1)
    else:
        avg_app_po = 0.0

    if 'BP_CREATED_ON' in merged.columns and 'PO_APPROVED_ON' in merged.columns:
        df_clean_bp = merged[merged['BP_CREATED_ON'].notnull()].copy()
        durasi_po_bp = df_clean_bp['BP_CREATED_ON'] - df_clean_bp['PO_APPROVED_ON']
        avg_po_bp = durasi_po_bp.mean() / pd.Timedelta(days=1)
    else:
        avg_po_bp = 0.0

    if 'TGLAPP' in sj_raw.columns and 'SJ_CREATED_ON' in sj_raw.columns:
        durasi_app_sj = sj_raw['TGLAPP'] - sj_raw['SJ_CREATED_ON']
        avg_app_sj = durasi_app_sj.mean() / pd.Timedelta(days=1)
    else:
        avg_app_sj = 0.0

    if 'SJ_CLOSED_ON' in sj_raw.columns and 'TGLAPP' in sj_raw.columns:
        durasi_proses_sj = sj_raw['SJ_CLOSED_ON'] - sj_raw['TGLAPP']
        avg_proses_sj = durasi_proses_sj.mean() / pd.Timedelta(days=1)
    else:
        avg_proses_sj = 0.0

    with col1:
        st.metric(label="Approval PO", value=f"{avg_app_po:.2f} Hari", 
                  help="Rata-rata waktu pembuatan hingga approval PO")
    with col2:
        st.metric(label="PO hingga BP", value=f"{avg_po_bp:.2f} Hari", 
                  help="Rata-rata waktu dari PO Approved hingga barang diterima (BP)")
    with col3:
        st.metric(label="Approval SJ", value=f"{avg_app_sj:.2f} Hari", 
                  help="Rata-rata waktu pembuatan hingga approval Surat Jalan")
    with col4:
        st.metric(label="SJ Appr ➔ Close", value=f"{avg_proses_sj:.2f} Hari", 
                  help="Rata-rata waktu dari Surat Jalan di-approve hingga diclose")

def analisis_top_barang_by_kategori(po_raw, file_consumable, file_nonconsumable, col_nama_barang='NAMA_BARANG', col_qty='JML_DISETUJUI'):

    st.markdown("---") 
    st.header("🏆 Top 5 Barang Terbanyak dipesan (Consumable & Non-Consumable)")
    if file_consumable is None or file_nonconsumable is None:
        st.warning("⚠️ Mohon upload file Master Consumable dan Non-Consumable di sidebar terlebih dahulu untuk melihat analisis ini.")
        return

    # 1. Cek keberadaan kolom
    if col_nama_barang not in po_raw.columns:
        st.error(f"Error: Kolom '{col_nama_barang}' tidak ditemukan di po_raw.")
        return

    if col_qty not in po_raw.columns:
        st.error(f"Error: Kolom '{col_qty}' (JML_DISETUJUI) tidak ditemukan di po_raw.")
        return

    # 2. Load Data Master
    try:
        df_master_con = pd.read_excel(file_consumable)
        df_master_noncon = pd.read_excel(file_nonconsumable)

        if col_nama_barang not in df_master_con.columns:
            st.error(f"Error: Kolom '{col_nama_barang}' tidak ditemukan di file {file_consumable}.")
            return

        if col_nama_barang not in df_master_noncon.columns: 
            st.error(f"Error: Kolom '{col_nama_barang}' tidak ditemukan di file {file_nonconsumable}.")
            return 

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file master Excel: {e}")
        return

    # 3. Tambah kategori ke master
    df_master_con["Kategori_Master"] = "Consumable"
    df_master_noncon["Kategori_Master"] = "Non-Consumable"

    # 4. Gabungkan master
    master_combined = pd.concat([
        df_master_con[[col_nama_barang, "Kategori_Master"]],
        df_master_noncon[[col_nama_barang, "Kategori_Master"]]
    ], ignore_index=True).drop_duplicates(subset=[col_nama_barang])

    # 5. Hitung total jumlah barang DIBELI (PCS/QTY)
    item_counts = (
        po_raw.groupby(col_nama_barang)[col_qty]
        .sum()
        .reset_index()
        .sort_values(by=col_qty, ascending=False)
    )

    item_counts.rename(columns={col_qty: "Total_Dipesan"}, inplace=True)

    # 6. Merge transaksi dengan kategori
    merged_items = pd.merge(
        item_counts,
        master_combined,
        on=col_nama_barang,
        how="left"
    )

    merged_items["Kategori_Master"] = merged_items["Kategori_Master"].fillna("Lainnya/Belum Terdaftar")

    # 7. Visualisasi
    # 7. Visualisasi (Atas-Bawah & Diperbesar)
    
    # ---- TOP 5 CONSUMABLE ----
    st.subheader("🛠️ Top 5 Consumable (berdasarkan total PCS dipesan)")

    top5_con = merged_items[merged_items["Kategori_Master"] == "Consumable"].head(5)

    if not top5_con.empty:
        fig_con = px.pie(
            top5_con,
            values="Total_Dipesan",
            names=col_nama_barang,
            hole=0.4, # Sedikit diperkecil lubangnya agar area warna lebih luas
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_con.update_traces(textposition="inside", textinfo="percent+label", textfont_size=14)
        
        # Memperbesar ukuran chart secara spesifik (Tinggi 550 pixel)
        fig_con.update_layout(
            height=550, 
            margin=dict(t=30, b=30, l=0, r=0),
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5) # Pindah legenda ke bawah agar chart lebih lega
        )
        st.plotly_chart(fig_con, use_container_width=True)
    else:
        st.info("Tidak ada barang Consumable ditemukan.")

    st.markdown("---") # Garis pembatas yang rapi

    # ---- TOP 5 NON-CONSUMABLE ----
    st.subheader("📦 Top 5 Non-Consumable (berdasarkan total PCS dipesan)")

    top5_noncon = merged_items[merged_items["Kategori_Master"] == "Non-Consumable"].head(5)

    if not top5_noncon.empty:
        fig_noncon = px.pie(
            top5_noncon,
            values="Total_Dipesan",
            names=col_nama_barang,
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Plotly
        )
        fig_noncon.update_traces(textposition="inside", textinfo="percent+label", textfont_size=14)
        
        # Memperbesar ukuran chart secara spesifik (Tinggi 550 pixel)
        fig_noncon.update_layout(
            height=550, 
            margin=dict(t=30, b=30, l=0, r=0),
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5) # Pindah legenda ke bawah
        )
        st.plotly_chart(fig_noncon, use_container_width=True)
    else:
        st.info("Tidak ada barang Non-Consumable ditemukan.")

def top_10_barang_bsm_by_kategori(bsm_raw, file_consumable, file_nonconsumable):
    st.header("🏆 Top 10 Barang Terbanyak (Berdasarkan Approval)")
    if file_consumable is None or file_nonconsumable is None:
        st.warning("⚠️ Mohon upload file Master Consumable dan Non-Consumable di sidebar terlebih dahulu untuk melihat analisis ini.")
        return

    # --- FITUR FILTER LOKASI ---
    if 'KODELOKASI' in bsm_raw.columns: 
        list_lokasi = bsm_raw['KODELOKASI'].dropna().unique().tolist()
        list_lokasi.sort()
        
        selected_lokasi = st.multiselect(
            "📍 Filter Berdasarkan Lokasi (KODELOKASI):",
            options=list_lokasi,
            default=list_lokasi,
            help="Pilih lokasi untuk melihat top barang spesifik di area tersebut."
        )
        bsm_filtered_loc = bsm_raw[bsm_raw['KODELOKASI'].isin(selected_lokasi)].copy()
    else:
        st.warning("Kolom 'KODELOKASI' tidak ditemukan. Menampilkan semua data.")
        bsm_filtered_loc = bsm_raw.copy()

    # --- LOAD MASTER BARANG ---
    try:
        df_master_con = pd.read_excel(file_consumable)
        df_master_noncon = pd.read_excel(file_nonconsumable)
        
        # Di data master namanya 'NAMA_BARANG', di BSM namanya 'NAMABRG'
        col_master_nama = 'NAMA_BARANG'
        if col_master_nama not in df_master_con.columns or col_master_nama not in df_master_noncon.columns:
            st.error(f"Error: Kolom '{col_master_nama}' tidak ditemukan di file master.")
            return

        df_master_con["Kategori_Master"] = "Consumable"
        df_master_noncon["Kategori_Master"] = "Non-Consumable"

        # Gabungkan dan hapus duplikat
        master_combined = pd.concat([
            df_master_con[[col_master_nama, "Kategori_Master"]],
            df_master_noncon[[col_master_nama, "Kategori_Master"]]
        ], ignore_index=True).drop_duplicates(subset=[col_master_nama])
        
        # Ubah nama kolom agar cocok saat di-merge dengan bsm_raw
        master_combined = master_combined.rename(columns={col_master_nama: 'NAMABRG'})
        
    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file master Excel: {e}")
        return

    # --- PROSES VISUALISASI ---
    if 'NAMABRG' in bsm_filtered_loc.columns and 'JMLDISETUJUI' in bsm_filtered_loc.columns:
        # Konversi ke angka & isi NaN dengan 0
        bsm_filtered_loc['JMLDISETUJUI'] = pd.to_numeric(bsm_filtered_loc['JMLDISETUJUI'], errors='coerce').fillna(0)
        
        if bsm_filtered_loc.empty:
            st.warning("Tidak ada data untuk lokasi yang dipilih.")
            return

        # Gabungkan data BSM dengan kategori Master
        merged_bsm = pd.merge(bsm_filtered_loc, master_combined, on='NAMABRG', how='left')
        merged_bsm['Kategori_Master'] = merged_bsm['Kategori_Master'].fillna('Lainnya/Belum Terdaftar')
        
        # Grouping berdasarkan nama barang dan kategori
        grouped_items = merged_bsm.groupby(['NAMABRG', 'Kategori_Master'])['JMLDISETUJUI'].sum().reset_index()
        
        # Pisahkan data Consumable & Non-Consumable
        con_items = grouped_items[grouped_items['Kategori_Master'] == 'Consumable'].sort_values(by='JMLDISETUJUI', ascending=False).head(10)
        noncon_items = grouped_items[grouped_items['Kategori_Master'] == 'Non-Consumable'].sort_values(by='JMLDISETUJUI', ascending=False).head(10)
        
        # --- BUAT 2 KOLOM UNTUK CHART ---
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🛠️ Top 10 Consumable")
            if not con_items.empty:
                fig_con = px.bar(
                    con_items, x='JMLDISETUJUI', y='NAMABRG', orientation='h',
                    text='JMLDISETUJUI', 
                    color='JMLDISETUJUI', color_continuous_scale='Teal'
                )
                fig_con.update_layout(yaxis=dict(autorange="reversed")) # Ranking 1 di atas
                st.plotly_chart(fig_con, use_container_width=True)
            else:
                st.info("Tidak ada data Consumable.")
                
        with col2:
            st.subheader("📦 Top 10 Non-Consumable")
            if not noncon_items.empty:
                fig_noncon = px.bar(
                    noncon_items, x='JMLDISETUJUI', y='NAMABRG', orientation='h',
                    text='JMLDISETUJUI', 
                    color='JMLDISETUJUI', color_continuous_scale='Purples'
                )
                fig_noncon.update_layout(yaxis=dict(autorange="reversed")) # Ranking 1 di atas
                st.plotly_chart(fig_noncon, use_container_width=True)
            else:
                st.info("Tidak ada data Non-Consumable.")
                
        # (Opsional) Tampilkan barang yang belum masuk di file Master
        unclassified = grouped_items[grouped_items['Kategori_Master'] == 'Lainnya/Belum Terdaftar'].sort_values(by='JMLDISETUJUI', ascending=False).head(10)
        if not unclassified.empty:
            with st.expander("⚠️ Lihat Top 10 Barang yang Belum Terdaftar di Master Excel"):
                fig_unclass = px.bar(
                    unclassified, x='JMLDISETUJUI', y='NAMABRG', orientation='h',
                    text='JMLDISETUJUI', 
                    color='JMLDISETUJUI', color_continuous_scale='Reds'
                )
                fig_unclass.update_layout(yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_unclass, use_container_width=True)

    else:
        st.error("Kolom 'NAMABRG' atau 'JMLDISETUJUI' tidak ditemukan. Pastikan tidak terhapus di cols_to_drop.")

# BSM

def analisis_fulfillment_rate_bsm(bsm_raw):
    st.markdown("---")
    st.header("⚖️ Rasio Pemenuhan Permintaan (Fulfillment Rate BSM)")

    # 1. Validasi Kolom
    kolom_wajib = ['NAMABRG', 'JUMLAH', 'JMLDISETUJUI']
    for col in kolom_wajib:
        if col not in bsm_raw.columns:
            st.error(f"⚠️ Kolom '{col}' tidak ditemukan. Pastikan kamu sudah menghapusnya dari daftar 'cols_to_drop' di fungsi process_excel_file2.")
            return

    # 2. Membersihkan Data (Pastikan format angka)
    df = bsm_raw.copy()
    df['JUMLAH'] = pd.to_numeric(df['JUMLAH'], errors='coerce').fillna(0)
    df['JMLDISETUJUI'] = pd.to_numeric(df['JMLDISETUJUI'], errors='coerce').fillna(0)

    # 3. Agregasi Data berdasarkan Nama Barang
    df_grouped = df.groupby('NAMABRG')[['JUMLAH', 'JMLDISETUJUI']].sum().reset_index()

    # Hitung Selisih (Kuantitas yang dipotong) dan Persentase Fulfillment
    df_grouped['Selisih_Pemotongan'] = df_grouped['JUMLAH'] - df_grouped['JMLDISETUJUI']
    df_grouped['Fulfillment_Rate_%'] = (df_grouped['JMLDISETUJUI'] / df_grouped['JUMLAH'] * 100).fillna(0).round(1)

    # 4. Metrik KPI Keseluruhan
    total_diminta = df_grouped['JUMLAH'].sum()
    total_disetujui = df_grouped['JMLDISETUJUI'].sum()
    overall_rate = (total_disetujui / total_diminta * 100) if total_diminta > 0 else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Material Diminta", f"{total_diminta:,.0f}")
    col2.metric("Total Material Disetujui", f"{total_disetujui:,.0f}")
    col3.metric("Overall Fulfillment Rate", f"{overall_rate:.1f}%", help="Persentase total barang yang disetujui dari total yang diminta")

    st.markdown("---")

    # 5. Visualisasi: Top 10 Barang Paling Banyak Diminta (Diminta vs Disetujui)
    st.subheader("📊 Top 10 Barang Paling Banyak Diminta (Diminta vs Disetujui)")
    
    # Ambil Top 10 berdasarkan JUMLAH (Permintaan tertinggi)
    top10_diminta = df_grouped.sort_values(by='JUMLAH', ascending=False).head(10)

    # Gunakan pd.melt agar data rapi untuk Grouped Bar Chart Plotly
    df_melted = top10_diminta.melt(
        id_vars='NAMABRG', 
        value_vars=['JUMLAH', 'JMLDISETUJUI'], 
        var_name='Status', 
        value_name='Kuantitas'
    )

    fig_bar = px.bar(
        df_melted, 
        x='Kuantitas', 
        y='NAMABRG', 
        color='Status', 
        barmode='group', # Menjadikan barnya berdampingan
        orientation='h',
        text='Kuantitas',
        color_discrete_map={'JUMLAH': '#EF553B', 'JMLDISETUJUI': '#00CC96'}, # Merah untuk Diminta, Hijau untuk Disetujui
    )
    # Atur urutan Y-axis agar yang paling besar ada di atas
    fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Kuantitas (Pcs/Unit)")
    fig_bar.update_traces(textposition='outside')
    st.plotly_chart(fig_bar, use_container_width=True)

    # 6. Tabel Deteksi: Barang dengan Pemotongan Kuantitas Terbesar
    st.subheader("✂️ Top 10 Barang dengan Pemotongan Kuantitas Terbesar")
    st.info("Tabel ini menunjukkan barang apa saja yang jumlah *approval*-nya paling jauh di bawah jumlah permintaannya.")
    
    top_selisih = df_grouped[df_grouped['Selisih_Pemotongan'] > 0].sort_values(by='Selisih_Pemotongan', ascending=False).head(10)
    
    if not top_selisih.empty:
        st.dataframe(
            top_selisih[['NAMABRG', 'JUMLAH', 'JMLDISETUJUI', 'Selisih_Pemotongan', 'Fulfillment_Rate_%']],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Fulfillment_Rate_%": st.column_config.ProgressColumn(
                    "Fulfillment Rate (%)",
                    help="Tingkat pemenuhan (0 - 100%)",
                    format="%.1f%%",
                    min_value=0,
                    max_value=100,
                ),
            }
        )
    else:
        st.success("✅ Hebat! Semua permintaan barang disetujui 100% tanpa ada pemotongan.")

def analisis_pengeluaran_abc(bsm_raw):
    st.markdown("---")
    st.header("💰 Spend Analysis (ABC Analysis)")
    st.info("""
    **Metode ABC (Prinsip Pareto 80/20):**
    * **Kelas A:** Barang yang menyerap ~80% dari total pengeluaran (Fokus utama efisiensi).
    * **Kelas B:** Barang yang menyerap ~15% dari total pengeluaran (Fokus menengah).
    * **Kelas C:** Barang yang menyerap ~5% dari total pengeluaran (Fokus rendah, biasanya pernak-pernik murah).
    """)

    # 1. Validasi Kolom
    if 'NAMABRG' not in bsm_raw.columns or 'TOTAL' not in bsm_raw.columns:
        st.error("⚠️ Kolom 'NAMABRG' atau 'TOTAL' tidak ditemukan. Pastikan tidak terhapus di cols_to_drop.")
        return

    df = bsm_raw.copy()
    
    # 2. Pembersihan Data (Pastikan TOTAL adalah angka)
    df['TOTAL'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
    
    # Filter data yang memiliki nilai (buang yang harganya 0 atau kosong)
    df = df[df['TOTAL'] > 0]

    if df.empty:
        st.warning("Tidak ada data pengeluaran (TOTAL) yang valid untuk dianalisis.")
        return

    # 3. Agregasi Total Pengeluaran per Barang
    df_spend = df.groupby('NAMABRG')['TOTAL'].sum().reset_index()
    
    # 4. Urutkan dari Pengeluaran Terbesar ke Terkecil
    df_spend = df_spend.sort_values(by='TOTAL', ascending=False).reset_index(drop=True)

    # 5. Hitung Persentase Kumulatif untuk Klasifikasi ABC
    total_semua_pengeluaran = df_spend['TOTAL'].sum()
    df_spend['Cum_Spend'] = df_spend['TOTAL'].cumsum()
    df_spend['Cum_Perc'] = (df_spend['Cum_Spend'] / total_semua_pengeluaran) * 100

    # 6. Tentukan Kelas ABC
    def tentukan_kelas(persentase):
        if persentase <= 80:
            return 'A (High Value)'
        elif persentase <= 95:
            return 'B (Medium Value)'
        else:
            return 'C (Low Value)'

    df_spend['Kelas_ABC'] = df_spend['Cum_Perc'].apply(tentukan_kelas)

    # --- KPI METRICS ---
    total_nilai = df_spend['TOTAL'].sum()
    item_kelas_a = len(df_spend[df_spend['Kelas_ABC'] == 'A (High Value)'])
    
    col1, col2 = st.columns(2)
    col1.metric("Total Pengeluaran BSM", f"Rp {total_nilai:,.0f}")
    col2.metric("Jumlah Item Kelas A (80% Budget)", f"{item_kelas_a} Barang", help="Fokuskan negosiasi harga pada barang-barang ini.")

    # --- VISUALISASI TREEMAP ---
    st.subheader("🗺️ Peta Pengeluaran (Treemap)")
    # Treemap sangat bagus untuk melihat proporsi uang yang dihabiskan
    fig_tree = px.treemap(
        df_spend.head(50), # Tampilkan top 50 agar grafik tidak terlalu padat
        path=['Kelas_ABC', 'NAMABRG'], 
        values='TOTAL',
        color='Kelas_ABC',
        color_discrete_map={
            'A (High Value)': '#EF553B', # Merah/Orange (Penting)
            'B (Medium Value)': '#FECB52', # Kuning
            'C (Low Value)': '#00CC96' # Hijau
        },
        title="Top 50 Barang Berdasarkan Pengeluaran"
    )
    fig_tree.update_traces(textinfo="label+value")
    st.plotly_chart(fig_tree, use_container_width=True)

    # --- TABEL DATA LENGKAP ---
    st.subheader("📋 Detail Data Kelas ABC")
    
    # Tampilkan Dataframe dengan format Rupiah
    st.dataframe(
        df_spend[['NAMABRG', 'Kelas_ABC', 'TOTAL', 'Cum_Perc']],
        use_container_width=True,
        hide_index=True,
        column_config={
            "TOTAL": st.column_config.NumberColumn(
                "Total Pengeluaran",
                format="Rp %d"
            ),
            "Cum_Perc": st.column_config.ProgressColumn(
                "Kumulatif %",
                format="%.1f%%",
                min_value=0,
                max_value=100
            )
        }
    )

    # Tombol Download
    csv_abc = df_spend.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="⬇️ Download Data Analisis ABC (CSV)",
        data=csv_abc,
        file_name='analisis_abc_pengeluaran.csv',
        mime='text/csv'
    )

    
def analisis_dead_stock(bp_raw, sj_raw):
    st.markdown("---")
    st.header("🕸️ Analisis Dead Stock (Barang Menganggur)")
    st.info("Mendeteksi barang yang sudah diterima di gudang (BP diterbitkan) namun belum didistribusikan (Surat Jalan / SJ belum diterbitkan).")

    # 1. Menyamakan nama kolom kunci (PO_NO)
    sj_temp = sj_raw.copy()
    if 'NO_PO' in sj_temp.columns:
        sj_temp.rename(columns={'NO_PO': 'PO_NO'}, inplace=True)

    if 'PO_NO' not in sj_temp.columns:
        st.error("⚠️ Kolom 'NO_PO' tidak ditemukan di data SJ. Pastikan kamu sudah menghapusnya dari `cols_to_drop_sj`.")
        return

    if 'NAMABRG' not in bp_raw.columns:
        st.error("⚠️ Kolom 'NAMABRG' tidak ditemukan di data BP. Pastikan tidak ter-drop.")
        return

    # 2. Siapkan data BP (Barang Masuk)
    # Kita ambil kolom penting saja agar merge tidak berat
    bp_subset = bp_raw[['PO_NO', 'BP_NO', 'BP_CREATED_ON', 'NAMABRG']].dropna(subset=['PO_NO'])

    # 3. Siapkan data SJ (Barang Keluar)
    # Hapus duplikat PO_NO di SJ agar hasil merge tidak terduplikasi (karena 1 PO bisa punya 1 SJ utama)
    sj_subset = sj_temp[['PO_NO', 'SJ_NO', 'SJ_CREATED_ON']].dropna(subset=['PO_NO']).drop_duplicates(subset=['PO_NO'])

    # 4. Gabungkan BP dan SJ (Left Join)
    merged_stock = pd.merge(bp_subset, sj_subset, on='PO_NO', how='left')

    # 5. Deteksi Dead Stock: Kondisi di mana BP_NO ada, tapi SJ_NO kosong (NaN)
    dead_stock = merged_stock[merged_stock['SJ_NO'].isna()].copy()

    if dead_stock.empty:
        st.success("✅ Luar Biasa! Semua barang yang masuk (BP) sudah diterbitkan Surat Jalan-nya (SJ). Tidak ada Dead Stock.")
        return

    # 6. Hitung Hari Menganggur (Idle Time)
    today_ts = pd.Timestamp(datetime.now().date())
    dead_stock['Lama_Menganggur_Hari'] = (today_ts - dead_stock['BP_CREATED_ON']).dt.days

    # Filter yang menganggur lebih dari 0 hari dan urutkan dari yang paling lama
    dead_stock = dead_stock[dead_stock['Lama_Menganggur_Hari'] >= 0]
    dead_stock = dead_stock.sort_values(by='Lama_Menganggur_Hari', ascending=False)

    # --- KPI METRICS ---
    total_tertahan = len(dead_stock)
    rata_lama = dead_stock['Lama_Menganggur_Hari'].mean()

    col1, col2 = st.columns(2)
    col1.metric("Total Item Tertahan di Gudang", f"{total_tertahan} Transaksi BP", help="Jumlah penerimaan barang yang belum ada Surat Jalan-nya.")
    col2.metric("Rata-rata Lama Menganggur", f"{rata_lama:.1f} Hari", help="Rata-rata waktu barang diam di gudang.")

    # --- VISUALISASI BAR CHART ---
    st.subheader("⚠️ Top 10 Barang Paling Lama Tertahan")
    
    top10_dead = dead_stock.head(10)
    fig_bar = px.bar(
        top10_dead,
        x='Lama_Menganggur_Hari',
        y='NAMABRG',
        orientation='h',
        text='Lama_Menganggur_Hari',
        color='Lama_Menganggur_Hari',
        color_continuous_scale='Reds', # Warna merah mengindikasikan bahaya/perhatian
        title="Berdasarkan Jumlah Hari Sejak BP Diterbitkan"
    )
    # Urutkan Y-axis agar bar terpanjang ada di atas
    fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Lama Menganggur (Hari)")
    st.plotly_chart(fig_bar, use_container_width=True)

    # --- TABEL DETAIL ---
    st.subheader("📋 Detail Data Barang Menganggur")
    kolom_tampil = ['PO_NO', 'BP_NO', 'BP_CREATED_ON', 'NAMABRG', 'Lama_Menganggur_Hari']
    
    st.dataframe(
        dead_stock[kolom_tampil],
        use_container_width=True,
        hide_index=True,
        column_config={
            "BP_CREATED_ON": st.column_config.DateColumn("Tanggal Masuk Gudang (BP)"),
            "Lama_Menganggur_Hari": st.column_config.NumberColumn("Menganggur (Hari)", format="%d Hari")
        }
    )

    # Tombol Download
    csv_dead = dead_stock[kolom_tampil].to_csv(index=False).encode('utf-8')
    st.download_button(
        label="⬇️ Download Data Dead Stock (CSV)",
        data=csv_dead,
        file_name='monitoring_dead_stock.csv',
        mime='text/csv'
    )   

def tampilkan_trend_dokumen(po_raw, bp_raw, sj_raw, bsm_raw):
    st.markdown("---")
    st.header("📈 Tren Penerbitan Dokumen (PO, BP, SJ, BSM)")
    
    # 1. Widget Pilihan Periode
    pilihan_waktu = st.selectbox(
        "Pilih Periode Waktu:",
        ["Harian", "Mingguan", "Bulanan"],
        index=0 # Default ke Harian
    )
    
    # Mapping pilihan ke format resample Pandas
    # 'D' = Daily, 'W-MON' = Weekly (Mulai Senin), 'MS' = Month Start (Awal Bulan)
    freq_map = {"Harian": "D", "Mingguan": "W-MON", "Bulanan": "MS"}
    freq = freq_map[pilihan_waktu]
    
    # 2. Fungsi Bantuan untuk Menghitung Trend
    def hitung_trend(df, date_col, doc_name):
        if date_col not in df.columns:
            return pd.DataFrame(columns=['Tanggal', 'Jumlah', 'Dokumen'])
        
        # Buat copy dan buang baris yang tanggalnya kosong (NaT)
        df_temp = df.dropna(subset=[date_col]).copy()
        
        # Set kolom tanggal sebagai index
        df_temp.set_index(date_col, inplace=True)
        
        # Lakukan resampling (pengelompokan berdasarkan waktu) dan hitung jumlah barisnya
        trend = df_temp.resample(freq).size().reset_index()
        trend.columns = ['Tanggal', 'Jumlah']
        
        # Jika mingguan/bulanan, pastikan hanya mengambil tanggal date-nya saja agar rapi di chart
        trend['Tanggal'] = trend['Tanggal'].dt.date
        trend['Dokumen'] = doc_name
        return trend

    # 3. Hitung Trend untuk Masing-Masing Dokumen
    po_trend = hitung_trend(po_raw, 'PO_CREATED_ON', 'PO')
    bp_trend = hitung_trend(bp_raw, 'BP_CREATED_ON', 'BP')
    sj_trend = hitung_trend(sj_raw, 'SJ_CREATED_ON', 'SJ')
    bsm_trend = hitung_trend(bsm_raw, 'BSM_CREATED_ON','BSM')
    
    # 4. Gabungkan Semua Dataframe Menjadi Satu
    all_trends = pd.concat([po_trend, bp_trend, sj_trend, bsm_trend], ignore_index=True)
    
    # Buang baris yang jumlahnya 0 (opsional: agar grafik tidak menukik ke 0 di hari libur)
    all_trends = all_trends[all_trends['Jumlah'] > 0]
    
    # Urutkan berdasarkan tanggal
    all_trends = all_trends.sort_values(by='Tanggal')

    # 5. Visualisasi dengan Plotly
    if not all_trends.empty:
        fig = px.line(
            all_trends, 
            x='Tanggal', 
            y='Jumlah',
            color='Dokumen', # Membedakan garis berdasarkan jenis dokumen
            markers=True, # Menambahkan titik di setiap data
            title=f'Trend Penerbitan Dokumen ({pilihan_waktu})'
        )

        # Mengatur tampilan agar rapi
        fig.update_layout(
            xaxis_title="Tanggal",
            yaxis_title="Total Dokumen Diterbitkan",
            hovermode="x unified", # Menampilkan tooltip gabungan saat kursor diarahkan ke satu titik X
            legend_title="Jenis Dokumen"
        )
        
        # Tampilkan grafik
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Tidak ada data tanggal yang valid untuk ditampilkan trend-nya.")

def trend_per_nama_barang(po_raw, bp_raw, sj_raw, file_consumable, file_nonconsumable):
    st.markdown("---")
    st.header("🔍 Tren Pergerakan Spesifik per Barang")
    
    # 1. Pastikan file master sudah di-upload di sidebar
    if file_consumable is None or file_nonconsumable is None:
        st.warning("⚠️ Mohon upload file Master Consumable dan Non-Consumable di sidebar terlebih dahulu.")
        return

    # 2. Reset pointer file uploader & Load data master
    try:
        file_consumable.seek(0)
        file_nonconsumable.seek(0)
        df_con = pd.read_excel(file_consumable)
        df_noncon = pd.read_excel(file_nonconsumable)
    except Exception as e:
        st.error(f"Gagal membaca file master: {e}")
        return
        
    col_master = 'NAMA_BARANG'
    
    # 3. UI: Pilihan Kategori & Periode menggunakan kolom sejajar
    col1, col2 = st.columns(2)
    with col1:
        kategori = st.radio("1. Pilih Kategori Master:", ["Consumable", "Non-Consumable"])
    
    with col2:
        pilihan_waktu = st.selectbox("2. Pilih Periode Tren:", ["Harian", "Mingguan", "Bulanan"], index=2)
        
    # Ambil list nama barang berdasarkan kategori yang dipilih
    if kategori == "Consumable" and col_master in df_con.columns:
        list_barang = df_con[col_master].dropna().astype(str).unique().tolist()
    elif kategori == "Non-Consumable" and col_master in df_noncon.columns:
        list_barang = df_noncon[col_master].dropna().astype(str).unique().tolist()
    else:
        st.error(f"Kolom '{col_master}' tidak ditemukan di file master yang dipilih.")
        return
            
    list_barang.sort()

    # 4. Input Pencarian Barang (Autocomplete Selectbox)
    barang_dicari = st.selectbox("3. Cari & Pilih Nama Barang:", options=["-- Pilih Barang --"] + list_barang)

    if barang_dicari == "-- Pilih Barang --":
        st.info("💡 Silakan ketik atau pilih nama barang di atas untuk melihat tren PO, BP, dan SJ-nya.")
        return

    # 5. Filter Data Transaksi berdasarkan barang yang dicari
    # (Penyesuaian nama kolom dinamis karena kadang penamaan di Excel berbeda)
    col_po = 'NAMA_BARANG' if 'NAMA_BARANG' in po_raw.columns else 'NAMABRG'
    col_bp = 'NAMABRG' if 'NAMABRG' in bp_raw.columns else 'NAMA_BARANG'
    
    po_filtered = po_raw[po_raw[col_po].astype(str) == barang_dicari] if col_po in po_raw.columns else pd.DataFrame()
    bp_filtered = bp_raw[bp_raw[col_bp].astype(str) == barang_dicari] if col_bp in bp_raw.columns else pd.DataFrame()
    
    # Penanganan khusus SJ (Jika SJ tidak punya nama barang, kita lacak lewat nomor PO-nya)
    sj_filtered = pd.DataFrame()
    if 'NAMABRG' in sj_raw.columns:
        sj_filtered = sj_raw[sj_raw['NAMABRG'].astype(str) == barang_dicari]
    elif 'NAMA_BARANG' in sj_raw.columns:
        sj_filtered = sj_raw[sj_raw['NAMA_BARANG'].astype(str) == barang_dicari]
    elif not po_filtered.empty:
        list_po = po_filtered['PO_NO'].unique().tolist()
        col_sj_po = 'NO_PO' if 'NO_PO' in sj_raw.columns else 'PO_NO'
        if col_sj_po in sj_raw.columns:
            sj_filtered = sj_raw[sj_raw[col_sj_po].isin(list_po)]

    # 6. Peringatan jika barang tidak pernah ditransaksikan
    if po_filtered.empty and bp_filtered.empty and sj_filtered.empty:
        st.warning(f"⚠️ Barang **'{barang_dicari}'** tidak ada datanya (Belum pernah diproses di PO, BP, maupun SJ).")
        return
        
    # 7. Fungsi Resampling & Kalkulasi Frekuensi
    freq_map = {"Harian": "D", "Mingguan": "W-MON", "Bulanan": "MS"}
    freq = freq_map[pilihan_waktu]

    def hitung_trend(df, date_col, doc_name):
        if df.empty or date_col not in df.columns:
            return pd.DataFrame(columns=['Tanggal', 'Frekuensi', 'Dokumen'])
        df_temp = df.dropna(subset=[date_col]).copy()
        df_temp.set_index(date_col, inplace=True)
        trend = df_temp.resample(freq).size().reset_index()
        trend.columns = ['Tanggal', 'Frekuensi']
        trend['Tanggal'] = trend['Tanggal'].dt.date
        trend['Dokumen'] = doc_name
        return trend

    po_trend = hitung_trend(po_filtered, 'PO_CREATED_ON', 'PO (Terbit)')
    bp_trend = hitung_trend(bp_filtered, 'BP_CREATED_ON', 'BP (Diterima)')
    sj_trend = hitung_trend(sj_filtered, 'SJ_CREATED_ON', 'SJ (Didistribusikan)')

    # Gabungkan semua data trend
    all_trends = pd.concat([po_trend, bp_trend, sj_trend], ignore_index=True)
    all_trends = all_trends[all_trends['Frekuensi'] > 0].sort_values(by='Tanggal')

    # 8. Visualisasi
    if not all_trends.empty:
        st.success(f"✅ Menampilkan riwayat transaksi untuk: **{barang_dicari}**")
        fig = px.line(
            all_trends, 
            x='Tanggal', 
            y='Frekuensi',
            color='Dokumen',
            markers=True,
            title=f'Tren Aktivitas Dokumen: {barang_dicari}',
            color_discrete_map={'PO (Terbit)': '#636EFA', 'BP (Diterima)': '#00CC96', 'SJ (Didistribusikan)': '#EF553B'}
        )
        fig.update_layout(
            xaxis_title="Tanggal",
            yaxis_title="Frekuensi Transaksi",
            hovermode="x unified"
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning(f"Data tanggal tidak valid untuk '{barang_dicari}'.")
#========================================================================
#========================================================================
#                           Main aplikasi
#========================================================================
#========================================================================

if uploaded_file is not None:

    # Proses file
    data = process_excel_file(uploaded_file)
    data2 = process_excel_file2(uploaded_file_2)
    if data is not None:
        # 1. Load Data Mentah
        po_raw = data['PO'].dropna()
        bp_raw = data['BP'].dropna() 
        sj_raw = data['SJ'].dropna()
        bsm_raw = data2['BSM'].dropna()
        po_raw['durasi_proses'] = po_raw['PO_APPROVED_ON'] - po_raw['PO_CREATED_ON']
        kolom_pilihan = ['PO_NO', 'PO_CREATED_ON', 'PO_APPROVED_ON', 'durasi_proses']
        # file_consumable = 'D:/Intern Jan 2025 - Jun 2025/Work_Directory/Barang_Consumable.xlsx'
        # file_nonconsumable = 'D:/Intern Jan 2025 - Jun 2025/Work_Directory/Barang_NonConsumable.xlsx'
        merged = pd.merge(po_raw, bp_raw, on="PO_NO", how="left")
                
        outstanding_bp = merged[
            (merged['PO_APPROVED_ON'].notnull()) & 
            (merged['BP_CREATED_ON'].isna())
        ].copy()

        tampilkan_kpi_dashboard(po_raw=po_raw, merged=merged, sj_raw=sj_raw)


        tab1, tab2, tab3, tab4, tab5, tab6= st.tabs(["📈 Analisis PO","📈 Analisis BP    ","📈 Analisis SJ","📈 Analisis BSM", "🔍 Trend", "📋 Detail Data"])
        with tab1:
            nama_kolom_barang = 'NAMA_BARANG'
            analisis_top_barang_by_kategori(
                po_raw=po_raw,
                file_consumable=uploaded_con,
                file_nonconsumable=uploaded_noncon,
                col_nama_barang=nama_kolom_barang
            )

            #========================================================================
            tampilkan_analisis_po(po_raw=po_raw, kolom_pilihan=kolom_pilihan)
            #========================================================================
            tracking_po_tanpa_bp(po_raw=po_raw, bp_raw=bp_raw)
            #========================================================================
            df_hasil_bersih, jumlah_po = analisis_durasi_po_ke_bp(merged=merged, bp_raw=bp_raw)
            #========================================================================
            
            #========================================================================
            deteksi_po_ekstrem(bp_raw=bp_raw)
            #========================================================================

        with tab2:
            #========================================================================
            tracking_bp_open(bp_raw=bp_raw)
            #========================================================================
            deteksi_po_multi_bp(bp_raw=bp_raw)
            #========================================================================
            analisis_perilaku_multi_bp(bp_raw=bp_raw)
            #========================================================================
            analisis_performa_supplier_bar(df_clean=df_hasil_bersih)

        with tab3:  
            #========================================================================
            sj_raw = analisis_durasi_approval_sj(sj_raw=sj_raw)
            #========================================================================
            deteksi_sj_telat(sj_raw=sj_raw)
            #========================================================================
            sj_raw = analisis_durasi_proses_sj(sj_raw=sj_raw)
            #========================================================================   

        with tab4:  
            top_10_barang_bsm_by_kategori(
                bsm_raw=bsm_raw, 
                file_consumable=uploaded_con, 
                file_nonconsumable=uploaded_noncon
            )
            #========================================================================   
            analisis_fulfillment_rate_bsm(bsm_raw=bsm_raw)
            # ========================================================================
            # Memanggil Spend Analysis (ABC)
            analisis_pengeluaran_abc(bsm_raw=bsm_raw)
            #========================================================================   
            analisis_dead_stock(bp_raw=bp_raw, sj_raw=sj_raw)


        with tab5:
            tampilkan_trend_dokumen(po_raw=po_raw, bp_raw=bp_raw, sj_raw=sj_raw, bsm_raw=bsm_raw)
            # ========================================================================
            trend_per_nama_barang(
                po_raw=po_raw, 
                bp_raw=bp_raw, 
                sj_raw=sj_raw, 
                file_consumable=uploaded_con,     
                file_nonconsumable=uploaded_noncon
            )
