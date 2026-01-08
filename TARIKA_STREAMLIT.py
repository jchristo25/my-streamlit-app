# import pandas as pd
# import matplotlib.pyplot as plt
# import matplotlib.dates as mdates
# import streamlit as st
# import plotly.graph_objects as go
# import plotly.express as px
# import io

# st.title("Dashboard Analisis")
# st.markdown("---")

# # ==========================================
# # 1. FILE UPLOADER
# # ==========================================
# st.sidebar.header("Upload Data")
# uploaded_file = st.sidebar.file_uploader("Upload File Excel (Format .xlsx)", type=["xlsx"])

# if uploaded_file is not None:
#     @st.cache_data
#     def load_data(file):
#         xls = pd.ExcelFile(file)
#         df_po = pd.read_excel(xls, 'PO', header=2)
#         df_bp = pd.read_excel(xls, 'BP', header=2)
#         df_sj = pd.read_excel(xls, 'SJ', header=3)
#         return df_po, df_bp, df_sj

    
#     df_po_raw, df_bp_raw, df_sj_raw = load_data(uploaded_file)
    
#     # Copy dataframe
#     df_po = df_po_raw.copy()
#     df_bp = df_bp_raw.copy()
#     df_sj = df_sj_raw.copy()

#     # ===========================================
#     # TIME SERIES PROCESSING (CLEANING)
#     # ===========================================
    
#     # 2. Data Cleaning & Date Conversion
#     # Lakukan konversi TGLAPP di sini agar aman
#     df_po['PO_CREATED_ON'] = pd.to_datetime(df_po['PO_CREATED_ON'], dayfirst=True, errors='coerce')
#     df_bp['BP_CREATED_ON'] = pd.to_datetime(df_bp['BP_CREATED_ON'], dayfirst=True, errors='coerce')
#     df_sj['SJ_CREATED_ON'] = pd.to_datetime(df_sj['SJ_CREATED_ON'], dayfirst=True, errors='coerce')
    
#     # [FIX] Konversi TGLAPP dilakukan di awal
#     # Pastikan nama kolom di Excel benar 'TGLAPP' (tanpa spasi)
#     if 'TGLAPP' in df_sj.columns:
#         df_sj['TGLAPP'] = pd.to_datetime(df_sj['TGLAPP'], dayfirst=True, errors='coerce')
#     else:
#         st.error("Kolom 'TGLAPP' tidak ditemukan di Sheet SJ. Cek nama kolom di Excel.")
#         df_sj['TGLAPP'] = pd.NaT # Isi kosong jika tidak ada agar tidak crash

#     # Filter valid dates untuk plotting harian
#     df_po_clean = df_po.dropna(subset=['PO_CREATED_ON'])
#     df_bp_clean = df_bp.dropna(subset=['BP_CREATED_ON'])
#     df_sj_clean_ts = df_sj.dropna(subset=['SJ_CREATED_ON']) # TS = Time Series

#     # 3. Aggregation (Daily)
#     po_daily = df_po_clean.groupby(df_po_clean['PO_CREATED_ON'].dt.to_period('D'))['PO_NO'].nunique().sort_index()
#     bp_daily = df_bp_clean.groupby(df_bp_clean['BP_CREATED_ON'].dt.to_period('D'))['BP_NO'].nunique().sort_index()
#     sj_daily = df_sj_clean_ts.groupby(df_sj_clean_ts['SJ_CREATED_ON'].dt.to_period('D'))['SJ_NO'].nunique().sort_index()

#     po_daily.index = po_daily.index.to_timestamp()
#     bp_daily.index = bp_daily.index.to_timestamp()
#     sj_daily.index = sj_daily.index.to_timestamp()

#     # ===========================================
#     # ANALISIS (LOGIC PROCESSING)
#     # ===========================================

#     # --- Analisis 1: Lead Time Supplier (PO ke BP) ---
#     first_receipt = df_bp.groupby('PO_NO')['BP_CREATED_ON'].min().reset_index()
#     po_perf = df_po[['PO_NO', 'NAMA_SUPPLIER', 'PO_APPROVED_ON','PO_CREATED_ON']].merge(first_receipt, on='PO_NO', how='inner')

#     po_perf['BP_CREATED_ON'] = pd.to_datetime(po_perf['BP_CREATED_ON'], errors='coerce')
#     po_perf['PO_APPROVED_ON'] = pd.to_datetime(po_perf['PO_APPROVED_ON'], errors='coerce')
#     po_perf['PO_CREATED_ON'] = pd.to_datetime(po_perf['PO_CREATED_ON'], errors='coerce')

#     po_perf['LEAD_TIME_DAYS'] = (po_perf['BP_CREATED_ON'] - po_perf['PO_APPROVED_ON']).dt.days
#     po_perf = po_perf[(po_perf['LEAD_TIME_DAYS'] >= 0) & (po_perf['LEAD_TIME_DAYS'] < 100)]

#     po_perf['APPROVAL_TIME_DAYS'] = (po_perf['PO_APPROVED_ON'] - po_perf['PO_CREATED_ON']).dt.days
    
#     # Agregasi performa per Supplier
#     supplier_stats = po_perf.groupby('NAMA_SUPPLIER').agg(
#         RATA_RATA_WAKTU=('LEAD_TIME_DAYS', 'mean'),
#         TOTAL_PO=('PO_NO', 'nunique')
#     ).sort_values(by='TOTAL_PO', ascending=False)

#     # --- Analisis 2: Kecepatan Gudang (BP ke SJ) & Approval SJ ---
#     avg_bp_date = df_bp.groupby('PO_NO')['BP_CREATED_ON'].mean().reset_index()
    
#     # [FIX] Sertakan TGLAPP saat grouping SJ
#     # Kita ambil 'max' untuk TGLAPP (asumsi tanggal approve terakhir) dan 'mean' untuk Created
#     avg_sj_date = df_sj.groupby('NO_PO').agg({
#         'SJ_CREATED_ON': 'mean',
#         'TGLAPP': 'max'  # Mengambil tanggal approval paling akhir jika ada double SJ
#     }).reset_index()

#     warehouse_flow = avg_bp_date.merge(avg_sj_date, left_on='PO_NO', right_on='NO_PO', how='inner')
    
#     # Hitung Waktu Proses Gudang
#     warehouse_flow['WAKTU_PROSES'] = (warehouse_flow['SJ_CREATED_ON'] - warehouse_flow['BP_CREATED_ON']).dt.days
#     warehouse_flow = warehouse_flow[warehouse_flow['WAKTU_PROSES'] >= 0]

#     # [FIX] Hitung Selisih SJ Approval (Sekarang TGLAPP sudah ada di warehouse_flow)
#     warehouse_flow['SJ_APPROVAL_TIME'] = (warehouse_flow['TGLAPP'] - warehouse_flow['SJ_CREATED_ON']).dt.total_seconds() / (24 * 3600)

#     # --- Analisis 3: Item Demand & Fulfillment ---
#     top_items = df_po.groupby('NAMA_BARANG')['JUMLAH'].sum().sort_values(ascending=False).head(10)

#     df_po['JUMLAH'] = pd.to_numeric(df_po['JUMLAH'], errors='coerce').fillna(0)
#     df_bp['JML_DITERIMA'] = pd.to_numeric(df_bp['JML_DITERIMA'], errors='coerce').fillna(0)
#     df_sj['JUMLAH'] = pd.to_numeric(df_sj['JUMLAH'], errors='coerce').fillna(0)

#     # Fulfillment Rate
#     po_target = df_po.groupby(['PO_NO', 'NAMA_SUPPLIER'])['JUMLAH'].sum().reset_index()
#     bp_actual = df_bp.groupby('PO_NO')['JML_DITERIMA'].sum().reset_index()

#     fulfillment = po_target.merge(bp_actual, on='PO_NO', how='left')
#     fulfillment['JML_DITERIMA'] = fulfillment['JML_DITERIMA'].fillna(0)
#     fulfillment['FILL_RATE'] = (fulfillment['JML_DITERIMA'] / fulfillment['JUMLAH']) * 100
#     fulfillment['FILL_RATE_CAPPED'] = fulfillment['FILL_RATE'].apply(lambda x: 100 if x > 100 else x)

#     supplier_fulfillment = fulfillment.groupby('NAMA_SUPPLIER').agg(
#         AVG_FILL_RATE=('FILL_RATE_CAPPED', 'mean'),
#         TOTAL_PO=('PO_NO', 'nunique')
#     ).query('TOTAL_PO >= 5').sort_values(by='AVG_FILL_RATE', ascending=True)

#     # Unfulfilled POs
#     unfulfilled = fulfillment[fulfillment['JML_DITERIMA'] == 0]
#     unfulfilled_count = unfulfilled.shape[0]
#     top_unfulfilled_suppliers = unfulfilled['NAMA_SUPPLIER'].value_counts()

#     # --- Analisis 4: Consumption Profiling ---
#     df_sj_clean_vessel = df_sj.dropna(subset=['NAMAKODEAKUN_BARANG'])
#     top_vessels_list = df_sj_clean_vessel['NAMA_KAPAL'].value_counts().head(5).index
#     df_sj_top = df_sj_clean_vessel[df_sj_clean_vessel['NAMA_KAPAL'].isin(top_vessels_list)].copy()
    
#     vessel_category_pivot = df_sj_top.groupby(['NAMAKODEAKUN_BARANG', 'NAMA_KAPAL'])['JUMLAH'].sum().unstack(fill_value=0)

#     # ===========================================
#     # TAMPILAN DASHBOARD
#     # ===========================================
    
#     st.subheader("Key Performance Indicators")

#     col1, col2, col3, col4 = st.columns(4)

#     with col1:
#         st.metric("Avg Approval Time (PO)", f"{po_perf['LEAD_TIME_DAYS'].mean():.2f} Days")
#         st.caption("PO Created -> PO Approved")

#     with col2:
#         st.metric("Avg Warehouse Process", f"{warehouse_flow['WAKTU_PROSES'].mean():.2f} Days")
#         st.caption("BP Created -> SJ Created")

#     with col3:
#         # Menampilkan metrik baru (SJ Approval)
#         avg_sj_app = warehouse_flow['SJ_APPROVAL_TIME'].mean()
#         # Handle jika hasil NaN (misal semua data kosong)
#         display_val = f"{avg_sj_app:.2f}" if pd.notnull(avg_sj_app) else "N/A"
#         st.metric("Avg SJ Approval Time", f"{display_val} Days")
#         st.caption("SJ Created -> SJ Approved")

#     with col4:
#         st.metric("Unfulfilled POs", f"{unfulfilled_count}")
#         st.caption("Black Hole (0% Delivery)")

#     tab1, tab2, tab3, tab4 = st.tabs(["📈 Tren Harian", "🚀 Analisis Consumption", "📋 Data Tabel", "🧠 Advanced Analytics"])

#     with tab1:
#         st.subheader("Analisis Tren Harian")
        
#         # Plotting Time Series
#         fig_ts, ax_ts = plt.subplots(2, 1, figsize=(14, 8))
        
#         ax_ts[0].plot(po_daily.index, po_daily.values, marker='o', linestyle='-', color='blue', label='PO Created')
#         ax_ts[0].set_title('Daily Procurement Activity', fontsize=14)
#         ax_ts[0].set_ylabel('Count')
#         ax_ts[0].legend()
#         ax_ts[0].grid(True, linestyle='--', alpha=0.5)

#         ax_ts[1].plot(bp_daily.index, bp_daily.values, color='green', label='BP (Inbound)', alpha=0.7)
#         ax_ts[1].plot(sj_daily.index, sj_daily.values, color='red', label='SJ (Outbound)', alpha=0.7)
#         ax_ts[1].set_title('Logistics Flow', fontsize=14)
#         ax_ts[1].set_ylabel('Count')
#         ax_ts[1].legend()
#         ax_ts[1].grid(True, linestyle='--', alpha=0.5)

#         plt.tight_layout()
#         st.pyplot(fig_ts)

#         st.subheader("Shipping Activity by Day of Week")
        
#         # Combo Chart Plotly
#         days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
#         df_sj_temp = df_sj.copy()
#         df_sj_temp['DayOfWeek'] = pd.Categorical(df_sj_temp['SJ_CREATED_ON'].dt.day_name(), categories=days_order, ordered=True)
        
#         sj_dow = df_sj_temp.groupby('DayOfWeek')['SJ_NO'].nunique().reindex(days_order).fillna(0).reset_index()
#         sj_dow['MA'] = sj_dow['SJ_NO'].rolling(window=2, min_periods=1).mean()

#         fig_combo = go.Figure()
#         fig_combo.add_trace(go.Bar(x=sj_dow['DayOfWeek'], y=sj_dow['SJ_NO'], name='Total SJ', marker_color='#FF4B4B', opacity=0.7))
#         fig_combo.add_trace(go.Scatter(x=sj_dow['DayOfWeek'], y=sj_dow['MA'], name='Moving Avg', mode='lines+markers', line=dict(color='blue')))
#         fig_combo.update_layout(title='Shipping Volume Trend', xaxis_title='Day', yaxis_title='Count', hovermode="x unified")
        
#         st.plotly_chart(fig_combo, use_container_width=True)

#     with tab2:
#         st.subheader("Supplier Fulfillment & Consumption")
        
#         # --- PREPARASI DATA UMUM ---
#         bottom_suppliers = supplier_fulfillment.head(10)
#         unfulfilled_counts = unfulfilled['NAMA_SUPPLIER'].value_counts().head(5)

#         # --- VISUALISASI ---
#         fig_di, axes_di = plt.subplots(2, 1, figsize=(12, 12)) # Cukup 2 baris untuk Matplotlib

#         # Plot 1: Supplier Fulfillment (Matplotlib)
#         bars = axes_di[0].barh(bottom_suppliers.index, bottom_suppliers['AVG_FILL_RATE'], color='salmon')
#         axes_di[0].set_title('Top 10 Lowest Fulfillment Suppliers', fontsize=14)
#         axes_di[0].bar_label(bars, fmt='%.1f%%')

#         # Plot 2: Unfulfilled POs (Matplotlib)
#         axes_di[1].bar(unfulfilled_counts.index, unfulfilled_counts.values, color='#d62728', alpha=0.8)
#         axes_di[1].set_title('Top Suppliers with Unfulfilled POs', fontsize=14)
#         axes_di[1].tick_params(axis='x', rotation=90)

#         plt.tight_layout()
#         st.pyplot(fig_di)

#         st.subheader("Top 5 Kapal Paling Sering Disupply")
#         top5_kapal = (
#             df_sj['NAMA_KAPAL']
#             .value_counts()
#             .head(5)
#             .reset_index()
#         )

#         top5_kapal.columns = ['Nama Kapal', 'Jumlah Supply (SJ)']

#         st.dataframe(top5_kapal)
#         # -------------------------------------------------------
#         # PLOT 3: Top Item consumption per Vessel (Bar Chart)
#         # -------------------------------------------------------
#         st.subheader("Top 5 Barang Paling Sering Dipesan per Kapal")
        
#         # 1. Siapkan Data (Group by Nama Asli)
#         # Ganti 'NAMAKODEAKUN_BARANG' dengan nama kolom yang sesuai di Excel Anda jika berbeda
#         col_item_name = 'NAMAKODEAKUN_BARANG' 
        
#         df_raw_item_data = df_sj_top.groupby(['NAMA_KAPAL', col_item_name])['JUMLAH'].sum().reset_index()

#         # 2. FILTER: Ambil Top 5 Item per Kapal
#         # Kita sort dulu berdasarkan Kapal (A-Z) dan Jumlah (Besar-Kecil)
#         df_raw_item_data = df_raw_item_data.sort_values(['NAMA_KAPAL', 'JUMLAH'], ascending=[True, False])
        
#         # Group by Kapal lalu ambil 5 baris teratas masing-masing grup
#         df_top5_per_vessel = df_raw_item_data.groupby('NAMA_KAPAL').head(5)

#         # 3. Buat Plot dengan Plotly Express (Bar Chart)
#         # Menggunakan Bar Chart Horizontal agar nama barang yang panjang bisa terbaca
#         fig_bar_multi = px.bar(
#             # margin='1'
#             df_top5_per_vessel, 
#             x='JUMLAH', 
#             y=col_item_name, 
#             color='JUMLAH',               # Warna gradasi berdasarkan jumlah
#             facet_col='NAMA_KAPAL',       # Memecah chart berdasarkan Nama Kapal
#             facet_col_wrap=1,             # Maksimal 2 chart per baris
#             orientation='h',              # Horizontal Bar
#             title='Top 5 Item Consumption per Vessel',
#             color_continuous_scale='Bluered'
#         )

#         # 4. Merapikan Layout (Sangat Penting untuk Facet Plot)
#         # matches=None: Agar sumbu Y (Nama Barang) tidak dipaksa sama antar kapal.
#         # showticklabels=True: Agar nama barang tetap muncul di setiap subplot.
#         fig_bar_multi.update_yaxes(matches=None, showticklabels=True)
#         fig_bar_multi.update_xaxes(matches=None)
        
#         # Menghilangkan tulisan "NAMA_KAPAL=" di judul subplot
#         fig_bar_multi.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
        
#         # Atur tinggi grafik (makin banyak kapal, mungkin perlu makin tinggi)
#         fig_bar_multi.update_layout(height=1000, margin=dict(t=70, l=30))

#         st.plotly_chart(fig_bar_multi, use_container_width=True)

#     with tab3:
#         st.markdown("#### Top 5 Supplier Teraktif")
#         st.dataframe(supplier_stats.head(5))
#         st.markdown("#### Top 5 Most Ordered Items")
#         st.dataframe(top_items.head(5))
#         st.markdown("#### Top Supplier Unfulfilled")
#         st.dataframe(top_unfulfilled_suppliers)

#     # --- TAB 4: ADVANCED ANALYTICS ---
#     with tab4:     
#         st.subheader("Potensi Dead Stock (Barang Masuk > Barang Keluar)")
        
#         # Hitung Total Masuk (BP)
#         inbound = df_bp.groupby('PO_NO')['JML_DITERIMA'].sum().reset_index()
#         # Kita butuh link ke nama barang. Ambil dari PO.
#         po_mapping = df_po[['PO_NO', 'NAMA_BARANG']].drop_duplicates()
#         inbound_merged = inbound.merge(po_mapping, on='PO_NO', how='left')
        
#         # Agregat per Barang
#         total_in = inbound_merged.groupby('NAMA_BARANG')['JML_DITERIMA'].sum().reset_index()
#         total_in.columns = ['NAMA_BARANG', 'TOTAL_MASUK']
        
#         # Hitung Total Keluar (SJ)
#         # Gunakan kolom barang yang sesuai di SJ
#         col_sj_barang = 'NAMABRG' if 'NAMABRG' in df_sj.columns else 'NAMAKODEAKUN_BARANG'
#         total_out = df_sj.groupby(col_sj_barang)['JUMLAH'].sum().reset_index()
#         total_out.columns = ['NAMA_BARANG', 'TOTAL_KELUAR']
        
#         # Gabungkan
#         inventory_flow = total_in.merge(total_out, on='NAMA_BARANG', how='left').fillna(0)
#         inventory_flow['SELISIH_STOK'] = inventory_flow['TOTAL_MASUK'] - inventory_flow['TOTAL_KELUAR']
        
#         # Filter: Hanya tampilkan yang stoknya menumpuk > 10 unit (bisa disesuaikan)
#         dead_stock = inventory_flow[inventory_flow['SELISIH_STOK'] > 0].sort_values(by='SELISIH_STOK', ascending=False).head(10)
        
#         # Bar Chart
#         fig_stock = px.bar(
#             dead_stock,
#             x='SELISIH_STOK',
#             y='NAMA_BARANG',
#             orientation='h',
#             title='Top 10 Barang dengan Penumpukan Stok Tertinggi (Gap Masuk vs Keluar)',
#             color='SELISIH_STOK',
#             color_continuous_scale='OrRd',
#             text_auto=True
#         )
#         fig_stock.update_layout(height=500)
#         st.plotly_chart(fig_stock, use_container_width=True)
        
#         st.caption("Catatan: Analisis ini mengasumsikan satuan unit (UoM) konsisten antara PO dan SJ. Jika beda satuan, data mungkin bias.")

# else:
#     st.info("Silakan upload file Excel pada sidebar.")


#=====================================================================================================================================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import warnings
import matplotlib.pyplot as plt
warnings.filterwarnings('ignore')

# Konfigurasi halaman
st.set_page_config(
    page_title="Purchase Order Analytics",
    page_icon="📊",
    layout="wide"
)

# Judul aplikasi
st.title("📊 Purchase Order Analytics Dashboard")
st.markdown("---")

# Sidebar untuk upload file
st.sidebar.header("📁 Upload Data")
uploaded_file = st.sidebar.file_uploader(
    "Upload Excel file dengan 3 sheet (PO, BP, SJ)",
    type=['xlsx', 'xls']
)

# Sidebar untuk filter waktu
st.sidebar.header("⏰ Filter Waktu")
time_filter = st.sidebar.selectbox(
    "Pilih Periode Analisis",
    ["Semua Data", "30 Hari Terakhir", "90 Hari Terakhir", "Tahun Ini", "Kustom"]
)

# Inisialisasi session state
if 'df_po' not in st.session_state:
    st.session_state.df_po = None
if 'df_bp' not in st.session_state:
    st.session_state.df_bp = None
if 'df_sj' not in st.session_state:
    st.session_state.df_sj = None

def load_data(file):
    """Memuat data dari file Excel"""
    try:
        df_po = pd.read_excel(file, 'PO', header=2)
        df_bp = pd.read_excel(file, 'BP', header=2)
        df_sj = pd.read_excel(file, 'SJ', header=3)
        
        # Konversi tanggal
        date_columns_po = ['PO_CREATED_ON', 'PO_APPROVED_ON', 'DELIVERY_TIME']
        date_columns_bp = ['BP_CREATED_ON']
        date_columns_sj = ['SJ_CREATED_ON', 'TGLAPP', 'SJ_CLOSED_ON', 'TGLINPG']
        
        for col in date_columns_po:
            if col in df_po.columns:
                df_po[col] = pd.to_datetime(df_po[col], errors='coerce')
        
        for col in date_columns_bp:
            if col in df_bp.columns:
                df_bp[col] = pd.to_datetime(df_bp[col], errors='coerce')
        
        for col in date_columns_sj:
            if col in df_sj.columns:
                df_sj[col] = pd.to_datetime(df_sj[col], errors='coerce')
        
        return df_po, df_bp, df_sj
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None, None, None

def apply_time_filter(df, date_column):
    """Menerapkan filter waktu berdasarkan pilihan"""
    if time_filter == "Semua Data":
        return df
    elif time_filter == "30 Hari Terakhir":
        cutoff_date = datetime.now() - timedelta(days=30)
        return df[df[date_column] >= cutoff_date]
    elif time_filter == "90 Hari Terakhir":
        cutoff_date = datetime.now() - timedelta(days=90)
        return df[df[date_column] >= cutoff_date]
    elif time_filter == "Tahun Ini":
        current_year = datetime.now().year
        return df[df[date_column].dt.year == current_year]
    else:
        return df

def analyze_po_insights(df_po):
    """Analisis insight dari data PO"""
    insights = []
    
    # 1. Volume PO per waktu
    df_po['PO_CREATED_ON_DATE'] = df_po['PO_CREATED_ON'].dt.date
    po_per_day = df_po.groupby('PO_CREATED_ON_DATE').size()
    
    if len(po_per_day) > 1:
        trend = "Naik" if po_per_day.iloc[-1] > po_per_day.iloc[0] else "Turun"
        avg_po_per_day = po_per_day.mean()
        insights.append(f"📈 **Tren PO**: {trend}, rata-rata {avg_po_per_day:.1f} PO/hari")
    
    # 2. Top suppliers
    top_suppliers = df_po['NAMA_SUPPLIER'].value_counts().head(5)
    if len(top_suppliers) > 0:
        top_supplier = top_suppliers.index[0]
        insights.append(f"🏆 **Supplier Teratas**: {top_supplier} ({top_suppliers.iloc[0]} PO)")
    
    # 3. Status distribusi
    status_dist = df_po['STATUS'].value_counts()
    if len(status_dist) > 0:
        main_status = status_dist.index[0]
        insights.append(f"📊 **Status Dominan**: {main_status} ({status_dist.iloc[0]} PO, {status_dist.iloc[0]/len(df_po)*100:.1f}%)")
    
    # 4. Purchase type analysis
    if 'PURCHASE_TYPE' in df_po.columns:
        purchase_type_dist = df_po['PURCHASE_TYPE'].value_counts()
        if len(purchase_type_dist) > 0:
            insights.append(f"🛒 **Tipe Pembelian Terbanyak**: {purchase_type_dist.index[0]}")
    
    # 5. Approval time analysis
    if 'PO_APPROVED_ON' in df_po.columns:
        df_po['APPROVAL_TIME'] = (df_po['PO_APPROVED_ON'] - df_po['PO_CREATED_ON']).dt.days
        avg_approval_time = df_po['APPROVAL_TIME'].mean()
        if not pd.isna(avg_approval_time):
            insights.append(f"⏱️ **Rata-rata Waktu Approve**: {avg_approval_time:.1f} hari")
    
    return insights

def analyze_bp_insights(df_bp, df_po):
    """Analisis insight dari data BP"""
    insights = []
    
    # 1. BP vs PO comparison
    if df_po is not None:
        total_po = len(df_po)
        total_bp = len(df_bp)
        bp_per_po = total_bp / total_po if total_po > 0 else 0
        insights.append(f"📋 **Rasio BP/PO**: {bp_per_po:.2f} ({total_bp} BP dari {total_po} PO)")
    
    # 2. BP creation trend
    df_bp['BP_CREATED_ON_DATE'] = df_bp['BP_CREATED_ON'].dt.date
    bp_per_day = df_bp.groupby('BP_CREATED_ON_DATE').size()
    
    if len(bp_per_day) > 1:
        avg_bp_per_day = bp_per_day.mean()
        insights.append(f"📅 **Rata-rata BP/hari**: {avg_bp_per_day:.1f}")
    
    # 3. Top suppliers in BP
    top_suppliers_bp = df_bp['NAMA_SUPPLIER'].value_counts().head(3)
    if len(top_suppliers_bp) > 0:
        insights.append(f"👥 **Supplier BP Teratas**: {', '.join(top_suppliers_bp.index[:3])}")
    
    return insights

def analyze_sj_insights(df_sj):
    """Analisis insight dari data SJ"""
    insights = []
    
    # 1. Delivery performance
    if 'STATUS' in df_sj.columns:
        status_counts = df_sj['STATUS'].value_counts()
        if len(status_counts) > 0:
            delivered_pct = status_counts.get('DELIVERED', 0) / len(df_sj) * 100
            insights.append(f"🚚 **Rate Pengiriman**: {delivered_pct:.1f}% terkirim")
    
    # 2. SJ creation trend
    df_sj['SJ_CREATED_ON_DATE'] = df_sj['SJ_CREATED_ON'].dt.date
    sj_per_day = df_sj.groupby('SJ_CREATED_ON_DATE').size()
    
    if len(sj_per_day) > 0:
        avg_sj_per_day = sj_per_day.mean()
        insights.append(f"📦 **Rata-rata SJ/hari**: {avg_sj_per_day:.1f}")
    
    # 3. Top vessel/alat berat
    if 'VESSELID' in df_sj.columns:
        top_vessel = df_sj['VESSELID'].value_counts().head(3)
        if len(top_vessel) > 0:
            insights.append(f"🚢 **Vessel Terbanyak**: {top_vessel.index[0]} ({top_vessel.iloc[0]} SJ)")
    
    return insights

def create_visualizations(df_po, df_bp, df_sj):
    """Membuat visualisasi data"""
    
    # Tab untuk visualisasi
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "📈 PO Analysis", "📋 BP Analysis", "🚚 SJ Analysis"])
    
    with tab1:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            total_po = len(df_po)
            st.metric("Total PO", total_po)
            
            # PO status distribution
            if 'STATUS' in df_po.columns:
                status_counts = df_po['STATUS'].value_counts()
                fig = px.pie(
                    values=status_counts.values,
                    names=status_counts.index,
                    title="Distribusi Status PO"
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            total_bp = len(df_bp)
            st.metric("Total BP", total_bp)
            
            # BP trend
            if 'BP_CREATED_ON' in df_bp.columns:
                bp_daily = df_bp.groupby(df_bp['BP_CREATED_ON'].dt.date).size()
                fig = px.line(
                    x=bp_daily.index,
                    y=bp_daily.values,
                    title="Trend BP per Hari",
                    labels={'x': 'Tanggal', 'y': 'Jumlah BP'}
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col3:
            total_sj = len(df_sj)
            st.metric("Total SJ", total_sj)
            
            # SJ status
            if 'STATUS' in df_sj.columns:
                sj_status = df_sj['STATUS'].value_counts()
                fig = px.bar(
                    x=sj_status.index,
                    y=sj_status.values,
                    title="Status SJ",
                    labels={'x': 'Status', 'y': 'Jumlah'}
                )
                st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        col1, col2 = st.columns(2)
        
        with col1:
            # PO over time
            if 'PO_CREATED_ON' in df_po.columns:
                po_daily = df_po.groupby(df_po['PO_CREATED_ON'].dt.date).size()
                fig = px.line(
                    x=po_daily.index,
                    y=po_daily.values,
                    title="PO per Hari",
                    labels={'x': 'Tanggal', 'y': 'Jumlah PO'}
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Top suppliers
            if 'NAMA_SUPPLIER' in df_po.columns:
                top_suppliers = df_po['NAMA_SUPPLIER'].value_counts().head(10)
                fig = px.bar(
                    x=top_suppliers.values,
                    y=top_suppliers.index,
                    orientation='h',
                    title="Top 10 Supplier",
                    labels={'x': 'Jumlah PO', 'y': 'Supplier'}
                )
                st.plotly_chart(fig, use_container_width=True)
        
        col3, col4 = st.columns(2)
        
        with col3:
            # Purchase type distribution
            if 'PURCHASE_TYPE' in df_po.columns:
                purchase_dist = df_po['PURCHASE_TYPE'].value_counts()
                fig = px.pie(
                    values=purchase_dist.values,
                    names=purchase_dist.index,
                    title="Distribusi Tipe Pembelian"
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col4:
            # Category analysis
            if 'CATEGORY' in df_po.columns:
                category_dist = df_po['CATEGORY'].value_counts().head(10)
                fig = px.bar(
                    x=category_dist.values,
                    y=category_dist.index,
                    orientation='h',
                    title="Top 10 Kategori",
                    labels={'x': 'Jumlah', 'y': 'Kategori'}
                )
                st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        col1, col2 = st.columns(2)
        
        with col1:
            # BP vs PO comparison
            if 'PO_NO' in df_bp.columns:
                bp_per_po = df_bp['PO_NO'].value_counts()
                fig = px.histogram(
                    x=bp_per_po.values,
                    title="Distribusi BP per PO",
                    labels={'x': 'Jumlah BP per PO', 'y': 'Frekuensi'}
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Supplier performance in BP
            if 'NAMA_SUPPLIER' in df_bp.columns:
                supplier_bp = df_bp['NAMA_SUPPLIER'].value_counts().head(10)
                fig = px.bar(
                    x=supplier_bp.values,
                    y=supplier_bp.index,
                    orientation='h',
                    title="Top 10 Supplier (BP)",
                    labels={'x': 'Jumlah BP', 'y': 'Supplier'}
                )
                st.plotly_chart(fig, use_container_width=True)
    
    with tab4:
        col1, col2 = st.columns(2)
        
        with col1:
            # SJ delivery trend
            if 'SJ_CREATED_ON' in df_sj.columns:
                sj_daily = df_sj.groupby(df_sj['SJ_CREATED_ON'].dt.date).size()
                fig = px.line(
                    x=sj_daily.index,
                    y=sj_daily.values,
                    title="SJ per Hari",
                    labels={'x': 'Tanggal', 'y': 'Jumlah SJ'}
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Vessel distribution
            if 'VESSELID' in df_sj.columns:
                vessel_dist = df_sj['VESSELID'].value_counts().head(10)
                fig = px.bar(
                    x=vessel_dist.values,
                    y=vessel_dist.index,
                    orientation='h',
                    title="Top 10 Vessel/Alat Berat",
                    labels={'x': 'Jumlah SJ', 'y': 'Vessel ID'}
                )
                st.plotly_chart(fig, use_container_width=True)

def show_data_summary(df_po, df_bp, df_sj):
    """Menampilkan ringkasan data"""
    st.subheader("📋 Data Summary")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.write("**PO Data**")
        st.write(f"Jumlah baris: {len(df_po)}")
        st.write(f"Jumlah kolom: {len(df_po.columns)}")
        st.write(f"Periode: {df_po['PO_CREATED_ON'].min().date()} hingga {df_po['PO_CREATED_ON'].max().date()}")
    
    with col2:
        st.write("**BP Data**")
        st.write(f"Jumlah baris: {len(df_bp)}")
        st.write(f"Jumlah kolom: {len(df_bp.columns)}")
        if 'BP_CREATED_ON' in df_bp.columns:
            st.write(f"Periode: {df_bp['BP_CREATED_ON'].min().date()} hingga {df_bp['BP_CREATED_ON'].max().date()}")
    
    with col3:
        st.write("**SJ Data**")
        st.write(f"Jumlah baris: {len(df_sj)}")
        st.write(f"Jumlah kolom: {len(df_sj.columns)}")
        if 'SJ_CREATED_ON' in df_sj.columns:
            st.write(f"Periode: {df_sj['SJ_CREATED_ON'].min().date()} hingga {df_sj['SJ_CREATED_ON'].max().date()}")

def main():
    """Fungsi utama aplikasi"""
    
    if uploaded_file is not None:
        # Load data
        with st.spinner("Memuat data..."):
            df_po, df_bp, df_sj = load_data(uploaded_file)
            
            if df_po is not None and df_bp is not None and df_sj is not None:
                # Simpan ke session state
                st.session_state.df_po = df_po
                st.session_state.df_bp = df_bp
                st.session_state.df_sj = df_sj
                
                # Terapkan filter waktu
                df_po_filtered = apply_time_filter(df_po, 'PO_CREATED_ON')
                df_bp_filtered = apply_time_filter(df_bp, 'BP_CREATED_ON')
                df_sj_filtered = apply_time_filter(df_sj, 'SJ_CREATED_ON')
                
                # Tampilkan summary
                show_data_summary(df_po_filtered, df_bp_filtered, df_sj_filtered)
                
                st.markdown("---")
                
                # Tampilkan insights
                st.subheader("🔍 Key Insights")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.write("**PO Insights**")
                    po_insights = analyze_po_insights(df_po_filtered)
                    for insight in po_insights:
                        st.info(insight)
                
                with col2:
                    st.write("**BP Insights**")
                    bp_insights = analyze_bp_insights(df_bp_filtered, df_po_filtered)
                    for insight in bp_insights:
                        st.success(insight)
                
                with col3:
                    st.write("**SJ Insights**")
                    sj_insights = analyze_sj_insights(df_sj_filtered)
                    for insight in sj_insights:
                        st.warning(insight)
                
                st.markdown("---")
                
                # Tampilkan visualisasi
                st.subheader("📊 Data Visualization")
                create_visualizations(df_po_filtered, df_bp_filtered, df_sj_filtered)
                
                # Data exploration section
                st.markdown("---")
                st.subheader("🔍 Data Exploration")
                
                explore_tab1, explore_tab2, explore_tab3 = st.tabs(["PO Data", "BP Data", "SJ Data"])
                
                with explore_tab1:
                    st.dataframe(df_po_filtered.head(100), use_container_width=True)
                
                with explore_tab2:
                    st.dataframe(df_bp_filtered.head(100), use_container_width=True)
                
                with explore_tab3:
                    st.dataframe(df_sj_filtered.head(100), use_container_width=True)
    
    else:
        # Tampilkan instruksi jika belum upload file
        st.info("👈 Silakan upload file Excel melalui sidebar di sebelah kiri")
        
        # Contoh struktur data yang diharapkan
        st.subheader("📋 Struktur Data yang Diperlukan")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.write("**Sheet PO (Purchase Order)**")
            st.code("""
PO_NO
PO_CREATED_ON
DOCKING
KODE_KATEGORI
KODE_BARANG
NAMA_BARANG
SATUAN
KODEAKUN_BARANG
NAMAKODEAKUN_BARANG
KODEAKUN_PO
NAMAKODEAKUN_PO
PURCHASE_TYPE
JUMLAH
JML_DISETUJUI
JML_DITERIMA
DELIVERY_TIME
VESSELID
VESSELNAME
KODEALATBERAT
KODE2
NAMAALATBERAT
HVE_OR_VESSEL
REF
REF2
CATEGORY
STATUS
USER_ID
NAMA_SUPPLIER
PAYMENT_TERM
DELIVERY_TERM
PO_APPROVED_BY
PO_APPROVED_ON
KODE_LOKASI
PO_KETERANGAN
PO_REMARK
PO_REMARK2
BPS_NO
BPS_KETERANGAN
            """)
        
        with col2:
            st.write("**Sheet BP (Bukti Penerimaan)**")
            st.code("""
BP_NO
KODE_SUPPLIER
NAMA_SUPPLIER
NAMABRG
SATUAN
KODEAKUN_BARANG
NAMAKODEAKUN_BARANG
BP_KETERANGAN
STATUS
BP_CREATED_ON
PO_NO
KODE_BARANG
DETAIL_KETERANGAN
TUJUAN
VESSELCODE
JML_DITERIMA
            """)
        
        with col3:
            st.write("**Sheet SJ (Surat Jalan)**")
            st.code("""
SJ_NO
KEPADA
VESSELID
NAMA_KAPAL
NO_PO
NAMA_KATEGORI
NAMABRG
SATUAN
KODEAKUN_BARANG
NAMAKODEAKUN_BARANG
JUMLAH
JMLDISETUJUI
STATUS
SJ_CREATED_ON
CREATED_BY
KODEBARANG
KETERANGAN
JML_DITERIMA
RELEASE
BPK_NO
USERIDAPP
TGLAPP
STATUSAPP
SJ_CLOSED_BY
NOSBT
SJ_CLOSED_ON
LOKASIBUAT
TGLINPG
            """)

if __name__ == "__main__":
    main()
