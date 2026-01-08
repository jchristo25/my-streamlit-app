import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import streamlit as st
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
# st.set_page_config(page_title="Dashboard Logistik & Procurement", layout="wide")

st.title("Dashboard Analisis Logistik & Financial (2023-2025)")
st.markdown("---")

# ==========================================
# 1. FILE UPLOADER (DUAL FILE)
# ==========================================
st.sidebar.header("Upload Data")

# File 1: Operasional (TARIKA)
file_tarika = st.sidebar.file_uploader("1. Upload File TARIKA (PO/BP/SJ)", type=["xlsx"], key="tarika")

# File 2: Finansial/Historis (BSM)
file_bsm = st.sidebar.file_uploader("2. Upload File BSM (2023-2025)", type=["xlsx"], key="bsm")

# Fungsi Load Data TARIKA
@st.cache_data
def load_tarika(file):
    xls = pd.ExcelFile(file)
    df_po = pd.read_excel(xls, 'PO', header=2)
    df_bp = pd.read_excel(xls, 'BP', header=2)
    df_sj = pd.read_excel(xls, 'SJ', header=3)
    return df_po, df_bp, df_sj

# Fungsi Load Data BSM
@st.cache_data
def load_bsm(file):
    try:
        df = pd.read_excel(file)
    except:
        df = pd.read_csv(file) # Fallback jika user upload CSV
    return df

# ==========================================
# MAIN LOGIC
# ==========================================
if file_tarika is not None:
    # --- LOAD TARIKA ---
    df_po_raw, df_bp_raw, df_sj_raw = load_tarika(file_tarika)
    df_po = df_po_raw.copy()
    df_bp = df_bp_raw.copy()
    df_sj = df_sj_raw.copy()

    # --- CLEANING TARIKA ---
    # Date Conversion (Global)
    df_po['PO_CREATED_ON'] = pd.to_datetime(df_po['PO_CREATED_ON'], dayfirst=True, errors='coerce')
    df_bp['BP_CREATED_ON'] = pd.to_datetime(df_bp['BP_CREATED_ON'], dayfirst=True, errors='coerce')
    df_sj['SJ_CREATED_ON'] = pd.to_datetime(df_sj['SJ_CREATED_ON'], dayfirst=True, errors='coerce')
    
    if 'TGLAPP' in df_sj.columns:
        df_sj['TGLAPP'] = pd.to_datetime(df_sj['TGLAPP'], dayfirst=True, errors='coerce')
    else:
        df_sj['TGLAPP'] = pd.NaT

    # Numeric Conversion
    df_po['JUMLAH'] = pd.to_numeric(df_po['JUMLAH'], errors='coerce').fillna(0)
    df_bp['JML_DITERIMA'] = pd.to_numeric(df_bp['JML_DITERIMA'], errors='coerce').fillna(0)
    df_sj['JUMLAH'] = pd.to_numeric(df_sj['JUMLAH'], errors='coerce').fillna(0)

    # ------------------------------------------
    # ANALISIS SELISIH (MISMATCH)
    # ------------------------------------------
    df_po['SELISIH_JML'] = df_po['JML_DISETUJUI'] - df_po['JML_DITERIMA']
    df_po['Status'] = np.where(df_po['SELISIH_JML'] == 0, 'Sesuai', 'Terdapat Selisih')
    
    # Filter hanya yang mismatch
    df_mismatch = df_po[df_po['SELISIH_JML'] != 0].copy()

    if df_mismatch.empty:
        st.success("✅ Luar biasa! Semua data pesanan sesuai dengan penerimaan (Tidak ada selisih).")
    else:
        # PENTING: Perbaikan Format Tanggal untuk Tampilan Tabel
        try:
            # Pastikan format tanggal benar sebelum diubah jadi string
            df_mismatch['PO_CREATED_ON'] = pd.to_datetime(df_mismatch['PO_CREATED_ON'], dayfirst=True)
            df_mismatch['PO_CREATED_ON'] = df_mismatch['PO_CREATED_ON'].dt.strftime('%d/%m/%Y')
        except Exception as e:
            pass # Lanjut saja jika ada error format

        st.warning(f"⚠️ Ditemukan {len(df_mismatch)} transaksi yang memiliki selisih.")

        # 1. Menampilkan Tabel Data Selisih
        kolom_pilihan = [
            'PO_NO', 'NAMA_BARANG', 'NAMA_SUPPLIER', 'SATUAN',
            'PO_CREATED_ON', 'PO_APPROVED_ON',
            'JML_DISETUJUI', 'JML_DITERIMA', 'SELISIH_JML'
        ]
        
        # Cek ketersediaan kolom sebelum menampilkan
        cols_to_show = [c for c in kolom_pilihan if c in df_mismatch.columns]
        
        st.dataframe(
            df_mismatch[cols_to_show].sort_values(by='SELISIH_JML', key=abs, ascending=False),
            use_container_width=True,
            hide_index=True
        )
        st.caption("*Tabel di atas hanya menunjukkan data yang JML_DISETUJUI ≠ JML_DITERIMA.*")
        
        st.divider()

        # 2. Visualisasi Grafik Selisih (LOOPING PER SUPPLIER)
        st.subheader("📊 Grafik Detail per Supplier")

        # Ambil daftar unik supplier yang bermasalah
        list_supplier = df_mismatch['NAMA_SUPPLIER'].unique()

        # Loop: Buat satu grafik untuk setiap supplier
        for supplier in list_supplier:
            
            # Filter data hanya untuk supplier ini
            df_per_supplier = df_mismatch[df_mismatch['NAMA_SUPPLIER'] == supplier]
            
            with st.container():
                st.markdown(f"### 🏢 Supplier: {supplier}")
                
                # Membuat Bar Chart Khusus Supplier Ini
                fig = px.bar(
                    df_per_supplier,
                    x='NAMA_BARANG',       # Sumbu X: Nama Barang
                    y='SELISIH_JML',       # Sumbu Y: Jumlah Selisih
                    color='PO_NO',         # Warna beda tiap PO
                    title=f"Analisis Selisih - {supplier}",
                    labels={
                        'SELISIH_JML': 'Selisih (Qty)', 
                        'NAMA_BARANG': 'Nama Barang', 
                        'PO_NO': 'No PO'
                    },
                    text_auto=True,        # Menampilkan angka di batang
                    height=400
                )
                
                # Garis 0 agar terlihat jelas positif/negatif
                fig.add_hline(y=0, line_dash="dot", line_color="black")
                
                st.plotly_chart(fig, use_container_width=True)
                st.markdown("---") # Garis pemisah antar supplier

    # ------------------------------------------
    # LANJUTAN LOGIKA ANALISIS DASHBOARD
    # ------------------------------------------

    # 1. Lead Time Analysis
    first_receipt = df_bp.groupby('PO_NO')['BP_CREATED_ON'].min().reset_index()
    po_perf = df_po[['PO_NO', 'NAMA_SUPPLIER', 'PO_APPROVED_ON','PO_CREATED_ON']].merge(first_receipt, on='PO_NO', how='inner')
    po_perf['BP_CREATED_ON'] = pd.to_datetime(po_perf['BP_CREATED_ON'], errors='coerce')
    po_perf['PO_APPROVED_ON'] = pd.to_datetime(po_perf['PO_APPROVED_ON'], errors='coerce')
    po_perf['PO_CREATED_ON'] = pd.to_datetime(po_perf['PO_CREATED_ON'], errors='coerce')
    
    po_perf['LEAD_TIME_DAYS'] = (po_perf['BP_CREATED_ON'] - po_perf['PO_APPROVED_ON']).dt.days
    po_perf = po_perf[(po_perf['LEAD_TIME_DAYS'] >= 0) & (po_perf['LEAD_TIME_DAYS'] < 100)]
    po_perf['APPROVAL_TIME_DAYS'] = (po_perf['PO_APPROVED_ON'] - po_perf['PO_CREATED_ON']).dt.days

    # 2. Warehouse Flow
    avg_bp_date = df_bp.groupby('PO_NO')['BP_CREATED_ON'].mean().reset_index()
    avg_sj_date = df_sj.groupby('NO_PO').agg({'SJ_CREATED_ON': 'mean', 'TGLAPP': 'max'}).reset_index()
    warehouse_flow = avg_bp_date.merge(avg_sj_date, left_on='PO_NO', right_on='NO_PO', how='inner')
    warehouse_flow['WAKTU_PROSES'] = (warehouse_flow['SJ_CREATED_ON'] - warehouse_flow['BP_CREATED_ON']).dt.days
    warehouse_flow = warehouse_flow[warehouse_flow['WAKTU_PROSES'] >= 0]
    warehouse_flow['SJ_APPROVAL_TIME'] = (warehouse_flow['TGLAPP'] - warehouse_flow['SJ_CREATED_ON']).dt.total_seconds() / (24 * 3600)

    # 3. Fulfillment & Unfulfilled
    po_target = df_po.groupby(['PO_NO', 'NAMA_SUPPLIER'])['JUMLAH'].sum().reset_index()
    bp_actual = df_bp.groupby('PO_NO')['JML_DITERIMA'].sum().reset_index()
    fulfillment = po_target.merge(bp_actual, on='PO_NO', how='left')
    fulfillment['JML_DITERIMA'] = fulfillment['JML_DITERIMA'].fillna(0)
    fulfillment['FILL_RATE'] = (fulfillment['JML_DITERIMA'] / fulfillment['JUMLAH']) * 100
    fulfillment['FILL_RATE_CAPPED'] = fulfillment['FILL_RATE'].apply(lambda x: 100 if x > 100 else x)
    
    supplier_fulfillment = fulfillment.groupby('NAMA_SUPPLIER').agg(
        AVG_FILL_RATE=('FILL_RATE_CAPPED', 'mean'),
        TOTAL_PO=('PO_NO', 'nunique')
    ).query('TOTAL_PO >= 5').sort_values(by='AVG_FILL_RATE', ascending=True)

    unfulfilled = fulfillment[fulfillment['JML_DITERIMA'] == 0]
    unfulfilled_count = unfulfilled.shape[0]
    top_unfulfilled_suppliers = unfulfilled['NAMA_SUPPLIER'].value_counts()

    # 4. Supplier Stats by VOLUME (UNIT)
    supplier_volume = df_po.groupby('NAMA_SUPPLIER')['JUMLAH'].sum().reset_index()
    supplier_volume = supplier_volume.sort_values(by='JUMLAH', ascending=False)
    
    # 5. Top Items & Vessels
    top_items = df_po.groupby('NAMA_BARANG')['JUMLAH'].sum().sort_values(ascending=False).head(10)
    df_sj_clean_vessel = df_sj.dropna(subset=['NAMAKODEAKUN_BARANG'])
    top_vessels_list = df_sj_clean_vessel['NAMA_KAPAL'].value_counts().head(5).index
    df_sj_top = df_sj_clean_vessel[df_sj_clean_vessel['NAMA_KAPAL'].isin(top_vessels_list)].copy()

    # ==========================================
    # LOAD BSM DATA (IF UPLOADED)
    # ==========================================
    df_bsm = None
    if file_bsm is not None:
        df_bsm = load_bsm(file_bsm)
        # Cleaning BSM
        df_bsm['BSM_CREATED_ON'] = pd.to_datetime(df_bsm['BSM_CREATED_ON'], dayfirst=True, errors='coerce')
        df_bsm['TOTAL'] = pd.to_numeric(df_bsm['TOTAL'], errors='coerce').fillna(0)
        df_bsm['JUMLAH'] = pd.to_numeric(df_bsm['JUMLAH'], errors='coerce').fillna(0)
        df_bsm['HRGSATUAN'] = pd.to_numeric(df_bsm['HRGSATUAN'], errors='coerce').fillna(0)
        # Tambah kolom Tahun/Bulan
        df_bsm['Year'] = df_bsm['BSM_CREATED_ON'].dt.year
        df_bsm['Month'] = df_bsm['BSM_CREATED_ON'].dt.to_period('M')

    # ==========================================
    # DASHBOARD LAYOUT
    # ==========================================
    
    # KPI Section
    st.subheader("Key Performance Indicators (Operational)")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Avg Approval Time", f"{po_perf['LEAD_TIME_DAYS'].mean():.2f} Days")
    c2.metric("Avg Warehouse Process", f"{warehouse_flow['WAKTU_PROSES'].mean():.2f} Days")
    val_sj = warehouse_flow['SJ_APPROVAL_TIME'].mean()
    c3.metric("Avg SJ Approval", f"{val_sj:.2f} Days" if pd.notnull(val_sj) else "N/A")
    c4.metric("Unfulfilled POs", f"{unfulfilled_count}", "Black Hole")

    # TABS
    tabs = st.tabs(["📈 Tren Operasional", "🚀 Analisis Consumption", "📋 Data Tabel", "🧠 Advanced Analytics", "💰 Financial & Historical (BSM)"])

    # --- TAB 1: Tren Harian ---
    with tabs[0]:
        st.subheader("Aktivitas Harian")
        
        # Agregasi Harian
        po_daily = df_po.groupby(df_po['PO_CREATED_ON'].dt.to_period('D'))['PO_NO'].nunique().sort_index().to_timestamp()
        
        fig_ts = plt.figure(figsize=(12, 5))
        plt.plot(po_daily.index, po_daily.values, marker='o', linestyle='-', color='blue')
        plt.title('Daily PO Created')
        plt.grid(True, alpha=0.3)
        st.pyplot(fig_ts)

        
        st.subheader("Shipping Activity by Day of Week")
        
        # Combo Chart Plotly
        days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        df_sj_temp = df_sj.copy()
        df_sj_temp['DayOfWeek'] = pd.Categorical(df_sj_temp['SJ_CREATED_ON'].dt.day_name(), categories=days_order, ordered=True)
        
        sj_dow = df_sj_temp.groupby('DayOfWeek')['SJ_NO'].nunique().reindex(days_order).fillna(0).reset_index()
        sj_dow['MA'] = sj_dow['SJ_NO'].rolling(window=2, min_periods=1).mean()

        fig_combo = go.Figure()
        fig_combo.add_trace(go.Bar(x=sj_dow['DayOfWeek'], y=sj_dow['SJ_NO'], name='Total SJ', marker_color='#FF4B4B', opacity=0.7))
        fig_combo.add_trace(go.Scatter(x=sj_dow['DayOfWeek'], y=sj_dow['MA'], name='Moving Avg', mode='lines+markers', line=dict(color='blue')))
        fig_combo.update_layout(title='Shipping Volume Trend', xaxis_title='Day', yaxis_title='Count', hovermode="x unified")
        
        st.plotly_chart(fig_combo, use_container_width=True)

        st.subheader("Scatter Plot: Target vs Realisasi") 
        
        # Gunakan df_po yang global (bukan mismatch saja) untuk scatter plot keseluruhan
        fig_scatter = px.scatter(
            df_po,
            x='JML_DISETUJUI',
            y='JML_DITERIMA',
            color='Status',
            color_discrete_map={'Sesuai': 'green', 'Terdapat Selisih': 'red'},
            hover_data=['PO_NO', 'NAMA_BARANG', 'SELISIH_JML'], 
            title="Peta Sebaran: Jumlah Disetujui vs Diterima",
            labels={
                'JML_DISETUJUI': 'Jumlah Disetujui (Target)',
                'JML_DITERIMA': 'Jumlah Diterima (Realisasi)'
            },
            height=500
        )

        # Garis referensi diagonal 1:1
        max_val = max(df_po['JML_DISETUJUI'].max(), df_po['JML_DITERIMA'].max())
        fig_scatter.add_shape(
            type="line",
            x0=0, y0=0, x1=max_val, y1=max_val,
            line=dict(color="gray", dash="dash"),
            opacity=0.5,
            name="Garis Referensi"
        )

        st.plotly_chart(fig_scatter, use_container_width=True)


    # --- TAB 2: Consumption ---
    with tabs[1]:
        col_con1, col_con2 = st.columns(2)
        with col_con1:
            st.subheader("Top 10 Lowest Fulfillment Suppliers")
            bottom_sup = supplier_fulfillment.head(10)
            fig_low = px.bar(bottom_sup, x='AVG_FILL_RATE', y=bottom_sup.index, orientation='h', 
                             title="Rata-rata Fill Rate (%)", color='AVG_FILL_RATE', color_continuous_scale='Reds_r')
            st.plotly_chart(fig_low, use_container_width=True)
        
        with col_con2:
            st.subheader("Top Supplier Unfulfilled (Macet)")
            top_unf = unfulfilled['NAMA_SUPPLIER'].value_counts().head(10)
            fig_unf = px.bar(x=top_unf.values, y=top_unf.index, orientation='h',
                             title="Jumlah PO Macet", color_discrete_sequence=['#d62728'])
            st.plotly_chart(fig_unf, use_container_width=True)

        st.markdown("---")
        st.subheader("Top 5 Item Consumption per Vessel (Barang Keluar/SJ)")
        
        col_item_name = 'NAMAKODEAKUN_BARANG' if 'NAMAKODEAKUN_BARANG' in df_sj.columns else 'NAMABRG'
        
        if col_item_name in df_sj_top.columns:
            df_raw_item_data = df_sj_top.groupby(['NAMA_KAPAL', col_item_name])['JUMLAH'].sum().reset_index()
            df_raw_item_data = df_raw_item_data.sort_values(['NAMA_KAPAL', 'JUMLAH'], ascending=[True, False])
            df_top5_per_vessel = df_raw_item_data.groupby('NAMA_KAPAL').head(5)
            
            fig_bar_multi = px.bar(
                df_top5_per_vessel, x='JUMLAH', y=col_item_name, color='JUMLAH',
                facet_col='NAMA_KAPAL', facet_col_wrap=2, orientation='h',
                title='Top 5 Items (Qty) per Vessel', color_continuous_scale='Bluered'
            )
            fig_bar_multi.update_yaxes(matches=None, showticklabels=True)
            fig_bar_multi.update_xaxes(matches=None)
            fig_bar_multi.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
            fig_bar_multi.update_layout(height=800)
            st.plotly_chart(fig_bar_multi, use_container_width=True)

    # --- TAB 3: Data Tabel (UPDATED METRIC) ---
    with tabs[2]:
        st.subheader("Ranking Supplier Berdasarkan Volume Barang")
        st.markdown("**Perubahan:** Ranking ini sekarang dihitung berdasarkan **Total Unit/Qty Barang** yang dipesan.")
        
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            st.markdown("#### Top 10 Supplier Terbesar (By Qty)")
            st.dataframe(supplier_volume.head(10).style.format({"JUMLAH": "{:,.0f}"}))
            
        with col_t2:
            st.markdown("#### Top 10 Barang Paling Laku")
            st.dataframe(top_items.head(10))

    # --- TAB 4: Advanced Analytics (Dead Stock) ---
    with tabs[3]:
        st.subheader("Potensi Dead Stock (Barang Masuk > Barang Keluar)")
        
        inbound = df_bp.groupby('PO_NO')['JML_DITERIMA'].sum().reset_index()
        po_mapping = df_po[['PO_NO', 'NAMA_BARANG']].drop_duplicates()
        inbound_merged = inbound.merge(po_mapping, on='PO_NO', how='left')
        total_in = inbound_merged.groupby('NAMA_BARANG')['JML_DITERIMA'].sum().reset_index()
        total_in.columns = ['NAMA_BARANG', 'TOTAL_MASUK']
        
        col_sj_barang = 'NAMABRG' if 'NAMABRG' in df_sj.columns else 'NAMAKODEAKUN_BARANG'
        total_out = df_sj.groupby(col_sj_barang)['JUMLAH'].sum().reset_index()
        total_out.columns = ['NAMA_BARANG', 'TOTAL_KELUAR']
        
        inventory_flow = total_in.merge(total_out, on='NAMA_BARANG', how='left').fillna(0)
        inventory_flow['SELISIH_STOK'] = inventory_flow['TOTAL_MASUK'] - inventory_flow['TOTAL_KELUAR']
        
        dead_stock = inventory_flow[inventory_flow['SELISIH_STOK'] > 0].sort_values(by='SELISIH_STOK', ascending=False).head(10)
        
        fig_stock = px.bar(
            dead_stock, x='SELISIH_STOK', y='NAMA_BARANG', orientation='h',
            title='Top 10 Barang Penumpukan (Gap Masuk vs Keluar)',
            color='SELISIH_STOK', color_continuous_scale='OrRd', text_auto=True
        )
        st.plotly_chart(fig_stock, use_container_width=True)

    # --- TAB 5: FINANCIAL & HISTORICAL (BSM DATA) ---
    with tabs[4]:
        if df_bsm is not None:
            st.success("✅ Data BSM (2023-2025) Berhasil Dimuat!")
            
            # --- KPI FINANCIAL ---
            total_spend = df_bsm['TOTAL'].sum()
            total_spend_2025 = df_bsm[df_bsm['Year'] == 2025]['TOTAL'].sum()
            total_spend_2024 = df_bsm[df_bsm['Year'] == 2024]['TOTAL'].sum()
            
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("Total Pemakaian (2023-2025)", f"Rp {total_spend:,.0f}")
            kpi2.metric("Pemakaian 2025 (YTD)", f"Rp {total_spend_2025:,.0f}")
            delta = ((total_spend_2025 - total_spend_2024)/total_spend_2024)*100 if total_spend_2024 > 0 else 0
            kpi3.metric("Pemakaian 2024", f"Rp {total_spend_2024:,.0f}", f"Vs 2025 (Run Rate): {delta:.1f}%")

            st.markdown("---")
            
            # --- 1. SPEND ANALYSIS (COST PER VESSEL) ---
            st.subheader("1. Analisis Biaya Realisasi per Kapal/Objek")
            col_b1, col_b2 = st.columns([2, 1])
            
            with col_b1:
                # Group by Object
                vessel_cost = df_bsm.groupby('NAMAOBJEK')['TOTAL'].sum().sort_values(ascending=False).head(15)
                fig_cost = px.bar(vessel_cost, x=vessel_cost.values, y=vessel_cost.index, orientation='h',
                                  title="Top 15 Kapal/Objek dengan Biaya Pemakaian Tertinggi",
                                  labels={'x': 'Total Rupiah', 'y': 'Nama Kapal'},
                                  color=vessel_cost.values, color_continuous_scale='Viridis')
                st.plotly_chart(fig_cost, use_container_width=True)
            
            with col_b2:
                st.info("""
                **Insight:** Grafik ini menunjukkan kapal mana yang paling "boros" (High Maintenance Cost).
                Gunakan data ini untuk mengevaluasi efisiensi mesin kapal atau investigasi kerusakan berulang.
                """)

            # --- 2. HISTORICAL TRENDS ---
            st.subheader("2. Tren Belanja Bulanan (2023-2025)")
            monthly_trend = df_bsm.groupby(df_bsm['BSM_CREATED_ON'].dt.to_period('M'))['TOTAL'].sum()
            monthly_trend.index = monthly_trend.index.to_timestamp()
            
            fig_trend = px.line(x=monthly_trend.index, y=monthly_trend.values, markers=True,
                                title="Tren Total Pemakaian Barang (Rupiah) per Bulan",
                                labels={'x': 'Bulan', 'y': 'Total Rupiah'})
            st.plotly_chart(fig_trend, use_container_width=True)

            # --- 3. DEPARTMENT ANALYSIS ---
            st.subheader("3. Analisis Biaya per Departemen (BAGIAN)")
            dept_cost = df_bsm.groupby('BAGIAN')['TOTAL'].sum().reset_index()
            fig_dept = px.pie(dept_cost, values='TOTAL', names='BAGIAN', title='Proporsi Biaya per Departemen',
                              hole=0.4)
            st.plotly_chart(fig_dept, use_container_width=True)

            # --- 4. PRICE TREND ANALYSIS ---
            st.subheader("4. Fluktuasi Harga Sparepart (Inflasi)")
            # Cari barang yang sering muncul
            top_bsm_items = df_bsm['NAMABRG'].value_counts().head(20).index.tolist()
            selected_item = st.selectbox("Pilih Barang untuk melihat tren harga:", top_bsm_items)
            
            item_trend = df_bsm[df_bsm['NAMABRG'] == selected_item].sort_values('BSM_CREATED_ON')
            if not item_trend.empty:
                fig_price = px.scatter(item_trend, x='BSM_CREATED_ON', y='HRGSATUAN', 
                                     title=f"Riwayat Harga Satuan: {selected_item}",
                                     trendline="lowess") # Menambahkan garis tren halus
                st.plotly_chart(fig_price, use_container_width=True)
            
        else:
            st.warning("⚠️ Silakan upload file BSM (2023-2025) di sidebar untuk melihat analisis finansial.")

else:
    st.info("👋 Silakan upload file TARIKA (Wajib) pada sidebar di sebelah kiri.")