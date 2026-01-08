import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Konfigurasi Halaman
st.set_page_config(page_title="Dashboard Biaya & Stok Material", layout="wide")

# Fungsi untuk memuat data
@st.cache_data
def load_data(file):
    try:
        df = pd.read_excel(file)

        # 1. Konversi Tanggal
        # Format di snippet terlihat seperti dd/mm/yyyy (e.g., 18/03/2023)
        df['BSM_CREATED_ON'] = pd.to_datetime(df['BSM_CREATED_ON'], dayfirst=True, errors='coerce')
        
        # 2. Pastikan kolom numerik benar
        numeric_cols = ['TOTAL', 'JUMLAH', 'JMLDISETUJUI', 'HRGSATUAN']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Buat Kategori Aset (Kapal vs Alat Berat vs Lainnya)
        def categorize_asset(row):
            if pd.notna(row['KODEKAPAL']):
                return 'Kapal'
            elif pd.notna(row['KODEALATBERAT']):
                return 'Alat Berat'
            else:
                return 'Lainnya'
        
        df['Tipe_Aset'] = df.apply(categorize_asset, axis=1)
        
        # 4. Hitung Fill Rate (Persentase Pemenuhan) per baris
        # Menghindari pembagian dengan nol
        df['Fill_Rate'] = df.apply(lambda x: (x['JMLDISETUJUI'] / x['JUMLAH'] * 100) if x['JUMLAH'] > 0 else 0, axis=1)

        # 5. Extract Bulan-Tahun untuk grouping
        df['Bulan_Tahun'] = df['BSM_CREATED_ON'].dt.to_period('M').astype(str)

        return df
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
        return None

# --- SIDEBAR ---
st.sidebar.title("Filter Data")

# Upload File di Sidebar untuk memudahkan penggantian data
uploaded_file = st.sidebar.file_uploader("2. Upload File BSM (2023-2025)", type=["xlsx"], key="bsm")

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    if df is not None:
        # Filter Tanggal
        min_date = df['BSM_CREATED_ON'].min().date()
        max_date = df['BSM_CREATED_ON'].max().date()
        
        start_date, end_date = st.sidebar.date_input(
            "Rentang Tanggal",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )

        # Filter Bagian (Departemen)
        all_bagian = ['Semua'] + sorted(df['BAGIAN'].dropna().unique().tolist())
        selected_bagian = st.sidebar.selectbox("Pilih Bagian / Departemen", all_bagian)

        # Filter Tipe Aset
        all_types = ['Semua'] + sorted(df['Tipe_Aset'].unique().tolist())
        selected_type = st.sidebar.selectbox("Pilih Tipe Aset", all_types)

        # Terapkan Filter
        mask = (df['BSM_CREATED_ON'].dt.date >= start_date) & (df['BSM_CREATED_ON'].dt.date <= end_date)
        df_filtered = df.loc[mask]

        if selected_bagian != 'Semua':
            df_filtered = df_filtered[df_filtered['BAGIAN'] == selected_bagian]
        
        if selected_type != 'Semua':
            df_filtered = df_filtered[df_filtered['Tipe_Aset'] == selected_type]

        # --- MAIN PAGE ---
        st.title("📊 Dashboard Analisis Maintenance & Stok")
        st.write(f"Menampilkan data dari **{start_date}** sampai **{end_date}**")

        # KPI Metrics
        col1, col2, col3, col4 = st.columns(4)
        
        total_cost = df_filtered['TOTAL'].sum()
        total_items_req = df_filtered['JUMLAH'].sum()
        total_items_app = df_filtered['JMLDISETUJUI'].sum()
        avg_fill_rate = (total_items_app / total_items_req * 100) if total_items_req > 0 else 0
        
        with col1:
            st.metric("Total Biaya (Maintenance)", f"Rp {total_cost:,.0f}")
        with col2:
            st.metric("Total Item Diminta", f"{total_items_req:,.0f}")
        with col3:
            st.metric("Total Item Disetujui", f"{total_items_app:,.0f}")
        with col4:
            st.metric("Rata-rata Fill Rate", f"{avg_fill_rate:.2f}%", help="Persentase permintaan yang dapat dipenuhi gudang")

        st.markdown("---")

        # TABS untuk memisahkan analisis
        tab1, tab2, tab3 = st.tabs(["💰 Analisis Biaya", "📦 Efisiensi Stok", "📋 Data Raw"])

        # --- TAB 1: ANALISIS BIAYA ---
        with tab1:
            st.subheader("Tren Biaya Operasional")
            
            # 1. Line Chart: Cost over Time
            cost_over_time = df_filtered.groupby('Bulan_Tahun')['TOTAL'].sum().reset_index()
            fig_trend = px.line(cost_over_time, x='Bulan_Tahun', y='TOTAL', markers=True, 
                                title="Tren Pengeluaran Biaya per Bulan")
            st.plotly_chart(fig_trend, use_container_width=True)

            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                # 2. Bar Chart: Top 10 Cost by Object (Kapal/Alat Berat)
                st.subheader("Top 10 Aset dengan Biaya Tertinggi")
                # Group by Nama Objek (Kapal/Alat)
                cost_by_obj = df_filtered.groupby('NAMAOBJEK')['TOTAL'].sum().reset_index().sort_values('TOTAL', ascending=False).head(10)
                fig_obj = px.bar(cost_by_obj, x='TOTAL', y='NAMAOBJEK', orientation='h', 
                                 title="Biaya per Objek (Kapal/Alat)", color='TOTAL', color_continuous_scale='Reds')
                fig_obj.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig_obj, use_container_width=True)

            with col_chart2:
                # 3. Pie Chart: Cost by Department (Bagian)
                st.subheader("Proporsi Biaya per Departemen")
                cost_by_dept = df_filtered.groupby('BAGIAN')['TOTAL'].sum().reset_index()
                fig_dept = px.pie(cost_by_dept, values='TOTAL', names='BAGIAN', title="Distribusi Biaya per Bagian")
                st.plotly_chart(fig_dept, use_container_width=True)

        # --- TAB 2: EFISIENSI STOK ---
        with tab2:
            st.subheader("Analisis Pergerakan Stok")

            col_stok1, col_stok2 = st.columns(2)

            with col_stok1:
                # 1. Barang Paling Sering Keluar (Frequency)
                st.write("**Top 10 Barang Sering Keluar (Frekuensi)**")
                top_freq = df_filtered['NAMABRG'].value_counts().head(10).reset_index()
                top_freq.columns = ['Nama Barang', 'Frekuensi Transaksi']
                st.dataframe(top_freq, use_container_width=True)

            with col_stok2:
                # 2. Barang Dengan Nilai Tertinggi (Value)
                st.write("**Top 10 Barang Pengeluaran Tertinggi (Rupiah)**")
                top_val = df_filtered.groupby('NAMABRG')['TOTAL'].sum().reset_index().sort_values('TOTAL', ascending=False).head(10)
                st.dataframe(top_val.style.format({"TOTAL": "Rp {:,.0f}"}), use_container_width=True)

            st.markdown("---")
            st.subheader("Analisis Supply vs Demand (Fill Rate)")
            
            # Scatter plot: Requested vs Approved
            # Agregasi per barang untuk melihat barang mana yang sering shortfall
            fill_rate_by_item = df_filtered.groupby('NAMABRG')[['JUMLAH', 'JMLDISETUJUI']].sum().reset_index()
            # Filter hanya barang yang diminta cukup banyak (>10 unit misalnya) agar chart tidak penuh noise
            fill_rate_by_item = fill_rate_by_item[fill_rate_by_item['JUMLAH'] > 10]
            
            fig_scatter = px.scatter(fill_rate_by_item, x='JUMLAH', y='JMLDISETUJUI', hover_name='NAMABRG',
                                     title="Permintaan (X) vs Disetujui (Y) per Barang",
                                     labels={'JUMLAH': 'Jumlah Diminta', 'JMLDISETUJUI': 'Jumlah Disetujui'})
            # Tambah garis diagonal (ideal)
            fig_scatter.add_shape(type="line", line=dict(dash='dash', color='gray'),
                                x0=0, y0=fill_rate_by_item['JUMLAH'].max(),
                                x1=0, y1=fill_rate_by_item['JUMLAH'].max())
            st.plotly_chart(fig_scatter, use_container_width=True)
            st.caption("*Titik yang berada jauh di bawah garis diagonal menunjukkan barang yang sering diminta tapi stok tidak mencukupi.*")

        # --- TAB 3: DATA RAW ---
        with tab3:
            st.subheader("Detail Data Transaksi")
            st.dataframe(df_filtered)

else:
    st.info("Silakan upload file CSV 'BSM' pada sidebar untuk memulai analisis.")