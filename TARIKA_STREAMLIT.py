import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Konfigurasi halaman
st.set_page_config(
    page_title="Analisis Data Pembelian & Inventori",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.8rem;
        color: #2563EB;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .metric-card {
        background-color: #F3F4F6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #3B82F6;
        margin-bottom: 1rem;
    }
    .insight-box {
        background-color: #EFF6FF;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #93C5FD;
        margin-bottom: 1rem;
    }
    .data-table {
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# Header aplikasi
st.markdown('<h1 class="main-header">📊 Analisis Data Pembelian & Inventori</h1>', unsafe_allow_html=True)

# Sidebar untuk upload file
with st.sidebar:
    st.header("📁 Upload Data")
    uploaded_file = st.file_uploader("Upload file Excel", type=['xlsx', 'xls'])
    
    st.markdown("---")
    st.header("⚙️ Pengaturan Analisis")
    analysis_type = st.selectbox(
        "Pilih Jenis Analisis",
        ["Analisis Per Item", "Analisis Satuan", "Analisis Kategori", "Analisis Supplier"]
    )
    
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

        # --- TAMBAHAN: DATA CLEANING SATUAN ---
        # Buat kamus perbaikan (Dictionary)
        # Kiri: Tulisan yang salah/singkatan
        # Kanan: Tulisan standar yang diinginkan
        # Baca semua sheet
        po_sheet = pd.read_excel(file, sheet_name='PO', header=2)
        bp_sheet = pd.read_excel(file, sheet_name='BP', header=2)
        sj_sheet = pd.read_excel(file, sheet_name='SJ', header=3)
        
        # Standardisasi nama kolom (Hapus spasi & Uppercase)
        po_sheet.columns = po_sheet.columns.str.strip().str.upper()
        bp_sheet.columns = bp_sheet.columns.str.strip().str.upper()
        sj_sheet.columns = sj_sheet.columns.str.strip().str.upper()

        # --- DATA CLEANING SATUAN (STANDARDISASI) ---
        # Kamus perbaikan: Kiri (Salah/Variasi) -> Kanan (Standar Baku)
        perbaikan_satuan = {
            # 1. Satuan Hitung (Count) -> PCS
            'PCS': 'PCS', 'PC': 'PCS', 'BJ': 'PCS', 'BIJI': 'PCS', 
            'BH': 'PCS', 'BUAH': 'PCS', 'BTG': 'PCS', 'BTNG': 'PCS', 
            'BATANG': 'PCS', 'BKS': 'PCS', 'BUNGKUS': 'PCS','LNJR':'PCS','PTG':'PCS',
            'BTL': 'PCS', 'BOTOL': 'PCS', 'UNIT': 'PCS', 'UNITS': 'PCS','PCS (G)':'PCS','PAX':'PCS','LJR':'PCS','EA':'PCS','TIN':'PCS',
            
            # 2. Satuan Set -> SET
            'SET': 'SET', 'SETS': 'SET', 'PSG': 'SET', 'PASANG': 'SET', 'PAIR': 'SET',
            
            # 3. Satuan Panjang -> METER
            'MTR': 'METER', 'MTR.': 'METER', 'METERS': 'METER', 'M': 'METER', 
            'M.': 'METER', 'METER LARI': 'METER',
            
            # 4. Satuan Volume -> LITER / DRUM / PAIL
            'LTR': 'LITER', 'L': 'LITER', 'LITRE': 'LITER','LTRS':'LITER',
            'DR': 'DRUM', 'DRM': 'DRUM',
            'PL': 'PAIL', 'PAL': 'PAIL', 'PEL': 'PAIL',
            
            # 5. Satuan Berat -> KG
            'KG': 'KG', 'KGS': 'KG', 'KILO': 'KG', 'KILOGRAM': 'KG',
            #Lain-lain
            'GLN':'GALON','LSN':'LUSIN','LUSIN':'LUSIN','DOZ':'LUSIN','PAK':'PACK','LMBR':'LEMBAR','LBR':'LEMBAR','LMB':'LEMBAR','SHT':'LEMBAR','SHEET':'LEMBAR',
            'KGR':'KARUNG','ZAK':'SAK','TABUNG':'CAN','TBLT':'TAB','STP':'TAB',
        }
        
        # Terapkan standardisasi ke semua sheet (PO, BP, SJ)
        # Kita pakai Loop supaya tidak menulis kode berulang 3 kali
        dataframes = [po_sheet, bp_sheet, sj_sheet]
        
        for df in dataframes:
            # Cek apakah kolom SATUAN ada di dataframe tersebut
            if 'SATUAN' in df.columns:
                # Langkah 1: Ubah ke String, Uppercase, dan Hapus Spasi (PENTING!)
                # Agar " Pcs " (ada spasi) tetap terbaca sebagai "PCS"
                df['SATUAN'] = df['SATUAN'].astype(str).str.strip().str.upper()
                
                # Langkah 2: Ganti nilai sesuai kamus perbaikan
                # Jika kata tidak ada di kamus, biarkan apa adanya
                df['SATUAN'] = df['SATUAN'].replace(perbaikan_satuan)
        
        return {
            'PO': po_sheet,
            'BP': bp_sheet,
            'SJ': sj_sheet
        }
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
        return None


def analyze_comparison_trend(po_df, bp_df, sj_df):
    # Helper function untuk memproses data per bulan
    def get_monthly_data(df, date_col, value_col):
        # Cek apakah kolom tanggal DAN kolom value ada di dataframe
        if date_col not in df.columns or value_col not in df.columns:
            # Debugging silent: print ini akan muncul di terminal jika error
            # print(f"Missing cols: {date_col} or {value_col}") 
            return None
        
        temp_df = df.copy()
        temp_df[date_col] = pd.to_datetime(temp_df[date_col], errors='coerce')
        temp_df = temp_df.dropna(subset=[date_col])
        return temp_df.set_index(date_col).resample('M')[value_col].sum()

    # 1. Proses Data PO (Kolom quantity: JUMLAH)
    po_trend = get_monthly_data(po_df, 'PO_CREATED_ON', 'JUMLAH')
    
    # 2. Proses Data BP (PERBAIKAN DISINI)
    # Gunakan 'BP_CREATED_ON' untuk tanggal
    # Gunakan 'JML_DITERIMA' untuk quantity (bukan JUMLAH)
    bp_trend = get_monthly_data(bp_df, 'BP_CREATED_ON', 'JML_DITERIMA')
    
    # 3. Proses Data SJ (Pastikan nama kolom quantity di SJ juga dicek)
    # Asumsi: Di SJ namanya masih 'JUMLAH', kalau beda sesuaikan juga
    sj_date_col = 'SJ_CREATED_ON' if 'SJ_CREATED_ON' in sj_df.columns else 'TANGGAL'
    sj_trend = get_monthly_data(sj_df, sj_date_col, 'JUMLAH')

    # Visualisasi
    fig = go.Figure()

    if po_trend is not None:
        fig.add_trace(go.Scatter(
            x=po_trend.index, y=po_trend.values,
            mode='lines+markers', name='Pembelian (PO)',
            line=dict(color='#1E3A8A', width=3)
        ))

    if bp_trend is not None:
        fig.add_trace(go.Scatter(
            x=bp_trend.index, y=bp_trend.values,
            mode='lines+markers', name='Penerimaan (BP)',
            line=dict(color='#10B981', width=3, dash='dash')
        ))

    if sj_trend is not None:
        fig.add_trace(go.Scatter(
            x=sj_trend.index, y=sj_trend.values,
            mode='lines+markers', name='Penggunaan (SJ)',
            line=dict(color='#F59E0B', width=3)
        ))

    fig.update_layout(
        title='Perbandingan Tren Supply Chain (PO vs BP vs SJ)',
        xaxis_title='Bulan',
        yaxis_title='Total Quantity',
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    return fig





def analyze_specific_item_trend(po_df, bp_df, sj_df, item_name):
    # 1. Filter Dataframe berdasarkan Nama Barang
    po_item = po_df[po_df['NAMA_BARANG'] == item_name].copy()
    
    # Perlu hati-hati: Pastikan nama kolom 'NAMA_BARANG' di BP dan SJ sama persis
    # Jika di BP namanya 'NAMABRG', sesuaikan kodenya:
    bp_col_name = 'NAMABRG' if 'NAMABRG' in bp_df.columns else 'NAMA_BARANG'
    bp_item = bp_df[bp_df[bp_col_name] == item_name].copy()
    
    sj_item = sj_df[sj_df['NAMABRG'] == item_name].copy()

    # 2. Helper Resample (Kembali ke .sum() karena kita lihat Stok Fisik)
    def get_monthly_qty(df, date_col, qty_col):
        if df.empty or date_col not in df.columns or qty_col not in df.columns:
            return None
        
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        return df.set_index(date_col).resample('M')[qty_col].sum()

    # 3. Proses Data
    po_trend = get_monthly_qty(po_item, 'PO_CREATED_ON', 'JUMLAH')
    # Ingat: Di BP kolom quantity namanya 'JML_DITERIMA'
    bp_trend = get_monthly_qty(bp_item, 'BP_CREATED_ON', 'JML_DITERIMA') 
    sj_trend = get_monthly_qty(sj_item, 'SJ_CREATED_ON', 'JUMLAH')

    # 4. Visualisasi
    fig = go.Figure()

    if po_trend is not None and not po_trend.empty:
        fig.add_trace(go.Scatter(
            x=po_trend.index, y=po_trend.values,
            mode='lines+markers', name='Dipesan (PO)',
            line=dict(color='#1E3A8A', width=3)
        ))

    if bp_trend is not None and not bp_trend.empty:
        fig.add_trace(go.Scatter(
            x=bp_trend.index, y=bp_trend.values,
            mode='lines+markers', name='Masuk Gudang (BP)',
            line=dict(color='#10B981', width=3, dash='dash')
        ))

    if sj_trend is not None and not sj_trend.empty:
        fig.add_trace(go.Scatter(
            x=sj_trend.index, y=sj_trend.values,
            mode='lines+markers', name='Dipakai (SJ)',
            line=dict(color='#F59E0B', width=3)
        ))

    fig.update_layout(
        title=f'Tren Stok Fisik: {item_name}',
        xaxis_title='Bulan',
        yaxis_title='Quantity (Satuan Barang)',
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    return fig






# Fungsi untuk analisis per item berdasarkan satuan
def analyze_items_by_unit(df):
    if 'SATUAN' not in df.columns or 'JUMLAH' not in df.columns or 'NAMA_BARANG' not in df.columns:
        return None
    
    # Group by nama barang dan satuan
    analysis = df.groupby(['NAMA_BARANG', 'SATUAN']).agg({
        'JUMLAH': 'sum',
        'PO_NO': 'nunique',
        'NAMA_SUPPLIER': pd.Series.nunique
    }).reset_index()
    
    analysis = analysis.rename(columns={
        'PO_NO': 'JUMLAH_PO',
        'NAMA_SUPPLIER': 'JUMLAH_SUPPLIER'
    })
    
    return analysis.sort_values('JUMLAH', ascending=False)

# Fungsi untuk analisis satuan
def analyze_units(df):
    if 'SATUAN' not in df.columns:
        return None
    
    unit_analysis = df['SATUAN'].value_counts().reset_index()
    unit_analysis.columns = ['SATUAN', 'JUMLAH_TRANSAKSI']
    
    # Hitung total quantity per satuan
    quantity_by_unit = df.groupby('SATUAN')['JUMLAH'].sum().reset_index()
    quantity_by_unit.columns = ['SATUAN', 'TOTAL_QUANTITY']
    
    # Merge kedua analisis
    result = pd.merge(unit_analysis, quantity_by_unit, on='SATUAN')
    
    return result.sort_values('TOTAL_QUANTITY', ascending=False)

# Fungsi untuk mencari item tertentu
def find_specific_item(df, item_name):
    item_name_lower = item_name.lower()
    
    # Cari item yang mengandung kata kunci
    mask = df['NAMA_BARANG'].str.lower().str.contains(item_name_lower, na=False)
    filtered_df = df[mask].copy()
    
    if filtered_df.empty:
        return None
    
    # Group by satuan
    result = filtered_df.groupby(['NAMA_BARANG', 'SATUAN']).agg({
        'JUMLAH': 'sum',
        'PO_NO': 'nunique',
        'NAMA_SUPPLIER': pd.Series.nunique
    }).reset_index()
    
    result = result.rename(columns={
        'PO_NO': 'JUMLAH_PO',
        'NAMA_SUPPLIER': 'JUMLAH_SUPPLIER'
    })
    
    return result.sort_values('JUMLAH', ascending=False)

# Fungsi untuk analisis insight
def generate_insights(df):
    insights = []
    
    # 1. Analisis satuan
    if 'SATUAN' in df.columns:
        unit_counts = df['SATUAN'].value_counts()
        most_common_unit = unit_counts.index[0]
        insights.append(f"📦 **Satuan paling umum**: {most_common_unit} ({unit_counts.iloc[0]} transaksi)")
    
    # 2. Analisi total quantity
    if 'JUMLAH' in df.columns:
        total_quantity = df['JUMLAH'].sum()
        avg_quantity = df['JUMLAH'].mean()
        insights.append(f"🧮 **Total quantity semua item**: {total_quantity:,.0f}")
        insights.append(f"📈 **Rata-rata quantity per item**: {avg_quantity:.2f}")
    
    # 3. Analisis supplier
    if 'NAMA_SUPPLIER' in df.columns:
        unique_suppliers = df['NAMA_SUPPLIER'].nunique()
        top_supplier = df['NAMA_SUPPLIER'].value_counts().index[0]
        insights.append(f"🏢 **Jumlah supplier unik**: {unique_suppliers}")
        insights.append(f"🥇 **Supplier teraktif**: {top_supplier}")
    
    # 4. Analisis kategori
    if 'KODE_KATEGORI' in df.columns:
        unique_categories = df['KODE_KATEGORI'].nunique()
        insights.append(f"🏷️ **Jumlah kategori barang**: {unique_categories}")
    
    # 5. Analisis trend waktu
    if 'PO_CREATED_ON' in df.columns:
        try:
            df['PO_CREATED_ON'] = pd.to_datetime(df['PO_CREATED_ON'], errors='coerce')
            monthly_trend = df.groupby(df['PO_CREATED_ON'].dt.to_period('M')).size()
            if len(monthly_trend) > 1:
                latest_month = monthly_trend.index[-1]
                latest_count = monthly_trend.iloc[-1]
                insights.append(f"📅 **Aktivitas bulan {latest_month}**: {latest_count} transaksi")
        except:
            pass
    
    return insights

# Main aplikasi
if uploaded_file is not None:
    # Proses file
    data = process_excel_file(uploaded_file)
    
    if data is not None:
        po_df = data['PO']
        
        # Tampilkan informasi dasar
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Jenis Barang", len(po_df))
        with col2:
            st.metric("Jumlah PO", po_df['PO_NO'].nunique())
        with col3:
            st.metric("Jumlah Supplier", po_df['NAMA_SUPPLIER'].nunique())
        with col4:
            st.metric("Total Quantity", f"{po_df['JUMLAH'].sum():,.0f}")
        
        # Tab untuk berbagai analisis
        tab1, tab2, tab3, tab4 = st.tabs(["📈 Analisis Utama", "🔍 Cari Item", "📋 Detail Data", "💡 Insight & Rekomendasi"])
        
        with tab1:
            st.markdown('<h2 class="sub-header">Analisis Berdasarkan Satuan</h2>', unsafe_allow_html=True)
            
            st.markdown("---") # Garis pemisah
            st.markdown('<h3 class="sub-header">📈 Tren Aktivitas (Supply Chain)</h3>', unsafe_allow_html=True)
            
            # Ambil data BP dan SJ juga
            bp_df = data.get('BP')
            sj_df = data.get('SJ')
            
            if bp_df is not None and sj_df is not None:
                # Panggil fungsi analisis trend
                fig_trend = analyze_comparison_trend(po_df, bp_df, sj_df)
                st.plotly_chart(fig_trend, use_container_width=True)
                
                # Tambahan Insight Cepat di bawah grafik
                col_t1, col_t2, col_t3 = st.columns(3)
                with col_t1:
                    st.info("💡 **Garis Biru (PO)**: Apa yang kita pesan.")
                with col_t2:
                    st.success("💡 **Garis Hijau (BP)**: Apa yang sudah datang di gudang.")
                with col_t3:
                    st.warning("💡 **Garis Kuning (SJ)**: Apa yang sudah dipakai/keluar.")
            else:
                st.warning("Data BP atau SJ tidak ditemukan/kosong. Pastikan file Excel memiliki sheet BP dan SJ.")




            st.markdown("---")
            st.markdown('<h3 class="sub-header">🔍 Deep Dive: Analisis Per Item</h3>', unsafe_allow_html=True)

            # Ambil daftar semua barang unik dari PO untuk dropdown
            item_list = sorted(po_df['NAMA_BARANG'].unique().tolist())
            
            # Dropdown Selection
            selected_item_trend = st.selectbox("Pilih Barang untuk melihat Tren Stok:", item_list)
            
            if selected_item_trend:
                # Panggil fungsi baru
                # Pastikan bp_df dan sj_df sudah diambil dari variable 'data' sebelumnya
                bp_df = data.get('BP')
                sj_df = data.get('SJ')
                
                if bp_df is not None and sj_df is not None:
                    fig_item = analyze_specific_item_trend(po_df, bp_df, sj_df, selected_item_trend)
                    st.plotly_chart(fig_item, use_container_width=True)
                    
                    # Tampilkan Info Satuan Barang Tersebut
                    satuan_info = po_df[po_df['NAMA_BARANG'] == selected_item_trend]['SATUAN'].iloc[0]
                    st.caption(f"ℹ️ Grafik di atas menampilkan jumlah fisik dalam satuan: **{satuan_info}**")







            # Analisis per item berdasarkan satuan
            if analysis_type == "Analisis Per Item":
                st.markdown('<h3 class="sub-header">Analisis Per Item</h3>', unsafe_allow_html=True)

                item_analysis = analyze_items_by_unit(po_df)
                if item_analysis is not None:
                    st.dataframe(
                        item_analysis.head(20),
                        use_container_width=True,
                        column_config={
                            "NAMA_BARANG": "Nama Barang",
                            "SATUAN": "Satuan",
                            "JUMLAH": st.column_config.NumberColumn("Total Quantity", format="%,.0f"),
                            "JUMLAH_PO": "Jumlah PO",
                            "JUMLAH_SUPPLIER": "Jumlah Supplier"
                        }
                    )
                    
                    # Visualisasi
                    col1, col2 = st.columns(2)
                    with col1:
                        top_items = item_analysis.head(10)
                        fig1 = px.bar(top_items, x='NAMA_BARANG', y='JUMLAH', 
                                     color='SATUAN', title='10 Item dengan Quantity Terbesar')
                        st.plotly_chart(fig1, use_container_width=True)
                    
                    with col2:
                        fig2 = px.pie(item_analysis.head(10), values='JUMLAH', 
                                     names='NAMA_BARANG', title='Distribusi Quantity Top 10 Item')
                        st.plotly_chart(fig2, use_container_width=True)
            
            # Analisis satuan
            elif analysis_type == "Analisis Satuan":
                st.markdown('<h3 class="sub-header">Analisis Satuan</h3>', unsafe_allow_html=True)

                unit_analysis = analyze_units(po_df)
                if unit_analysis is not None:
                    st.dataframe(
                        unit_analysis,
                        use_container_width=True,
                        column_config={
                            "SATUAN": "Satuan",
                            "JUMLAH_TRANSAKSI": "Jumlah Transaksi",
                            "TOTAL_QUANTITY": st.column_config.NumberColumn("Total Quantity", format="%,.0f")
                        }
                    )
                    
                    # Visualisasi
                    fig = px.bar(unit_analysis, x='SATUAN', y='TOTAL_QUANTITY',
                                title='Total Quantity per Satuan')
                    st.plotly_chart(fig, use_container_width=True)
            
            # Analisis kategori
            elif analysis_type == "Analisis Kategori":
                st.markdown('<h3 class="sub-header">Analisis Kategori</h3>', unsafe_allow_html=True)

                if 'KODE_KATEGORI' in po_df.columns:
                    category_analysis = po_df.groupby('KODE_KATEGORI').agg({
                        'JUMLAH': 'sum',
                        'PO_NO': 'nunique',
                        'NAMA_BARANG': pd.Series.nunique
                    }).reset_index()
                    
                    category_analysis = category_analysis.rename(columns={
                        'PO_NO': 'JUMLAH_PO',
                        'NAMA_BARANG': 'JENIS_BARANG'
                    })
                    
                    st.dataframe(
                        category_analysis.sort_values('JUMLAH', ascending=False),
                        use_container_width=True
                    )
            
            # Analisis supplier
            elif analysis_type == "Analisis Supplier":
                st.markdown('<h3 class="sub-header">Analisis Supplier</h3>', unsafe_allow_html=True)
                if 'NAMA_SUPPLIER' in po_df.columns:
                    supplier_analysis = po_df.groupby('NAMA_SUPPLIER').agg({
                        'JUMLAH': 'sum',
                        'PO_NO': 'nunique',
                        'NAMA_BARANG': pd.Series.nunique
                    }).reset_index()
                    
                    supplier_analysis = supplier_analysis.rename(columns={
                        'PO_NO': 'JUMLAH_PO',
                        'NAMA_BARANG': 'JENIS_BARANG'
                    })
                    
                    st.dataframe(
                        supplier_analysis.sort_values('JUMLAH', ascending=False).head(20),
                        use_container_width=True
                    )
        
        with tab2:
            st.markdown('<h2 class="sub-header">Cari Item Spesifik</h2>', unsafe_allow_html=True)
            
            # Search box
            search_term = st.text_input("Masukkan nama item (contoh: OLI TELLUS S2 VX 32):")
            
            if search_term:
                result = find_specific_item(po_df, search_term)
                
                if result is not None and not result.empty:
                    st.success(f"Ditemukan {len(result)} hasil untuk '{search_term}'")
                    
                    # Tampilkan hasil
                    for idx, row in result.iterrows():
                        with st.expander(f"{row['NAMA_BARANG']} - {row['SATUAN']}"):
                            col1, col2, col3, col4 = st.columns(4)
                            col1.metric("Total Quantity", f"{row['JUMLAH']:,.0f}")
                            col2.metric("Satuan", row['SATUAN'])
                            col3.metric("Jumlah PO", row['JUMLAH_PO'])
                            col4.metric("Jumlah Supplier", row['JUMLAH_SUPPLIER'])
                            
                            # Tampilkan detail transaksi untuk item ini
                            item_transactions = po_df[
                                # po_df['NAMA_BARANG'].str.lower().str.contains(search_term.lower(), na=False) &
                                # (po_df['SATUAN'] == row['SATUAN'])
                                (po_df['NAMA_BARANG'] == row['NAMA_BARANG']) & 
                                (po_df['SATUAN'] == row['SATUAN'])
                            ]
                            item_transactions = item_transactions.sort_values(by='NAMA_SUPPLIER', ascending=True)
                            
                            st.dataframe(
                                item_transactions[['PO_NO', 'NAMA_BARANG', 'PO_CREATED_ON', 'JUMLAH', 'NAMA_SUPPLIER', 'STATUS']],
                                use_container_width=True
                            )
                else:
                    st.warning(f"Item '{search_term}' tidak ditemukan")
        
        with tab3:
            st.markdown('<h2 class="sub-header">Detail Data</h2>', unsafe_allow_html=True)
            
            # Filter data
            col1, col2 = st.columns(2)
            with col1:
                if 'SATUAN' in po_df.columns:
                    selected_units = st.multiselect(
                        "Filter Satuan",
                        options=po_df['SATUAN'].unique(),
                        default=po_df['SATUAN'].unique()[:3]
                    )
                else:
                    selected_units = []
            
            with col2:
                if 'KODE_KATEGORI' in po_df.columns:
                    selected_categories = st.multiselect(
                        "Filter Kategori",
                        options=po_df['KODE_KATEGORI'].unique(),
                        default=po_df['KODE_KATEGORI'].unique()[:3]
                    )
                else:
                    selected_categories = []
            
            # Terapkan filter
            filtered_df = po_df.copy()
            if selected_units:
                filtered_df = filtered_df[filtered_df['SATUAN'].isin(selected_units)]
            if selected_categories:
                filtered_df = filtered_df[filtered_df['KODE_KATEGORI'].isin(selected_categories)]
            
            # Tampilkan data
            st.dataframe(
                filtered_df[['PO_NO', 'NAMA_BARANG', 'SATUAN', 'JUMLAH', 
                           'KODE_KATEGORI', 'NAMA_SUPPLIER', 'STATUS']],
                use_container_width=True,
                height=400
            )
            
            # Opsi ekspor
            if st.button("📥 Ekspor Data ke Excel"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Filtered_Data')
                st.download_button(
                    label="Download Excel",
                    data=output.getvalue(),
                    file_name="data_terfilter.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with tab4:
            st.markdown('<h2 class="sub-header">Insight & Rekomendasi</h2>', unsafe_allow_html=True)
            
            # Generate insights
            insights = generate_insights(po_df)
            
            # Tampilkan insights
            for insight in insights:
                st.markdown(f'<div class="insight-box">{insight}</div>', unsafe_allow_html=True)
            
            # Analisis masalah dan solusi
            st.markdown("### 🎯 Identifikasi Masalah & Solusi")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                **Masalah yang Ditemukan:**
                
                1. **Variasi Satuan yang Tidak Standar**
                   - Item sama menggunakan satuan berbeda (DRUM, LITER, PAIL)
                   - Menyulitkan perhitungan inventory
                
                2. **Item dengan Banyak Satuan**
                   - Contoh: OLI TELLUS tersedia dalam DRUM, LITER, PAIL
                   - Kompleksitas dalam pengelolaan stok
                
                3. **Data Tidak Konsisten**
                   - Nama item tidak seragam
                   - Satuan tidak terstandarisasi
                
                4. **Supplier Terlalu Banyak**
                   - Item sama dibeli dari supplier berbeda
                   - Potensi inkonsistensi kualitas
                """)
            
            with col2:
                st.markdown("""
                **Rekomendasi Solusi:**
                
                1. **Standardisasi Satuan**
                   - Buat master list satuan yang diperbolehkan
                   - Konversi semua satuan ke unit standar
                
                2. **Konsolidasi Item**
                   - Group item yang sama dengan satuan berbeda
                   - Buat konversi antar satuan (1 DRUM = 208 LITER)
                
                3. **Implementasi Master Data**
                   - Buat katalog barang terpusat
                   - Standarisasi nama dan satuan
                
                4. **Optimasi Supplier**
                   - Konsolidasi pembelian ke supplier terpilih
                   - Negosiasi harga berdasarkan volume
                
                5. **Sistem Tracking**
                   - Implementasi barcode/RFID
                   - Real-time inventory tracking
                """)
            
            # Contoh analisis untuk OLI TELLUS
            st.markdown("### 🔍 Contoh Analisis: OLI TELLUS")
            
            # Cari semua item OLI TELLUS
            oli_items = find_specific_item(po_df, "OLI TELLUS")
            if oli_items is not None and not oli_items.empty:
                st.dataframe(oli_items, use_container_width=True)
                
                # Analisis detail
                st.markdown("**Analisis:**")
                st.write(f"OLI TELLUS ditemukan dalam {len(oli_items)} variasi satuan:")
                for idx, row in oli_items.iterrows():
                    st.write(f"- {row['SATUAN']}: {row['JUMLAH']:,.0f} unit dari {row['JUMLAH_PO']} PO")
            
            else:
                st.info("Contoh: Untuk item seperti OLI TELLUS S2 VX 32, sistem akan menampilkan total quantity per satuan (DRUM/LITER/PAIL) beserta detail transaksinya.")
    
    else:
        st.error("Gagal memproses file. Pastikan format file sesuai.")
else:
    # Tampilan default saat belum ada file
    st.info("👈 Silakan upload file Excel melalui sidebar di sebelah kiri.")
    
    # Contoh struktur data yang diharapkan
    st.markdown("""
    ### 📋 Struktur File Excel yang Diperlukan:
    
    **Sheet PO (Purchase Order):**
    - PO_NO, PO_CREATED_ON, NAMA_BARANG, SATUAN, JUMLAH, NAMA_SUPPLIER, dll.
    
    **Sheet BP (Bon Permintaan):**
    - BPS_NO, KODE_BARANG, NAMA_BARANG, SATUAN, JUMLAH, dll.
    
    **Sheet SJ (Surat Jalan):**
    - SJ_NO, PO_NO, NAMA_BARANG, SATUAN, JUMLAH_DIKIRIM, dll.
    
    ### 🎯 Fitur Utama Aplikasi:
    1. **Analisis per Item**: Melihat total quantity per item berdasarkan satuan
    2. **Pencarian Item**: Cari item spesifik seperti "OLI TELLUS S2 VX 32"
    3. **Analisis Satuan**: Distribusi penggunaan satuan pengukuran
    4. **Insight Otomatis**: Identifikasi pola dan masalah
    5. **Filter Interaktif**: Filter data berdasarkan berbagai kriteria
    """)
    
    # Contoh data
    example_data = pd.DataFrame({
        'NAMA_BARANG': ['OLI TELLUS S2 VX 32', 'OLI TELLUS S2 VX 32', 'BAUT L 8 X 25', 'MUR BAUT BAJA'],
        'SATUAN': ['DRUM', 'LITER', 'PCS', 'PCS'],
        'JUMLAH': [5, 1000, 25, 25],
        'PO_NO': ['PO/001', 'PO/002', 'PO/003', 'PO/003'],
        'NAMA_SUPPLIER': ['Supplier A', 'Supplier B', 'Supplier C', 'Supplier C']
    })
    
    st.dataframe(example_data, use_container_width=True)
