import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import streamlit as st

# 1. Load file 
file_path = '01. TARIKA PO, BP dan SJ 2025.xlsx'
xls = pd.ExcelFile(file_path)

# 2. Lihat nama sheet yang ada 
# print(xls.sheet_names) 

# 3. Buat variabel berbeda berdasarkan nama sheet
df_po = pd.read_excel(xls, 'PO', header=2)
df_bp = pd.read_excel(xls, 'BP', header=2)
df_sj = pd.read_excel(xls, 'SJ', header=3)

#===========================================
# TIME SERIES
#===========================================

# 2. Data Cleaning & Date Conversion
# PO
df_po['PO_CREATED_ON'] = pd.to_datetime(df_po['PO_CREATED_ON'], dayfirst=True, errors='coerce')
# BP
df_bp['BP_CREATED_ON'] = pd.to_datetime(df_bp['BP_CREATED_ON'], dayfirst=True, errors='coerce')
# SJ
df_sj['SJ_CREATED_ON'] = pd.to_datetime(df_sj['SJ_CREATED_ON'], dayfirst=True, errors='coerce')

# Filter valid dates
df_po = df_po.dropna(subset=['PO_CREATED_ON'])
df_bp = df_bp.dropna(subset=['BP_CREATED_ON'])
df_sj = df_sj.dropna(subset=['SJ_CREATED_ON'])

# 3. Aggregation (Daily)
# We count unique Document Numbers to measure "Activity Volume"
po_daily = df_po.groupby(df_po['PO_CREATED_ON'].dt.to_period('D'))['PO_NO'].nunique().sort_index()
bp_daily = df_bp.groupby(df_bp['BP_CREATED_ON'].dt.to_period('D'))['BP_NO'].nunique().sort_index()
sj_daily = df_sj.groupby(df_sj['SJ_CREATED_ON'].dt.to_period('D'))['SJ_NO'].nunique().sort_index()

# Convert index back to timestamp for plotting
po_daily.index = po_daily.index.to_timestamp()
bp_daily.index = bp_daily.index.to_timestamp()
sj_daily.index = sj_daily.index.to_timestamp()

# 4. Plotting
plt.figure(figsize=(14, 8))

# Subplot 1: PO Trend (Procurement Activity)
plt.subplot(2, 1, 1)
plt.plot(po_daily.index, po_daily.values, marker='o', linestyle='-', color='blue', label='PO Created (Count)')
plt.title('Daily Procurement Activity (PO Created)', fontsize=14)
plt.ylabel('Number of POs')
plt.grid(True, linestyle='--', alpha=0.7)
plt.legend()
# Format x-axis
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d-%b'))
plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=30))

# Subplot 2: Logistics Flow (Inbound BP vs Outbound SJ)
plt.subplot(2, 1, 2)
plt.plot(bp_daily.index, bp_daily.values, marker='.', linestyle='-', color='green', label='Inbound (BP Created)', alpha=0.7)
plt.plot(sj_daily.index, sj_daily.values, marker='.', linestyle='-', color='red', label='Outbound (SJ Created)', alpha=0.7)
plt.title('Logistics Flow: Inbound vs Outbound', fontsize=14)
plt.ylabel('Number of Documents')
plt.grid(True, linestyle='--', alpha=0.7)
plt.legend()
# Format x-axis
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d-%b'))
plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=5))

plt.tight_layout()
# plt.savefig('trend_analysis.png')
plt.show()

# 5. Additional Time Series Insights (Day of Week)
df_sj['DayOfWeek'] = df_sj['SJ_CREATED_ON'].dt.day_name()
sj_dow = df_sj.groupby('DayOfWeek')['SJ_NO'].nunique().reindex(
    ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
)

# Print insights
print("--- TIME SERIES INSIGHTS ---")
print("\n1. Peak PO Dates (Top 3 Days):")
print(po_daily.sort_values(ascending=False).head(3))
print("\n2. Peak Shipping (SJ) Dates (Top 3 Days):")
print(sj_daily.sort_values(ascending=False).head(3))
print("\n3. Shipping Activity by Day of Week (SJ Count):")
print(sj_dow)
print("")


#===========================================
#ANALISIS
#===========================================

# 3. Analisis Lead Time Supplier (PO ke BP)
# Ambil tanggal penerimaan paling awal (first receipt) untuk setiap PO
first_receipt = df_bp.groupby('PO_NO')['BP_CREATED_ON'].min().reset_index()
po_perf = df_po[['PO_NO', 'NAMA_SUPPLIER', 'PO_APPROVED_ON','PO_CREATED_ON']].merge(first_receipt, on='PO_NO', how='inner')

po_perf['BP_CREATED_ON'] = pd.to_datetime(po_perf['BP_CREATED_ON'], errors='coerce')
po_perf['PO_APPROVED_ON'] = pd.to_datetime(po_perf['PO_APPROVED_ON'], errors='coerce')
po_perf['PO_CREATED_ON'] = pd.to_datetime(po_perf['PO_CREATED_ON'], errors='coerce')

# Hitung selisih hari
po_perf['LEAD_TIME_DAYS'] = (po_perf['BP_CREATED_ON'] - po_perf['PO_APPROVED_ON']).dt.days
po_perf = po_perf[(po_perf['LEAD_TIME_DAYS'] >= 0) & (po_perf['LEAD_TIME_DAYS'] < 100)] # Filter error


po_perf['APPROVAL_TIME_DAYS'] = (po_perf['PO_APPROVED_ON'] - po_perf['PO_CREATED_ON']).dt.days
po_perf = po_perf[(po_perf['APPROVAL_TIME_DAYS'] >= 0) & (po_perf['APPROVAL_TIME_DAYS'] < 100)] # Filter error

# Agregasi performa per Supplier
supplier_stats = po_perf.groupby('NAMA_SUPPLIER').agg(
    RATA_RATA_WAKTU=('LEAD_TIME_DAYS', 'mean'),
    TOTAL_PO=('PO_NO', 'nunique')
).sort_values(by='TOTAL_PO', ascending=False)

# 4. Analisis Kecepatan Gudang (BP ke SJ)
# Rata-rata tanggal BP vs Rata-rata tanggal SJ per PO
avg_bp_date = df_bp.groupby('PO_NO')['BP_CREATED_ON'].mean().reset_index()
avg_sj_date = df_sj.groupby('NO_PO')['SJ_CREATED_ON'].mean().reset_index()

warehouse_flow = avg_bp_date.merge(avg_sj_date, left_on='PO_NO', right_on='NO_PO', how='inner')
warehouse_flow['WAKTU_PROSES'] = (warehouse_flow['SJ_CREATED_ON'] - warehouse_flow['BP_CREATED_ON']).dt.days
warehouse_flow = warehouse_flow[warehouse_flow['WAKTU_PROSES'] >= 0]

# --- Analysis 3: Item Demand (Top Items) ---
top_items = df_po.groupby('NAMA_BARANG')['JUMLAH'].sum().sort_values(ascending=False).head(10)

# 5. Output Hasil
print(f"Average Approval Time  (created PO -> approval PO): {po_perf['LEAD_TIME_DAYS'].mean():.2f} hari")
print(f"Average Waktu Proses Gudang (created BP -> created SJ): {warehouse_flow['WAKTU_PROSES'].mean():.2f} hari")

print("\nTop 5 Supplier Teraktif by PO Volume & Their Speed:")
print(supplier_stats.head(5))

print("\nTop 5 Most Ordered Items (by Quantity in PO):")
print(top_items.head(5))

print("\nTop 5 Kapal Paling Sering Disupply by Frequency of Supply (SJ count):")
print(df_sj['NAMA_KAPAL'].value_counts().head(5))

#======================================
# Numeric conversions
df_po['JUMLAH'] = pd.to_numeric(df_po['JUMLAH'], errors='coerce').fillna(0)
df_bp['JML_DITERIMA'] = pd.to_numeric(df_bp['JML_DITERIMA'], errors='coerce').fillna(0)
df_sj['JUMLAH'] = pd.to_numeric(df_sj['JUMLAH'], errors='coerce').fillna(0)

# --- Insight 1: Supplier Fulfillment Rate (Reliability) ---
# Aggregate Target Qty per PO
po_target = df_po.groupby(['PO_NO', 'NAMA_SUPPLIER'])['JUMLAH'].sum().reset_index()
# Aggregate Received Qty per PO
bp_actual = df_bp.groupby('PO_NO')['JML_DITERIMA'].sum().reset_index()

# Merge
fulfillment = po_target.merge(bp_actual, on='PO_NO', how='left')
fulfillment['JML_DITERIMA'] = fulfillment['JML_DITERIMA'].fillna(0)
fulfillment['FILL_RATE'] = (fulfillment['JML_DITERIMA'] / fulfillment['JUMLAH']) * 100
# Cap at 100% (handle over-delivery anomalies for metric purity)
fulfillment['FILL_RATE_CAPPED'] = fulfillment['FILL_RATE'].apply(lambda x: 100 if x > 100 else x)

# Avg Fill Rate per Supplier (min 5 POs to be significant)
supplier_fulfillment = fulfillment.groupby('NAMA_SUPPLIER').agg(
    AVG_FILL_RATE=('FILL_RATE_CAPPED', 'mean'),
    TOTAL_PO=('PO_NO', 'nunique')
).query('TOTAL_PO >= 5').sort_values(by='AVG_FILL_RATE', ascending=True)

# --- Insight 2: Unfulfilled POs (The "Black Hole" Orders) ---
# POs with 0 received quantity
unfulfilled = fulfillment[fulfillment['JML_DITERIMA'] == 0]
unfulfilled_count = unfulfilled.shape[0]
top_unfulfilled_suppliers = unfulfilled['NAMA_SUPPLIER'].value_counts()

# --- Insight 3: Consumption Profiling (What are they fixing?) ---
# Group SJ by Category and Vessel
# Clean Category Names first (remove nan)
df_sj = df_sj.dropna(subset=['NAMAKODEAKUN_BARANG'])

# Filter Top 5 Vessels
top_vessels_list = df_sj['NAMA_KAPAL'].value_counts().head(5).index
df_sj_top = df_sj[df_sj['NAMA_KAPAL'].isin(top_vessels_list)]

# Pivot for Heatmap
vessel_category_pivot = df_sj_top.groupby(['NAMAKODEAKUN_BARANG', 'NAMA_KAPAL'])['JUMLAH'].sum().unstack(fill_value=0)

# --- Insight 4: Sparepart Specifics ---
# Check specific keywords in Item Name for top vessels
def classify_item(name):
    name = str(name).lower()
    if 'filter' in name: return 'Filters'
    if 'oli' in name or 'lubricant' in name: return 'Lubricants'
    if 'bearing' in name: return 'Bearings'
    if 'seal' in name or 'ring' in name: return 'Seals/Gaskets'
    if 'baut' in name or 'bolt' in name: return 'Bolts/Nuts'
    return 'Others'

df_sj_top['ITEM_CLASS'] = df_sj_top['NAMABRG'].apply(classify_item)
item_class_pivot = df_sj_top.groupby(['ITEM_CLASS', 'NAMA_KAPAL'])['JUMLAH'].sum().unstack(fill_value=0)


# --- Visualizations ---
fig, axes = plt.subplots(2, 1, figsize=(12, 12))

# Plot 1: Supplier Fulfillment (Bottom 10 - The "Problem" Suppliers)
# We plot the worst performers to highlight risk
bottom_suppliers = supplier_fulfillment.head(10)
bars = axes[0].barh(bottom_suppliers.index, bottom_suppliers['AVG_FILL_RATE'], color='salmon')
axes[0].set_title('Top 10 Suppliers with Lowest Fulfillment Rate (Partial/No Delivery)', fontsize=14)
axes[0].set_xlabel('Average Fill Rate (%)')
axes[0].axvline(x=100, color='grey', linestyle='--', alpha=0.5)
axes[0].bar_label(bars, fmt='%.1f%%')

# Plot 2: Vessel Consumption Profile (Stacked Bar)
# Normalize to see percentage distribution
vessel_category_pct = vessel_category_pivot.div(vessel_category_pivot.sum(axis=0), axis=1) * 100
vessel_category_pct.T.plot(kind='bar', stacked=True, ax=axes[1], colormap='tab20')
axes[1].set_title('Consumption Profile: Spending Category per Top Vessel', fontsize=14)
axes[1].set_ylabel('Percentage of Items Shipped (%)')
axes[1].legend(title='Category', bbox_to_anchor=(1.05, 1), loc='upper left')

plt.tight_layout()
plt.savefig('deep_insights.png')

# Output Text
print("--- INSIGHTS LANJUTAN ---")
print(f"1. Total PO yang belum diterima sama sekali (0% Delivery): {unfulfilled_count} PO")
print("\n2. Top Supplier dengan PO 'Macet' (Unfulfilled):")
print(top_unfulfilled_suppliers)
print(top_unfulfilled_suppliers.count())

print("\n3. Profil Konsumsi Sparepart (Top 5 Kapal):")
print(item_class_pivot)