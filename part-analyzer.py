"""
Part-Rank Calculator
Python script untuk menganalisis dan mengelompokkan pemakaian sparepart.

Sumber data:
1. master.XLS - Daftar master part dari SAP.
2. usage.XLS  - Data pemakaian dari SAP MB51 (Material Document List).

Pastikan:
- Storage location: "All Storage exclude disposal"
- Movement type: 201, 202, 261, 262, 601, 602, 937, 938
- Periode: bulanan, tahunan, atau sesuai kebutuhan
"""

import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta


# Fungsi agregasi khusus: jumlah negatif dijadikan positif, positif jadi 0
def custom_sum(series):
    total = series.sum()
    return 0 if total > 0 else abs(total)

# STEP 1: Persiapan Data
used_cols = ['Material', 'SLoc', 'MvT', 'Quantity', 'Pstng Date']
master_cols = ['Material', 'Material Description', 'Basic material', 'Created On']

# Import data master dan usage
master_df = pd.read_csv("sumberdata/master.XLS", encoding="utf_16", dialect="excel-tab", skiprows=3, usecols=master_cols)
usage_df = pd.read_csv("sumberdata/usage.XLS", encoding="utf_16", dialect="excel-tab", skiprows=3, usecols=used_cols)

# Konversi kolom tanggal & jumlah
usage_df['Pstng Date'] = pd.to_datetime(usage_df['Pstng Date'], errors="coerce", format="%d.%m.%Y")
usage_df["Quantity"] = pd.to_numeric(usage_df["Quantity"], errors="coerce")
master_df["Created On"] = pd.to_datetime(master_df["Created On"], errors="coerce", format="%d.%m.%Y")

# Tambahkan kolom tahun dan bulan
usage_df["Year"] = usage_df["Pstng Date"].dt.year
usage_df["Month"] = usage_df["Pstng Date"].dt.strftime("%m")

# 12 Bulan lalu
today = pd.Timestamp.today()
twelve_months_ago = today - relativedelta(months=12)


# STEP 2: Buat Tabel Utama

# Pivot data pemakaian
pivot_usage = pd.pivot_table(
    usage_df,
    index="Material",
    columns=["Year", "Month"],
    values="Quantity",
    aggfunc=custom_sum,
    fill_value=0
)

# Buat nama kolom menjadi 'YYYY_MM'
pivot_usage.columns = [f"{year}_{month}" for (year, month) in pivot_usage.columns]

# Gabungkan data master dengan usage
master_table = pd.merge(master_df, pivot_usage, how="inner", on="Material")

# Hitung jumlah kolom bulan yang tersedia
usage_columns = pivot_usage.columns
last_6_months = 6 # Part Rank ~ C
last_12_months = 12 # Part Rank ~ F
last_24_months = 24 # Part Rank ~ G

if len(usage_columns) < last_24_months:
    print("Analisa pemakaian harus minimal 2 Tahun data.\nData yang tersedia masih kurang!")
    raise SystemExit


# Ambil rata-rata pemakaian 6, 12, 24 bulan terakhir
master_table["Part Usage 6 Mo"] = master_table[usage_columns[-last_6_months:]].mean(axis=1)

# Hitung jumlah bulan dalam 6,12,24 bulan terakhir yang ada pemakaian (> 0)
master_table["Part Move in 6 Mo"] = (master_table[usage_columns[-last_6_months:]] > 0).values.sum(axis=1)
master_table["Part Move in 12 Mo"] = (master_table[usage_columns[-last_12_months:-last_6_months]] > 0).values.sum(axis=1)
master_table["Part Move in 24 Mo"] = (master_table[usage_columns[-last_24_months:-last_12_months]] > 0).values.sum(axis=1)

# Tentukan ranking berdasarkan aturan logika
ranks_conditions = [
    # N = Semua part yang create on nya dibawah 12 bulan
    (master_table["Created On"] >= twelve_months_ago),   # Kategori N: part baru (< 12 bulan)
    # S1 = usia part > 12 bulan, Pemakian selama 6 bulan terakhir > 50, dan setiap bulan ada penggunaan
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Usage 6 Mo"] >= 50) & (master_table["Part Move in 6 Mo"] == 6),
    # S2 = usia part > 12 bulan, Pemakian selama 6 bulan terakhir < 50 , dan setiap bulan ada penggunaan
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Usage 6 Mo"] < 50) & (master_table["Part Move in 6 Mo"] == 6),
    # A1 = usia part > 12 bulan, Pemakian selama 6 bulan terakhir>= 50 , dan ada penggunaan 5x
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Usage 6 Mo"] >= 50) & (master_table["Part Move in 6 Mo"] == 5),
    # A2 = usia part > 12 bulan, Pemakian 6 bulan terakhir < 50 tapi > 6, dan ada penggunaan 5x
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Usage 6 Mo"].between(7, 49)) & (master_table["Part Move in 6 Mo"] == 5),
    # A3 = usia part > 12 bulan, Pemakian selama 6 bulan terakhir <= 6, dan ada penggunaan 5x
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Usage 6 Mo"] < 6) & (master_table["Part Move in 6 Mo"] == 5),
    # B1 = usia part > 12 bulan, Pemakian selama 6 bulan terakhir ada penggunaan 4x 
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Move in 6 Mo"] == 4),
    # B2 = usia part > 12 bulan, Pemakian selama 6 bulan terakhir ada penggunaan 3x 
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Move in 6 Mo"] == 3),
    # B3 = usia part > 12 bulan, Pemakian selama 6 bulan terakhir ada penggunaan 2x 
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Move in 6 Mo"] == 2),
    # C = usia part > 12 bulan, Pemakian selama 6 bulan terakhir ada penggunaan 1x 
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Move in 6 Mo"] == 1),
    # D = usia part > 12 bulan, Pemakian selama 12 bulan terakhir ada penggunaan 1x 
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Move in 12 Mo"] >= 1),
    # E = usia part > 12 bulan, Pemakian selama 24 bulan terakhir ada penggunaan 1x 
    (master_table["Created On"] <= twelve_months_ago & master_table["Part Move in 24 Mo"] >= 1),
]

ranks_labels = ['N', 'S1', 'S2', 'A1', 'A2', 'A3', 'B1', 'B2', 'B3', 'C', 'D', 'E']

master_table["Part Ranks"] = np.select(ranks_conditions, ranks_labels, default="G")

# hapus kolom temporari
master_table = master_table.drop(columns=["Part Usage 6 Mo", "Part Move in 6 Mo", "Part Move in 12 Mo", "Part Move in 24 Mo"])

# Output akhir
master_table.to_excel("PartRank.xlsx")
