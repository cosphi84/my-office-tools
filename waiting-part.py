'''
Waiting Part checker
Python script untuk menganalisis LR status 30 vs Reservasi vs DO list.

Sumber data:
1. status30.XLS - Daftar job LR status 30 waiitng part dari SAP.
2. reservasi.XLS  - Data Reservasi part dari SAP.
3. do-list.XLS - Data list DO dari SAP
'''

import pandas as pd
from pathlib import Path
import os


# check file dan pastikan ada
file_30 = Path("source/status30.XLS")
file_res = Path("source/reservasi.XLS")
file_DO = Path("source/do-list.XLS")

if not file_30.exists:
    print('File source status30.XLS tidak ada!')
    raise SystemExit

if not file_res.exists:
    print("File Reservasi.XLS tidak ditemukan!")
    raise SystemExit

if not file_DO.exists:
    print("File do-list.XLS tidak ada!")
    raise SystemExit
stt_30_columns = ["Notifictn", "Notif.date", "Mn.wk.ctr","List name", "Addit. device data"]
res_columns = ["Reserv.No", "Item", "Material No.", "Reqmt Qty", "RcvSloc", "Base Date", "Recipient", "Text"]

df_30 = pd.read_csv(file_30, dialect="excel-tab", encoding="utf_16", skiprows=3, usecols=stt_30_columns)
df_do = pd.read_csv(file_DO, encoding="utf_16", dialect="excel-tab", skiprows=3)
df_rs = pd.read_csv(file_res, encoding="utf_16", dialect="excel-tab", skiprows=3, usecols=res_columns)

# Convert some stuff
df_30["Notif.date"] = pd.to_datetime(df_30["Notif.date"], errors="coerce", format="%d.%m.%Y")
df_rs["Base Date"] = pd.to_datetime(df_rs["Base Date"], errors="coerce", format="%d.%m.%Y")
# Remove Refurbish
df_rs = df_rs[~df_rs["RcvSloc"].str.endswith('91')]