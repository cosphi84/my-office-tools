'''
csopp.py
Tools untuk melakukan ekstraksi pencapaian pekerjan, berdasarkan target KPI CS
Sumber data:
1. Ots.XLS : File csv job orders status Outstanding
2. Completed.XLS : File csv completed data
'''
import pandas as pd
import os
from pathlib import Path
from tkinter import messagebox
import warnings
from datetime import datetime

# Disable deprecated warning
warnings.simplefilter(action="ignore", category=FutureWarning)
warnings.simplefilter(action="ignore", category=UserWarning)


# File path constants
FILE_CONFIG = Path("Config/csopp.xlsx")
FILE_OTS = Path('DataSource/ots.csv')
FILE_COMPLETED = Path("DataSource/completed.csv")
FILE_RESULT = Path("Result/Result.xlsx")
FILE_DETAIL = Path("Result/Detail.xlsx")
FILE_ERROR = Path("Result/Error.xlsx")
FILE_TEKNISI = Path("Result/JadwalTeknisi.xlsx")

def check_file_exists(path: Path, required=True):
    """Cek apakah file ada, dan hentikan program jika diperlukan."""
    if not path.exists():
        msg = f"File '{path}' tidak ditemukan!"
        print(msg)
        if required:
            raise SystemExit(msg)

def safe_remove(path: Path):
    """Hapus file jika ada, dengan penanganan error."""
    if path.exists():
        try:
            os.remove(path)
            print(f"File {path} dihapus.")
        except Exception as e:
            msg = f"Gagal menghapus file {path}: {e}"
            print(msg)
            raise SystemExit(msg)

def load_config(path: Path):
    """Muat konfigurasi dari file Excel dan kembalikan sebagai dictionary."""
    config_excel = pd.ExcelFile(path)
    config = {}
    for sheet in config_excel.sheet_names:
        if sheet.strip().lower() == "baca saya dulu":
            continue
        config[sheet] = config_excel.parse(sheet)
    return config

def extract_setting(config_sheet):
    """Ekstrak sheet 'seting' menjadi dictionary key-value."""
    setting_records = config_sheet.to_dict(orient='records')
    return {row.get('Seting'): row.get('Value') for row in setting_records}

def load_source(path: Path):
    cols = ["Typ", "Notifctn", "Notif.date", "Time", "Req. start", "Req. End", "Changed on", "Completion", "PG",  "Mn.wk.ctr", "UserStatus", "List name", "Street", "Telephone", "Material", "Serial number", "Description", "Addit. device data"]
    isOts = True if path.name == 'ots.csv' else False
    sekarang = datetime.now()
    
    df = pd.read_csv(path, usecols=cols, dialect="excel-tab", encoding="utf_16", skiprows=3)
    # exclude Type LR
    df = df[~df["Typ"].isin(["Z8", "ZZ"])]

    # Convert some stuff
    df["Notif.date"] = pd.to_datetime(df["Notif.date"] + " " + df["Time"], format="%d.%m.%Y %H:%M:%S", errors="coerce")
    df["Completion"] = pd.to_datetime(df["Completion"], errors="coerce", format="%d.%m.%Y")
    df["Changed on"] = pd.to_datetime(df["Changed on"], format="%d.%m.%Y", errors="coerce")
    df["UserStatus"] = pd.to_numeric(df["UserStatus"], errors="coerce")
    
    df.drop(columns=["Time"], inplace=True)

    # add some stuff    
    if isOts :
        df["isOTS"] = 1
        df["Req. End"] = sekarang
        df["e_no_req_end"] = 0
    else:
        df["isOTS"] = 0
        df = df[~df["UserStatus"].isin([51, 52])]
        df["Req. End"] = pd.to_datetime(df["Req. End"], errors="coerce", format="%d.%m.%Y")
        df["e_no_req_end"] = df['Req. End'].isna().astype(int)
        df["Req. End"] = df["Req. End"].fillna(sekarang)
        
    
    df["Lo"] = (df["Req. End"] - df["Notif.date"]).dt.days
    return df


# --- Proses Utama ---

# Cek file konfigurasi wajib
check_file_exists(FILE_CONFIG)

# Cek file sumber lainnya 
check_file_exists(FILE_OTS)
check_file_exists(FILE_COMPLETED)

# Hapus file hasil jika sudah ada (hapus komentar jika diperlukan)
# for file in [FILE_RESULT, FILE_DETAIL, FILE_ERROR]:
#     safe_remove(file)

# Load dan parsing konfigurasi
config_data = load_config(FILE_CONFIG)
if 'seting' not in config_data:
    raise SystemExit("Sheet 'seting' tidak ditemukan dalam file konfigurasi.")

config_apps = extract_setting(config_data['seting'])

df_ots = load_source(FILE_OTS)
df_completed = load_source(FILE_COMPLETED)
source = pd.concat([df_ots, df_completed])
source.to_excel(FILE_RESULT)