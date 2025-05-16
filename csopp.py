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

warnings.simplefilter(action="ignore", category=FutureWarning)


# File path constants
FILE_CONFIG = Path("Config/csopp.xlsx")
FILE_OTS = Path('DataSource/ots.XLS')
FILE_COMPLETED = Path("DataSource/completed.XLS")
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

# Tampilkan konfigurasi
print(config_apps)