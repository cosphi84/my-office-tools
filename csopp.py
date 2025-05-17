'''
csopp.py
Tools untuk melakukan ekstraksi pencapaian pekerjan, berdasarkan target KPI CS
Sumber data:
1. Ots.XLS : File csv job orders status Outstanding
2. Completed.XLS : File csv completed data
'''
import pandas as pd
import numpy as np
import os
from pathlib import Path
from tkinter import messagebox
import warnings
from datetime import datetime
from functools import reduce


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

def load_config():
    """Muat konfigurasi dari file Excel dan kembalikan sebagai dictionary."""
    config_excel = pd.ExcelFile(FILE_CONFIG)
    config = {}
    for sheet in config_excel.sheet_names:
        if sheet.strip().lower() == "baca saya dulu":
            continue
        config[sheet] = config_excel.parse(sheet)
    return config

def extract_setting(item: str):
    cfg_sheet = load_config()
    setting_records = cfg_sheet.to_dict(orient='records')
    if item not in cfg_sheet:
        raise SystemExit("Sheet 'seting' tidak ditemukan dalam file konfigurasi.")
    
    return {row.get(item): row.get('Value') for row in setting_records}

def load_source(path: Path):
    cols = ["Typ", "Notifctn", "Notif.date", "Req. start", "Req. End", "Changed on", "Completion", "PG",  "Mn.wk.ctr", "UserStatus", "List name", "Street", "Telephone", "Material", "Serial number", "Description", "Addit. device data"]
    isOts = True if path.name == 'ots.csv' else False
    sekarang = datetime.now()

    # Load beberapa config data
    cfg_excel = pd.ExcelFile(FILE_CONFIG)
    with cfg_excel as xls:
        pg = pd.read_excel(xls, "pg")
        mwc = pd.read_excel(xls, "mwc").rename(columns={'Kode':'Mn.wk.ctr', "Teknisi": "Work Center"})
        teknisi = pd.read_excel(xls, "teknisi", usecols=["id", "name"]).rename(columns={"id": "idCSMS", "name":"Teknisi"})
        excludeMwc = pd.read_excel(xls, "exclude", usecols=["mwc"])
    
    df = pd.read_csv(path, usecols=cols, dialect="excel-tab", encoding="utf_16", skiprows=3)
    # exclude Type LR
    df = df[~df["Typ"].isin(["Z8", "ZZ"])]

    # Convert some stuff
    df["Notif.date"] = pd.to_datetime(df["Notif.date"], format="%d.%m.%Y", errors="coerce")
    df["Completion"] = pd.to_datetime(df["Completion"], errors="coerce", format="%d.%m.%Y")
    df["Changed on"] = pd.to_datetime(df["Changed on"], format="%d.%m.%Y", errors="coerce")
    df["UserStatus"] = pd.to_numeric(df["UserStatus"], errors="coerce")
    df["PG"] = pd.to_numeric(df["PG"], errors="coerce")
    df["Mn.wk.ctr"] = df["Mn.wk.ctr"].astype(str)
    df["bulan_complet"] = df["Completion"].fillna(sekarang).values.astype("datetime64[M]")
    df["RcvdThisMo"] = (df["Notif.date"] >= df["bulan_complet"]).astype(int)
    df.drop(columns=["bulan_complet"], inplace=True)

    df["Req. End"] = pd.to_datetime(df["Req. End"], errors="coerce", format="%d.%m.%Y")

    # add some stuff    
    if isOts :
        df["isOTS"] = 1
        df["Req. End"] = df["Req. End"].fillna(sekarang)
        df["e_no_req_end"] = 0
    else:
        df["isOTS"] = 0
        df["e_no_req_end"] = df['Req. End'].isna().astype(int)
        df["Req. End"] = df["Req. End"].fillna(sekarang)
        df = df[~df["Mn.wk.ctr"].isin(excludeMwc)]
    
    df = df[~df["UserStatus"].isin([51, 52])]
    df["Lo"] = (df["Req. End"] - df["Notif.date"]).dt.days
    df["1D"] = (df["Lo"] <= 1).astype(int)
    df["1W"] = (df["Lo"] <= 7).astype(int)
    df["Cashless"] = df["UserStatus"].isin([94, 95, 96, 98]).astype(int)
    # CDR = Cabang, SDSS, SSR
    cdr_kriteria = [
        df["Mn.wk.ctr"].str.startswith("ST"),
        df["Mn.wk.ctr"].str.startswith("SR"),
    ]
    cdr = ["SDSS", "SSR"]
    df["CDR"] = np.select(cdr_kriteria, cdr, default="Cabang")
    df["idCSMS"] = df["Addit. device data"].fillna('').str.split(r"[;:/ ]").str[0]

    

    pg["PG"] = pd.to_numeric(pg["PG"], errors="coerce")
    mwc["Mn.wk.ctr"] = mwc["Mn.wk.ctr"].astype(str)
    df = df.merge(pg, how="left", on="PG")
    df = df.merge(mwc, how="left", on="Mn.wk.ctr")
    df = df.merge(teknisi, how="left", on="idCSMS")

    return df


def proses_data(dfData: pd.DataFrame):
    pv = pd.pivot_table(dfData, index=["Regional", "Cabang"], values=["Notifctn"], aggfunc="count", fill_value=0).rename(columns={"Notifctn": "ALL LR"})
    
    # pecah lagi data OTS & data Complete
    dfOTS = dfData[dfData["isOTS"].isin([1])]
    dfComplete = dfData[dfData["isOTS"].isin([0])]
    dfCash = dfData[dfData["UserStatus"].isin([93])]
    dfCash = dfCash[~dfCash["Typ"].isin(["ZX"])]

    # Pivoting
    pv_cmplt = pd.pivot_table(dfComplete, index=["Regional", "Cabang"], values=["Notifctn"], aggfunc="count", fill_value=0).rename(columns={"Notifctn": "CMPLT"})
    pv_ots = pd.pivot_table(dfOTS, index=["Regional", "Cabang"], values=["Notifctn"], aggfunc="count", fill_value=0).rename(columns={"Notifctn": "OTS"})
    pv_1D = pd.pivot_table(dfComplete, index=["Regional", "Cabang"], values=["1D"], aggfunc="sum", fill_value=0)
    pv_1W = pd.pivot_table(dfComplete, index=["Regional", "Cabang"], values=["1W"], aggfunc="sum", fill_value=0)
    pv_cash = pd.pivot_table(dfCash, index=["Regional", "Cabang"], values=["Notifctn"], aggfunc="count", fill_value=0).rename(columns={"Notifctn":"Cash"})
    pv_cashless = pd.pivot_table(dfComplete, index=["Regional", "Cabang"], values=["Cashless"], aggfunc="sum", fill_value=0)

    pv = reduce(lambda left, right: pd.merge(left, right, on=["Regional", "Cabang"], how="left"), [pv, pv_ots, pv_cmplt, pv_1D, pv_1W, pv_cash, pv_cashless]).fillna(0)

    pv["Rasio Cmplt"] = (pv["CMPLT"] / pv["ALL LR"])*100
    pv["Rasio 1D"] = (pv["1D"] / pv["CMPLT"]) * 100
    pv["Rasio 1W"] = (pv["1W"] / pv["CMPLT"]) * 100
    pv["Rasio Cassless"] = (pv["Cashless"] / ( pv["Cashless"] + pv["Cash"])) *100
    return pv

# --- Proses Utama ---

# Cek file konfigurasi wajib
check_file_exists(FILE_CONFIG)

# Cek file sumber lainnya 
check_file_exists(FILE_OTS)
check_file_exists(FILE_COMPLETED)

# Hapus file hasil jika sudah ada
for file in [FILE_RESULT, FILE_DETAIL, FILE_ERROR]:
    safe_remove(file)

# Load Source file & merge
df_ots = load_source(FILE_OTS)
df_completed = load_source(FILE_COMPLETED)
source = pd.concat([df_ots, df_completed])

# Prosess data
result = proses_data(source)

result.to_excel(FILE_RESULT)