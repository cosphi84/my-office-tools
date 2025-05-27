import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from functools import reduce
from typing import Dict
from Config.config import csopp_config

def fill_data(data: Dict[str, pd.DataFrame]) -> dict:
    '''
    fill_data: modifikasi dataframe

    Args:
        data Dict Dict data, harus berisi 'ots' => dataframe OTS, 'complited' => dataframe complited

    Returns:
        dict Dataframe OTS dan COMPLETED yang sudah siap
    '''
    sekarang = datetime.now()
    
    result_df = {}
    df_error = {}

    for tabel, df in data.items():
        df = df.copy()  # menghindari SettingWithCopyWarning di awal loop

        # Konversi dan isian awal
        df["bulan_complet"] = pd.to_datetime(df["Completn date"].fillna(sekarang)).values.astype("datetime64[M]")
        df["RcvdThisMo"] = (df["Notif.date"] >= df["bulan_complet"]).astype(int)
        df.drop(columns=["bulan_complet"], inplace=True)

        if tabel == 'ots':
            df["Required End"] = pd.to_datetime(df["Required End"].fillna(sekarang))
            df["Lo"] = (df["Required End"] - df["Notif.date"]).dt.days
        else:
            # Simpan baris error ke dict
            df_error["no_req_end"] = df[df["Required End"].isna()]

            # Filter baris yang memiliki Required End, lalu salin aman
            df = df[~df["Required End"].isna()]

            # Konversi tanggal pastikan konsisten
            df["Required End"] = pd.to_datetime(df["Required End"])
            df["Lo"] = (df["Required End"] - df["Notif.date"]).dt.days
            df["1D"] = (df["Lo"] <= 1).astype(int)
            df["1W"] = (df["Lo"] <= 7).astype(int)

            df["Cashless"] = df["User status"].isin(csopp_config().get("csopp_cashless_status", [])).astype(int)
            df["Cash"] = df["User status"].isin(csopp_config().get("csopp_cash_status", [])).astype(int)

            # Filter Main WorkCtr, lalu salin ulang agar aman
            exclude_mwc = [item[0] for item in csopp_config().get('csopp_exclude_mwc', [])]
            df = df[~df["Main WorkCtr"].isin(exclude_mwc)]
        
            # Hanya bulan ini kompletion datenya
            startDate = sekarang.replace(day=1)
            df = df[df["Completn date"] >= startDate]

        # Penentuan CDR berdasarkan prefix kode Main WorkCtr
        cdr_kriteria = [
            df["Main WorkCtr"].str.startswith("ST"),
            df["Main WorkCtr"].str.startswith("SR"),
            df["Main WorkCtr"].str.startswith("OT"),
        ]
        cdr = ["SDSS", "SSR", "OVJ"]
        df["CDR"] = np.select(cdr_kriteria, cdr, default="Cabang")

        # Ekstraksi idCSMS dari kolom Device data
        df["idCSMS"] = df["Device data"].fillna('').str.split(r"[;:/ ]").str[0]

        # Load konfigurasi tambahan
        pg = csopp_config().get('csopp_pg', pd.DataFrame())        
        mwc = csopp_config().get('csopp_mwc', pd.DataFrame())
        teknisi = csopp_config().get('csopp_idcsms', pd.DataFrame())

        # Join tambahan - aktifkan jika dibutuhkan
        if isinstance(pg, pd.DataFrame) and not pg.empty and "Planner group" in pg.columns:
            pg["Planner group"] = pg["Planner group"].astype(str)
        if isinstance(mwc, pd.DataFrame) and not mwc.empty and "Main WorkCtr" in mwc.columns:
            mwc["Main WorkCtr"] = mwc["Main WorkCtr"].astype(str)
        

        df = df.merge(pg, how="left", on="Planner group").fillna("")
        df = df.merge(mwc, how="left", on="Main WorkCtr").fillna("")
        df = df.merge(teknisi, how="left", on="idCSMS").fillna("")

        result_df[tabel] = df

    return result_df

def apply_filter(dfData: Dict[str, pd.DataFrame], CDR: str = "Cabang") -> Dict[str, pd.DataFrame]:
    dfData = dfData.copy()
    cfg = csopp_config().get('csopp_setting', {})

    for tipe, df in dfData.items():
        df = df.copy()

        # Filter berdasarkan 'Regional' dan 'Cabang'
        for col in ["Regional", "Cabang"]:
            filter_val = cfg.get(col, "All")
            if filter_val != "All" and col in df.columns:
                df = df[df[col] == filter_val]

        # Filter berdasarkan 'CDR'
        if CDR != "All" and "CDR" in df.columns:
            df = df[df["CDR"] == CDR]

        dfData[tipe] = df

    return dfData
    
def get_error_notif(dfData: Dict[str, pd.DataFrame])->dict:
    '''
    Fungsi untuk memisahkan error notif dari sumber data mentah

    Parameters:
        - dfData: Dict dataframe sumber
    
    Returns:
        - dict data error dan data bersih
    '''
    returned_data = {'OK': {}, 'error': {}}
    for label, df in dfData.items():
        df = df.copy()

        # Notif masih di xx23 (belum ada penugasan teknisi)
        returned_data['error'][label] = {'xx23': df[df['Main WorkCtr'].str.endswith('23')]}
        # Error Notif status completed
        if label == 'completed':
            # Buang LR error degan MWC masih di xx23
            df = df[~df["Main WorkCtr"].str.endswith('23')]

            # Status < 90 
            returned_data['error'][label] = {'status': df[df['User status'] < 90]}
            df = df[~(df['User status'] < 90)]
            
            # Required end kosong
            returned_data['error'][label] = {'no_req_end': df[df['Required End'].isna()]}
            df = df[~df["Required End"].isna()]

            # Required end lebih awal dari req start
            returned_data['error'][label] = {'min_req_end': df[df['Required End'] < df['Notif.date']]}
            df = df[~(df["Required End"] < df["Notif.date"])]
        
        returned_data["OK"][label] = df

    return returned_data

def calc_achivement(dfData: Dict[str, pd.DataFrame], GlobalResult : bool = False, cdr: str = "All")->pd.DataFrame:
    dfData = dfData.copy()
    idx = ["Regional"] if GlobalResult == True else ["Regional", "Cabang"]
    if GlobalResult :
        idx = ["Regional"]
    elif not GlobalResult and cdr == 'Cabang':
        idx = ["Regional", "Cabang"]
    else:
        idx = ["Regional", "Work Center"]

    dfData = apply_filter(dfData, CDR=cdr)
    
    dfOTS = dfData.get('ots', pd.DataFrame())
    dfComplete = dfData.get('completed', pd.DataFrame())
    
    # pecah lagi data OTS & data Complete
    df30 = dfOTS[dfOTS["User status"].isin([30])]
    dfLo = dfOTS[dfOTS["Lo"] >= 60]
    dfCash = dfComplete[dfComplete["User status"].isin(csopp_config().get('csopp_cash_status', []))]
    # Abaikan type ZX untuk perhitungan Cashless
    dfCash = dfCash[~dfCash["Notifictn type"].isin(csopp_config().get('csopp_exclude_cash', []))]

    # Pivoting
    pv_ots = pd.pivot_table(dfOTS, index=idx, values=["Notification"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notification": "OTS"})
    pv_cmplt = pd.pivot_table(dfComplete, index=idx, values=["Notification"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notification": "Komplit"})
    pv_1D = pd.pivot_table(dfComplete, index=idx, values=["1D"], aggfunc="sum", fill_value=0, margins=True).rename(columns={"1D": "1 Day"})
    pv_1W = pd.pivot_table(dfComplete, index=idx, values=["1W"], aggfunc="sum", fill_value=0, margins=True).rename(columns={"1W": "1 Week"})
    pv_cash = pd.pivot_table(dfCash, index=idx, values=["Notification"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notification":"Cash"})
    pv_cashless = pd.pivot_table(dfComplete, index=idx, values=["Cashless"], aggfunc="sum", fill_value=0, margins=True)
    pv_30 = pd.pivot_table(df30, index=idx, values=["Notification"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notification": "STT 30"})
    pv_Lo = pd.pivot_table(dfLo, index=idx, values=["Notification"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notification": "LO"})
    pv_TAT = pd.pivot_table(dfComplete, index=idx, values=["Lo"], aggfunc="mean", fill_value=0, margins=True).rename(columns={"Lo": "TAT"})
    pv = reduce(lambda left, right: pd.merge(left, right, on=idx, how="left"), [pv_ots, pv_cmplt, pv_1D, pv_1W, pv_cash, pv_cashless, pv_30, pv_Lo, pv_TAT]).fillna(0)

    pv["Total LK"] = pv["OTS"] + pv["Komplit"]
    pv["Cplt Ratio"] = (pv["Komplit"] / pv["Total LK"])
    pv["LO Ratio"] = (pv["LO"] / pv["OTS"])
    pv["1D Ratio"] = (pv["1 Day"] / pv["Komplit"]) 
    pv["1W Ratio"] = (pv["1 Week"] / pv["Komplit"]) 
    pv["Cashless Ratio"] = (pv["Cashless"] / ( pv["Cashless"] + pv["Cash"]))
    pv["STT 30 VS OTS"] = (pv["STT 30"] / pv["OTS"])

    pv = pv[['OTS', 'STT 30', 'LO', 'Komplit', 'Total LK', 'TAT', '1 Day', '1 Week', 'Cash', 'Cashless', 'Cplt Ratio', '1D Ratio', '1W Ratio', 'Cashless Ratio', 'LO Ratio', 'STT 30 VS OTS']]
    
    return pv

def calc_productivity(dfData: Dict[str, pd.DataFrame], byMWC : bool = True, cdr:str= "Cabang")->pd.DataFrame:
    dfData = dfData.copy()
    idx = ["Regional","Cabang", "Work Center"] if byMWC == True else ["Regional","Cabang", "Main WorkCtr", "Teknisi"]   
    config = csopp_config().get('csopp_setting')
    hari = config.get('Hari')
    
    dfData = apply_filter(dfData, cdr)
    
    dfComplete = dfData.get('completed', pd.DataFrame())
    
    # pecah lagi data OTS & data Complete
    dfCash = dfComplete[dfComplete["User status"].isin(csopp_config().get('csopp_cash_status', []))]
    # Abaikan type ZX untuk perhitungan Cashless
    dfCash = dfCash[~dfCash["Notifictn type"].isin(csopp_config().get('csopp_exclude_cash', []))]

    # Pivoting
    pv_cmplt = pd.pivot_table(dfComplete, index=idx, values=["Notification"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notification": "Komplit"})
    pv_1D = pd.pivot_table(dfComplete, index=idx, values=["1D"], aggfunc="sum", fill_value=0, margins=True).rename(columns={"1D": "1 Day"})
    pv_1W = pd.pivot_table(dfComplete, index=idx, values=["1W"], aggfunc="sum", fill_value=0, margins=True).rename(columns={"1W": "1 Week"})
    pv_cash = pd.pivot_table(dfCash, index=idx, values=["Notification"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notification":"Cash"})
    pv_cashless = pd.pivot_table(dfComplete, index=idx, values=["Cashless"], aggfunc="sum", fill_value=0, margins=True)
    pv_TAT = pd.pivot_table(dfComplete, index=idx, values=["Lo"], aggfunc="mean", fill_value=0, margins=True).rename(columns={"Lo": "TAT"})
    pv = reduce(lambda left, right: pd.merge(left, right, on=idx, how="left"), [pv_cmplt, pv_1D, pv_1W, pv_cash, pv_cashless, pv_TAT]).fillna(0)
    pv["Produktifitas"] = pv["Komplit"] / hari
    pv["1D Ratio"] = (pv["1 Day"] / pv["Komplit"]) 
    pv["1W Ratio"] = (pv["1 Week"] / pv["Komplit"]) 
    pv["Cashless Ratio"] = (pv["Cashless"] / ( pv["Cashless"] + pv["Cash"]))

    pv = pv[['Komplit','TAT', '1 Day', '1 Week', 'Cash', 'Cashless', 'Produktifitas', '1D Ratio', '1W Ratio', 'Cashless Ratio']]
    
    return pv


def get_table_position(result: Dict[str, Dict[str, pd.DataFrame]]) -> dict:
    kordinat_table = {}

    for sheet, tables in result.items():
        kordinat_table[sheet] = {}
        space = 2  # spasi antar tabel
        srow = 0   # mulai dari baris 1 (Excel-like)
        scol = 0   # mulai dari kolom 1 (Excel-like)

        vertical_layout = sheet == "Pencapaian"

        for name, table in tables.items():
            table = table.copy()
            #table = table.reset_index(drop=False)  # pastikan index jadi kolom
            nrow, ncol = table.shape  # termasuk kolom index yang baru

            erow = srow + nrow
            ecol = scol + ncol - 1
            nindex = len(table.index.names)

            kordinat_table[sheet][name] = {
                'start_row': srow,
                'end_row': erow,
                'start_col': scol,
                'end_col': ecol,
                'nindex':  nindex
            }

            if vertical_layout:
                srow = erow + space + 1
            else:
                scol = ecol + space + nindex

    return kordinat_table
