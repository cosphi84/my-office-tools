import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from functools import reduce

def fill_data(data: dict, add_data: dict) -> dict:
    '''
    fill_data: modifikasi dataframe

    Args:
        data Dict Dict data, harus berisi 'ots' => dataframe OTS, 'complited' => dataframe complited
        add_data Dict add_data, berisi seting jenis lr dan status yang akan di exclude

    Returns:
        dict Dataframa OTS dan COMPLITED yang sudah ready
    '''
    the_file = Path(__file__).resolve()
    config_file = the_file.parent.parent / "Config" / "csopp.xlsx"
    sekarang = datetime.now()
    
    with config_file as cfg:
        try:
            pg = pd.read_excel(cfg, "pg")
            mwc = pd.read_excel(cfg, "mwc").rename(columns={'Kode':'Mn.wk.ctr', "Teknisi": "Work Center"})
            teknisi = pd.read_excel(cfg, "teknisi", usecols=["id", "name"]).rename(columns={"id": "idCSMS", "name":"Teknisi"})
            excludeMwc = pd.read_excel(cfg, "exclude", usecols=["mwc"])
        except Exception as e:
            raise SystemExit(e)
    
    result_df = {}
    
    for tabel, df in data.items():
        df = pd.DataFrame(df)
        cols = ["Typ", "Notifctn", "Notif.date", "Req. start", "Req. End", "Changed on", "Completion", "PG",  "Mn.wk.ctr", "UserStatus", "List name", "Street", "Telephone", "Material", "Serial number", "Description", "Addit. device data", "Changed on"]
        missing_cols = [col for col in cols if col not in df.columns]
        if missing_cols:
            raise SystemExit(f'Kolom berikut tidak ada di file source: {missing_cols}')
        
        is_ots = False if tabel != 'ots' else True
        
        if len(add_data['exclude_lr']) > 0:
            df = df[~df["Typ"].isin([add_data["exclude_lr"]])]

        if len(add_data["exclude_stt"]) > 0:
            df = df[~df["UserStatus"].isin([add_data["exclude_stt"]])]
        
        # Convert some stuff
        df["Notif.date"] = pd.to_datetime(df["Notif.date"], dayfirst=True, errors="coerce")
        df["Completion"] = pd.to_datetime(df["Completion"], errors="coerce", format="%d.%m.%Y")
        df["Changed on"] = pd.to_datetime(df["Changed on"], dayfirst=True, errors="coerce")
        df["UserStatus"] = pd.to_numeric(df["UserStatus"], errors="coerce")
        df["PG"] = pd.to_numeric(df["PG"], errors="coerce")
        df["Mn.wk.ctr"] = df["Mn.wk.ctr"].astype(str)
        df["bulan_complet"] = df["Completion"].fillna(sekarang).values.astype("datetime64[M]")
        df["RcvdThisMo"] = (df["Notif.date"] >= df["bulan_complet"]).astype(int)
        df.drop(columns=["bulan_complet"], inplace=True)
        df["Req. End"] = pd.to_datetime(df["Req. End"], errors="coerce", format="%d.%m.%Y")

        # add some stuff    
        if is_ots :
            df["isOTS"] = 1
            df["Req. End"] = df["Req. End"].fillna(sekarang)
            df["e_no_req_end"] = 0
        else:
            df["isOTS"] = 0
            df["e_no_req_end"] = df['Req. End'].isna().astype(int)
            df["Req. End"] = df["Req. End"].fillna(sekarang)
            # Buang beberapa Main work yang ingin dibuang
            df = df[~df["Mn.wk.ctr"].isin(excludeMwc)]
        
        df["Lo"] = (df["Req. End"] - df["Notif.date"]).dt.days
        df["1D"] = (df["Lo"] <= 1).astype(int)
        df["1W"] = (df["Lo"] <= 7).astype(int)
        df["Cashless"] = df["UserStatus"].isin(add_data["cashless_stt"]).astype(int)
        
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
        df = df.merge(pg, how="left", on="PG").fillna("")
        df = df.merge(mwc, how="left", on="Mn.wk.ctr").fillna("")
        df = df.merge(teknisi, how="left", on="idCSMS").fillna("")

        result_df[tabel] = df
    
    return result_df    
    
def calc_achivement(dfData: pd.DataFrame, GlobalResult : bool = False, Cabang : bool= True, additional_data: dict = {}):
    idx = ["Regional"] if GlobalResult == True else ["Regional", "Cabang"]
    if GlobalResult :
        idx = ["Regional"]
    elif not GlobalResult and Cabang:
        idx = ["Regional", "Cabang"]
    else:
        idx = ["Regional", "Work Center"]
    pv = pd.pivot_table(dfData, index=idx, values=["Notifctn"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notifctn": "Total LK"})
    
    # pecah lagi data OTS & data Complete
    dfOTS = dfData[dfData["isOTS"].isin([1])]
    df30 = dfOTS[dfOTS["UserStatus"].isin([30])]
    dfLo = dfOTS[dfOTS["Lo"] >= 60]
    dfComplete = dfData[dfData["isOTS"].isin([0])]
    dfCash = dfData[dfData["UserStatus"].isin([93])]
    # Abaikan type ZX untuk perhitungan Cashless
    dfCash = dfCash[~dfCash["Typ"].isin(["ZX"])]

    # Pivoting
    pv_cmplt = pd.pivot_table(dfComplete, index=idx, values=["Notifctn"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notifctn": "Komplit"})
    pv_ots = pd.pivot_table(dfOTS, index=idx, values=["Notifctn"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notifctn": "OTS"})
    pv_1D = pd.pivot_table(dfComplete, index=idx, values=["1D"], aggfunc="sum", fill_value=0, margins=True).rename(columns={"1D": "1 Day"})
    pv_1W = pd.pivot_table(dfComplete, index=idx, values=["1W"], aggfunc="sum", fill_value=0, margins=True).rename(columns={"1W": "1 Week"})
    pv_cash = pd.pivot_table(dfCash, index=idx, values=["Notifctn"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notifctn":"Cash"})
    pv_cashless = pd.pivot_table(dfComplete, index=idx, values=["Cashless"], aggfunc="sum", fill_value=0, margins=True)
    pv_30 = pd.pivot_table(df30, index=idx, values=["Notifctn"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notifctn": "STT 30"})
    pv_Lo = pd.pivot_table(dfLo, index=idx, values=["Notifctn"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notifctn": "LO"})
    pv_TAT = pd.pivot_table(dfComplete, index=idx, values=["Lo"], aggfunc="mean", fill_value=0, margins=True).rename(columns={"Lo": "TAT"})
    pv = reduce(lambda left, right: pd.merge(left, right, on=idx, how="left"), [pv, pv_ots, pv_cmplt, pv_1D, pv_1W, pv_cash, pv_cashless, pv_30, pv_Lo, pv_TAT]).fillna(0)

    pv["Cplt Ratio"] = (pv["Komplit"] / pv["Total LK"])
    pv["LO Ratio"] = (pv["LO"] / pv["OTS"])
    pv["1D Ratio"] = (pv["1 Day"] / pv["Komplit"]) 
    pv["1W Ratio"] = (pv["1 Week"] / pv["Komplit"]) 
    pv["Cashless Ratio"] = (pv["Cashless"] / ( pv["Cashless"] + pv["Cash"]))
    pv["STT 30 VS OTS"] = (pv["STT 30"] / pv["OTS"]) 
    return pv

def calc_productifitas(dfData: pd.DataFrame, byMWC : bool = True, additional_data: dict = {}):
    the_file = Path(__file__).resolve()
    config_file = the_file.parent.parent / "Config" / "csopp.xlsx"
    idx = ["Regional","Cabang", "Work Center"] if byMWC == True else ["Regional","Cabang", "Mn.wk.ctr", "Teknisi"]
    cfg = pd.read_excel(config_file, "seting").set_index("Seting")["Value"].to_dict()
    
    # pecah lagi data OTS & data Complete
    dfComplete = dfData[dfData["isOTS"].isin([0])]
    dfCash = dfData[dfData["UserStatus"].isin(additional_data["cash_stt"])]
    # Abaikan type ZX untuk perhitungan Cashless
    dfCash = dfCash[~dfCash["Typ"].isin(additional_data["exclude_cash"])]

    # Pivoting
    pv = pd.pivot_table(dfComplete, index=idx, values=["Notifctn"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notifctn": "Total LK"})
    pv_1D = pd.pivot_table(dfComplete, index=idx, values=["1D"], aggfunc="sum", fill_value=0, margins=True).rename(columns={"1D": "1 Day"})
    pv_1W = pd.pivot_table(dfComplete, index=idx, values=["1W"], aggfunc="sum", fill_value=0, margins=True).rename(columns={"1W": "1 Week"})
    pv_cash = pd.pivot_table(dfCash, index=idx, values=["Notifctn"], aggfunc="count", fill_value=0, margins=True).rename(columns={"Notifctn":"Cash"})
    pv_cashless = pd.pivot_table(dfComplete, index=idx, values=["Cashless"], aggfunc="sum", fill_value=0, margins=True)
    pv_TAT = pd.pivot_table(dfComplete, index=idx, values=["Lo"], aggfunc="mean", fill_value=0, margins=True).rename(columns={"Lo": "TAT"})
    pv = reduce(lambda left, right: pd.merge(left, right, on=idx, how="left"), [pv, pv_1D, pv_1W, pv_cash, pv_cashless, pv_TAT]).fillna(0)

    pv["Produktifitas"] = pv["Total LK"] / cfg["Hari"]
    pv["1D Ratio"] = (pv["1 Day"] / pv["Total LK"]) 
    pv["1W Ratio"] = (pv["1 Week"] / pv["Total LK"]) 
    pv["Cashless Ratio"] = (pv["Cashless"] / ( pv["Cashless"] + pv["Cash"]))
    return pv

def apply_filter(dfData: pd.DataFrame, CDR = "Cabang"):
    the_file = Path(__file__).resolve()
    config_file = the_file.parent.parent / "Config" / "csopp.xlsx"
    cfg = pd.read_excel(config_file, "seting").set_index("Seting")["Value"].to_dict()
    for col in ["Regional", "Cabang"]:
        if cfg.get(col, "All") != "All":
            dfData = dfData[dfData[col] == cfg[col]]

    dfData = dfData[dfData["CDR"] == CDR]
    return dfData