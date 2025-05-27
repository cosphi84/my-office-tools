from typing import Dict
from tkinter import messagebox
import pandas as pd


def get_detail_lr(souredata):
    sourcedata = souredata.copy()
    detail = {}

    try:
        dfCompleted: pd.DataFrame = souredata.get('completed')
        dfOts: pd.DataFrame= souredata.get('ots')
    except Exception as e:
        messagebox.showinfo('Load Detail Kosong', f'Load source data untuk detail error: {e}')
        return None
    
    # stt 30 dari ots
    dfsvc = dfOts[dfOts["User status"] < 30]
    # stt 30 dari ots
    df30 = dfOts[dfOts["User status"] == 30]
    # stt 32
    df32 = dfOts[dfOts['User Status'] == 32]
    # stt 33
    df33 = dfOts[dfOts['User Status'] == 33]
    # stt Subtitusi
    df34 = dfOts[dfOts['User Status'] == 34]
    # stt OK
    dfOK = dfOts[dfOts['User Status'] >= 53]
    # Z2 
    dfZ2 = dfOts[dfOts["Notifictn type"] == 'Z2']


    

    return df