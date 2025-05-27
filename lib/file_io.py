import pandas as pd
import os
from pathlib import Path
from tkinter import messagebox
from Config.config import csopp_config
from typing import Dict



def check_file(path: Path, required: bool = True):
    '''
    Fungsi ini akan mengecek apakah file yang ditunjuk ada atau tidak.

    Args :
        path: Path file yang ditunjuk
        required: Boolean, True akan membuat system keluar jika file tidak ditemukan

    Returns:
        Path file
    '''
    if not path.exists():
        msg = f"File '{path}' tidak ditemukan!"
        messagebox.showerror('Error', f'{msg}')
        if required:
            raise SystemExit(msg)
    return path

def remove_file(path: Path):
    """
    Hapus file jika ada

    Args:
        path : Path file yang akan dihapus

    Returns: 
        True ketika sukses, False sebaliknya
    """
    if path.exists():
        try:
            os.remove(path)
            print(f"File {path} dihapus.")
        except Exception as e:
            msg = f"Gagal menghapus file {path}: {e}"
            print(msg)
            return False
        
        return True

def load_source() -> dict:
    '''
    load_source: Memuat file SourceData menjadi dic dataframe. Support SourceFile hanya ots dan completed dengan extensi XLS, XLSX dan CSV saja

    Returns:
        dic df["ots", "completed"] atau None
    '''
    cols = ["Notifictn type", "Notification", "Notif.date", "Req. start","Required End","Changed on","Completn date","Planner group",
            "Main WorkCtr","User status","List name", "Street", "Telephone", "Material", "Serial number", "Description", "Changed on",
            "Device data"]    
    sources = csopp_config().get('csopp_files')
    s_files = {
        'ots': sources['FILE_OTS'],
        'completed': sources['FILE_COMPLETED']
    }
    
    dfResult = {}
    for label, file in s_files.items():
        file = Path(file)

        try:
            df = pd.read_excel(
                file, 
                usecols=cols,
                engine="openpyxl"
            )
        except Exception as e:
            messagebox.showerror("Error", f'Gagal memuat Source File: {e}')
            raise SystemExit()
        # Buang semua LR yang tidak dipakai
        df[~df['Notifictn type'].isin([csopp_config().get('csopp_exclude_typ',[])])]
        
        # fix column data type
        df["Notif.date"] = pd.to_datetime(df["Notif.date"], errors='coerce')
        df["Req. start"] = pd.to_datetime(df["Req. start"], errors="coerce")
        df["Required End"] = pd.to_datetime(df["Required End"], errors="coerce")
        df["Changed on"] = pd.to_datetime(df["Changed on"], errors="coerce")
        df["Completn date"] = pd.to_datetime(df["Completn date"], errors="coerce")
        df["Planner group"] = df["Planner group"].astype(str)
        df["User status"] = pd.to_numeric(df["User status"], errors="coerce")
        df["Main WorkCtr"] = df["Main WorkCtr"].astype(str)
        df["Notification"] = df["Notification"].astype(str)

        dfResult[label] = df
    return dfResult

def print_error(df: Dict[str, dict] ):
    files = csopp_config().get('csopp_files')
    for jenis, tables in df.items():
        with pd.ExcelWriter(files.get('FILE_ERROR')) as xls:
            if jenis == 'error':
                for i, datas in tables.items():
                    for sheet, obj in datas.items():
                        obj.to_excel(xls, sheet_name=i+"_"+sheet, index=False)
            else:
                continue

def export_to_excel(result: Dict[str, pd.DataFrame])->bool:
    files = csopp_config().get('csopp_file')


'''
def write_excel(source: dict) -> bool:
    with pd.ExcelWriter(FILE_RESULT, engine="openpyxl") as writer:
    row = 0
    nasional.to_excel(writer, sheet_name=SHEET_NAMES["Result"], startrow=row)
    row += len(nasional) + 5
    excel_tables[SHEET_NAMES["Result"]] = {
        "nasional" : {
            "start_row": 1,
            "end_row": len(nasional)+1,
            "start_col": 1,
            "end_col": len(nasional.columns)+1
        }
    }

    for label, df in result.items():
        df.to_excel(writer, sheet_name=SHEET_NAMES["Result"], startrow=row)
        start_row = row
        row += len(df) + 5
        excel_tables[SHEET_NAMES["Result"]].update({
            label: {
                "start_row": start_row+1,
                "end_row" : len(df)+1 + start_row,
                "start_col": 1,
                "end_col": len(df.columns)+2
            }
        })

    col = 0
    excel_tables[SHEET_NAMES["Prod_C"]] = {}
    for label, df in result_tech.items():
        shift = 3 if label == 'ByMWC' else 4
        excel_tables[SHEET_NAMES["Prod_C"]].update({
            label: {
                "start_row": 1,
                "end_row": len(df)+1,
                "start_col": col+1,
                "end_col": len(df.columns)+col+shift,
            }
        })
        df.to_excel(writer, sheet_name=SHEET_NAMES["Prod_C"], startcol=col, startrow=0)
        col += len(df.columns) + 5
        

    col = 0
    excel_tables[SHEET_NAMES["Prod_D"]] = {}
    for label, df in result_SDSS.items():
        shift = 3 if label == 'ByMWC' else 4
        excel_tables[SHEET_NAMES["Prod_D"]].update({
            label: {
                "start_row": 1,
                "end_row": len(df)+1,
                "start_col": col+1,
                "end_col": len(df.columns)+col+shift
            }
        })
        df.to_excel(writer, sheet_name=SHEET_NAMES["Prod_D"], startcol=col, startrow=0)
        col += len(df.columns) + 5
        

    col = 0
    excel_tables[SHEET_NAMES["Prod_R"]] = {}
    result_SSR.to_excel(writer, sheet_name=SHEET_NAMES["Prod_R"], startcol=col, startrow=0)
    excel_tables[SHEET_NAMES["Prod_R"]].update({
            "ssr": {
                "start_row": 1,
                "end_row": len(result_SSR)+1,
                "start_col": col+1,
                "end_col": len(result_SSR.columns)+3,
            }
        })

    return True
'''