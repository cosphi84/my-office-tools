import pandas as pd
import os
from pathlib import Path



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
        print(msg)
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

def load_source(source = []) -> dict:
    '''
    load_source: Memuat file SourceData menjadi dic dataframe. Support SourceFile hanya ots dan completed dengan extensi XLS, XLSX dan CSV saja

    Returns:
        dic df["ots", "completed"] atau None
    '''
    cols = {
            ".csv" : [   "Typ", "Notifctn", "Notif.date", "Req. start","Req. End","Changed on","Completion","PG","Mn.wk.ctr","UserStatus",
                        "List name", "Street", "Telephone", "Material", "Serial number", "Description","Addit. device data", 'Changed on'],
            ".xlsx" : [  "Notifictn type", "Notification", "Notif.date", "Req. start","Required End","Changed on","Completn date","Planner group",
                        "Main WorkCtr","User status","List name", "Street", "Telephone", "Material", "Serial number", "Description", 'Changed on'
                        "Device data"]
    }
    rename_cols = {
        'Notifictn type': 'Typ',
        'Notification': 'Notifctn', 
        'Required End': 'Req. End', 
        'User status' : 'UserStatus', 
        'Device data' : 'Addit. device data',
        'Main WorkCtr': 'Mn.wk.ctr',
        'Planner group':'PG',
        'Completn date' : 'Completion'
    }
    if len(source) <= 0:
        raise SystemExit("File Source mutlak dibutuhkan")
    dfResult = {}
    for label, file in source.items():
        file = Path(file)
        ext = file.suffix.lower()
        if ext == '.csv' or ext == '.xls':
            try:
              dfResult[label]   = pd.read_csv(file, dialect='excel-tab', encoding='utf_16', skiprows=3, usecols=cols[ext])
            except Exception as e:
                raise SystemExit(f'Gagal meload source file: {e}')
        elif ext == '.xlsx':
            try:
              dfResult[label]   = pd.read_excel(file, usecols=cols[ext]).rename(columns=rename_cols)
            except Exception as e:
                raise SystemExit(f'Gagal meload source file: {e}')
        else:
            raise SystemExit("File source tidak didukung.") 
    return dfResult

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