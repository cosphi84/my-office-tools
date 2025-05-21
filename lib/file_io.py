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

def load_file(path: Path, required_cols = []):
    '''
    load_file Memuat file yang ditunjuk menjadi dataframe. support hanya file XLS, XLSX dan CSV saja

    Args:
        path : Path file yang ingin di load
        required_cols: Array kolom apa saja yang dibutuhkan

    Returns:
        dataframe on success. False on fail
    '''

    file_ext = (path.suffix).lower()
    df = None
    try:
        if file_ext == '.csv' or file_ext == '.xls':
            df = pd.read_csv(path, dialect="excel-tab", encoding="utf_16", skiprows=3)
        elif file_ext == '.xlsx':
            df = pd.read_excel(path)
        else: 
            raise SystemExit(f'File {path} tidak didukung.')
        
        df_columns = [col.lower() for col in df.columns]
        if not set([c.lower() for c in required_cols]).issubset(df_columns):
            missing = set([c.lower() for c in required_cols]) - set(df_columns)
            raise SystemExit(f'Kolom beirkut tidak ada di datasource: {missing}')
        
    except Exception as e:
        print(f"Kesalahan tak terduga saat membaca file: {e}")
        dfData = None
            
    return df
