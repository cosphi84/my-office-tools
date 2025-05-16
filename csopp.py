'''
csopp.py
Tools untuk melakukan ekstraksi pencapaian pekerjan, berdasarkan target KPI CS
Sumber data:
1. Ots.XLS : File csv job orders status Outstanding
2. Completed.XLS : File csv completed data
'''

import pandas as pd
import os
from tkinter import messagebox
from pathlib import Path

file_config = Path("Config/csopp.xlsx")
file_ots = Path('DataSource/ots.XLS')
file_completed = Path("DataSource/completed.XLS")
file_result = Path("Result/Result.xlsx")
file_detail = Path("Result/Detail.xlsx")
file_error = Path("Result/Error.xlsx")
file_teknisi = Path("Result/JadwalTeknisi.xlsx")

if not file_config.exists():
    print(f'File {file_config} tidak ada!')
    raise SystemExit

if not file_ots.exists():
    print(f'File sumber {file_ots} tidak ada!')
    raise SystemExit

if not file_completed.exists():
    print(f'File sumber {file_completed} tidak ada!')
    raise SystemExit

if file_result.exists():
    try:
        os.remove(file_result)
    except:
        print(f'File {file_result} sudah ada dan gagal di hapus!')
        raise SystemExit
    
if file_detail.exists():
    try:
        os.remove(file_detail)
    except:
        print(f'File {file_detail} sudah ada dan gagal di hapus!')
        raise SystemExit

if file_error.exists():
    try:
        os.remove(file_error)
    except:
        print(f'File {file_error} sudah ada dan gagal di hapus!')
        raise SystemExit


