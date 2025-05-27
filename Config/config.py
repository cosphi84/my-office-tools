from pathlib import Path
import pandas as pd
from tkinter import messagebox
from typing import Dict

def csopp_config() -> Dict[str, Dict]:
    the_file = Path(__file__).resolve().parent
    csopp = the_file / "csopp.xlsx"
    # missing
    if not csopp.exists:
        messagebox.showerror('Config Missing', f'File config {csopp.name} tidak berada di tempatnya\nProses dibatalkan.' )
        raise SystemExit
    
    with pd.ExcelFile(csopp) as cfg:
        base_dir = the_file.parent
        try:
            config_val = {
                "csopp_setting": pd.read_excel(cfg, sheet_name='seting').set_index('Seting')['Value'].to_dict(),
                'csopp_bp': pd.read_excel(cfg, sheet_name='bp'),
                'csopp_pg': pd.read_excel(cfg, sheet_name='pg').rename(columns={'PG': 'Planner group'}),
                'csopp_mwc' : pd.read_excel(cfg, sheet_name='mwc').rename(columns={'Kode': 'Main WorkCtr', 'Teknisi':'Work Center'}),
                'csopp_idcsms': pd.read_excel(cfg, sheet_name='teknisi', usecols=['id', 'name']).rename(columns={'id': 'idCSMS', 'name':'Teknisi'}),
                'csopp_exclude_mwc': pd.read_excel(cfg, sheet_name='exclude').values.tolist(),
                'csopp_exclude_typ': ['Z8'],
                'csopp_exclude_cash': ['ZX'],
                'csopp_exclude_stt': [34, 97],
                'csopp_cash_status': [93],
                'csopp_cashless_status': [94,95,96,98],
                'csopp_files' : {
                    # File path constants
                    'FILE_OTS' : Path(base_dir / 'DataSource/ots.xlsx'),
                    'FILE_COMPLETED' : Path(base_dir / "DataSource/completed.xlsx"),
                    'FILE_RESULT' : Path(base_dir / "Result/Result.xlsx"),
                    'FILE_DETAIL' : Path(base_dir / "Result/Detail.xlsx"),
                    'FILE_ERROR': Path(base_dir / "Result/Cek_Notif.xlsx"),
                    'FILE_TEKNISI':  Path(base_dir / "Result/JadwalTeknisi.xlsx"),
                    'SHEET_NAMES' : {
                        "Result": "Pencapaian",
                        "Prod_C": "Produktifitas Cabang",
                        "Prod_D": "Produktifitas SDSS",
                        "Prod_R": "Produktifitas SSR"
                    }
                }
            }
        except Exception as e:
            messagebox.showerror('Error', f'Error:{e}')
            raise SystemExit
        
        return config_val