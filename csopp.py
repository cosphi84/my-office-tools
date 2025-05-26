'''
csopp.py
Tools untuk melakukan ekstraksi pencapaian pekerjan, berdasarkan target KPI CS
Sumber data:
1. Ots.XLS : File csv job orders status Outstanding
2. Completed.XLS : File csv completed data
'''
import warnings

from lib.file_io import load_source, print_error
from Config.config import csopp_config
from lib.processor import fill_data, get_error_notif, calc_achivement, calc_productifitas, get_kordinat_table
from lib.formater import format_result
import pandas as pd

# Disable deprecated warning
warnings.simplefilter(action="ignore", category=FutureWarning)
warnings.simplefilter(action="ignore", category=UserWarning)

def main()->bool:
    print('CS Operational Achivement Processor')

    # Load konfigurasi operasi
    cfg = csopp_config()
    file = cfg.get('csopp_files')
    path_result = file.get('FILE_RESULT')

    # load sumber data
    source_data = load_source()

    # Clean error notif
    fixed_data = get_error_notif(source_data)
    print_error(fixed_data)

    # Tambahkan data tambahan
    final_data = fill_data(fixed_data["OK"])
   
    
    # Tabel Result
    result = {
        # Pencapaian untuk Nasional, Cabang, SDSS dan SSR
        "Pencapaian": {
            "Nasional": calc_achivement(final_data, True).fillna(0),
            "Cabang": calc_achivement(final_data, GlobalResult=False, cdr="Cabang").fillna(0),
            "SDSS": calc_achivement(final_data, False, "SDSS").fillna(0),
            "SSR": calc_achivement(final_data, False, "SSR").fillna(0),
        },
        # Produktifitas by Main work center dan by ID CSMS
        "Produktifitas Cabang": {
            'ByMWC': calc_productifitas(final_data, byMWC=True, cdr="Cabang" ).fillna(0),
            'ByTech': calc_productifitas(final_data, byMWC=False, cdr="Cabang").fillna(0),
        },
        "Produktifitas SDSS": {
            'ByMWC': calc_productifitas(final_data, byMWC=True, cdr="SDSS" ).fillna(0),
            'ByTech': calc_productifitas(final_data, byMWC=False, cdr="SDSS").fillna(0),
        },
        "Produktifitas SSR": {
            'ByMWC': calc_productifitas(final_data, byMWC=True, cdr="SSR" ).fillna(0),
            'ByTech': calc_productifitas(final_data, byMWC=False, cdr="SSR").fillna(0),
        }
    }

    # Kordinat Tabel
    kordinat_table = get_kordinat_table(result)
        
    with pd.ExcelWriter(path_result) as file:
        for sheet, tables in result.items():
            for name, table in tables.items():
                table.to_excel(file, sheet_name=sheet, startcol=kordinat_table[sheet][name]["start_col"], startrow=kordinat_table[sheet][name]["start_row"], )
    
    format_result(kordinat_table)    
    return True


if __name__ == '__main__':
    main()