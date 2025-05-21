import openpyxl
from openpyxl.styles import Border, Side, Alignment, PatternFill
from pathlib import Path
import pandas as pd

def format_result(path: Path, table_pos: dict):
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as e:
        msg = f"Gagal membuka file {path}: {e}"
        print(msg)
        raise SystemExit(msg)

    the_file = Path(__file__).resolve()
    config_file = the_file.parent.parent / "Config" / "csopp.xlsx"
    cfg = pd.read_excel(config_file, "bp")
    config = {
        row["item"]: {"mode": row["kondisi"], "value": row["bp"]}
        for _, row in cfg.iterrows()
    }

    #print(table_pos)

    # Gaya umum
    border = Border(
        bottom=Side(style='thin'), top=Side(style='thin'),
        left=Side(style='thin'), right=Side(style='thin')
    )
    align_left = Alignment(horizontal='left', vertical='center')
    align_center = Alignment(horizontal='center', vertical='center')
    fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fill_ng = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    sky_blue = PatternFill(fill_type='solid', fgColor='FFC0FFC0')

    #print(excel_tables)
    
    col_map = {
        'nasional': {'LO': 9, 'TAT': 10, 'Cplt Ratio': 11, 'LO Ratio': 12, '1D Ratio': 13, '1W Ratio': 14, 'Cashless Ratio': 15, 'STT 30 VS OTS': 16},
        'Cabang': {'LO': 10, 'TAT': 11, 'Cplt Ratio': 12, 'LO Ratio': 13, '1D Ratio': 14, '1W Ratio': 15, 'Cashless Ratio': 16, 'STT 30 VS OTS': 17},
        'SDSS': {'LO': 10, 'TAT': 11, 'Cplt Ratio': 12, 'LO Ratio': 13, '1D Ratio': 14, '1W Ratio': 15, 'Cashless Ratio': 16, 'STT 30 VS OTS': 17},
        'SSR': {'LO': 10, 'TAT': 11, 'Cplt Ratio': 12, 'LO Ratio': 13, '1D Ratio': 14, '1W Ratio': 15, 'Cashless Ratio': 16, 'STT 30 VS OTS': 17},
        'ByMWC': {'TAT': 8, 'Produktifitas': 9 , '1D Ratio': 10, '1W Ratio': 11, 'Cashless Ratio': 12}, 
        'ByTech': {'TAT': 9, 'Produktifitas': 10 , '1D Ratio': 11, '1W Ratio': 12, 'Cashless Ratio': 13},
        'ssr': {'TAT': 8, 'Produktifitas': 9 , '1D Ratio': 10, '1W Ratio': 11, 'Cashless Ratio': 12},
    }

    for sheet in wb.worksheets:
        # Loop untuk setiap tabel dalam setiap sheet
        for table, coord in table_pos[sheet.title].items():
            # Format Header
            for col in sheet.iter_cols(
                min_row=coord["start_row"], max_row=coord["start_row"],
                min_col=coord["start_col"], max_col=coord["end_col"]
            ):
                 for idx, cell in enumerate(col):
                     cell.fill = sky_blue
                     cell.alignment = align_center
                     cell.border = border

            # Format Footer
            for col in sheet.iter_cols(
                min_row=coord["end_row"], max_row=coord["end_row"],
                min_col=coord["start_col"], max_col=coord["end_col"]
            ):
                 for idx, cell in enumerate(col):
                     cell.fill = sky_blue
                     cell.alignment = align_center
                     cell.border = border

            # Format body tabel
            for row in sheet.iter_rows(
                min_row=coord["start_row"]+1, max_row=coord["end_row"],
                min_col=coord["start_col"], max_col=coord["end_col"]
            ):
                for idx, cell in enumerate(row):
                    for i, val in col_map[table].items():
                        if table == 'nasional' :
                            judul = 1
                        elif table == 'ByTech':
                            judul = 4
                        else:
                            judul = 3

                        if idx <= judul:    
                            cell.alignment = align_left
                        
                        if idx == col_map[table][i]:
                            cell.alignment = align_center
                            cell.border = border
                            cell.number_format = '0.00' if i in ["LO", "TAT", "Produktifitas"] else '0.00%'                            
                            cell_value = float(cell.value)
                            if i in config:
                                if config[i]["mode"] == 'min':
                                    cell.fill = fill_ok if cell_value >= config[i]["value"] else fill_ng
                                else: 
                                    cell.fill = fill_ok if cell_value <= config[i]["value"] else fill_ng
                            else:
                                continue
                        else:
                            cell.border = border
    wb.save(path)