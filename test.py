from lib.file_io import load_source
from lib.processor import fill_data
from lib.formater import format_result
from pathlib import Path

# file sumber
source = { "ots" :Path('DataSource/ots.csv'),
            "completed" : Path('DataSource/completed.csv')
}

result = {
    "Result": Path('Result/result.xlsx'),
    "Detail": Path("Result/detail.xlsx"),
    "Error": Path("Result/error.xlsx")
}

# data tambahan
add_data = {
    "exclude_lr": ["Z6","Z8", "ZX", "ZZ"],
    "exclude_stt": [51, 52, 97],
    "cashless_stt": [94, 95, 96, 98],
    "cash_stt": [93],
    "exclude_cash": ["ZX"]
}


def main():
    # Ambil file source
    df = load_source(source)

    # tambah data data
    df = fill_data(df, add_data)

    # Format result
    format_result(result["Result"])

if __name__ == '__main__':
    main()