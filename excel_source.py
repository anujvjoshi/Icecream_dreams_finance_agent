import os
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = "Ice Cream Dreams - Expense Report.xlsx"


def get_workbook_path():
    load_dotenv()
    configured_path = os.getenv("WORKBOOK_PATH", DEFAULT_WORKBOOK).strip()
    path = Path(configured_path).expanduser()
    if not path.is_absolute():
        path = BASE_DIR / path
    return path.resolve()


def read_workbook():
    """Read workbook worksheet values without modifying the source file."""
    path = get_workbook_path()
    if not path.is_file():
        raise FileNotFoundError(f"Configured workbook was not found: {path}")

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = {
            worksheet.title: list(worksheet.iter_rows(values_only=True))
            for worksheet in workbook.worksheets
        }
    finally:
        workbook.close()

    return path.stem, sheets