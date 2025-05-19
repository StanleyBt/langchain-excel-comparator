# excel_comparator/utils/file_loader.py
import pandas as pd

def detect_header_row(path, sheet_name, max_scan_rows=10):
    """
    Detect the row index likely containing the actual headers.
    Chooses the row with the most non-empty cells within the first few rows.
    """
    preview = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=max_scan_rows)
    non_empty_counts = preview.notna().sum(axis=1)
    likely_header_row = non_empty_counts.idxmax()  # returns index of most-filled row
    return likely_header_row

def load_headers(path, sheet_name):
    """
    Load the header row and return the header row index along with cleaned column names.
    """
    header_row = detect_header_row(path, sheet_name)
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row, nrows=1)
    headers = [col.strip().lower() for col in df.columns]
    return df, headers, header_row