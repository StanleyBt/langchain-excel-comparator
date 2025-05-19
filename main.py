# excel_comparator/main.py
# === Entrypoint for the Excel Header Comparison Pipeline ===
from utils.file_loader import load_headers
from utils.header_matching import get_exact_matches, get_semantic_mapping
from utils.export_writer import export_results
from utils.row_comparison import compare_rows
from config import EXCEL_PATH, VENDOR_SHEET, SYSTEM_SHEET
import pandas as pd

def main():
    df_vendor_preview, headers_vendor, vendor_header_row = load_headers(EXCEL_PATH, VENDOR_SHEET)
    df_system_preview, headers_system, system_header_row = load_headers(EXCEL_PATH, SYSTEM_SHEET)

    exact_matches, unmatched_vendor, unmatched_system = get_exact_matches(headers_vendor, headers_system)
    semantic_matches = get_semantic_mapping(unmatched_vendor, unmatched_system)

    export_results(exact_matches, semantic_matches, unmatched_vendor, unmatched_system)

    final_mapping = {
        **exact_matches,
        **{k: v for k, v in semantic_matches.items() if v and v.lower() != "no match"}
    }

    df_vendor = pd.read_excel(EXCEL_PATH, sheet_name=VENDOR_SHEET, header=vendor_header_row)
    df_system = pd.read_excel(EXCEL_PATH, sheet_name=SYSTEM_SHEET, header=system_header_row)

    compare_rows(df_vendor, df_system, final_mapping)

if __name__ == "__main__":
    main()