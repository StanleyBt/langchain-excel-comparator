# excel_comparator/utils/export_writer.py
import pandas as pd
import xlsxwriter
from config import OUTPUT_FILE

def export_results(matched, semantic, unmatched_vendor, unmatched_system):
    # Prepare rows
    exact_rows = [
        {"Vendor Header": k, "System Header": v, "Match Type": "Exact Match"}
        for k, v in matched.items()
    ]
    semantic_rows = [
        {"Vendor Header": k, "System Header": v, "Match Type": "Semantic Match"}
        for k, v in semantic.items() if v and v.lower() != "no match"
    ]
    matched_vendor_keys = set(matched.keys()) | {k for k, v in semantic.items() if v and v.lower() != "no match"}
    unmatched_vendor_only = [v for v in unmatched_vendor if v not in matched_vendor_keys]
    unmatched_vendor_rows = [
        {"Vendor Header": v, "System Header": None, "Match Type": "Not Matched"}
        for v in unmatched_vendor_only
    ]
    matched_system_values = set(matched.values()) | {v for v in semantic.values() if v and v.lower() != "no match"}
    unmatched_system_only = [s for s in unmatched_system if s not in matched_system_values]
    unmatched_system_rows = [
        {"Vendor Header": None, "System Header": s, "Match Type": "System Only"}
        for s in unmatched_system_only
    ]

    # Combine and write
    df_combined = pd.DataFrame(
        exact_rows + semantic_rows + unmatched_vendor_rows + unmatched_system_rows
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        df_combined.to_excel(writer, sheet_name="Header Mapping", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Header Mapping"]

        # Styles
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        exact_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1})
        semantic_fmt = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1})
        unmatched_fmt = workbook.add_format({'bg_color': '#FCE4D6', 'border': 1})
        system_only_fmt = workbook.add_format({'bg_color': '#E0E0E0', 'border': 1})

        # Format header row
        for col_num, value in enumerate(df_combined.columns):
            worksheet.write(0, col_num, value, header_fmt)

        # Style Match Type column
        match_col_index = df_combined.columns.get_loc("Match Type")
        for row_num, match_type in enumerate(df_combined["Match Type"], start=1):
            if match_type == "Exact Match":
                fmt = exact_fmt
            elif match_type == "Semantic Match":
                fmt = semantic_fmt
            elif match_type == "Not Matched":
                fmt = unmatched_fmt
            else:
                fmt = system_only_fmt
            worksheet.write(row_num, match_col_index, match_type, fmt)

        # Autofit and freeze
        for i, col in enumerate(df_combined.columns):
            max_len = max(df_combined[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(df_combined), len(df_combined.columns) - 1)

    print(f"🎨 Header mapping exported to {OUTPUT_FILE}")
