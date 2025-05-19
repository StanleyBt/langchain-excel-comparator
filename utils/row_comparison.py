# excel_comparator/utils/row_comparison.py
import pandas as pd

def compare_rows(df_vendor_raw, df_system_raw, final_mapping):
    # === Normalize headers ===
    df_vendor_raw.columns = [col.strip().lower() for col in df_vendor_raw.columns]
    df_system_raw.columns = [col.strip().lower() for col in df_system_raw.columns]

    # === Identify key column from vendor-side mapping ===
    possible_ids = [col for col in final_mapping if any(
        key in col.replace(" ", "").lower() for key in ["employeeid", "employeeno", "bluetreeid", "cemsemployeeid", "empid"]
    )]

    if not possible_ids:
        print("❌ Could not automatically detect an Employee ID column.")
        print("Available mapped columns:", list(final_mapping.keys()))
        raise ValueError("Please manually specify the employee ID column.")

    if len(possible_ids) > 1:
        print("🤖 Multiple possible Employee ID columns found:")
        for i, col in enumerate(possible_ids, start=1):
            print(f"{i}. {col}")
        choice = int(input("Please select the Employee ID column (1/2/...): ")) - 1
        key_col = possible_ids[choice]
    else:
        key_col = possible_ids[0]

    print(f"🔑 Using '{key_col}' as Employee ID column")

    # === Align system columns to vendor names ===
    reverse_mapping = {v: k for k, v in final_mapping.items() if v}
    df_system = df_system_raw.rename(columns=reverse_mapping)

    # === Set index for join ===
    df_vendor = df_vendor_raw.set_index(key_col)
    df_system = df_system.set_index(key_col)

    # === Join DataFrames ===
    joined = df_vendor.join(df_system, how="outer", lsuffix="_vendor", rsuffix="_system")

    diffs = []
    summary_counts = {}

    # === Compare all mapped columns (not just what's in df_vendor) ===
    for vendor_col in final_mapping:
        vendor_field = f"{vendor_col}_vendor"
        system_field = f"{vendor_col}_system"

        if vendor_field in joined.columns and system_field in joined.columns:
            for emp_id, row in joined.iterrows():
                val_v = row.get(vendor_field, None)
                val_s = row.get(system_field, None)

                if isinstance(val_v, pd.Series):
                    val_v = val_v.iloc[0]
                if isinstance(val_s, pd.Series):
                    val_s = val_s.iloc[0]

                try:
                    if pd.isna(val_v) and pd.isna(val_s):
                        match_flag = "Yes"
                    elif pd.isna(val_v) or pd.isna(val_s):
                        match_flag = "No"
                    elif str(val_v).strip() == str(val_s).strip():
                        match_flag = "Yes"
                    else:
                        match_flag = "No"
                except Exception:
                    match_flag = "No"

                diffs.append({
                    "Employee ID": emp_id,
                    "Column": vendor_col,
                    "Vendor Value": val_v,
                    "System Value": val_s,
                    "Match?": match_flag
                })
                if match_flag == "No":
                    summary_counts[vendor_col] = summary_counts.get(vendor_col, 0) + 1

    # === Find missing Employee IDs ===
    vendor_ids = set(df_vendor.index)
    system_ids = set(df_system.index)

    missing = []
    for emp_id in sorted(vendor_ids - system_ids):
        missing.append({"Employee ID": emp_id, "Missing In": "System"})
    for emp_id in sorted(system_ids - vendor_ids):
        missing.append({"Employee ID": emp_id, "Missing In": "Vendor"})

    # === Export results ===
    from config import ROW_OUTPUT_FILE

    with pd.ExcelWriter(ROW_OUTPUT_FILE) as writer:
        if diffs:
            df_diffs = pd.DataFrame(diffs)
            df_diffs.to_excel(writer, sheet_name="Row Differences", index=False)
        if missing:
            df_missing = pd.DataFrame(missing)
            df_missing.to_excel(writer, sheet_name="Missing Employees", index=False)
        if summary_counts:
            df_summary = pd.DataFrame.from_dict(summary_counts, orient="index", columns=["Mismatch Count"])
            df_summary.index.name = "Column"
            df_summary.reset_index(inplace=True)
            df_summary.to_excel(writer, sheet_name="Summary", index=False)

    print(f"✅ Row comparison report exported: {ROW_OUTPUT_FILE}")
