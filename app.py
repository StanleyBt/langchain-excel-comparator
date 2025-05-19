# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from utils.header_matching import find_header_row, run_header_mapping

st.set_page_config(page_title="Paysheet Comparator (Batch Mode)", layout="wide")
st.title("📁 Multi-Vendor Paysheet Comparator")

# === File Upload ===
vendor_file = st.file_uploader("Upload Vendor Paysheet (Multiple Sheets)", type=["xlsx"])
system_file = st.file_uploader("Upload System Paysheet", type=["xlsx"])

def is_match(val1, val2):
    try:
        v = str(val1).strip().split(".")[0]
        s = str(val2).strip().split(".")[0]
        return v == s
    except:
        return False

if vendor_file and system_file and st.button("🔍 Run Batch Comparison"):
    try:
        preview_system = pd.read_excel(system_file, sheet_name=0, nrows=10, header=None)
        system_header_row = find_header_row(preview_system)
        df_system_full = pd.read_excel(system_file, sheet_name=0, header=system_header_row)
        df_system_full.columns = [str(col).strip().lower() for col in df_system_full.columns]

        vendor_xl = pd.ExcelFile(vendor_file)
        headcount_data = []
        all_comparisons = {}

        for sheet_name in vendor_xl.sheet_names:
            final_mapping, df_mapping, vendor_header_row, _ = run_header_mapping(
                vendor_file, system_file, sheet_name, 0
            )

            df_vendor = pd.read_excel(vendor_file, sheet_name=sheet_name, header=vendor_header_row)
            df_vendor.columns = [str(col).strip().lower() for col in df_vendor.columns]

            emp_col = None
            system_emp_col = None
            for vendor_col, system_col in final_mapping.items():
                norm = vendor_col.replace(" ", "").lower()
                if any(key in norm for key in ["employeeid", "employeeno", "employeenumber", "bluetreeid"]):
                    emp_col = vendor_col
                    system_emp_col = system_col
                    break

            if not emp_col or not system_emp_col:
                st.warning(f"⚠️ No employee ID mapping for vendor sheet '{sheet_name}'. Skipping.")
                continue

            vendor_name = None
            possible_vendor_cols = [col for col in df_vendor.columns if "vendor" in col or "contractor" in col]
            if possible_vendor_cols:
                vendor_col = possible_vendor_cols[0]
                vendor_name = df_vendor[vendor_col].dropna().astype(str).str.lower().str.strip().mode().iloc[0]

            if not vendor_name:
                st.warning(f"⚠️ Could not detect vendor name from data in sheet '{sheet_name}'. Skipping.")
                continue

            system_filtered = df_system_full.copy()
            system_vendor_col = next((col for col in df_system_full.columns if "vendor" in col or "contractor" in col), None)
            if system_vendor_col:
                system_filtered = df_system_full[df_system_full[system_vendor_col].astype(str).str.lower().str.strip() == vendor_name]

            vendor_emp_ids = set(df_vendor[emp_col].dropna().astype(str))
            system_emp_ids = set(system_filtered[system_emp_col].dropna().astype(str))
            matched_ids = vendor_emp_ids & system_emp_ids
            only_in_vendor = vendor_emp_ids - system_emp_ids
            only_in_system = system_emp_ids - vendor_emp_ids

            headcount_data.append({
                "Vendor": vendor_name.title(),
                "Vendor Count": len(vendor_emp_ids),
                "System Count": len(system_emp_ids),
                "Matching": len(matched_ids),
                "Only in Vendor": len(only_in_vendor),
                "Only in System": len(only_in_system)
            })

            df_vendor_match = df_vendor[df_vendor[emp_col].astype(str).isin(matched_ids)].set_index(emp_col)
            df_system_match = system_filtered[system_filtered[system_emp_col].astype(str).isin(matched_ids)].set_index(system_emp_col)

            all_comparisons[sheet_name] = {
                "vendor": df_vendor_match,
                "system": df_system_match,
                "mapping": final_mapping,
                "df_mapping": df_mapping,
                "emp_col": emp_col,
                "system_emp_col": system_emp_col
            }

        st.subheader("👥 Headcount Summary")
        df_headcount = pd.DataFrame(headcount_data)
        st.dataframe(df_headcount, use_container_width=True)

        st.session_state["comparisons"] = all_comparisons
        st.session_state["df_headcount"] = df_headcount

    except Exception as e:
        st.error(f"❌ Error during batch processing: {e}")

# === Step 7: Column Comparison ===
if "comparisons" in st.session_state:
    st.markdown("---")
    st.subheader("📊 Column Comparison Preview")
    vendor_list = list(st.session_state["comparisons"].keys())
    selected_vendor = st.selectbox("Choose a vendor to compare:", vendor_list)

    data = st.session_state["comparisons"][selected_vendor]
    df_vendor = data["vendor"]
    df_system = data["system"]
    mapping = data["mapping"]
    emp_col = data["emp_col"]

    valid_cols = [k for k, v in mapping.items() if k in df_vendor.columns and v in df_system.columns]
    selected_cols = st.multiselect("Select columns to compare:", valid_cols, default=valid_cols[:5])

    if selected_cols:
        rows = []
        for emp_id in df_vendor.index:
            for col in selected_cols:
                v_val = df_vendor.at[emp_id, col] if col in df_vendor.columns else None
                s_col = mapping[col]
                s_val = df_system.at[emp_id, s_col] if s_col in df_system.columns else None

                match = is_match(v_val, s_val)
                diff = None
                try:
                    v_num = float(str(v_val).strip().split(".")[0])
                    s_num = float(str(s_val).strip().split(".")[0])
                    diff = v_num - s_num
                except:
                    pass

                rows.append({
                    "Employee ID": emp_id,
                    f"{col} | Vendor Value": v_val,
                    f"{col} | System Value": s_val,
                    f"{col} | Difference": diff,
                    f"{col} | Match": "✔️" if match else "❌",
                    "": ""  # Spacer
                })

        df_final = pd.DataFrame(rows)
        show_only_mismatches = st.toggle("Show only mismatches", value=True)

        if show_only_mismatches:
            mismatch_cols = [col for col in df_final.columns if col.endswith("| Match")]
            mask = df_final[mismatch_cols].apply(lambda row: any(val == "❌" for val in row), axis=1)
            df_final = df_final[mask]

        st.dataframe(df_final, use_container_width=True)

        match_cols = [col for col in df_final.columns if col.endswith("| Match")]
        if match_cols:
            total = len(df_final)
            correct = df_final[match_cols].apply(lambda row: sum(val == "✔️" for val in row), axis=1).sum()
            st.info(f"✅ Overall Match Rate: {correct / (total * len(match_cols)):.2%}")
