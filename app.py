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

@st.cache_data(show_spinner=False)
def load_system_file(file):
    preview = pd.read_excel(file, sheet_name=0, nrows=10, header=None)
    header_row = find_header_row(preview)
    df_full = pd.read_excel(file, sheet_name=0, header=header_row)
    df_full.columns = [str(col).strip().lower() for col in df_full.columns]
    return df_full

@st.cache_data(show_spinner=False)
def load_vendor_sheet(file, sheet_name, header_row):
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    df.columns = [str(col).strip().lower() for col in df.columns]
    return df


def normalize_text(val):
    return " ".join(str(val).strip().lower().split())


def is_match(val1, val2, col_name):
    if col_name.lower() == "employee name":
        return normalize_text(val1) == normalize_text(val2)
    try:
        v = float(val1)
        s = float(val2)
        return abs(v - s) <= 2
    except:
        return False


if vendor_file and system_file and st.button("🔍 Run Batch Comparison"):
    try:
        df_system_full = load_system_file(system_file)
        vendor_xl = pd.ExcelFile(vendor_file)
        headcount_data = []
        all_comparisons = {}

        for sheet_name in vendor_xl.sheet_names:
            final_mapping, df_mapping, vendor_header_row, _ = run_header_mapping(
                vendor_file, system_file, sheet_name, 0
            )

            df_vendor = load_vendor_sheet(vendor_file, sheet_name, vendor_header_row)

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

            # === Vendor name detection ===
            vendor_col = next((col for col in df_vendor.columns if any(k in col for k in ["vendor", "contractor"])), None)
            vendor_name = None
            if vendor_col:
                try:
                    vendor_name = df_vendor[vendor_col].dropna().astype(str).str.lower().str.strip().mode().iloc[0]
                except:
                    pass

            if not vendor_name:
                st.warning(f"⚠️ Could not detect vendor name from data in sheet '{sheet_name}'. Skipping.")
                continue

            system_filtered = df_system_full.copy()
            system_vendor_col = next((col for col in df_system_full.columns if any(k in col for k in ["vendor", "contractor"])), None)
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

        st.session_state["comparisons"] = all_comparisons
        st.session_state["df_headcount"] = pd.DataFrame(headcount_data)

    except Exception as e:
        st.error(f"❌ Error during batch processing: {e}")

if "df_headcount" in st.session_state:
    st.subheader("👥 Headcount Summary")
    st.dataframe(st.session_state["df_headcount"], use_container_width=True)

# === Step 7: Column Comparison ===
if "comparisons" in st.session_state:
    st.markdown("---")
    st.subheader("📊 Column Comparison Preview")
    vendor_list = list(st.session_state["comparisons"].keys())
    selected_vendor = st.selectbox("Choose a vendor to compare:", vendor_list, key="vendor_selector")

    data = st.session_state["comparisons"][selected_vendor]
    df_vendor = data["vendor"]
    df_system = data["system"]
    mapping = data["mapping"]
    emp_col = data["emp_col"]

    valid_cols = [k for k, v in mapping.items() if k in df_vendor.columns and v in df_system.columns]
    default_selection = "invoice value" if "invoice value" in valid_cols else valid_cols[:1]
    selected_cols = st.multiselect("Select columns to compare:", valid_cols, default=default_selection, key="col_selector")

    if selected_cols:
        results = []
        for emp_id in df_vendor.index:
            vendor_row = df_vendor.loc[emp_id]
            system_row = df_system.loc[emp_id] if emp_id in df_system.index else pd.Series()

            for col in selected_cols:
                v_val = vendor_row.get(col)
                s_col = mapping[col]
                s_val = system_row.get(s_col)

                match = is_match(v_val, s_val, col)
                diff = None
                try:
                    diff = float(v_val) - float(s_val)
                except:
                    pass

                results.append({
                    "Employee ID": emp_id,
                    f"{col} | Vendor Value": v_val,
                    f"{col} | System Value": s_val,
                    f"{col} | Difference": diff,
                    f"{col} | Match": "✔️" if match else "❌",
                    "": ""  # Spacer
                })

        df_final = pd.DataFrame(results)
        st.dataframe(df_final, use_container_width=True)

        match_cols = [col for col in df_final.columns if col.endswith("| Match")]
        if match_cols:
            total = len(df_final)
            correct = df_final[match_cols].apply(lambda row: sum(val == "✔️" for val in row), axis=1).sum()
            st.info(f"✅ Overall Match Rate: {correct / (total * len(match_cols)):.2%}")
