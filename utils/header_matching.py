# excel_comparator/utils/header_matching.py
import pandas as pd
import json
import re
from langchain_openai import AzureChatOpenAI
from langchain.schema import HumanMessage
from config import AZURE_DEPLOYMENT_NAME, AZURE_ENDPOINT, AZURE_API_KEY, AZURE_API_VERSION

def find_header_row(df_raw, max_rows=10):
    """
    Detects the likely header row by scanning first N rows.
    Picks the row with the most non-empty values.
    """
    for i in range(max_rows):
        row = df_raw.iloc[i]
        non_empty = row.notna().sum()
        if non_empty / len(row) > 0.5:
            return i
    return 0  # fallback if no strong candidate


def detect_header_row(df_preview):
    non_empty_counts = df_preview.notna().sum(axis=1)
    return non_empty_counts.idxmax()


def run_header_mapping(vendor_file, system_file, vendor_sheet, system_sheet, *_):
    # === Step 1: Detect header row for both files ===
    preview_vendor = pd.read_excel(vendor_file, sheet_name=vendor_sheet, nrows=10, header=None)
    preview_system = pd.read_excel(system_file, sheet_name=system_sheet, nrows=10, header=None)

    vendor_header_row = detect_header_row(preview_vendor)
    system_header_row = detect_header_row(preview_system)

    # === Step 2: Load headers ===
    df_vendor = pd.read_excel(vendor_file, sheet_name=vendor_sheet, header=vendor_header_row, nrows=1)
    df_system = pd.read_excel(system_file, sheet_name=system_sheet, header=system_header_row, nrows=1)

    headers_vendor = [col.strip().lower() for col in df_vendor.columns]
    headers_system = [col.strip().lower() for col in df_system.columns]

    exact_matches = {}
    unmatched_vendor = []

    for h in headers_vendor:
        if h in headers_system:
            exact_matches[h] = h
        else:
            unmatched_vendor.append(h)

    unmatched_system = [h for h in headers_system if h not in exact_matches.values()]

    # === GPT Matching ===
    llm = AzureChatOpenAI(
        deployment_name=AZURE_DEPLOYMENT_NAME,
        temperature=0,
        azure_endpoint=AZURE_ENDPOINT,
        api_key=AZURE_API_KEY,
        api_version=AZURE_API_VERSION
    )

    prompt = f"""
You are a data assistant comparing column headers between two Excel sheets.

These are headers from a Vendor Paysheet (File A) that had no exact match:
{unmatched_vendor}

And these are remaining headers from the System Paysheet (File B):
{unmatched_system}

Your task:
- Match each Vendor header to the most semantically similar System header.
- If no good match exists, set the value to null.
- Use domain knowledge of payroll systems. For example:
  - "employee id" could be "employee number", "cems employee id", or "blue tree id".
  - "gross salary" could be "fixed gross" or "ctc".

⚠️ Return output as JSON only.
"""

    response = llm.invoke([HumanMessage(content=prompt)])
    raw_output = response.content

    try:
        semantic_matches = json.loads(raw_output)
    except json.JSONDecodeError:
        match = re.search(r"{.*}", raw_output, re.DOTALL)
        if match:
            semantic_matches = json.loads(match.group())
        else:
            raise ValueError("❌ GPT did not return valid JSON or parsable output.")

    final_mapping = {
        **exact_matches,
        **{k: v for k, v in semantic_matches.items() if v and v.lower() != "no match"}
    }

    # === Prepare Mapping Table for Display ===
    rows = []
    for vendor_col in headers_vendor:
        if vendor_col in exact_matches:
            rows.append({"Vendor Header": vendor_col, "System Header": exact_matches[vendor_col], "Match Type": "Exact"})
        elif vendor_col in semantic_matches and semantic_matches[vendor_col]:
            rows.append({"Vendor Header": vendor_col, "System Header": semantic_matches[vendor_col], "Match Type": "Semantic"})
        else:
            rows.append({"Vendor Header": vendor_col, "System Header": None, "Match Type": "Not Matched"})

    df_mapping = pd.DataFrame(rows)
    return final_mapping, df_mapping, vendor_header_row, system_header_row
