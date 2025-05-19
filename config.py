# excel_comparator/config.py
# === Configuration Constants ===

import streamlit as st

# Input Excel file path
EXCEL_PATH = "data\Regular paysheet - feb'25.xlsx"

# Sheet names
VENDOR_SHEET = "Vendor 1"
SYSTEM_SHEET = "system"

# Output report path
OUTPUT_FILE = "output/header_comparison_report.xlsx"
ROW_OUTPUT_FILE = "output/row_comparison_report.xlsx"


# Azure OpenAI configuration
AZURE_API_KEY = st.secrets["AZURE_API_KEY"]
AZURE_ENDPOINT = st.secrets["AZURE_ENDPOINT"]
AZURE_DEPLOYMENT_NAME = st.secrets["AZURE_DEPLOYMENT_NAME"]
AZURE_API_VERSION = st.secrets["AZURE_API_VERSION"]
