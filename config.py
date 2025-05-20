# excel_comparator/config.py
# === Configuration Constants ===

import streamlit as st

# Azure OpenAI configuration
AZURE_API_KEY = st.secrets["AZURE_API_KEY"]
AZURE_ENDPOINT = st.secrets["AZURE_ENDPOINT"]
AZURE_DEPLOYMENT_NAME = st.secrets["AZURE_DEPLOYMENT_NAME"]
AZURE_API_VERSION = st.secrets["AZURE_API_VERSION"]
