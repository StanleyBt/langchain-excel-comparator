# excel_comparator/config.py
# === Configuration Constants ===

import os
from typing import Optional, List, Dict
from dotenv import load_dotenv

# Load environment variables from .env file if it exists
load_dotenv()

# Azure OpenAI configuration
# Load from environment variables (.env file or system environment)
AZURE_API_KEY = os.getenv("AZURE_API_KEY")
AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT")
AZURE_DEPLOYMENT_NAME = os.getenv("AZURE_DEPLOYMENT_NAME")
AZURE_API_VERSION = os.getenv("AZURE_API_VERSION")

# ============================================================================
# CONSTANT HEADERS CONFIGURATION
# ============================================================================
# These are the only headers that will be matched and returned in the response.
# You can add, remove, or modify these headers as needed.
# 
# Format: List of header names (case-insensitive, will be normalized)
# ============================================================================
CONSTANT_HEADERS: List[str] = [
    "EMPLOYEE_NUMBER",
    "NET_PAY",
    "INVOICE_VALUE"
]

# ============================================================================
# HEADER MATCHING PRIORITY CONFIGURATION
# ============================================================================
# Define priority keywords for matching vendor headers to constant headers.
# Headers matching higher priority keywords (lower index) are preferred.
# 
# Format: Dict[constant_header_name, List[List[keywords]]]
# - Each inner list represents a priority level (index 0 = highest priority)
# - Headers containing ALL keywords in a list get that priority level
# ============================================================================
HEADER_MATCHING_PRIORITY: Dict[str, List[List[str]]] = {
    "EMPLOYEE_NUMBER": [
        ["employee", "number"],      # Best: "employee" and "number"
        ["employee", "no"],          # Good: "employee" and "no"
        ["employee", "id"]           # Less preferred: "employee" and "id"
    ],
    "NET_PAY": [
        ["net", "pay"],              # Best: "net" and "pay"
        ["netpay"],                  # Alternative: "netpay"
        ["net", "salary"]            # Alternative: "net" and "salary"
    ],
    "INVOICE_VALUE": [
        ["invoice", "value"],        # Best: "invoice" and "value"
        ["invoice", "amount"],       # Alternative: "invoice" and "amount"
        ["invoicevalue"]             # Alternative: "invoicevalue"
    ]
    # Add more priority rules for other constant headers as needed
}

# ============================================================================
# HEADER VARIATION MAPPINGS
# ============================================================================
# Define possible variations for each constant header.
# Used for matching vendor headers to system headers.
# 
# Format: Dict[constant_header_name, List[variation_strings]]
# ============================================================================
HEADER_VARIATIONS: Dict[str, List[str]] = {
    "EMPLOYEE_NUMBER": ["employee_number", "employee number", "employeenumber", "employee id", "employeeid", "employee no", "employee_no"],
    "NET_PAY": ["net_pay", "net pay", "netpay", "net salary", "netsalary", "net_amount", "net amount"],
    "INVOICE_VALUE": ["invoice_value", "invoice value", "invoicevalue", "invoice_amount", "invoice amount", "invoiceamount"]
}

# ============================================================================
# AI MATCHING KEYWORDS
# ============================================================================
# Keywords used for AI semantic matching when exact matches fail.
# These help identify potential matches for each constant header.
# 
# Format: Dict[constant_header_name, List[keyword_strings]]
# ============================================================================
AI_MATCHING_KEYWORDS: Dict[str, List[str]] = {
    "EMPLOYEE_NUMBER": ["employee", "emp", "id", "number", "no"],
    "NET_PAY": ["net", "pay", "salary", "amount", "netpay"],
    "INVOICE_VALUE": ["invoice", "value", "amount"]
}

# ============================================================================
# EXCLUDED MATCHING PATTERNS
# ============================================================================
# Patterns that should NOT be auto-matched to constant headers.
# If a vendor header contains these patterns, it will be excluded from
# automatic matching and require manual mapping.
# 
# Format: Dict[constant_header_name, List[exclusion_patterns]]
# - Patterns are checked in normalized header names
# - If vendor header contains any exclusion pattern, it won't be auto-matched
# ============================================================================
EXCLUDED_MATCHING_PATTERNS: Dict[str, List[str]] = {
    "EMPLOYEE_NUMBER": [
        "name"          # Exclude "employee name" - we want number/id, not name
    ],
    "NET_PAY": [
        "gross",        # Exclude "gross pay" - we want net pay
        "basic",        # Exclude "basic pay" - we want net pay
        "total",        # Exclude "total pay", "total payable" - we want net pay (after deductions)
        "payable"       # Exclude "total payable", "amount payable" - different from net pay
    ]
    # Add more exclusion patterns as needed
}

# ============================================================================
# TEXT-ONLY COLUMNS CONFIGURATION
# ============================================================================
# Columns that should ALWAYS be treated as text, never as numbers.
# This prevents numeric conversion issues (e.g., leading zeros lost, 
# scientific notation, numeric tolerance applied).
# 
# These columns will be compared as exact text matches only.
# 
# Format: List of column name patterns (case-insensitive, will be normalized)
# - Can use exact matches or partial patterns
# - Employee numbers are included by default to preserve leading zeros
# ============================================================================
TEXT_ONLY_COLUMNS: List[str] = [
    "EMPLOYEE_NUMBER",           # Always treat employee numbers as text
    "employee_number",           # Lowercase variant
    "employee number",           # Space-separated variant
    "employeeid",                # No space variant
    "employee_id",               # Underscore variant
    "employee no",               # "no" variant
    "employee_no",               # Underscore "no" variant
    "CONTRACTOR",                # Always treat contractor names as text
    "contractor",                # Lowercase variant  
    "contractor name",           # Space-separated variant
    "contractorname",            # No space variant
    "vendor",                    # Alternative name
    "vendor name",               # Alternative with space
    "vendorname",                # Alternative no space
    # Add more text-only columns as needed
    # Examples: "CONTRACTOR_ID", "VENDOR_CODE", "REFERENCE_NUMBER", etc.
]
