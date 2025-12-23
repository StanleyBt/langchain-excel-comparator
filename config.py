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
    "INVOICE_AMOUNT",
    "PREVIOUS_MONTH_CALENDAR_DAYS",
    "BASE_VALUE",
    "CONTRACTOR"
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
    "CONTRACTOR": [
        ["contractor", "name"],      # Highest priority: contains both "contractor" and "name"
        ["contractor"],              # Medium priority: just "contractor"
        ["contractor", "id"],        # Lower priority: "contractor" and "id"
        ["vendor", "name"]           # Alternative: "vendor" and "name"
    ],
    "EMPLOYEE_NUMBER": [
        ["employee", "number"],      # Best: "employee" and "number"
        ["employee", "no"],          # Good: "employee" and "no"
        ["employee", "id"]           # Less preferred: "employee" and "id"
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
    "INVOICE_AMOUNT": ["invoice_amount", "invoice amount", "invoiceamount"],
    "PREVIOUS_MONTH_CALENDAR_DAYS": [],
    "BASE_VALUE": ["base_value", "base value", "basevalue"],
    "CONTRACTOR": ["contractor", "contractor name", "contractorname", "contractor_id", "contractor id", "vendor", "vendor name"]
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
    "INVOICE_AMOUNT": ["invoice", "amount"],
    "PREVIOUS_MONTH_CALENDAR_DAYS": ["previous", "month", "calendar", "days", "payable", "work"],
    "BASE_VALUE": ["base", "value"],
    "CONTRACTOR": ["contractor", "vendor"]
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
    "CONTRACTOR": [
        "id",           # Exclude "contractor id", "contractor_id" - these are identifiers, not names
        "number",       # Exclude "contractor number"
        "code"          # Exclude "contractor code"
    ],
    "EMPLOYEE_NUMBER": [
        "name"          # Exclude "employee name" - we want number/id, not name
    ]
    # Add more exclusion patterns as needed
}
