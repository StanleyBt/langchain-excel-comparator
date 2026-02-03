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

# Azure OpenAI timeout configuration (in seconds)
# Default: 60 seconds. Increase if you have very large prompts (>50 headers) or slow network.
# Typical response time: 5-20 seconds. 60s provides buffer for network delays and service load.
AZURE_OPENAI_TIMEOUT = int(os.getenv("AZURE_OPENAI_TIMEOUT", "60"))

# Azure OpenAI retry configuration
# Number of retries for timeout errors (with exponential backoff)
AZURE_OPENAI_MAX_RETRIES = int(os.getenv("AZURE_OPENAI_MAX_RETRIES", "2"))

# Logging configuration
# Log level: DEBUG, INFO, WARNING, ERROR, CRITICAL
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
# Use JSON structured logging (set to False for development)
LOG_JSON = os.getenv("LOG_JSON", "true").lower() == "true"
# Optional log file path
LOG_FILE = os.getenv("LOG_FILE", None)

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
# IMPORTANT: This is used ONLY for generating hints in Phase 1 of header matching.
# The AI (Phase 2) makes the final matching decision autonomously based on semantic
# understanding. These priorities help the AI prioritize likely candidates, but
# the AI can override them if it finds a better semantic match.
# 
# Format: Dict[constant_header_name, List[List[keywords]]]
# - Each inner list represents a priority level (index 0 = highest priority)
# - Headers containing ALL keywords in a list get that priority level
# - Lower index = higher priority (used for sorting hints)
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
# COLUMN DATA TYPE CONFIGURATION
# ============================================================================
# Define which constant headers should be treated as text-only (never numeric).
# This prevents numeric conversion issues (e.g., leading zeros lost, 
# scientific notation, numeric tolerance applied).
# 
# Format: Set of constant header names that should be text-only
# - These headers will be compared as exact text matches only
# - All variations of these headers (from HEADER_VARIATIONS) are automatically included
# ============================================================================
TEXT_ONLY_CONSTANT_HEADERS: set = {
    "EMPLOYEE_NUMBER",  # IDs should always be text to preserve leading zeros
    # Add more constant headers here that should be text-only
    # Examples: "CONTRACTOR", "VENDOR_CODE", "REFERENCE_NUMBER", etc.
}

# ============================================================================
# ADDITIONAL TEXT-ONLY COLUMN PATTERNS
# ============================================================================
# Additional column name patterns (not in CONSTANT_HEADERS) that should be text-only.
# Use this for columns that aren't constant headers but still need text treatment.
# 
# Format: List of normalized patterns (case-insensitive, spaces/underscores removed)
# ============================================================================
ADDITIONAL_TEXT_ONLY_PATTERNS: List[str] = [
    "contractor",    # Contractor name (if not a constant header)
    "vendor",        # Vendor name (if not a constant header)
    # Add more patterns as needed
    # Examples: "contractorid", "vendorcode", "referencenumber", etc.
]

# ============================================================================
# AUTO-GENERATED: TEXT-ONLY COLUMNS (DO NOT EDIT MANUALLY)
# ============================================================================
# This list is automatically generated from:
# 1. TEXT_ONLY_CONSTANT_HEADERS + their variations from HEADER_VARIATIONS
# 2. ADDITIONAL_TEXT_ONLY_PATTERNS
# ============================================================================
def _generate_text_only_columns() -> List[str]:
    """Generate TEXT_ONLY_COLUMNS list from configuration"""
    patterns = set()
    
    # Add patterns from constant headers marked as text-only
    for const_header in TEXT_ONLY_CONSTANT_HEADERS:
        if const_header in HEADER_VARIATIONS:
            # Add all variations for this constant header
            for variation in HEADER_VARIATIONS[const_header]:
                # Normalize: lowercase, remove spaces/underscores
                normalized = variation.strip().lower().replace(" ", "").replace("_", "")
                patterns.add(normalized)
        else:
            # If no variations defined, normalize the header name itself
            normalized = const_header.strip().lower().replace(" ", "").replace("_", "")
            patterns.add(normalized)
    
    # Add additional patterns
    for pattern in ADDITIONAL_TEXT_ONLY_PATTERNS:
        normalized = pattern.strip().lower().replace(" ", "").replace("_", "")
        patterns.add(normalized)
    
    return sorted(list(patterns))

TEXT_ONLY_COLUMNS: List[str] = _generate_text_only_columns()
