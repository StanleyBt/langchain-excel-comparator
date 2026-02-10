# config.py
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
# COLUMNS TO CHECK (single source of truth)
# ============================================================================
# Add or remove columns here. Each entry defines one column that will be
# matched between vendor and system paysheets and included in comparison.
#
# To ADD a column: append a new dict with name, text_only, variations, priority_keywords.
# To REMOVE a column: delete its dict.
#
# Fields per column:
#   name: Canonical header name (case-insensitive). Used in API responses.
#   text_only: True = compare as text (e.g. IDs, codes); False = allow numeric comparison/tolerance.
#   variations: Possible names in vendor files (used for matching). Include common spellings.
#   priority_keywords: List of [keyword_lists]. Each keyword_list is a list of words; vendor headers
#       containing all words in a list get that priority (index 0 = best). Used for AI hints only.
# ============================================================================
CONSTANT_HEADER_CONFIG: List[Dict] = [
    {
        "name": "EMPLOYEE_NUMBER",
        "text_only": True,
        "variations": ["employee_number", "employee number", "employeenumber", "employee id", "employeeid", "employee no", "employee_no"],
        "priority_keywords": [
            ["employee", "number"],
            ["employee", "no"],
            ["employee", "id"],
        ],
    },
    {
        "name": "NET_PAY",
        "text_only": False,
        "variations": ["net_pay", "net pay", "netpay", "net salary", "netsalary", "net_amount", "net amount"],
        "priority_keywords": [
            ["net", "pay"],
            ["netpay"],
            ["net", "salary"],
        ],
    },
    {
        "name": "INVOICE_VALUE",
        "text_only": False,
        "variations": ["invoice_value", "invoice value", "invoicevalue", "invoice_amount", "invoice amount", "invoiceamount"],
        "priority_keywords": [
            ["invoice", "value"],
            ["invoice", "amount"],
            ["invoicevalue"],
        ],
    },
]

# Derived from CONSTANT_HEADER_CONFIG (do not edit these directly)
CONSTANT_HEADERS: List[str] = [c["name"] for c in CONSTANT_HEADER_CONFIG]
HEADER_VARIATIONS: Dict[str, List[str]] = {c["name"]: c["variations"] for c in CONSTANT_HEADER_CONFIG}
HEADER_MATCHING_PRIORITY: Dict[str, List[List[str]]] = {c["name"]: c["priority_keywords"] for c in CONSTANT_HEADER_CONFIG}
TEXT_ONLY_CONSTANT_HEADERS: set = {c["name"] for c in CONSTANT_HEADER_CONFIG if c.get("text_only", False)}

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
