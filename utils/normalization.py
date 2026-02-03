# utils/normalization.py
"""
Centralized normalization functions for the paysheet comparison system.

This module provides unified normalization logic to ensure consistency
across header matching, column comparison, and data processing.
"""
from typing import Any
import pandas as pd


def normalize_header_name(header: str, for_matching: bool = False) -> str:
    """
    Normalize header name for consistent comparison and matching.
    
    Args:
        header: Header name to normalize
        for_matching: If True, use matching normalization (underscores instead of removal)
                     If False, use general normalization (remove spaces/underscores)
    
    Returns:
        Normalized header name
        
    Examples:
        normalize_header_name("EMPLOYEE_NUMBER") -> "employeenumber"
        normalize_header_name("employee number", for_matching=True) -> "employee_number"
    """
    if not header:
        return ""
    
    header_str = str(header).strip().lower()
    
    if for_matching:
        # For matching: replace spaces/dashes/dots with underscores
        # This preserves word boundaries for semantic matching
        return header_str.replace(" ", "_").replace("-", "_").replace(".", "")
    else:
        # For general use: remove spaces and underscores completely
        # This creates a compact identifier
        return header_str.replace(" ", "").replace("_", "")


def normalize_text_value(val: Any) -> str:
    """
    Normalize text value for comparison.
    
    Converts value to string, strips whitespace, and normalizes spaces.
    Used for text-based column comparisons.
    
    Args:
        val: Value to normalize (can be any type)
        
    Returns:
        Normalized text string (empty string if None/NaN)
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    
    # Convert to string, strip whitespace, normalize spaces
    text = str(val).strip()
    if not text:
        return ""
    
    # Normalize multiple spaces to single space, convert to lowercase
    return " ".join(text.lower().split())


def normalize_column_name_for_check(column_name: str) -> str:
    """
    Normalize column name for checking against TEXT_ONLY_COLUMNS list.
    
    This creates a compact identifier by removing all spaces and underscores.
    Used specifically for pattern matching against text-only column configurations.
    
    Args:
        column_name: Column name to normalize
        
    Returns:
        Normalized column name (lowercase, no spaces/underscores)
    """
    if not column_name:
        return ""
    
    # Convert to lowercase, remove spaces/underscores completely
    return str(column_name).strip().lower().replace(" ", "").replace("_", "")


def normalize_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize all column names in a DataFrame to lowercase.
    
    This is the standard normalization applied to all DataFrames
    when they are first loaded. All subsequent operations use
    these normalized column names.
    
    Args:
        df: DataFrame with columns to normalize
        
    Returns:
        DataFrame with normalized column names (same DataFrame, modified in place)
    """
    # Normalize column names: strip whitespace, convert to lowercase
    df.columns = [str(col).strip().lower() for col in df.columns]
    return df


def normalize_for_comparison(header: str) -> str:
    """
    Normalize header for column name comparison operations.
    
    Used when comparing column names that may have different
    formatting (spaces vs underscores, case differences).
    
    Args:
        header: Header name to normalize
        
    Returns:
        Normalized header (lowercase, spaces replaced with underscores)
    """
    if not header:
        return ""
    
    return str(header).strip().lower().replace(" ", "_")
