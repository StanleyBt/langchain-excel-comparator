# utils/comparison_engine.py
import pandas as pd
import re
from typing import Dict, List, Optional, Any, Tuple
from models.schemas import RowComparisonResult, ColumnComparison
from config import TEXT_ONLY_COLUMNS


def format_column_name(column_name: str) -> str:
    """
    Convert a column name to a clean, readable format.
    
    Examples:
        "EMPLOYEE_NUMBER" -> "Employee Number"
        "employee number" -> "Employee Number"
        "invoice_value" -> "Invoice Value"
        "CONTRACTOR" -> "Contractor"
        "previous_month_calendar_days" -> "Previous Month Calendar Days"
    
    Args:
        column_name: Raw column name (can be vendor or system column name)
        
    Returns:
        Formatted column name with proper capitalization
    """
    if not column_name:
        return ""
    
    # Convert to string and strip
    col = str(column_name).strip()
    
    # Replace underscores and multiple spaces with single space
    col = re.sub(r'[_\s]+', ' ', col)
    
    # Split by spaces and capitalize each word
    words = col.split()
    formatted_words = []
    
    for word in words:
        # Convert to lowercase first, then capitalize first letter
        word_lower = word.lower()
        # Capitalize first letter of each word
        formatted_word = word_lower.capitalize()
        formatted_words.append(formatted_word)
    
    return ' '.join(formatted_words)


def normalize_text(val: Any) -> str:
    """Normalize text for comparison"""
    if val is None or pd.isna(val):
        return ""
    # Convert to string, strip whitespace, normalize spaces
    text = str(val).strip() 
    if not text:
        return ""
    return " ".join(text.lower().split())


def normalize_column_name_for_check(column_name: str) -> str:
    """
    Normalize column name for checking against TEXT_ONLY_COLUMNS list.
    Converts to lowercase, removes spaces/underscores for flexible matching.
    
    Args:
        column_name: Column name to normalize
        
    Returns:
        Normalized column name for comparison
    """
    if not column_name:
        return ""
    # Convert to lowercase, replace spaces/underscores with nothing, strip
    normalized = str(column_name).strip().lower().replace(" ", "").replace("_", "")
    return normalized


def is_text_only_column(column_name: str) -> bool:
    """
    Check if a column should always be treated as text (never as numeric).
    
    This ensures columns like EMPLOYEE_NUMBER are always compared as text,
    preserving leading zeros and preventing numeric conversion issues.
    
    Args:
        column_name: Column name to check
        
    Returns:
        True if column should be treated as text-only, False otherwise
    """
    if not column_name:
        return False
    
    # Normalize the input column name
    col_normalized = normalize_column_name_for_check(column_name)
    
    # Check against all text-only column patterns
    for text_only_pattern in TEXT_ONLY_COLUMNS:
        pattern_normalized = normalize_column_name_for_check(text_only_pattern)
        
        # Check for exact match or if pattern is contained in column name
        if (col_normalized == pattern_normalized or 
            pattern_normalized in col_normalized or 
            col_normalized in pattern_normalized):
            return True
    
    return False


def is_numeric_match(val1: Any, val2: Any, tolerance: float = 2.00) -> Tuple[bool, Optional[float]]:
    """
    Check if two numeric values match within tolerance.
    
    Returns:
        Tuple of (is_match, difference)
    """
    try:
        v1 = float(val1) if val1 is not None and not pd.isna(val1) else None
        v2 = float(val2) if val2 is not None and not pd.isna(val2) else None
        
        if v1 is None or v2 is None:
            return False, None
        
        diff = v1 - v2
        is_match = abs(diff) <= tolerance
        return is_match, diff
    except (ValueError, TypeError):
        return False, None


def is_text_match(val1: Any, val2: Any) -> bool:
    """Check if two text values match (normalized)"""
    norm1 = normalize_text(val1)
    norm2 = normalize_text(val2)
    # Both empty/null are considered matching
    if not norm1 and not norm2:
        return True
    # Both must be non-empty and equal
    return norm1 == norm2 and norm1 != ""


def compare_column_values(
    vendor_value: Any,
    system_value: Any,
    column_name: str
) -> ColumnComparison:
    """
    Compare values from vendor and system for a single column.
    
    Employee numbers and other text-only columns are always treated as text
    to preserve leading zeros and prevent numeric conversion issues.
    
    Args:
        vendor_value: Value from vendor paysheet
        system_value: Value from system paysheet
        column_name: Name of the column being compared
        
    Returns:
        ColumnComparison object with comparison results
    """
    # Check if this column should always be treated as text (e.g., EMPLOYEE_NUMBER)
    force_text = is_text_only_column(column_name)
    
    # Determine if column is numeric based on values (only if not forced to text)
    is_numeric = False
    if not force_text:
        try:
            if vendor_value is not None and not pd.isna(vendor_value):
                float(vendor_value)
            if system_value is not None and not pd.isna(system_value):
                float(system_value)
            is_numeric = True
        except (ValueError, TypeError):
            pass
    
    # Handle None/NaN values
    vendor_is_none = vendor_value is None or pd.isna(vendor_value)
    system_is_none = system_value is None or pd.isna(system_value)
    
    if vendor_is_none and system_is_none:
        # Both None/NaN - consider as match
        return ColumnComparison(
            columnName=column_name,
            vendorValue=vendor_value,
            systemValue=system_value,
            difference=None,
            isMatch=True,
            matchType="both_none"
        )
    
    if is_numeric and not force_text:
        is_match, difference = is_numeric_match(vendor_value, system_value)
        match_type = "numeric_tolerance" if is_match else "numeric_mismatch"
    else:
        # Always use text matching for text-only columns or non-numeric values
        # For text-only columns, preserve original string representation
        if force_text:
            # For text-only columns, compare as exact strings (preserve case/format)
            vendor_str = str(vendor_value).strip() if vendor_value is not None and not pd.isna(vendor_value) else ""
            system_str = str(system_value).strip() if system_value is not None and not pd.isna(system_value) else ""
            is_match = vendor_str == system_str
            match_type = "text_exact" if is_match else "text_mismatch"
        else:
            # For regular text columns, use normalized comparison
            is_match = is_text_match(vendor_value, system_value)
            match_type = "text_normalized" if is_match else "text_mismatch"
        difference = None
    
    return ColumnComparison(
        columnName=column_name,
        vendorValue=vendor_value,
        systemValue=system_value,
        difference=difference,
        isMatch=is_match,
        matchType=match_type
    )


def compare_rows(
    df_vendor: pd.DataFrame,
    df_system: pd.DataFrame,
    header_mapping: Dict[str, str],
    employee_number_vendor_col: Optional[str] = None,
    employee_number_system_col: Optional[str] = None
) -> Dict[str, Any]:
    """
    Compare rows between vendor and system DataFrames row by row.
    
    Returns:
        Dictionary with:
        - "results": List of RowComparisonResult for matched employees
        - "only_in_vendor_count": Count of employees only in vendor
        - "only_in_system_count": Count of employees only in system
    
    Args:
        df_vendor: Vendor paysheet DataFrame
        df_system: System paysheet DataFrame
        header_mapping: Dictionary mapping vendor_header -> system_header
        employee_number_vendor_col: Vendor column name for employee number (auto-detect if None)
        employee_number_system_col: System column name for employee number (auto-detect if None)
        
    Returns:
        List of RowComparisonResult objects
    """
    # Auto-detect employee number columns if not provided
    if not employee_number_vendor_col:
        for col in df_vendor.columns:
            col_lower = str(col).lower().replace(" ", "_")
            if "employee" in col_lower and ("number" in col_lower or "id" in col_lower):
                employee_number_vendor_col = col
                break
    
    if not employee_number_system_col:
        # Try to find from header mapping first
        if employee_number_vendor_col and employee_number_vendor_col in header_mapping:
            employee_number_system_col = header_mapping[employee_number_vendor_col]
        else:
            for col in df_system.columns:
                col_lower = str(col).lower().replace(" ", "_")
                if "employee" in col_lower and ("number" in col_lower or "id" in col_lower):
                    employee_number_system_col = col
                    break
    
    if not employee_number_vendor_col or not employee_number_system_col:
        raise ValueError("Could not find employee number columns in vendor or system data")
    
    # Convert employee numbers to string for matching
    df_vendor[employee_number_vendor_col] = df_vendor[employee_number_vendor_col].astype(str)
    df_system[employee_number_system_col] = df_system[employee_number_system_col].astype(str)
    
    # Get sets of employee numbers
    vendor_emp_ids = set(df_vendor[employee_number_vendor_col].dropna().astype(str))
    system_emp_ids = set(df_system[employee_number_system_col].dropna().astype(str))
    
    matched_emp_ids = vendor_emp_ids & system_emp_ids
    only_in_vendor = vendor_emp_ids - system_emp_ids
    only_in_system = system_emp_ids - vendor_emp_ids
    
    results = []
    # Log employee matching summary
    print(f"📊 Employee Matching: {len(matched_emp_ids)} matched, {len(only_in_vendor)} only in vendor, {len(only_in_system)} only in system")
    
    # Show detailed debug for first 3 employees only, or if there are mismatches
    show_detailed_debug = len(matched_emp_ids) <= 3
    detailed_count = 0
    mismatch_count = 0
    
    for idx, emp_id in enumerate(matched_emp_ids):
        vendor_rows = df_vendor[df_vendor[employee_number_vendor_col] == emp_id]
        system_rows = df_system[df_system[employee_number_system_col] == emp_id]
        
        if len(vendor_rows) > 1:
            print(f"⚠️  WARNING: Employee {emp_id} has {len(vendor_rows)} rows in vendor data! Using first row.")
        if len(system_rows) > 1:
            print(f"⚠️  WARNING: Employee {emp_id} has {len(system_rows)} rows in system data! Using first row.")
        
        vendor_row = vendor_rows.iloc[0]
        system_row = system_rows.iloc[0]
        
        # Only show detailed debug for very small datasets
        should_show_debug = show_detailed_debug
        
        column_comparisons = []
        all_match = True
        employee_mismatches = []
        
        # Add employee number as the first column comparison
        emp_num_comp = compare_column_values(
            str(emp_id),
            str(emp_id),
            "Employee Number"
        )
        column_comparisons.append(emp_num_comp)
        
        # Compare each mapped column
        for vendor_col, system_col in header_mapping.items():
            # Skip employee number column if it's in the mapping (already added manually)
            # Normalize column names for comparison
            vendor_col_normalized = str(vendor_col).strip().lower().replace(" ", "_")
            system_col_normalized = str(system_col).strip().lower().replace(" ", "_")
            emp_vendor_normalized = str(employee_number_vendor_col).strip().lower().replace(" ", "_")
            emp_system_normalized = str(employee_number_system_col).strip().lower().replace(" ", "_")
            
            if (vendor_col_normalized == emp_vendor_normalized or 
                system_col_normalized == emp_system_normalized):
                continue
            
            if vendor_col in df_vendor.columns and system_col in df_system.columns:
                # Get values using .get() with explicit column access
                vendor_val = vendor_row[vendor_col] if vendor_col in vendor_row.index else None
                system_val = system_row[system_col] if system_col in system_row.index else None
                
                # Use clean, formatted system column name for display
                display_col_name = format_column_name(system_col)
                
                # Check if columns exist in DataFrames (warn only)
                if vendor_col not in df_vendor.columns:
                    print(f"⚠️  Warning: Vendor column '{vendor_col}' not found")
                if system_col not in df_system.columns:
                    print(f"⚠️  Warning: System column '{system_col}' not found")
                
                col_comp = compare_column_values(vendor_val, system_val, display_col_name)
                
                if not col_comp.isMatch:
                    all_match = False
                    employee_mismatches.append(f"{display_col_name}: vendor={vendor_val}, system={system_val}, diff={col_comp.difference}")
                
                column_comparisons.append(col_comp)
        
        # Silent processing - no per-employee debug output
        
        results.append(RowComparisonResult(
            employeeNumber=str(emp_id),
            columnComparisons=column_comparisons,
            overallMatch=all_match,
            rowStatus="matched"
        ))
    
    # Log comparison summary
    perfect_matches = sum(1 for r in results if r.overallMatch)
    mismatches = sum(1 for r in results if not r.overallMatch)
    print(f"📊 Comparison: {perfect_matches} perfect matches, {mismatches} with mismatches")
    
    # Log unmatched employees (for information only, not included in response)
    if only_in_vendor:
        print(f"ℹ️  {len(only_in_vendor)} employees only in vendor (not included in response)")
    if only_in_system:
        print(f"ℹ️  {len(only_in_system)} employees only in system (not included in response)")
    
    # Return matched employees and counts of unmatched employees
    return {
        "results": results,
        "only_in_vendor_count": len(only_in_vendor),
        "only_in_system_count": len(only_in_system)
    }

