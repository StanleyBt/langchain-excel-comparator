# utils/comparison_engine.py
import pandas as pd
import re
from typing import Dict, List, Optional, Any, Tuple
from models.schemas import RowComparisonResult, ColumnComparison
from config import TEXT_ONLY_COLUMNS
from utils.normalization import normalize_text_value, normalize_column_name_for_check, normalize_for_comparison
from utils.logger import get_logger

logger = get_logger("comparison_engine")


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


# Use centralized normalization functions
# normalize_text -> normalize_text_value (from utils.normalization)
# normalize_column_name_for_check -> normalize_column_name_for_check (from utils.normalization)


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
    norm1 = normalize_text_value(val1)
    norm2 = normalize_text_value(val2)
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
    header_mapping: Dict[Any, str],  # Supports both str (single) and tuple (multiple) keys
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
            col_normalized = normalize_for_comparison(col)
            if "employee" in col_normalized and ("number" in col_normalized or "id" in col_normalized):
                employee_number_vendor_col = col
                break
    
    if not employee_number_system_col:
        # Try to find from header mapping first
        if employee_number_vendor_col and employee_number_vendor_col in header_mapping:
            employee_number_system_col = header_mapping[employee_number_vendor_col]
        else:
            for col in df_system.columns:
                col_normalized = normalize_for_comparison(col)
                if "employee" in col_normalized and ("number" in col_normalized or "id" in col_normalized):
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
    # Log employee matching summary (without sensitive employee IDs)
    logger.info(
        f"Employee matching: {len(matched_emp_ids)} matched, {len(only_in_vendor)} only in vendor, {len(only_in_system)} only in system",
        extra={
            "matched_count": len(matched_emp_ids),
            "only_in_vendor_count": len(only_in_vendor),
            "only_in_system_count": len(only_in_system)
        }
    )
    
    # Show detailed debug for first 3 employees only, or if there are mismatches
    show_detailed_debug = len(matched_emp_ids) <= 3
    detailed_count = 0
    mismatch_count = 0
    
    for idx, emp_id in enumerate(matched_emp_ids):
        vendor_rows = df_vendor[df_vendor[employee_number_vendor_col] == emp_id]
        system_rows = df_system[df_system[employee_number_system_col] == emp_id]
        
        if len(vendor_rows) > 1:
            logger.warning(
                f"Employee has {len(vendor_rows)} rows in vendor data, using first row",
                extra={"row_count": len(vendor_rows)}  # Don't log employee ID
            )
        if len(system_rows) > 1:
            logger.warning(
                f"Employee has {len(system_rows)} rows in system data, using first row",
                extra={"row_count": len(system_rows)}  # Don't log employee ID
            )
        
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
            "Employee Head Count"
        )
        column_comparisons.append(emp_num_comp)
        
        # Compare each mapped column
        for vendor_col_key, system_col in header_mapping.items():
            # Handle both single column (str) and multiple columns (tuple)
            is_multi_column = isinstance(vendor_col_key, tuple)
            
            if is_multi_column:
                # Multiple vendor columns - sum them
                vendor_cols = list(vendor_col_key)
                vendor_val = None
                vendor_values = []
                
                # Sum all vendor columns
                for col in vendor_cols:
                    if col in df_vendor.columns and col in vendor_row.index:
                        val = vendor_row[col]
                        if val is not None and not pd.isna(val):
                            try:
                                vendor_values.append(float(val))
                            except (ValueError, TypeError):
                                pass
                
                if vendor_values:
                    vendor_val = sum(vendor_values)
                else:
                    vendor_val = None
                
                # Check if any vendor columns exist
                vendor_cols_exist = any(col in df_vendor.columns for col in vendor_cols)
            else:
                # Single vendor column (backward compatible)
                vendor_col = vendor_col_key
                vendor_cols = [vendor_col]
                vendor_cols_exist = vendor_col in df_vendor.columns
                
                if vendor_cols_exist and vendor_col in vendor_row.index:
                    vendor_val = vendor_row[vendor_col]
                else:
                    vendor_val = None
            
            # Skip employee number column if it's in the mapping (already added manually)
            system_col_normalized = normalize_for_comparison(system_col)
            emp_vendor_normalized = normalize_for_comparison(employee_number_vendor_col)
            emp_system_normalized = normalize_for_comparison(employee_number_system_col)
            
            # Check if any vendor column matches employee number (for multi-column case)
            if is_multi_column:
                vendor_cols_normalized = [normalize_for_comparison(col) for col in vendor_cols]
                skip_employee = any(vc_norm == emp_vendor_normalized for vc_norm in vendor_cols_normalized)
            else:
                vendor_col_normalized = normalize_for_comparison(vendor_col_key)
                skip_employee = vendor_col_normalized == emp_vendor_normalized
            
            if (skip_employee or system_col_normalized == emp_system_normalized):
                continue
            
            # Get system value
            if system_col in df_system.columns and system_col in system_row.index:
                system_val = system_row[system_col]
            else:
                system_val = None
                if system_col not in df_system.columns:
                    logger.warning(f"System column not found", extra={"column": "[MASKED]"})
            
            # Check if vendor columns exist
            if not vendor_cols_exist:
                missing_cols = [col for col in vendor_cols if col not in df_vendor.columns]
                if missing_cols:
                    logger.warning(f"Vendor column(s) not found", extra={"column_count": len(missing_cols)})
                continue
            
            # Use clean, formatted system column name for display
            display_col_name = format_column_name(system_col)
            
            # Compare values
            col_comp = compare_column_values(vendor_val, system_val, display_col_name)
            
            if not col_comp.isMatch:
                all_match = False
                # Don't log actual values (sensitive data) - just track mismatch
                employee_mismatches.append(f"{display_col_name}: mismatch detected")
            
            column_comparisons.append(col_comp)
        
        # Silent processing - no per-employee debug output
        
        results.append(RowComparisonResult(
            employeeNumber=str(emp_id),
            columnComparisons=column_comparisons,
            overallMatch=all_match,
            rowStatus="matched"
        ))
    
    # Log comparison summary (without sensitive data)
    perfect_matches = sum(1 for r in results if r.overallMatch)
    mismatches = sum(1 for r in results if not r.overallMatch)
    logger.info(
        f"Comparison completed: {perfect_matches} perfect matches, {mismatches} with mismatches",
        extra={
            "perfect_matches": perfect_matches,
            "mismatches": mismatches,
            "total_employees": len(results)
        }
    )
    
    # Log unmatched employees count (without IDs)
    if only_in_vendor:
        logger.info(f"{len(only_in_vendor)} employees only in vendor", extra={"count": len(only_in_vendor)})
    if only_in_system:
        logger.info(f"{len(only_in_system)} employees only in system", extra={"count": len(only_in_system)})
    
    # Return matched employees and counts of unmatched employees
    return {
        "results": results,
        "only_in_vendor_count": len(only_in_vendor),
        "only_in_system_count": len(only_in_system)
    }

