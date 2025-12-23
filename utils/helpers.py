# utils/helpers.py
"""Helper functions for common operations"""
import pandas as pd
from typing import Dict, Tuple, Optional, List
from models.schemas import RowComparisonResult


def find_contractor_column(df: pd.DataFrame) -> Optional[str]:
    """Find contractor/vendor column in DataFrame"""
    for col in df.columns:
        col_lower = str(col).lower()
        if "contractor" in col_lower or "vendor" in col_lower:
            return col
    return None


def filter_by_contractor(df: pd.DataFrame, contractor_name: str) -> pd.DataFrame:
    """Filter DataFrame by contractor name"""
    contractor_col = find_contractor_column(df)
    if contractor_col:
        return df[
            df[contractor_col].astype(str).str.lower().str.strip() == 
            contractor_name.lower().strip()
        ]
    return df


def normalize_header_mapping(
    header_mapping: Dict[str, str],
    df_vendor: pd.DataFrame,
    df_system: pd.DataFrame
) -> Dict[str, str]:
    """Normalize header mapping to match actual DataFrame column names (case-insensitive)"""
    vendor_cols_lower = {str(col).strip().lower(): col for col in df_vendor.columns}
    system_cols_lower = {str(col).strip().lower(): col for col in df_system.columns}
    
    normalized = {}
    for vendor_key, system_key in header_mapping.items():
        vendor_col = vendor_cols_lower.get(str(vendor_key).strip().lower())
        system_col = system_cols_lower.get(str(system_key).strip().lower())
        
        if vendor_col and system_col:
            normalized[vendor_col] = system_col
        elif vendor_key in df_vendor.columns and system_key in df_system.columns:
            normalized[vendor_key] = system_key
    
    return normalized


def calculate_match_rate(row_comparisons: List[RowComparisonResult]) -> float:
    """Calculate overall match rate from row comparisons"""
    matched_employees = [r for r in row_comparisons if r.rowStatus == "matched"]
    if not matched_employees:
        return 0.0
    
    total_comparisons = sum(len(r.columnComparisons) for r in matched_employees)
    if total_comparisons == 0:
        return 0.0
    
    matched_comparisons = sum(
        sum(1 for cc in r.columnComparisons if cc.isMatch)
        for r in matched_employees
    )
    
    return matched_comparisons / total_comparisons


def calculate_basic_summary(
    row_comparisons: List[RowComparisonResult], 
    columns_compared: int,
    only_in_vendor_count: int = 0,
    only_in_system_count: int = 0
) -> Dict[str, any]:
    """Calculate basic summary statistics"""
    total_rows = len(row_comparisons)
    matched_rows = sum(1 for r in row_comparisons if r.overallMatch)
    unmatched_rows = total_rows - matched_rows
    match_rate = calculate_match_rate(row_comparisons)
    
    return {
        "totalRows": total_rows,
        "matchedRows": matched_rows,
        "unmatchedRows": unmatched_rows,
        "onlyInVendor": only_in_vendor_count,
        "onlyInSystem": only_in_system_count,
        "matchRate": match_rate,
        "columnsCompared": columns_compared
    }

