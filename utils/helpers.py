# utils/helpers.py
"""Helper functions for common operations"""
import pandas as pd
from typing import Dict, Union, Tuple, Any
from utils.normalization import normalize_header_name


def normalize_header_mapping(
    header_mapping: Dict[Any, str],  # Supports both str (single) and tuple (multiple) keys
    df_vendor: pd.DataFrame,
    df_system: pd.DataFrame
) -> Dict[Any, str]:
    """
    Normalize header mapping to match actual DataFrame column names (case-insensitive).
    
    Supports both single column (str key) and multiple columns (tuple key) mappings.
    Uses centralized normalization functions for consistency.
    """
    # Create lookup dictionaries: normalized_name -> original_name
    vendor_cols_lower = {normalize_header_name(str(col)): col for col in df_vendor.columns}
    system_cols_lower = {normalize_header_name(str(col)): col for col in df_system.columns}
    
    normalized = {}
    for vendor_key, system_key in header_mapping.items():
        # Check if vendor_key is a tuple (multiple columns) or string (single column)
        is_multi_column = isinstance(vendor_key, tuple)
        
        if is_multi_column:
            # Multiple vendor columns - normalize each and create tuple
            normalized_vendor_cols = []
            for v_col in vendor_key:
                v_col_norm = normalize_header_name(str(v_col))
                vendor_col = vendor_cols_lower.get(v_col_norm)
                if vendor_col:
                    normalized_vendor_cols.append(vendor_col)
                elif v_col in df_vendor.columns:
                    # Fallback: use original if exists as-is
                    normalized_vendor_cols.append(v_col)
            
            # Normalize system column
            system_key_norm = normalize_header_name(str(system_key))
            system_col = system_cols_lower.get(system_key_norm)
            if not system_col and system_key in df_system.columns:
                system_col = system_key
            
            if normalized_vendor_cols and system_col:
                # Use tuple as key for multi-column mapping
                normalized[tuple(normalized_vendor_cols)] = system_col
        else:
            # Single vendor column (backward compatible)
            vendor_key_norm = normalize_header_name(str(vendor_key))
            system_key_norm = normalize_header_name(str(system_key))
            
            vendor_col = vendor_cols_lower.get(vendor_key_norm)
            system_col = system_cols_lower.get(system_key_norm)
            
            if vendor_col and system_col:
                normalized[vendor_col] = system_col
            elif vendor_key in df_vendor.columns and system_key in df_system.columns:
                # Fallback: use original keys if they exist as-is
                normalized[vendor_key] = system_key
    
    return normalized