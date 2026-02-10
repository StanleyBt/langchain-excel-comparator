# utils/helpers.py
"""Helper functions for common operations"""
import pandas as pd
from typing import Dict, Union, Tuple, Any
from utils.normalization import normalize_header_name


def _mapping_system_header(value: Any) -> str:
    """Extract system header string from mapping value (str or dict with system_header/systemColumn)."""
    if isinstance(value, dict):
        return value.get("system_header") or value.get("systemColumn") or ""
    return str(value)


def _mapping_formula_steps(value: Any) -> Any:
    """Extract formula steps list from mapping value (supports formula_steps and formulaSteps)."""
    if not isinstance(value, dict):
        return None
    return value.get("formula_steps") or value.get("formulaSteps")


def normalize_header_mapping(
    header_mapping: Dict[Any, Any],  # Keys: tuple of vendor headers. Values: str or dict with system_header, formula_steps
    df_vendor: pd.DataFrame,
    df_system: pd.DataFrame
) -> Dict[Any, Any]:
    """
    Normalize header mapping to match actual DataFrame column names (case-insensitive).

    Keys are tuples of vendor header names; values are either a system header string
    or a dict with "system_header" and "formula_steps". Formula steps are preserved as-is.
    """
    vendor_cols_lower = {normalize_header_name(str(col)): col for col in df_vendor.columns}
    system_cols_lower = {normalize_header_name(str(col)): col for col in df_system.columns}

    normalized = {}
    for vendor_key, mapping_value in header_mapping.items():
        system_key = _mapping_system_header(mapping_value)
        if not system_key:
            continue
        is_multi_column = isinstance(vendor_key, tuple)
        if not is_multi_column:
            vendor_key = (vendor_key,)
        normalized_vendor_cols = []
        for v_col in vendor_key:
            v_col_norm = normalize_header_name(str(v_col))
            vendor_col = vendor_cols_lower.get(v_col_norm)
            if vendor_col:
                normalized_vendor_cols.append(vendor_col)
            elif v_col in df_vendor.columns:
                normalized_vendor_cols.append(v_col)
        system_key_norm = normalize_header_name(str(system_key))
        system_col = system_cols_lower.get(system_key_norm)
        if not system_col and system_key in df_system.columns:
            system_col = system_key
        if not normalized_vendor_cols or not system_col:
            continue
        out_key = tuple(normalized_vendor_cols)
        if isinstance(mapping_value, dict):
            steps = _mapping_formula_steps(mapping_value)
            normalized[out_key] = {"system_header": system_col, "formula_steps": steps}
        else:
            normalized[out_key] = system_col
    return normalized