# utils/data_processor.py
import pandas as pd
from typing import List, Dict, Any
from models.schemas import SystemDataItem
from utils.normalization import normalize_dataframe_columns


def system_json_to_dataframe(system_data: List[SystemDataItem]) -> pd.DataFrame:
    """
    Convert system data JSON to pandas DataFrame.
    
    System data has nested structure:
    - Each item has 'details' object with all fields
    - Each item has 'employeeNumber' at top level
    
    Args:
        system_data: List of SystemDataItem objects
        
    Returns:
        DataFrame with all fields from 'details' + 'employeeNumber' as columns
    """
    rows = []
    for item in system_data:
        row = item.details.copy()  # Copy all details fields
        # Add employeeNumber from top level
        row['employeeNumber'] = item.employeeNumber
        rows.append(row)
    
    df = pd.DataFrame(rows)
    
    # Normalize column names using centralized function
    normalize_dataframe_columns(df)
    
    return df


def vendor_json_to_dataframe(vendor_data: List[Dict[str, Any]]) -> pd.DataFrame:
    """
    Convert vendor paysheet JSON to pandas DataFrame.
    
    Vendor data has flat structure - all fields at top level.
    
    Args:
        vendor_data: List of dictionaries with vendor paysheet data
        
    Returns:
        DataFrame with normalized column names
    """
    df = pd.DataFrame(vendor_data)
    
    # Normalize column names using centralized function
    normalize_dataframe_columns(df)
    
    return df

