"""
Excel Reader Service - Handles reading Excel files.

Provides a clean interface for reading orders, contacts, and holidays
from Excel files with team-specific configurations.
"""
from typing import List, Optional
import pandas as pd
import numpy as np

from teams.team_config import TeamConfig, ColumnMapping
from logClass.log import Log


class ExcelReader:
    """
    Service for reading Excel files with team-specific configurations.
    
    Handles column mapping, type conversion, and data normalization.
    """
    
    def __init__(self, log: Log):
        """
        Initialize the Excel reader.
        
        Args:
            log: Logger instance.
        """
        self._log = log
    
    def read_orders(
        self,
        file_path: str,
        sheet_name: str,
        config: TeamConfig,
        skip_rows: int = 0,
    ) -> pd.DataFrame:
        """
        Read orders from an Excel file.
        
        Args:
            file_path: Path to the Excel file.
            sheet_name: Name of the sheet to read.
            config: Team configuration with column mappings.
            skip_rows: Number of rows to skip at the beginning.
            
        Returns:
            DataFrame with standardized column names.
        """
        try:
            mapping = config.get_orders_column_mapping()
            
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                dtype=mapping.type_map,
                header=0,
                skiprows=skip_rows if skip_rows > 0 else None,
            )
            
            if mapping.rename_map:
                df.rename(columns=mapping.rename_map, inplace=True)
            
            return df
            
        except Exception as e:
            self._log.add_error_log(f"Error reading orders Excel: {e}")
            return pd.DataFrame()
    
    def read_contacts(
        self,
        file_path: str,
        sheet_name: str,
        config: TeamConfig,
    ) -> pd.DataFrame:
        """
        Read contacts from an Excel file.
        
        Args:
            file_path: Path to the Excel file.
            sheet_name: Name of the sheet to read.
            config: Team configuration with column mappings.
            
        Returns:
            DataFrame with standardized column names.
        """
        try:
            mapping = config.get_contacts_column_mapping()
            
            # Read as string first for proper type conversion
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                dtype=str,
                header=0,
            )
            
            # Apply type conversions
            df = self._apply_type_conversions(df, mapping.type_map)
            
            if mapping.rename_map:
                df.rename(columns=mapping.rename_map, inplace=True)
            
            return df
            
        except Exception as e:
            self._log.add_error_log(f"Error reading contacts Excel: {e}")
            return pd.DataFrame()
    
    def read_holidays(
        self,
        file_path: str,
        sheet_name: str,
        config: TeamConfig,
    ) -> List:
        """
        Read non-working days from an Excel file.
        
        Args:
            file_path: Path to the Excel file.
            sheet_name: Name of the sheet to read.
            config: Team configuration with column mappings.
            
        Returns:
            List of holiday dates.
        """
        try:
            mapping = config.get_holidays_column_mapping()
            
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                dtype=mapping.type_map,
                header=0,
            )
            
            if mapping.rename_map:
                df.rename(columns=mapping.rename_map, inplace=True)
            
            return df["DATE"].tolist()
            
        except Exception as e:
            self._log.add_error_log(f"Error reading holidays Excel: {e}")
            return []
    
    def _apply_type_conversions(
        self,
        df: pd.DataFrame,
        type_map: dict,
    ) -> pd.DataFrame:
        """Apply type conversions to DataFrame columns."""
        
        def convert_value(value, target_type):
            if pd.isna(value):
                return np.nan
            try:
                if target_type == int or target_type == 'int':
                    return int(value)
                elif target_type == float or target_type == 'float':
                    return float(value)
                elif target_type == str or target_type == 'str':
                    return str(value)
                elif target_type == bool or target_type == 'bool':
                    return bool(int(value))
                else:
                    return value
            except (ValueError, TypeError):
                return np.nan
        
        for column, target_type in type_map.items():
            if column in df.columns:
                df[column] = df[column].apply(
                    lambda x: convert_value(x, target_type)
                )
        
        return df
