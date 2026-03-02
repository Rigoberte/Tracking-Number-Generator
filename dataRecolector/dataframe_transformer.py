"""
DataFrame transformation utilities.

Handles column normalization, type corrections, and DataFrame merging operations.
"""
import datetime as dt
from typing import List
import pandas as pd

from .column_config import COLUMNS, VALID_VALUES


class DataFrameTransformer:
    """Transforms and normalizes DataFrames for orders and contacts."""
    
    # --- Column Correction Methods ---
    
    @staticmethod
    def normalize_date_column(df: pd.DataFrame, column: str) -> pd.Series:
        """
        Normalize a date column to datetime format.
        
        Args:
            df: The DataFrame.
            column: The column name to normalize.
            
        Returns:
            Normalized datetime Series.
        """
        df[column] = df[column].astype("datetime64[ns]")
        df[column] = pd.to_datetime(df[column], format='%d/%m/%Y', errors='coerce')
        return df[column]
    
    @staticmethod
    def normalize_time_column(df: pd.DataFrame, column: str) -> pd.Series:
        """
        Normalize a time column to HH:MM format.
        
        Args:
            df: The DataFrame.
            column: The column name to normalize.
            
        Returns:
            Normalized time string Series.
        """
        return pd.to_datetime(
            df[column], format='%H:%M:%S', errors='coerce'
        ).dt.strftime('%H:%M')
    
    @staticmethod
    def ensure_columns_exist(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
        """
        Ensure all specified columns exist in the DataFrame.
        
        Missing columns are added with empty string values.
        
        Args:
            df: The DataFrame to modify.
            columns: List of required column names.
            
        Returns:
            DataFrame with all required columns.
        """
        for column in columns:
            if column not in df.columns:
                df[column] = ""
        return df
    
    # --- Orders Table Transformations ---
    
    def normalize_orders_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Apply standard normalizations to orders DataFrame columns.
        
        Args:
            df: Orders DataFrame.
            
        Returns:
            Normalized DataFrame.
        """
        # Date columns
        df["SHIP_DATE"] = self.normalize_date_column(df, "SHIP_DATE")
        df["DELIVERY_DATE"] = self.normalize_date_column(df, "DELIVERY_DATE")
        
        # Text cleanup
        df["TEMPERATURE"] = df["TEMPERATURE"].str.strip()
        df["SITE#"] = df["SITE#"].astype(object)
        
        # Numeric columns
        df["AMOUNT_OF_BOXES_TO_SEND"] = (
            df["AMOUNT_OF_BOXES_TO_SEND"]
            .replace('', '0')
            .fillna('0')
            .astype(int)
        )
        
        # Ship time normalization
        df = self._normalize_ship_times(df)
        
        # Return boxes
        df = self._normalize_return_boxes(df)
        
        # Default type of return
        df["TYPE_OF_RETURN"] = "NA"
        
        return df
    
    def _normalize_ship_times(self, df: pd.DataFrame) -> pd.DataFrame:
        """Normalize ship time columns."""
        df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].replace(VALID_VALUES.SHIP_TIME_MAPPINGS)
        df["SHIP_TIME_FROM"] = pd.to_datetime(
            df["SHIP_TIME_FROM"], format='%H:%M:%S', errors='coerce'
        )
        df["SHIP_TIME_TO"] = df["SHIP_TIME_FROM"] + dt.timedelta(minutes=30)
        df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].dt.strftime('%H:%M')
        df["SHIP_TIME_TO"] = df["SHIP_TIME_TO"].dt.strftime('%H:%M')
        return df
    
    def _normalize_return_boxes(self, df: pd.DataFrame) -> pd.DataFrame:
        """Normalize return boxes column."""
        if "AMOUNT_OF_BOXES_TO_RETURN" in df.columns:
            df["AMOUNT_OF_BOXES_TO_RETURN"] = (
                df["AMOUNT_OF_BOXES_TO_RETURN"]
                .replace('', '0')
                .fillna('0')
                .astype(int)
            )
        else:
            df["AMOUNT_OF_BOXES_TO_RETURN"] = 0
        return df
    
    def format_dates_for_output(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Format date columns as strings for output.
        
        Args:
            df: DataFrame with datetime columns.
            
        Returns:
            DataFrame with formatted date strings.
        """
        df["SHIP_DATE"] = df["SHIP_DATE"].dt.strftime('%d/%m/%Y')
        df["DELIVERY_DATE"] = df["DELIVERY_DATE"].dt.strftime('%d/%m/%Y')
        
        df["RETURN_DATE"] = self.normalize_date_column(df, "RETURN_DATE")
        df["RETURN_DATE"] = df["RETURN_DATE"].dt.strftime('%d/%m/%Y')
        
        return df
    
    def add_return_delivery_hours(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add default return delivery hours based on HAS_RETURN flag."""
        df["RETURN_DELIVERY_HOUR_FROM"] = df["HAS_RETURN"].apply(
            lambda x: "09:00" if x else ""
        )
        df["RETURN_DELIVERY_HOUR_TO"] = df["HAS_RETURN"].apply(
            lambda x: "16:00" if x else ""
        )
        return df
    
    # --- Contacts Table Transformations ---
    
    def normalize_contacts_columns(self, df: pd.DataFrame, team_email: str) -> pd.DataFrame:
        """
        Apply standard normalizations to contacts DataFrame columns.
        
        Args:
            df: Contacts DataFrame.
            team_email: Team email to add to all rows.
            
        Returns:
            Normalized DataFrame.
        """
        df["DELIVERY_TIME_FROM"] = self.normalize_time_column(df, "DELIVERY_TIME_FROM")
        df["DELIVERY_TIME_TO"] = self.normalize_time_column(df, "DELIVERY_TIME_TO")
        df["TEAM_EMAILS"] = team_email
        return df
    
    def transform_material_receiving_options(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Transform boolean material receiving columns into rows.
        
        Converts columns like CAN_RECEIVE_MEDICINES into a TYPE_OF_MATERIAL_CAN_RECEIVE column.
        
        Args:
            df: Contacts DataFrame with boolean material columns.
            
        Returns:
            Transformed DataFrame with material type column.
        """
        receiving_columns = list(COLUMNS.MATERIAL_RECEIVING_COLUMNS)
        other_columns = df.columns.difference(receiving_columns)
        
        # Melt the receiving columns into rows
        df_melted = df.melt(
            id_vars=other_columns,
            value_vars=receiving_columns,
            var_name='Option',
            value_name='Chosen'
        )
        
        # Keep only rows where the option is True
        df_filtered = df_melted[df_melted['Chosen']].drop(columns='Chosen')
        
        # Extract material type from column name
        df_filtered['TYPE_OF_MATERIAL_CAN_RECEIVE'] = (
            df_filtered['Option']
            .str.replace('CAN_RECEIVE_', '')
            .replace(VALID_VALUES.MATERIAL_TYPE_MAPPINGS)
        )
        
        return df_filtered.drop(columns='Option')
    
    # --- Merge Operations ---
    
    def merge_orders_with_contacts(
        self,
        orders_df: pd.DataFrame,
        contacts_df: pd.DataFrame
    ) -> pd.DataFrame:
        """
        Merge orders with contacts, using contact delivery times when orders don't have them.
        
        Args:
            orders_df: Orders DataFrame.
            contacts_df: Contacts DataFrame.
            
        Returns:
            Merged DataFrame with resolved delivery times.
        """
        merged = pd.merge(
            orders_df,
            contacts_df,
            left_on=["STUDY", "SITE#", "TYPE_OF_MATERIAL"],
            right_on=["STUDY", "SITE#", "TYPE_OF_MATERIAL_CAN_RECEIVE"],
            how="left"
        )
        
        # Resolve delivery times: use contact times if order times are empty
        merged = self._resolve_delivery_times(merged)
        
        # Clean up duplicate columns
        merged = merged.drop(columns=[
            "DELIVERY_TIME_FROM_x", "DELIVERY_TIME_TO_x",
            "DELIVERY_TIME_FROM_y", "DELIVERY_TIME_TO_y"
        ])
        
        return merged
    
    def _resolve_delivery_times(self, df: pd.DataFrame) -> pd.DataFrame:
        """Resolve delivery times by preferring order times over contact times."""
        empty_values = ('', '00:00', None)
        
        df["DELIVERY_TIME_FROM"] = df.apply(
            lambda row: (
                row["DELIVERY_TIME_FROM_y"]
                if row["DELIVERY_TIME_FROM_x"] in empty_values
                else row["DELIVERY_TIME_FROM_x"]
            ),
            axis=1
        )
        
        df["DELIVERY_TIME_TO"] = df.apply(
            lambda row: (
                row["DELIVERY_TIME_TO_y"]
                if row["DELIVERY_TIME_TO_x"] in empty_values
                else row["DELIVERY_TIME_TO_x"]
            ),
            axis=1
        )
        
        return df
