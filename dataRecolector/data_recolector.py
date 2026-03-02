"""
Data Recolector - Orchestrates order and contact data collection.

This module is responsible for loading, transforming, and validating
shipping orders and contact information from Excel files.

Now uses TeamConfig + ExcelReader instead of legacy Team class.
"""
import datetime as dt
import queue
from typing import List, Any, Protocol, runtime_checkable, Optional, TYPE_CHECKING

import pandas as pd

from teams.team_config import TeamConfig
from dataPathController.dataPathController import DataPathController
from logClass.log import Log

from .column_config import COLUMNS
from .order_validator import OrderValidator
from .date_calculator import DateCalculator
from .dataframe_transformer import DataFrameTransformer

if TYPE_CHECKING:
    from services.excel_reader import ExcelReader


@runtime_checkable
class MessageQueue(Protocol):
    """Protocol for message queue compatibility."""
    def put(self, item: Any) -> None: ...


class DataRecolector:
    """
    Collects and processes order and contact data for shipping operations.
    
    Coordinates the loading of data from Excel files, applies transformations,
    validates orders, and merges orders with contact information.
    """
    
    def __init__(
        self,
        team_config: TeamConfig,
        message_queue: MessageQueue = None,
        log: Log = None
    ):
        """
        Initialize the DataRecolector.
        
        Args:
            team_config: The team configuration for column mappings and transformations.
            message_queue: Optional queue for sending progress updates.
            log: Optional logger for error reporting.
        """
        # Import here to avoid circular dependency
        from services.excel_reader import ExcelReader
        
        self._team_config = team_config
        self._queue = message_queue or queue.Queue()
        self._log = log or Log()
        
        self._excel_reader = ExcelReader(self._log)
        self._transformer = DataFrameTransformer()
        self._validator = OrderValidator()
        
        # Load team configuration from DataPathController
        self._data_config = DataPathController().get_config_of_a_team(team_config.name)
    
    # --- Public Methods ---
    
    def collect_orders_and_contacts(self, ship_date: dt.datetime) -> pd.DataFrame:
        """
        Collect and process all orders and contacts for a given ship date.
        
        Args:
            ship_date: The date to filter orders by.
            
        Returns:
            DataFrame containing validated orders merged with contact info.
        """
        try:
            orders = self._load_orders(ship_date)
            contacts = self._load_contacts()
            
            merged = self._transformer.merge_orders_with_contacts(orders, contacts)
            merged["HAS_AN_ERROR"] = merged.apply(self._validator.validate, axis=1)
            merged.fillna("", inplace=True)
            
            result = merged[COLUMNS.merged_columns_list]
            self._queue.put(result)
            
            return result
            
        except Exception as e:
            self._log.add_error_log(f"Error collecting orders and contacts: {e}")
            return self.get_empty_dataframe()
    
    def get_empty_dataframe(self) -> pd.DataFrame:
        """Return an empty DataFrame with the correct columns."""
        return pd.DataFrame(columns=COLUMNS.merged_columns_list)
    
    def get_empty_orders_dataframe(self) -> pd.DataFrame:
        """Return an empty orders DataFrame."""
        return pd.DataFrame(columns=COLUMNS.order_columns_list)
    
    def get_empty_contacts_dataframe(self) -> pd.DataFrame:
        """Return an empty contacts DataFrame."""
        return pd.DataFrame(columns=COLUMNS.contact_columns_list)
    
    # --- Private Loading Methods ---
    
    def _load_orders(self, ship_date: dt.datetime) -> pd.DataFrame:
        """Load and transform orders for a given ship date."""
        df = self._load_raw_orders()
        
        if df.empty:
            return self.get_empty_orders_dataframe()
        
        # Filter by ship date
        df = df[df["SHIP_DATE"] == ship_date]
        
        # Apply transformations
        df = self._transformer.normalize_orders_columns(df)
        df = self._team_config.transform_orders(df)
        df = self._transformer.ensure_columns_exist(df, COLUMNS.order_columns_list)
        
        # Calculate return dates
        df = self._calculate_return_dates(df)
        df = self._transformer.format_dates_for_output(df)
        df = self._transformer.add_return_delivery_hours(df)
        
        return df[COLUMNS.order_columns_list]
    
    def _load_raw_orders(self) -> pd.DataFrame:
        """Load raw orders from Excel with column normalization."""
        excel_path = self._data_config.get("team_excel_path", "")
        sheet_name = self._data_config.get("team_orders_sheet", "")
        
        if not excel_path or not sheet_name:
            self._log.add_error_log("Orders Excel path or sheet not configured")
            return pd.DataFrame()
        
        df = self._excel_reader.read_orders(
            file_path=excel_path,
            sheet_name=sheet_name,
            config=self._team_config,
        )
        
        return df
    
    def _load_contacts(self) -> pd.DataFrame:
        """Load and transform contacts data."""
        df = self._load_raw_contacts()
        
        if df.empty:
            return self.get_empty_contacts_dataframe()
        
        # Get team email
        team_email = self._data_config.get("team_email", "")
        
        # Apply transformations
        df = self._transformer.normalize_contacts_columns(df, team_email)
        df = self._team_config.transform_contacts(df)
        df = self._transformer.transform_material_receiving_options(df)
        
        # Remove duplicates
        df = df.drop_duplicates(
            subset=["STUDY", "SITE#", "TYPE_OF_MATERIAL_CAN_RECEIVE"],
            keep='last'
        )
        
        df = self._transformer.ensure_columns_exist(df, COLUMNS.contact_columns_list)
        
        return df[COLUMNS.contact_columns_list]
    
    def _load_raw_contacts(self) -> pd.DataFrame:
        """Load raw contacts from Excel with column normalization."""
        excel_path = self._data_config.get("team_excel_path", "")
        sheet_name = self._data_config.get("team_contacts_sheet", "")
        
        if not excel_path or not sheet_name:
            self._log.add_error_log("Contacts Excel path or sheet not configured")
            return pd.DataFrame()
        
        df = self._excel_reader.read_contacts(
            file_path=excel_path,
            sheet_name=sheet_name,
            config=self._team_config,
        )
        
        return df
    
    def _load_non_working_days(self) -> List[dt.datetime]:
        """Load list of non-working days (holidays)."""
        excel_path = self._data_config.get("team_excel_path", "")
        sheet_name = self._data_config.get("team_not_working_days_sheet", "")
        
        if not excel_path or not sheet_name:
            return []
        
        return self._excel_reader.read_holidays(
            file_path=excel_path,
            sheet_name=sheet_name,
            config=self._team_config,
        )
    
    # --- Private Calculation Methods ---
    
    def _calculate_return_dates(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate transit days and return dates for orders."""
        non_working_days = self._load_non_working_days()
        date_calculator = DateCalculator(non_working_days)
        
        df["TRANSIT"] = df.apply(
            lambda row: date_calculator.calculate_transit_days(
                row["SHIP_DATE"],
                row["DELIVERY_DATE"]
            ),
            axis=1
        )
        
        df["RETURN_DATE"] = df.apply(
            lambda row: date_calculator.calculate_return_date(
                row["HAS_RETURN"],
                row["DELIVERY_DATE"],
                row["TRANSIT"]
            ),
            axis=1
        )
        
        return df
