"""
Order Service - Main orchestrator for order processing.

Coordinates the entire workflow of loading, validating, and processing
shipping orders using the configured team, carrier, and email services.
"""
import datetime as dt
import time
import queue
from typing import Optional
import pandas as pd

from logClass.log import Log
from dataPathController.dataPathController import DataPathController
from utils.create_folder import create_folder
from utils.getFolderPathToDownload import getFolderPathToDownload
from utils.renameReturnPDFFile import renameReturnPDFFile
from utils.export_to_excel import export_to_excel
from utils.merge_PDFs import merge_PDFs

from teams.team_config import TeamConfig
from teams.team_config_factory import TeamConfigFactory
from dataRecolector.column_config import COLUMNS

from .carrier_client import CarrierClient, ShippingOrder, ReturnOrder
from .carrier_factory import CarrierClientFactory
from .email_service import EmailService, OrderSummary, SenderInfo
from .excel_reader import ExcelReader


class OrderService:
    """
    Main service for order processing operations.
    
    Orchestrates:
    - Loading orders and contacts from Excel
    - Creating shipping orders via carrier client
    - Printing shipping documents
    - Sending summary emails
    """
    
    def __init__(
        self,
        message_queue: Optional[queue.Queue] = None,
        log: Optional[Log] = None,
    ):
        """
        Initialize the order service.
        
        Args:
            message_queue: Queue for sending UI updates.
            log: Logger instance.
        """
        self._queue = message_queue or queue.Queue()
        self._log = log or Log()
        
        self._excel_reader = ExcelReader(self._log)
        self._email_service = EmailService(self._log)
        
        self._team_config: Optional[TeamConfig] = None
        self._carrier_client: Optional[CarrierClient] = None
        self._download_path: str = ""
    
    # --- Configuration ---
    
    def configure_for_team(self, team_name: str, ship_date: str) -> None:
        """
        Configure the service for a specific team and date.
        
        Args:
            team_name: Name of the team to process orders for.
            ship_date: Ship date string (YYYY-MM-DD format).
        """
        self._team_config = TeamConfigFactory.create(team_name)
        
        date_for_folder = ship_date.replace("/", "_").replace("-", "_")
        self._download_path = getFolderPathToDownload(team_name, date_for_folder)
        
        self._carrier_client = CarrierClientFactory.create(
            carrier_name=self._team_config.carrier_name,
            download_path=self._download_path,
            log=self._log,
        )
    
    def set_email_sender(self, sender: SenderInfo) -> None:
        """Set the email sender information."""
        self._email_service.set_sender(sender)
        self._email_service.save_sender_config()
    
    def load_email_sender_config(self) -> SenderInfo:
        """Load saved email sender configuration."""
        return self._email_service.load_sender_config()
    
    @property
    def team_config(self) -> Optional[TeamConfig]:
        """Get the current team configuration."""
        return self._team_config
    
    @property
    def download_path(self) -> str:
        """Get the current download path."""
        return self._download_path
    
    # --- Order Loading ---
    
    def load_orders(self, ship_date: dt.datetime) -> pd.DataFrame:
        """
        Load orders and contacts for a ship date.
        
        Uses TeamConfig + DataRecolector for loading and transforming data.
        
        Args:
            ship_date: Date to load orders for.
            
        Returns:
            DataFrame with merged orders and contacts.
        """
        if not self._team_config:
            self._log.add_error_log("No team configured")
            return self._get_empty_dataframe()
        
        # Import here to avoid circular dependency
        from dataRecolector.data_recolector import DataRecolector
        
        # Use the new DataRecolector with TeamConfig (not legacy Team)
        recolector = DataRecolector(self._team_config, self._queue, self._log)
        return recolector.collect_orders_and_contacts(ship_date)
    
    # --- Order Processing ---
    
    def authenticate(self, username: str, password: str) -> bool:
        """
        Authenticate with the carrier service.
        
        Args:
            username: User credentials.
            password: Password credentials.
            
        Returns:
            True if authentication successful.
        """
        if not self._carrier_client:
            self._log.add_error_log("No carrier configured")
            return False
        
        self._carrier_client.connect()
        return self._carrier_client.authenticate(username, password)
    
    def process_orders(
        self,
        orders_df: pd.DataFrame,
        ship_date: dt.datetime,
    ) -> pd.DataFrame:
        """
        Process all orders in the DataFrame.
        
        Args:
            orders_df: DataFrame with orders to process.
            ship_date: The ship date.
            
        Returns:
            Updated DataFrame with tracking numbers.
        """
        if not self._carrier_client or not self._team_config:
            self._log.add_error_log("Service not configured")
            return orders_df
        
        try:
            create_folder(self._download_path)
            
            # Process each order
            orders_df = self._process_all_orders(orders_df)
            
            time.sleep(2)  # Wait for downloads to finish
            
            # Post-processing
            self._rename_return_files(orders_df)
            self._group_and_merge_pdfs(orders_df)
            self._export_orders_to_excel(orders_df)
            
        except Exception as e:
            self._log.add_error_log(f"Error processing orders: {e}")
        
        finally:
            try:
                self._carrier_client.disconnect()
            except Exception as e:
                self._log.add_error_log(f"Error disconnecting carrier: {e}")
        
        return orders_df
    
    def send_summary_email(
        self,
        orders_df: pd.DataFrame,
        ship_date: str,
    ) -> bool:
        """
        Send summary email to the team.
        
        Args:
            orders_df: Processed orders DataFrame.
            ship_date: Ship date string.
            
        Returns:
            True if email sent successfully.
        """
        if not self._team_config:
            return False
        
        team_email = self._get_team_email()
        if not team_email:
            return False
        
        total = len(orders_df)
        processed = len(orders_df[orders_df['TRACKING_NUMBER'] != ""])
        pending = total - processed
        
        summary = OrderSummary(
            team_name=self._team_config.name,
            ship_date=ship_date,
            total_orders=total,
            processed_orders=processed,
            pending_orders=pending,
        )
        
        return self._email_service.send_team_summary_email(
            to_email=team_email,
            order_summary=summary,
            attachments_folder=self._download_path,
        )
    
    # --- Private Methods ---
    
    def _process_all_orders(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process all orders that need processing."""
        mask = (df["TRACKING_NUMBER"] == "") & (df["HAS_AN_ERROR"] == "No error")
        
        for idx in df[mask].index:
            row = df.loc[idx]
            
            try:
                # Create shipping order
                result = self._create_shipping_order(row)
                
                if result.success:
                    df.at[idx, "TRACKING_NUMBER"] = result.tracking_number
                    df.at[idx, "CONTACTS"] = result.contacts
                    
                    # Print documents
                    self._print_shipping_documents(result.tracking_number, row)
                    
                    # Create return order if needed
                    if row.get("HAS_RETURN", False):
                        return_result = self._create_return_order(row, result)
                        if return_result.success:
                            df.at[idx, "RETURN_TRACKING_NUMBER"] = return_result.tracking_number
                            
                            if row.get("PRINT_RETURN_DOCUMENT", False):
                                self._print_return_documents(return_result.tracking_number)
                
                # Notify UI of progress
                self._queue.put({
                    "INDEX": idx,
                    "TRACKING_NUMBER": df.at[idx, "TRACKING_NUMBER"],
                    "RETURN_TRACKING_NUMBER": df.at[idx, "RETURN_TRACKING_NUMBER"],
                    "CONTACTS": df.at[idx, "CONTACTS"],
                })
                
            except Exception as e:
                self._log.add_error_log(f"Error processing order {idx}: {e}")
        
        return df
    
    def _create_shipping_order(self, row: pd.Series):
        """Create a shipping order from a DataFrame row."""
        order = ShippingOrder(
            carrier_id=str(row.get("CARRIER_ID", "")),
            reference=f"{row.get('SYSTEM_NUMBER', '')} {row.get('IVRS_NUMBER', '')}"[:50],
            ship_date=str(row.get("SHIP_DATE", "")),
            ship_time_from=str(row.get("SHIP_TIME_FROM", "")),
            ship_time_to=str(row.get("SHIP_TIME_TO", "")),
            delivery_date=str(row.get("DELIVERY_DATE", "")),
            delivery_time_from=str(row.get("DELIVERY_TIME_FROM", "")),
            delivery_time_to=str(row.get("DELIVERY_TIME_TO", "")),
            type_of_material=str(row.get("TYPE_OF_MATERIAL", "")),
            temperature=str(row.get("TEMPERATURE", "")),
            contacts=str(row.get("CONTACTS", "")),
            amount_of_boxes=int(row.get("AMOUNT_OF_BOXES_TO_SEND", 1)),
        )
        
        return self._carrier_client.create_shipping_order(order)
    
    def _create_return_order(self, row: pd.Series, shipping_result):
        """Create a return order."""
        order = ReturnOrder(
            carrier_id=str(row.get("CARRIER_ID", "")),
            reference=f"RET {row.get('SYSTEM_NUMBER', '')} {row.get('IVRS_NUMBER', '')}"[:50],
            delivery_date=str(row.get("RETURN_DATE", "")),
            return_time_from=str(row.get("RETURN_DELIVERY_HOUR_FROM", "09:00")),
            return_time_to=str(row.get("RETURN_DELIVERY_HOUR_TO", "16:00")),
            type_of_return=str(row.get("TYPE_OF_RETURN", "CREDO")),
            contacts=shipping_result.contacts,
            amount_of_boxes=int(row.get("AMOUNT_OF_BOXES_TO_RETURN", 1)),
            return_to_depot=bool(row.get("RETURN_TO_CARRIER_DEPOT", False)),
            original_tracking_number=shipping_result.tracking_number,
        )
        
        return self._carrier_client.create_return_order(order)
    
    def _print_shipping_documents(self, tracking_number: str, row: pd.Series) -> None:
        """Print shipping documents (waybill and label)."""
        boxes = int(row.get("AMOUNT_OF_BOXES_TO_SEND", 1))
        self._carrier_client.print_waybill(tracking_number, boxes)
        self._carrier_client.print_label(tracking_number)
    
    def _print_return_documents(self, tracking_number: str) -> None:
        """Print return waybill document."""
        self._carrier_client.print_return_waybill(tracking_number, 1)
    
    def _rename_return_files(self, df: pd.DataFrame) -> None:
        """Rename downloaded return PDF files."""
        for idx, row in df.iterrows():
            if row.get("RETURN_TRACKING_NUMBER") and row.get("PRINT_RETURN_DOCUMENT"):
                try:
                    renameReturnPDFFile(
                        self._download_path,
                        str(row["TRACKING_NUMBER"]),
                        str(row["RETURN_TRACKING_NUMBER"]),
                    )
                except Exception as e:
                    self._log.add_warning_log(f"Error renaming return file: {e}")
    
    def _group_and_merge_pdfs(self, df: pd.DataFrame) -> None:
        """Group and merge PDF files by study."""
        try:
            studies = df["STUDY"].unique()
            for study in studies:
                merge_PDFs(self._download_path, study)
        except Exception as e:
            self._log.add_warning_log(f"Error merging PDFs: {e}")
    
    def _export_orders_to_excel(self, df: pd.DataFrame) -> None:
        """Export processed orders to Excel."""
        try:
            export_to_excel(df, self._download_path, "orders")
        except Exception as e:
            self._log.add_error_log(f"Error exporting to Excel: {e}")
    
    def _get_team_email(self) -> str:
        """Get the team email from configuration."""
        if not self._team_config:
            return ""
        
        config = DataPathController().get_config_of_a_team(self._team_config.name)
        return config.get("team_email", "")
    
    def _get_empty_dataframe(self) -> pd.DataFrame:
        """Get an empty orders DataFrame."""
        return pd.DataFrame(columns=COLUMNS.merged_columns_list)
    
    # --- Static Methods ---
    
    @staticmethod
    def get_team_names() -> list:
        """Get list of available team names."""
        return TeamConfigFactory.get_team_names()
