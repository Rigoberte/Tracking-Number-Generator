import pandas as pd
import datetime as dt
import time
import os
import threading

from dataPathController.dataPathController import DataPathController
from logClass.log import Log
from services.event_queue import EventQueue
from services.ui_events import UIEvent, UIEventType
from services.order_service import OrderService
from services.email_service import SenderInfo
from dataRecolector import COLUMNS


class Model:
    """
    Model layer - Business logic and data management.
    
    Uses OrderService as the main orchestrator for order processing.
    """
    
    def __init__(self):
        pd.set_option('future.no_silent_downcasting', True)
        self._event_queue = EventQueue()
        self.log = Log()
        
        # Main service orchestrator
        self._order_service = OrderService(
            message_queue=self._event_queue,
            log=self.log,
        )
        
        # State
        self._selected_team_name: str = ""
        self._selected_date: dt.datetime = None
        self.ordersAndContactsDataframe = self._get_empty_dataframe()
        
        self.add_info_log("Application started")

    # Event Queue Access (encapsulated)
    def poll_event(self) -> UIEvent:
        """
        Poll for the next UI event (non-blocking).
        
        Returns:
            UIEvent if available
            
        Raises:
            queue.Empty if no events available
        """
        return self._event_queue.get_nowait()
    
    def has_pending_events(self) -> bool:
        """Check if there are pending events."""
        return not self._event_queue.empty()

    # DataRecolector methods
    def on_loadOrders_btn_click(self, selected_team_name, selected_date) -> None:
        self.log.add_separator()

        self.ordersAndContactsDataframe = self._get_empty_dataframe()
        self._event_queue.emit_update_orders(self.ordersAndContactsDataframe)

        thread = threading.Thread(
            target=self._load_orders_with_timing,
            args=(selected_team_name, selected_date),
            daemon=True
        )
        thread.start()

    def on_clearOrders_btn_click(self) -> None:
        self.log.add_separator()

        self.ordersAndContactsDataframe = self._get_empty_dataframe()
        self._event_queue.emit_update_orders(self.ordersAndContactsDataframe)

        self.log.add_info_log("Orders table cleaned")

    def _get_empty_dataframe(self) -> pd.DataFrame:
        """Return an empty DataFrame with the correct columns."""
        return pd.DataFrame(columns=COLUMNS.merged_columns_list)

    def get_empty_ordersAndContactsData(self) -> pd.DataFrame:
        """Public method for backward compatibility."""
        return self._get_empty_dataframe()

    def get_team_names(self) -> list:
        """Get list of available team names."""
        return OrderService.get_team_names()

    def validate_login(self, username: str, password: str) -> bool:
        """Authenticate with the carrier service."""
        return self._order_service.authenticate(username, password)

    # OrderProcessor methods
    def on_login_successful(self) -> None:
        """Start processing orders after successful login."""
        thread = threading.Thread(
            target=self._process_orders_with_timing,
            daemon=True
        )
        thread.start()

    def on_login_failed(self) -> None:
        """Handle failed login - disconnect carrier."""
        self._order_service._carrier_client.disconnect()

    # Log methods
    def add_error_log(self, text: str) -> None:
        self.log.add_error_log(text)

    def add_warning_log(self, text: str) -> None:
        self.log.add_warning_log(text)

    def add_info_log(self, text: str) -> None:
        self.log.add_info_log(text)

    def print_logs(self) -> pd.DataFrame:
        return self.log.print_logs()

    def print_last_n_logs(self, n: int) -> pd.DataFrame:
        return self.log.print_last_n_logs(n)

    def on_export_logs_to_csv(self) -> None:
        self.log.export_to_csv(os.path.expanduser("~\\Downloads"))

    # EmailSender methods
    def confirm_email_sender(self, full_name, job_position, address, phone_number, email_address) -> None:
        """Set the email sender information."""
        sender = SenderInfo(
            full_name=full_name,
            job_position=job_position,
            site_address=address,
            phone_number=phone_number,
            email_address=email_address,
        )
        self._order_service.set_email_sender(sender)

    def get_last_sender_config(self) -> dict:
        """Get the last saved sender configuration."""
        sender = self._order_service.load_email_sender_config()
        return {
            "full_name": sender.full_name,
            "job_position": sender.job_position,
            "address": sender.site_address,
            "phone_number": sender.phone_number,
            "email_address": sender.email_address,
        }

    # Configs methods
    def get_config_of_a_team(self, teamName: str) -> str:
        return DataPathController().get_config_of_a_team(teamName)
    
    def selected_team_must_send_email_to_medical_centers(self) -> bool:
        """Check if the selected team should send emails to medical centers."""
        try:
            config = DataPathController().get_config_of_a_team(self._selected_team_name)
            return config.get("team_send_email_to_medical_centers", False)
        except:
            return False
        
    def on_click_save_config_button(self, 
                                    team_name : str, 
                                    team_excel_path : str, 
                                    team_orders_sheet : str, 
                                    team_contacts_sheet : str, 
                                    team_not_working_days_sheet : str, 
                                    team_send_email_to_medical_centers : bool,
                                    team_email: str) -> None:
        
        DataPathController().redefine_a_config_of_a_team(team_name, 
            {
            "team_excel_path": team_excel_path,
            "team_orders_sheet": team_orders_sheet,
            "team_contacts_sheet": team_contacts_sheet,
            "team_not_working_days_sheet": team_not_working_days_sheet,
            "team_send_email_to_medical_centers": team_send_email_to_medical_centers,
            "team_email": team_email
            }
        )

    # Other mainUserForm methods
    def on_open_excel_double_btn_click(self, temporal_selected_team_name: str) -> None:
        """Open the Excel file for a team."""
        try:
            config = DataPathController().get_config_of_a_team(temporal_selected_team_name)
            excel_path = config.get("team_excel_path", "")

            if not excel_path or not os.path.exists(excel_path) or not excel_path.endswith(".xlsx"):
                self.add_warning_log("Excel file path not found")
                return
                
            os.startfile(excel_path)
        except Exception as e:
            self.add_error_log(f"Error opening Excel file: {e}")

    # Private methods - Order Loading
    def _load_orders_with_timing(self, selected_team_name: str, selected_date: str) -> None:
        """Load orders with timing and UI updates."""
        self._event_queue.emit_block_widgets()
        
        self.add_info_log(f"Start loading orders for {selected_team_name} team")
        
        time0 = time.time()
        try:
            self._load_orders(selected_team_name, selected_date)
        finally:
            time1 = time.time()
            total_time = round(time1 - time0, 2)

            self.add_info_log(f"End loading orders")
            self.add_info_log(f"Total loading time: {total_time} s")

            self._event_queue.emit_unblock_widgets()

    def _load_orders(self, selected_team_name: str, selected_date: str) -> None:
        """Load orders using OrderService."""
        self._selected_team_name = selected_team_name
        self._selected_date = dt.datetime.strptime(selected_date, '%Y-%m-%d')
        
        # Configure the service for this team
        self._order_service.configure_for_team(selected_team_name, selected_date)
        
        # Load orders
        self.ordersAndContactsDataframe = self._order_service.load_orders(self._selected_date)

    # Private methods - Order Processing
    def _process_orders_with_timing(self) -> None:
        """Process orders with timing and UI updates."""
        self._event_queue.emit_block_widgets()
        
        orders_to_process = len(
            self.ordersAndContactsDataframe[
                (self.ordersAndContactsDataframe["TRACKING_NUMBER"] == "") & 
                (self.ordersAndContactsDataframe["HAS_AN_ERROR"] == "No error")
            ]
        )
        
        self.add_info_log(f"Start processing orders")
        
        time0 = time.time()
        try:
            # Process orders
            self.ordersAndContactsDataframe = self._order_service.process_orders(
                self.ordersAndContactsDataframe,
                self._selected_date,
            )
            
            # Send summary email
            ship_date_str = self._selected_date.strftime("%Y-%m-%d")
            self._order_service.send_summary_email(
                self.ordersAndContactsDataframe,
                ship_date_str,
            )
            
        finally:
            time1 = time.time()
            total_time = round(time1 - time0, 2)
            
            self.add_info_log(f"End processing orders")
            self.add_info_log(f"Total processing time: {total_time} s for {orders_to_process} orders")

            if orders_to_process > 0:
                avg_time = round(total_time / orders_to_process, 2)
                self.add_info_log(f"Average processing time: {avg_time} s")

            self._event_queue.emit_update_orders(self.ordersAndContactsDataframe)
            self._event_queue.emit_unblock_widgets()