"""
Transportes Ambientales Carrier Client Adapter.

Wraps the existing TransportesAmbientales implementation
to implement the CarrierClient interface.
"""
from typing import Optional

from services.carrier_client import (
    CarrierClient,
    ShippingOrder,
    ReturnOrder,
    ShippingResult,
)
from logClass.log import Log


class TransportesAmbientalesAdapter(CarrierClient):
    """
    Adapter for Transportes Ambientales carrier service.
    
    Wraps the legacy Selenium-based implementation to provide
    the new CarrierClient interface.
    """
    
    def __init__(self, download_path: str, log: Log, use_http: bool = False):
        """
        Initialize the adapter.
        
        Args:
            download_path: Path for file downloads.
            log: Logger instance.
            use_http: If True, use HTTP API instead of Selenium.
        """
        super().__init__(download_path, log)
        self._use_http = use_http
        self._legacy_client = None
    
    @property
    def name(self) -> str:
        return "Transportes Ambientales"
    
    def connect(self) -> None:
        """Initialize the legacy client and build the driver."""
        self._legacy_client = self._create_legacy_client()
        self._legacy_client.build_driver()
    
    def disconnect(self) -> None:
        """Quit the driver."""
        if self._legacy_client:
            self._legacy_client.quit_driver()
    
    def authenticate(self, username: str, password: str) -> bool:
        """Authenticate using the legacy client."""
        if not self._legacy_client:
            self.connect()
        return self._legacy_client.check_if_user_and_password_are_correct(username, password)
    
    def create_shipping_order(self, order: ShippingOrder) -> ShippingResult:
        """Create a shipping order using the legacy client."""
        try:
            tracking_number, contacts = self._legacy_client.complete_shipping_order_form(
                carrier_id=order.carrier_id,
                reference=order.reference,
                ship_date=order.ship_date,
                ship_time_from=order.ship_time_from,
                ship_time_to=order.ship_time_to,
                delivery_date=order.delivery_date,
                delivery_time_from=order.delivery_time_from,
                delivery_time_to=order.delivery_time_to,
                type_of_material=order.type_of_material,
                temperature=order.temperature,
                contacts=order.contacts,
                amount_of_boxes=order.amount_of_boxes,
            )
            
            return ShippingResult(
                tracking_number=tracking_number,
                contacts=contacts if contacts else order.contacts,
                success=bool(tracking_number),
            )
            
        except Exception as e:
            self._log.add_error_log(f"Error creating shipping order: {e}")
            return ShippingResult(
                tracking_number="",
                contacts=order.contacts,
                success=False,
                error_message=str(e),
            )
    
    def create_return_order(self, order: ReturnOrder) -> ShippingResult:
        """Create a return order using the legacy client."""
        try:
            tracking_number = self._legacy_client.complete_shipping_order_return_form(
                carrier_id=order.carrier_id,
                reference_return=order.reference,
                delivery_date=order.delivery_date,
                return_time_from=order.return_time_from,
                return_time_to=order.return_time_to,
                type_of_return=order.type_of_return,
                contacts=order.contacts,
                amount_of_boxes_to_return=order.amount_of_boxes,
                return_to_carrier_depot=order.return_to_depot,
                tracking_number=order.original_tracking_number,
            )
            
            return ShippingResult(
                tracking_number=tracking_number,
                contacts=order.contacts,
                success=tracking_number != "ERROR",
            )
            
        except Exception as e:
            self._log.add_error_log(f"Error creating return order: {e}")
            return ShippingResult(
                tracking_number="ERROR",
                contacts=order.contacts,
                success=False,
                error_message=str(e),
            )
    
    def print_waybill(self, tracking_number: str, copies: int = 1) -> None:
        """Print waybill document."""
        self._legacy_client.print_wayBill_document(tracking_number, copies)
    
    def print_label(self, tracking_number: str) -> None:
        """Print label document."""
        self._legacy_client.print_label_document(tracking_number)
    
    def print_return_waybill(self, tracking_number: str, copies: int = 1) -> None:
        """Print return waybill document."""
        self._legacy_client.print_return_wayBill_document(tracking_number, copies)
    
    def get_contacts(self, carrier_id: str) -> str:
        """Get contacts from carrier system."""
        return self._legacy_client.get_contacts(carrier_id)
    
    def _create_legacy_client(self):
        """Create the appropriate legacy client based on configuration."""
        if self._use_http:
            from carriersWebpage.TransportesAmbientales_requests import TransportesAmbientalesHTTP
            return TransportesAmbientalesHTTP(self._download_path, self._log)
        else:
            from carriersWebpage.TransportesAmbientales import TransportesAmbientales
            return TransportesAmbientales(self._download_path, self._log)
