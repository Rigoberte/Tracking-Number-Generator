"""
Carrier Client Factory.

Creates carrier client instances based on carrier name.
"""
from typing import Dict, Type

from logClass.log import Log
from .carrier_client import CarrierClient, ShippingResult
from .transportes_ambientales_adapter import TransportesAmbientalesAdapter


class CarrierClientFactory:
    """Factory for creating carrier clients."""
    
    @classmethod
    def create(
        cls,
        carrier_name: str,
        download_path: str,
        log: Log
    ) -> CarrierClient:
        """
        Create a carrier client by name.
        
        Args:
            carrier_name: Name of the carrier service.
            download_path: Path for file downloads.
            log: Logger instance.
            
        Returns:
            CarrierClient instance.
        """
        # Map carrier names to their configurations
        carrier_configs = {
            "Transportes Ambientales": {"use_http": False},
            "Transportes Ambientales HTTP": {"use_http": True},
        }
        
        config = carrier_configs.get(carrier_name, {"use_http": False})
        
        if "Transportes Ambientales" in carrier_name:
            return TransportesAmbientalesAdapter(
                download_path=download_path,
                log=log,
                use_http=config.get("use_http", False),
            )
        
        if carrier_name == "Carrier Webpage For Testing":
            return TestCarrierAdapter(download_path, log)
        
        # Default: return a no-op carrier
        return NullCarrierClient(download_path, log)


class TestCarrierAdapter(CarrierClient):
    """Adapter for the testing carrier."""
    
    @property
    def name(self) -> str:
        return "Carrier Webpage For Testing"
    
    def connect(self) -> None:
        """Initialize the test carrier."""
        from carriersWebpage.CarrierWebpageForTesting import CarrierWebpageForTesting
        self._legacy_client = CarrierWebpageForTesting(self._download_path, self._log)
        self._legacy_client.build_driver()
    
    def disconnect(self) -> None:
        if hasattr(self, '_legacy_client') and self._legacy_client:
            self._legacy_client.quit_driver()
    
    def authenticate(self, username: str, password: str) -> bool:
        if not hasattr(self, '_legacy_client'):
            self.connect()
        return self._legacy_client.check_if_user_and_password_are_correct(username, password)
    
    def create_shipping_order(self, order):
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
            self._log.add_error_log(f"Error creating test shipping order: {e}")
            return ShippingResult(
                tracking_number="",
                contacts=order.contacts,
                success=False,
                error_message=str(e),
            )
    
    def create_return_order(self, order):
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
            self._log.add_error_log(f"Error creating test return order: {e}")
            return ShippingResult(
                tracking_number="ERROR",
                contacts=order.contacts,
                success=False,
                error_message=str(e),
            )
    
    def print_waybill(self, tracking_number: str, copies: int = 1) -> None:
        self._legacy_client.print_wayBill_document(tracking_number, copies)
    
    def print_label(self, tracking_number: str) -> None:
        self._legacy_client.print_label_document(tracking_number)
    
    def print_return_waybill(self, tracking_number: str, copies: int = 1) -> None:
        self._legacy_client.print_return_wayBill_document(tracking_number, copies)


class NullCarrierClient(CarrierClient):
    """Null carrier client for when no carrier is selected."""
    
    @property
    def name(self) -> str:
        return "No Carrier"
    
    def connect(self) -> None:
        pass
    
    def disconnect(self) -> None:
        pass
    
    def authenticate(self, username: str, password: str) -> bool:
        return True
    
    def create_shipping_order(self, order):
        from .carrier_client import ShippingResult
        return ShippingResult(
            tracking_number="",
            contacts="",
            success=False,
            error_message="No carrier configured",
        )
    
    def create_return_order(self, order):
        from .carrier_client import ShippingResult
        return ShippingResult(
            tracking_number="",
            contacts="",
            success=False,
            error_message="No carrier configured",
        )
    
    def print_waybill(self, tracking_number: str, copies: int = 1) -> None:
        pass
    
    def print_label(self, tracking_number: str) -> None:
        pass
    
    def print_return_waybill(self, tracking_number: str, copies: int = 1) -> None:
        pass
