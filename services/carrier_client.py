"""
Carrier Client - Interface for carrier service interactions.

Defines the contract for interacting with carrier services (web, API, etc.)
without exposing implementation details like Selenium or HTTP.
"""
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Tuple, Optional

from logClass.log import Log


@dataclass
class ShippingOrder:
    """Data class for shipping order information."""
    carrier_id: str
    reference: str
    ship_date: str
    ship_time_from: str
    ship_time_to: str
    delivery_date: str
    delivery_time_from: str
    delivery_time_to: str
    type_of_material: str
    temperature: str
    contacts: str
    amount_of_boxes: int


@dataclass
class ReturnOrder:
    """Data class for return order information."""
    carrier_id: str
    reference: str
    delivery_date: str
    return_time_from: str
    return_time_to: str
    type_of_return: str
    contacts: str
    amount_of_boxes: int
    return_to_depot: bool
    original_tracking_number: str


@dataclass
class ShippingResult:
    """Result from creating a shipping order."""
    tracking_number: str
    contacts: str
    success: bool
    error_message: Optional[str] = None


class CarrierClient(ABC):
    """
    Abstract base class for carrier service clients.
    
    Implementations handle the actual communication with carrier services,
    whether via web automation (Selenium), HTTP API, or other methods.
    """
    
    def __init__(self, download_path: str, log: Log):
        """
        Initialize the carrier client.
        
        Args:
            download_path: Path where files should be downloaded.
            log: Logger instance.
        """
        self._download_path = download_path
        self._log = log
    
    # --- Connection Management ---
    
    @abstractmethod
    def connect(self) -> None:
        """Establish connection to the carrier service."""
        pass
    
    @abstractmethod
    def disconnect(self) -> None:
        """Close connection to the carrier service."""
        pass
    
    @abstractmethod
    def authenticate(self, username: str, password: str) -> bool:
        """
        Authenticate with the carrier service.
        
        Args:
            username: User credentials.
            password: Password credentials.
            
        Returns:
            True if authentication was successful.
        """
        pass
    
    # --- Order Operations ---
    
    @abstractmethod
    def create_shipping_order(self, order: ShippingOrder) -> ShippingResult:
        """
        Create a shipping order in the carrier's system.
        
        Args:
            order: Shipping order details.
            
        Returns:
            Result with tracking number and status.
        """
        pass
    
    @abstractmethod
    def create_return_order(self, order: ReturnOrder) -> ShippingResult:
        """
        Create a return order in the carrier's system.
        
        Args:
            order: Return order details.
            
        Returns:
            Result with tracking number and status.
        """
        pass
    
    # --- Document Operations ---
    
    @abstractmethod
    def print_waybill(self, tracking_number: str, copies: int = 1) -> None:
        """Print the waybill document for a shipment."""
        pass
    
    @abstractmethod
    def print_label(self, tracking_number: str) -> None:
        """Print the label document for a shipment."""
        pass
    
    @abstractmethod
    def print_return_waybill(self, tracking_number: str, copies: int = 1) -> None:
        """Print the return waybill document."""
        pass
    
    # --- Utility Methods ---
    
    def get_contacts(self, carrier_id: str) -> str:
        """
        Get contacts for a site from the carrier system.
        
        Default implementation returns empty string.
        Override in subclasses that support contact lookup.
        """
        return ""
    
    @property
    def name(self) -> str:
        """Return the carrier service name."""
        return self.__class__.__name__
