"""
Services Package.

Provides business logic services for the application.
"""
from .carrier_client import CarrierClient, ShippingOrder, ReturnOrder, ShippingResult
from .carrier_factory import CarrierClientFactory
from .email_service import EmailService, SenderInfo, OrderSummary
from .excel_reader import ExcelReader
from .order_service import OrderService
from .ui_events import UIEvent, UIEventType
from .event_queue import EventQueue

__all__ = [
    # Carrier
    'CarrierClient',
    'ShippingOrder',
    'ReturnOrder',
    'ShippingResult',
    'CarrierClientFactory',
    
    # Email
    'EmailService',
    'SenderInfo',
    'OrderSummary',
    
    # Excel
    'ExcelReader',
    
    # Main Service
    'OrderService',
    
    # Events
    'UIEvent',
    'UIEventType',
    'EventQueue',
]
