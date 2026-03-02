"""
Column configuration constants for orders and contacts data.

This module centralizes all column definitions to avoid hardcoding
column names throughout the codebase.
"""
from dataclasses import dataclass
from typing import List, Dict


@dataclass(frozen=True)
class ColumnConfig:
    """Immutable configuration for DataFrame columns."""
    
    # Main merged DataFrame columns
    MERGED_COLUMNS: tuple = (
        'SYSTEM_NUMBER', 'IVRS_NUMBER', 'CUSTOMER', 'STUDY', 'SITE#',
        'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO',
        'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO',
        'TYPE_OF_MATERIAL', 'TEMPERATURE', 'AMOUNT_OF_BOXES_TO_SEND',
        'HAS_RETURN', 'RETURN_TO_CARRIER_DEPOT', 'TYPE_OF_RETURN',
        'RETURN_DATE', 'RETURN_DELIVERY_HOUR_FROM', 'RETURN_DELIVERY_HOUR_TO',
        'AMOUNT_OF_BOXES_TO_RETURN',
        'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT',
        'CONTACTS', 'TYPE_OF_MATERIAL_CAN_RECEIVE',
        'MEDICAL_CENTER_EMAILS', 'CUSTOMER_EMAIL', 'CRA_EMAILS', 'TEAM_EMAILS',
        'CARRIER_ID', 'HAS_AN_ERROR'
    )
    
    # Orders table columns
    ORDER_COLUMNS: tuple = (
        'SYSTEM_NUMBER', 'IVRS_NUMBER', 'CUSTOMER', 'STUDY', 'SITE#',
        'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO',
        'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO',
        'TYPE_OF_MATERIAL', 'TEMPERATURE', 'AMOUNT_OF_BOXES_TO_SEND',
        'HAS_RETURN', 'RETURN_TO_CARRIER_DEPOT', 'TYPE_OF_RETURN',
        'RETURN_DATE', 'RETURN_DELIVERY_HOUR_FROM', 'RETURN_DELIVERY_HOUR_TO',
        'AMOUNT_OF_BOXES_TO_RETURN',
        'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT'
    )
    
    # Contacts table columns
    CONTACT_COLUMNS: tuple = (
        'STUDY', 'SITE#', 'CARRIER_ID',
        'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO',
        'CONTACTS', 'TYPE_OF_MATERIAL_CAN_RECEIVE',
        'MEDICAL_CENTER_EMAILS', 'CUSTOMER_EMAIL', 'CRA_EMAILS', 'TEAM_EMAILS'
    )
    
    # Material receiving columns for transformation
    MATERIAL_RECEIVING_COLUMNS: tuple = (
        'CAN_RECEIVE_MEDICINES',
        'CAN_RECEIVE_ANCILLARIES_TYPE1',
        'CAN_RECEIVE_ANCILLARIES_TYPE2',
        'CAN_RECEIVE_EQUIPMENTS'
    )
    
    @property
    def merged_columns_list(self) -> List[str]:
        return list(self.MERGED_COLUMNS)
    
    @property
    def order_columns_list(self) -> List[str]:
        return list(self.ORDER_COLUMNS)
    
    @property
    def contact_columns_list(self) -> List[str]:
        return list(self.CONTACT_COLUMNS)


class ValidValues:
    """Valid values for categorical columns."""
    
    MATERIAL_TYPES: tuple = (
        "Medicine",
        "Ancillary Type 1",
        "Ancillary Type 2",
        "Equipment"
    )
    
    TEMPERATURES: tuple = (
        "Ambient",
        "Controlled Ambient",
        "Refrigerated",
        "Frozen",
        "Refrigerated with Dry Ice",
        "Frozen with Liquid Nitrogen"
    )
    
    RETURN_TYPES: tuple = (
        "CREDO",
        "DATALOGGER",
        "CREDO AND DATALOGGER",
        "NA"
    )
    
    # Mapping for ship time normalization
    SHIP_TIME_MAPPINGS: Dict[str, str] = {
        "8": "08:00:00",
        "16.3": "16:30:00",
        "19": "19:00:00"
    }
    
    # Mapping for material type transformation
    MATERIAL_TYPE_MAPPINGS: Dict[str, str] = {
        "MEDICINES": "Medicine",
        "ANCILLARIES_TYPE1": "Ancillary Type 1",
        "ANCILLARIES_TYPE2": "Ancillary Type 2",
        "EQUIPMENTS": "Equipment"
    }


# Singleton instance for easy access
COLUMNS = ColumnConfig()
VALID_VALUES = ValidValues()
