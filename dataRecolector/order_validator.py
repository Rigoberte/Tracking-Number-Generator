"""
Order validation module.

Validates order rows against business rules and returns descriptive error messages.
"""
import datetime as dt
from typing import List, Optional
import pandas as pd
import numpy as np

from .column_config import VALID_VALUES


class OrderValidator:
    """Validates individual order rows against business rules."""
    
    def validate(self, row: pd.Series) -> str:
        """
        Validate an order row and return error messages.
        
        Args:
            row: A pandas Series representing an order row.
            
        Returns:
            "No error" if valid, otherwise a semicolon-separated list of errors.
        """
        errors = self._collect_errors(row)
        return "No error" if not errors else "; ".join(errors) + "; "
    
    def _collect_errors(self, row: pd.Series) -> List[str]:
        """Collect all validation errors for a row."""
        errors = []
        
        # Required field validations
        errors.extend(self._validate_required_fields(row))
        
        # Date validations
        errors.extend(self._validate_dates(row))
        
        # Time validations
        errors.extend(self._validate_times(row))
        
        # Material and temperature validations
        errors.extend(self._validate_material_and_temperature(row))
        
        # Return order validations
        errors.extend(self._validate_return_fields(row))
        
        return errors
    
    def _validate_required_fields(self, row: pd.Series) -> List[str]:
        """Validate that required fields are not empty."""
        errors = []
        
        required_fields = {
            'SYSTEM_NUMBER': "No system number",
            'CUSTOMER': "No customer",
            'STUDY': "No study",
            'SITE#': "No site",
            'CARRIER_ID': "No carrier ID",
        }
        
        for field, error_msg in required_fields.items():
            if not self._is_not_null(row.get(field)):
                errors.append(error_msg)
        
        return errors
    
    def _validate_dates(self, row: pd.Series) -> List[str]:
        """Validate date fields."""
        errors = []
        
        ship_date_valid = self._is_not_null(row.get('SHIP_DATE'))
        delivery_date_valid = self._is_not_null(row.get('DELIVERY_DATE'))
        
        if not ship_date_valid:
            errors.append("No ship date")
        
        if not delivery_date_valid:
            errors.append("No delivery date")
        
        if ship_date_valid and delivery_date_valid:
            if not self._are_valid_dates(row):
                errors.append("Invalid dates")
        
        return errors
    
    def _validate_times(self, row: pd.Series) -> List[str]:
        """Validate time fields."""
        errors = []
        
        # Ship times
        ship_time_from_valid = self._is_not_null(row.get('SHIP_TIME_FROM'))
        ship_time_to_valid = self._is_not_null(row.get('SHIP_TIME_TO'))
        
        if not ship_time_from_valid:
            errors.append("No ship time from")
        if not ship_time_to_valid:
            errors.append("No ship time to")
        if ship_time_from_valid and ship_time_to_valid:
            if row['SHIP_TIME_FROM'] > row['SHIP_TIME_TO']:
                errors.append("Invalid ship times")
        
        # Delivery times
        delivery_time_from_valid = self._is_not_null(row.get('DELIVERY_TIME_FROM'))
        delivery_time_to_valid = self._is_not_null(row.get('DELIVERY_TIME_TO'))
        
        if not delivery_time_from_valid:
            errors.append("No delivery time from")
        if not delivery_time_to_valid:
            errors.append("No delivery time to")
        if delivery_time_from_valid and delivery_time_to_valid:
            if row['DELIVERY_TIME_FROM'] > row['DELIVERY_TIME_TO']:
                errors.append("Invalid delivery times")
        
        return errors
    
    def _validate_material_and_temperature(self, row: pd.Series) -> List[str]:
        """Validate material type, temperature, and box count."""
        errors = []
        
        # Type of material
        if not self._is_not_null(row.get('TYPE_OF_MATERIAL')):
            errors.append("No type of material")
        elif row['TYPE_OF_MATERIAL'] not in VALID_VALUES.MATERIAL_TYPES:
            errors.append("Invalid type of material")
        
        # Temperature
        if not self._is_not_null(row.get('TEMPERATURE')):
            errors.append("No temperature")
        elif row['TEMPERATURE'] not in VALID_VALUES.TEMPERATURES:
            errors.append("Invalid temperature")
        
        # Amount of boxes
        boxes = row.get('AMOUNT_OF_BOXES_TO_SEND')
        if not isinstance(boxes, int) or boxes <= 0:
            errors.append("Invalid amount of boxes")
        
        return errors
    
    def _validate_return_fields(self, row: pd.Series) -> List[str]:
        """Validate return-related fields."""
        errors = []
        
        if not self._is_not_null(row.get('HAS_RETURN')):
            errors.append("No has return")
        
        if not self._is_not_null(row.get('RETURN_TO_CARRIER_DEPOT')):
            errors.append("No return to carrier depot")
        
        # Type of return validation
        type_of_return = row.get('TYPE_OF_RETURN')
        has_return = row.get('HAS_RETURN')
        
        if type_of_return not in VALID_VALUES.RETURN_TYPES:
            errors.append("Invalid type of return")
        elif has_return and type_of_return == "NA":
            errors.append("Invalid type of return")
        
        # Boxes to return validation
        if has_return:
            boxes_to_return = row.get('AMOUNT_OF_BOXES_TO_RETURN')
            boxes_to_send = row.get('AMOUNT_OF_BOXES_TO_SEND')
            
            if not self._is_valid_return_box_count(boxes_to_return, boxes_to_send):
                errors.append("Invalid number of boxes to return")
        
        return errors
    
    def _is_not_null(self, value) -> bool:
        """Check if a value is not null/empty."""
        if value is None:
            return False
        
        if isinstance(value, str):
            if value == "" or value.lower() == "nan":
                return False
        
        if pd.isna(value):
            return False
        
        try:
            return not np.isnan(float(value))
        except (ValueError, TypeError):
            return True
    
    def _are_valid_dates(self, row: pd.Series) -> bool:
        """Check if dates are valid and in correct order."""
        try:
            today = dt.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
            
            ship_date = self._parse_date(str(row['SHIP_DATE']))
            delivery_date = self._parse_date(str(row['DELIVERY_DATE']))
            
            if ship_date is None or delivery_date is None:
                return False
            
            return today <= ship_date <= delivery_date
            
        except (ValueError, TypeError):
            return False
    
    def _parse_date(self, date_str: str) -> Optional[dt.datetime]:
        """Parse a date string in dd/mm/yyyy format."""
        try:
            parts = date_str.split("/")
            return dt.datetime(
                year=int(parts[2]),
                month=int(parts[1]),
                day=int(parts[0])
            )
        except (ValueError, IndexError):
            return None
    
    def _is_valid_return_box_count(self, boxes_to_return, boxes_to_send) -> bool:
        """Check if the return box count is valid."""
        if not isinstance(boxes_to_return, int):
            return False
        if not isinstance(boxes_to_send, int):
            return False
        return 0 <= boxes_to_return <= boxes_to_send
