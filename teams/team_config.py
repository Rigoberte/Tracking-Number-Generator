"""
Team Configuration - Base class for team-specific settings.

Teams now only contain configuration and data transformation rules,
NOT carrier/email/driver logic.
"""
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Dict, Tuple, Optional
import pandas as pd


@dataclass
class ColumnMapping:
    """Maps source Excel columns to standard column names."""
    rename_map: Dict[str, str]
    type_map: Dict[str, type]


class TeamConfig(ABC):
    """
    Abstract base class for team configurations.
    
    Teams define:
    - Column mappings for Excel files
    - Data transformation rules
    - Team-specific business logic for orders/contacts
    
    Teams DO NOT handle:
    - Carrier/web interactions (use CarrierClient)
    - Email sending (use EmailService)
    - WebDriver management
    """
    
    @property
    @abstractmethod
    def name(self) -> str:
        """Return the team's display name."""
        pass
    
    @property
    @abstractmethod
    def carrier_name(self) -> str:
        """Return the carrier service name this team uses."""
        pass
    
    @property
    def customer_name(self) -> str:
        """Return the customer/company name."""
        return self.name
    
    # --- Column Mappings ---
    
    @abstractmethod
    def get_orders_column_mapping(self) -> ColumnMapping:
        """Get column mapping configuration for orders table."""
        pass
    
    @abstractmethod
    def get_contacts_column_mapping(self) -> ColumnMapping:
        """Get column mapping configuration for contacts table."""
        pass
    
    def get_holidays_column_mapping(self) -> ColumnMapping:
        """Get column mapping for non-working days table."""
        return ColumnMapping(
            rename_map={"date": "DATE"},
            type_map={"date": "datetime64[ns]"}
        )
    
    # --- Data Transformations ---
    
    @abstractmethod
    def transform_orders(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Apply team-specific transformations to orders DataFrame.
        
        Args:
            df: Orders DataFrame with standardized column names.
            
        Returns:
            Transformed DataFrame with team-specific business rules applied.
        """
        pass
    
    @abstractmethod
    def transform_contacts(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Apply team-specific transformations to contacts DataFrame.
        
        Args:
            df: Contacts DataFrame with standardized column names.
            
        Returns:
            Transformed DataFrame with team-specific rules applied.
        """
        pass
    
    # --- Helper Methods ---
    
    def ensure_contact_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Ensure required contact columns exist with defaults."""
        defaults = {
            "CONTACTS": "No contact",
            "MEDICAL_CENTER_EMAILS": "",
            "CUSTOMER_EMAIL": "",
            "CRA_EMAILS": "",
        }
        
        for col, default in defaults.items():
            if col not in df.columns:
                df[col] = default
            else:
                df[col] = df[col].fillna(default).replace("", default)
        
        return df
    
    def set_material_receiving_flags(
        self,
        df: pd.DataFrame,
        medicines: bool = False,
        ancillaries_type1: bool = False,
        ancillaries_type2: bool = False,
        equipment: bool = False
    ) -> pd.DataFrame:
        """Set material receiving capability flags."""
        df["CAN_RECEIVE_MEDICINES"] = medicines
        df["CAN_RECEIVE_ANCILLARIES_TYPE1"] = ancillaries_type1
        df["CAN_RECEIVE_ANCILLARIES_TYPE2"] = ancillaries_type2
        df["CAN_RECEIVE_EQUIPMENTS"] = equipment
        return df
