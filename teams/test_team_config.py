"""
Test Team Configuration - For testing purposes.
"""
import pandas as pd
import datetime as dt
from typing import Tuple

from .team_config import TeamConfig, ColumnMapping


class TestTeamConfig(TeamConfig):
    """Configuration for the test team."""
    
    @property
    def name(self) -> str:
        return "Team_for_testings"
    
    @property
    def carrier_name(self) -> str:
        return "Carrier Webpage For Testing"
    
    @property
    def customer_name(self) -> str:
        return "Test Customer"
    
    def get_orders_column_mapping(self) -> ColumnMapping:
        return ColumnMapping(
            rename_map={},
            type_map={
                "SITE#": str,
                "RETURN_TO_CARRIER_DEPOT": bool,
                "SHIP_TIME_TO": str,
            }
        )
    
    def get_contacts_column_mapping(self) -> ColumnMapping:
        return ColumnMapping(
            rename_map={},
            type_map={
                "STUDY": str,
                "SITE#": str,
                "CARRIER_ID": str,
            }
        )
    
    def transform_orders(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply test team specific order transformations."""
        df["PRINT_RETURN_DOCUMENT"] = True
        df["HAS_RETURN"] = (
            (df["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & 
            (df["TEMPERATURE"] != "Ambient")
        )
        df["TYPE_OF_RETURN"] = df["HAS_RETURN"].apply(lambda x: "CREDO" if x else "NA")
        return df
    
    def transform_contacts(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply test team specific contact transformations."""
        df["TYPE_OF_MATERIAL_CAN_RECEIVE"] = "Can receive medicines"
        df["MEDICAL_CENTER_EMAILS"] = "test@example.com"
        df["CUSTOMER_EMAIL"] = ""
        df["CRA_EMAILS"] = ""
        return df
