"""
Null/Default Team Configuration.
"""
import pandas as pd

from .team_config import TeamConfig, ColumnMapping


class NoTeamConfig(TeamConfig):
    """Default configuration when no team is selected."""
    
    @property
    def name(self) -> str:
        return "No Selected Team"
    
    @property
    def carrier_name(self) -> str:
        return "NoCarrier"
    
    def get_orders_column_mapping(self) -> ColumnMapping:
        return ColumnMapping(rename_map={}, type_map={})
    
    def get_contacts_column_mapping(self) -> ColumnMapping:
        return ColumnMapping(rename_map={}, type_map={})
    
    def transform_orders(self, df: pd.DataFrame) -> pd.DataFrame:
        return df
    
    def transform_contacts(self, df: pd.DataFrame) -> pd.DataFrame:
        return df
