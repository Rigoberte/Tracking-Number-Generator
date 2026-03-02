"""
Eli Lilly Argentina Team Configuration.
"""
import pandas as pd
from typing import Tuple

from .team_config import TeamConfig, ColumnMapping


class EliLillyArgentinaConfig(TeamConfig):
    """Configuration for Eli Lilly Argentina team."""
    
    @property
    def name(self) -> str:
        return "Eli Lilly Argentina"
    
    @property
    def carrier_name(self) -> str:
        return "Transportes Ambientales HTTP"
    
    @property
    def customer_name(self) -> str:
        return "Eli Lilly and Company"
    
    def get_orders_column_mapping(self) -> ColumnMapping:
        return ColumnMapping(
            rename_map={
                "CT-WIN": "SYSTEM_NUMBER",
                "IWRS": "IVRS_NUMBER",
                "Trial Alias": "STUDY",
                "Site ": "SITE#",
                "Order received": "ENTER DATE",
                "Ship date": "SHIP_DATE",
                "Horario de Despacho": "SHIP_TIME_FROM",
                "Delivery Date": "DELIVERY_DATE",
                "Destination": "DESTINATION",
                "CONDICION": "TEMPERATURE",
                "TT4": "AMOUNT_OF_BOXES_TO_SEND",
                "AWB": "TRACKING_NUMBER",
                "Shipper return AWB": "RETURN_TRACKING_NUMBER",
            },
            type_map={
                "CT-WIN": str,
                "IWRS": str,
                "Trial Alias": str,
                "Site ": str,
                "Order received": str,
                "Horario de Despacho": str,
                "Destination": str,
                "CONDICION": str,
                "TT4": str,
                "AWB": str,
                "Shipper return AWB": str,
            }
        )
    
    def get_contacts_column_mapping(self) -> ColumnMapping:
        return ColumnMapping(
            rename_map={
                "Protocolo": "STUDY",
                "Codigo": "CARRIER_ID",
                "Site": "SITE#",
                "Horario inicio": "DELIVERY_TIME_FROM",
                "Horario fin": "DELIVERY_TIME_TO",
            },
            type_map={
                "Protocolo": str,
                "Site": str,
                "Codigo": str,
                "Horario inicio": str,
                "Horario fin": str,
            }
        )
    
    def transform_orders(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply Eli Lilly specific order transformations."""
        df["TYPE_OF_MATERIAL"] = "Medicine"
        df["CUSTOMER"] = self.customer_name
        
        # Temperature mappings
        df["TEMPERATURE"] = df["TEMPERATURE"].replace(self._get_temperature_mappings())
        
        # Adjust temperatures based on return requirements
        df = self._adjust_temperatures_for_returns(df)
        
        # Calculate boxes to return
        df["Cajas (Carton)"] = df["Cajas (Carton)"].replace("", 0).fillna(0).astype(int)
        df["AMOUNT_OF_BOXES_TO_RETURN"] = df["AMOUNT_OF_BOXES_TO_SEND"] - df["Cajas (Carton)"]
        
        # Return flags
        df["RETURN_TO_CARRIER_DEPOT"] = False
        df["HAS_RETURN"] = (
            (df["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & 
            (df["TEMPERATURE"] != "Ambient")
        )
        df.loc[df["HAS_RETURN"], "TYPE_OF_RETURN"] = "CREDO"
        df["PRINT_RETURN_DOCUMENT"] = df["HAS_RETURN"]
        
        return df
    
    def transform_contacts(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply Eli Lilly specific contact transformations."""
        df = self.ensure_contact_columns(df)
        df = self.set_material_receiving_flags(df, medicines=True)
        return df
    
    def _get_temperature_mappings(self) -> dict:
        """Get temperature code to name mappings."""
        return {
            "L": "Ambient",
            "M": "Controlled Ambient",
            "M + L": "Controlled Ambient, Ambient",
            "H": "Controlled Ambient",
            "H + M": "Controlled Ambient",
            "H + L": "Controlled Ambient, Ambient",
            "H + M + L": "Controlled Ambient, Ambient",
            "REF": "Refrigerated",
            "REF + H": "Refrigerated, Controlled Ambient",
            "REF + M": "Refrigerated, Controlled Ambient",
            "REF + L": "Refrigerated, Ambient",
            "REF + H + M": "Refrigerated, Controlled Ambient",
            "REF + H + L": "Refrigerated, Controlled Ambient, Ambient",
            "REF + M + L": "Refrigerated, Controlled Ambient, Ambient",
            "REF + H + M + L": "Refrigerated, Controlled Ambient, Ambient",
        }
    
    def _adjust_temperatures_for_returns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Adjust temperature values when returns are required."""
        has_return = df["RETURN_TRACKING_NUMBER"] != "N"
        
        adjustments = {
            "Ambient": "Controlled Ambient",
            "Controlled Ambient, Ambient": "Controlled Ambient",
            "Refrigerated, Ambient": "Refrigerated",
            "Refrigerated, Controlled Ambient, Ambient": "Refrigerated, Controlled Ambient",
        }
        
        for old_temp, new_temp in adjustments.items():
            mask = (df["TEMPERATURE"] == old_temp) & has_return
            df.loc[mask, "TEMPERATURE"] = new_temp
        
        return df
