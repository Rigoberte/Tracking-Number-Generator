"""
GPM Argentina Team Configuration.
"""
import pandas as pd

from .team_config import TeamConfig, ColumnMapping


class GPMArgentinaConfig(TeamConfig):
    """Configuration for GPM Argentina team."""
    
    @property
    def name(self) -> str:
        return "GPM Argentina"
    
    @property
    def carrier_name(self) -> str:
        return "Transportes Ambientales"
    
    @property
    def customer_name(self) -> str:
        return "GPM"
    
    @property
    def orders_skip_rows(self) -> int:
        """Number of rows to skip when reading orders Excel."""
        return 7
    
    def get_orders_column_mapping(self) -> ColumnMapping:
        # GPM uses standard column names, minimal renaming needed
        return ColumnMapping(
            rename_map={},
            type_map={
                "SYSTEM_NUMBER": str,
                "IVRS_NUMBER": str,
                "STUDY": str,
                "SITE#": str,
                "TYPE_OF_MATERIAL": str,
                "TEMPERATURE": str,
                "AMOUNT_OF_BOXES_TO_SEND": str,
                "HAS_RETURN": bool,
                "RETURN_TO_CARRIER_DEPOT": bool,
                "TYPE_OF_RETURN": str,
                "DISPOSABLE_BOXES": str,
                "TRACKING_NUMBER": str,
                "RETURN_TRACKING_NUMBER": str,
                "PRINT_RETURN_DOCUMENT": bool,
                "CONTACTS": str,
                "CARRIER_ID": str,
            }
        )
    
    def get_contacts_column_mapping(self) -> ColumnMapping:
        return ColumnMapping(
            rename_map={
                "STUDY": "STUDY",
                "Site": "SITE#",
                "Site ID": "CARRIER_ID",
                "CONTACTOS": "CONTACTS",
                "EMAILS": "MEDICAL_CENTER_EMAILS",
                "Emails2": "CUSTOMER_EMAIL",
                "Emails3": "CRA_EMAILS",
            },
            type_map={
                "STUDY": str,
                "Site": str,
                "Site ID": str,
                "CONTACTOS": str,
                "EMAILS": str,
                "CAN_RECEIVE_MEDICINES": str,
                "CAN_RECEIVE_ANCILLARIES_TYPE1": str,
                "CAN_RECEIVE_ANCILLARIES_TYPE2": str,
                "CAN_RECEIVE_EQUIPMENTS": str,
                "CUSTOMER_EMAIL": str,
            }
        )
    
    def get_holidays_column_mapping(self) -> ColumnMapping:
        return ColumnMapping(
            rename_map={"Date": "DATE"},
            type_map={"Date": "datetime64[ns]"}
        )
    
    def transform_orders(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply GPM specific order transformations."""
        # Temperature normalization
        df["TEMPERATURE"] = df["TEMPERATURE"].replace(self._get_temperature_mappings())
        
        # Material type normalization
        df["TYPE_OF_MATERIAL"] = df["TYPE_OF_MATERIAL"].replace(self._get_material_mappings())
        
        # Calculate boxes to return
        df["DISPOSABLE_BOXES"] = df["DISPOSABLE_BOXES"].replace("", 0).fillna(0).astype(int)
        df["AMOUNT_OF_BOXES_TO_RETURN"] = df["AMOUNT_OF_BOXES_TO_SEND"] - df["DISPOSABLE_BOXES"]
        
        # Return flags
        df["HAS_RETURN"] = (
            (df["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & 
            (df["TEMPERATURE"] != "Ambient")
        )
        df.loc[df["HAS_RETURN"], "TYPE_OF_RETURN"] = "CREDO"
        df["RETURN_TO_CARRIER_DEPOT"] = df["HAS_RETURN"]
        
        return df
    
    def transform_contacts(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply GPM specific contact transformations."""
        # Strip whitespace
        df["STUDY"] = df["STUDY"].str.strip()
        df["SITE#"] = df["SITE#"].str.strip()
        
        # Convert string flags to boolean
        df["CAN_RECEIVE_MEDICINES"] = df["CAN_RECEIVE_MEDICINES"] != ""
        df["CAN_RECEIVE_ANCILLARIES_TYPE1"] = df["CAN_RECEIVE_ANCILLARIES_TYPE1"] != ""
        df["CAN_RECEIVE_ANCILLARIES_TYPE2"] = df["CAN_RECEIVE_ANCILLARIES_TYPE2"] != ""
        df["CAN_RECEIVE_EQUIPMENTS"] = df["CAN_RECEIVE_EQUIPMENTS"] != ""
        
        # Clear email fields (GPM handles emails differently)
        df["CUSTOMER_EMAIL"] = ""
        df["CRA_EMAILS"] = ""
        
        return df
    
    def _get_temperature_mappings(self) -> dict:
        return {
            "Ambiente": "Ambient",
            "AMB": "Ambient",
            "Ambiente Controlado": "Controlled Ambient",
            "CON": "Controlled Ambient",
            "Refrigerado": "Refrigerated",
            "REF": "Refrigerated",
            "Congelado": "Frozen",
            "FRO": "Frozen",
            "Refrigerado con Hielo Seco": "Refrigerated with Dry Ice",
            "RHS": "Refrigerated with Dry Ice",
            "Congelado con Nitrogeno Liquido": "Frozen with Liquid Nitrogen",
            "FNL": "Frozen with Liquid Nitrogen",
        }
    
    def _get_material_mappings(self) -> dict:
        return {
            "MED": "Medicine",
            "ANC": "Ancillaries Type 1",
            "ANC1": "Ancillary Type 1",
            "ANC2": "Ancillary Type 2",
            "EQUIP": "Equipment",
        }
