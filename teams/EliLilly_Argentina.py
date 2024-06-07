import os

import pandas as pd
import datetime as dt
from typing import Tuple

from .team import Team

class EliLillyArgentinaTeam(Team):
    def __init__(self, folder_path_to_download: str = ""):
        self.carrierWebpage = self.__build_carrier_Webpage__("Transportes Ambientales", folder_path_to_download)

    def getTeamName(self) -> str:
        return "Eli Lilly Argentina"

    def getTeamEmail(self) -> str:
        return "guido.hendl@thermofisher.com; florencia.acosta@thermofisher.com"
        return "AR.Lilly.logistics@fishersci.com"

    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            ordersDataFrame = pd.read_excel(path_from_get_data, sheet_name=orders_sheet, dtype=columns_types, header=0)
        except Exception as e:
            raise e
        return ordersDataFrame
    
    def readSitesExcel(self, path_from_get_data: str, contacts_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            contactsDataFrame = pd.read_excel(path_from_get_data, sheet_name=contacts_sheet, dtype=columns_types, header=0)
        except Exception as e:
            raise e
        return contactsDataFrame

    def get_column_rename_type_config_for_contacts_table(self) -> Tuple[dict, dict]:
        columns_names = {"Protocolo": "STUDY", "Codigo": "CARRIER_ID", "Site": "SITE#",
                        "Horario inicio": "DELIVERY_TIME_FROM", "Horario fin": "DELIVERY_TIME_TO"}
        columns_types = {"Protocolo": str, "Site": str, "Codigo": str, "Horario inicio": str, "Horario fin": str}
        return columns_names, columns_types
    
    def apply_team_specific_changes_for_contacts_table(self, contactsDataFrame: pd.DataFrame) -> pd.DataFrame:
        contactsDataFrame["CONTACTS"] = "No contact"
        contactsDataFrame["CAN_RECEIVE_MEDICINES"] = True
        contactsDataFrame["CAN_RECEIVE_ANCILLARIES"] = False
        contactsDataFrame["MEDICAL_CENTER_EMAILS"] = ""
        contactsDataFrame["CUSTOMER_EMAIL"] = ""
        contactsDataFrame["CRA_EMAILS"] = ""
        return contactsDataFrame
    
    def get_data_path(self) -> Tuple[str, str, str]:
        path_from_get_data = os.path.expanduser("~\\Thermo Fisher Scientific\Power BI Lilly Argentina - General\Share Point Lilly Argentina.xlsx")
        orders_sheet = "Shipments"
        contacts_sheet = "Dias y Destinos"
        return path_from_get_data, orders_sheet, contacts_sheet
    
    def get_column_rename_type_config_for_orders_tables(self) -> Tuple[dict, dict]:
        columns_names = {"CT-WIN": "SYSTEM_NUMBER", "IVRS": "IVRS_NUMBER",
                        "Trial Alias": "STUDY", "Site ": "SITE#",
                        "Order received": "ENTER DATE", "Ship date": "SHIP_DATE",
                        "Horario de Despacho": "SHIP_TIME_FROM",  
                        "Dia de entrega": "DELIVERY_DATE", "Destination": "DESTINATION",
                        "CONDICION": "TEMPERATURE", "TT4": "AMOUNT_OF_BOXES_TO_SEND",  
                        "AWB": "TRACKING_NUMBER", "Shipper return AWB": "RETURN_TRACKING_NUMBER"}
        columns_types = {"CT-WIN": str, "IVRS": str, 
                        "Trial Alias": str, "Site ": str, 
                        "Order received": str, "Ship date": dt.datetime,
                        "Horario de Despacho": str,
                        "Dia de entrega": dt.datetime, "Destination": str,
                        "CONDICION": str, "TT4": str, 
                        "AWB": str, "Shipper return AWB": str}
        return columns_names, columns_types
    
    def apply_team_specific_changes_for_orders_tables(self, ordersDataFrame: pd.DataFrame) -> pd.DataFrame:
        ordersDataFrame["TYPE_OF_MATERIAL"] = "Medicine"
        ordersDataFrame["CUSTOMER"] = "Eli Lilly and Company"

        temperatures = {"L": "Ambient",
                        "M": "Controlled Ambient", "M + L": "Controlled Ambient",
                        "H": "Controlled Ambient", "H + M": "Controlled Ambient", "H + L": "Controlled Ambient", "H + M + L": "Controlled Ambient",
                        "REF": "Refrigerated", "REF + H": "Refrigerated", "REF + M": "Refrigerated", "REF + L": "Refrigerated",
                        "REF + H + M": "Refrigerated", "REF + H + L": "Refrigerated", "REF + M + L": "Refrigerated",
                        "REF + H + M + L": "Refrigerated"}
        ordersDataFrame["TEMPERATURE"] = ordersDataFrame["TEMPERATURE"].replace(temperatures)
        ordersDataFrame.loc[(ordersDataFrame["TEMPERATURE"] == "Ambient") & (ordersDataFrame["RETURN_TRACKING_NUMBER"] != "N"), "TEMPERATURE"] = "Controlled Ambient"
        
        ordersDataFrame["Cajas (Carton)"] = ordersDataFrame["Cajas (Carton)"].replace("", 0).fillna(0).astype(int)
        ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] = ordersDataFrame["AMOUNT_OF_BOXES_TO_SEND"] - ordersDataFrame["Cajas (Carton)"]
        
        ordersDataFrame["RETURN_TO_CARRIER_DEPOT"] = False
        
        ordersDataFrame["HAS_RETURN"] = (ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & (ordersDataFrame["TEMPERATURE"] != "Ambient")
        ordersDataFrame.loc[ordersDataFrame["HAS_RETURN"], "TYPE_OF_RETURN"] = "CREDO"

        ordersDataFrame["PRINT_RETURN_DOCUMENT"] = ordersDataFrame["HAS_RETURN"]

        return ordersDataFrame

    def sendEmailWithOrdersToTeam(self, folder_path_with_orders_files: str, date: str):
        self.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "inaki.costa@thermofisher")

    def build_driver(self):
        self.__build_driver__(self.carrierWebpage)

    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        return self.__check_if_user_and_password_are_correct__(self.carrierWebpage, username, password)
    
    def quit_driver(self):
        self.__quit_driver__(self.carrierWebpage)

    def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                    ship_date: str, ship_time_from: str, ship_time_to: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        return self.__complete_shipping_order_form__(self.carrierWebpage, carrier_id, reference,
                                    ship_date, ship_time_from, ship_time_to,
                                    delivery_date, delivery_time_from, delivery_time_to,
                                    type_of_material, temperature,
                                    contacts, amount_of_boxes)
    
    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                            delivery_date: str, return_time_from: str,
                                            return_time_to: str, type_of_return: str,
                                            contacts: str, amount_of_boxes_to_return: int,
                                            return_to_TA: bool, tracking_number: str) -> str:
        return self.__complete_shipping_order_return_form__(self.carrierWebpage, carrier_id, reference_return,
                                            delivery_date, return_time_from,
                                            return_time_to, type_of_return,
                                            contacts, amount_of_boxes_to_return,
                                            return_to_TA, tracking_number)
    
    def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
        self.__printWayBillDocument__(self.carrierWebpage, tracking_number, amount_of_copies)

    def printLabelDocument(self, tracking_number: str):
        self.__printLabelDocument__(self.carrierWebpage, tracking_number)

    def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
        self.__printReturnWayBillDocument__(self.carrierWebpage, return_tracking_number, amount_of_copies)
