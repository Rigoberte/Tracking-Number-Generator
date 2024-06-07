import os

import pandas as pd
import datetime as dt
from typing import Tuple

from .team import Team

class GPMArgentinaTeam(Team):
    def __init__(self, folder_path_to_download: str = ""):
        self.carrierWebpage = self.__build_carrier_Webpage__("Transportes Ambientales", folder_path_to_download)

    def getTeamName(self) -> str:
        return "GPM Argentina"

    def getTeamEmail(self) -> str:
        return "inaki.costa@thermofisher.com"

    def get_column_rename_type_config_for_contacts_table(self) -> Tuple[dict, dict]:
        columns_names = {"Linea de facturacion" : "STUDY", "Site": "SITE#", "Site ID": "CARRIER_ID",
                        "Persona de contacto" : "CONTACTS", "Emails": "MEDICAL_CENTER_EMAILS", 
                        "Emails2": "CUSTOMER_EMAIL", "Emails3": "CRA_EMAILS"}
        columns_types = {"Linea de facturacion": str, "Site": str, "Site ID": str, "Persona de contacto": str}
        return columns_names, columns_types
    
    def apply_team_specific_changes_for_contacts_table(self, contactsDataFrame: pd.DataFrame) -> pd.DataFrame:
        contactsDataFrame["STUDY"] = contactsDataFrame["STUDY"].str.strip()
        contactsDataFrame["SITE#"] = contactsDataFrame["SITE#"].str.strip()

        contactsDataFrame["CAN_RECEIVE_MEDICINES"] = contactsDataFrame["CAN_RECEIVE_MEDICINES"] != ""
        contactsDataFrame["CAN_RECEIVE_ANCILLARIES"] = contactsDataFrame["CAN_RECEIVE_ANCILLARIES"] != ""
        
        contactsDataFrame["CUSTOMER_EMAIL"] = ""
        contactsDataFrame["CRA_EMAILS"] = ""
        return contactsDataFrame
    
    def get_data_path(self) -> Tuple[str, str, str]:
        path_from_get_data = os.path.expanduser("~\\Downloads\orderTracker_GPM.xlsx")
        orders_sheet = "Ordenes"
        contacts_sheet = "Contactos"
        return path_from_get_data, orders_sheet, contacts_sheet
    
    def get_column_rename_type_config_for_orders_tables(self) -> Tuple[dict, dict]:
        columns_names = {}
        columns_types = {"SYSTEM_NUMBER": str, "IVRS_NUMBER": str,"STUDY": str, "SITE#": str, 
                            "SHIP_DATE": dt.datetime, "SHIP_TIME_FROM": dt.datetime, "SHIP_TIME_TO": dt.datetime,
                            "DELIVERY_DATE": dt.datetime, "DELIVERY_TIME_FROM": dt.datetime, "DELIVERY_TIME_TO": dt.datetime,
                            "TYPE_OF_MATERIAL": str, "TEMPERATURE": str, 
                            "AMOUNT_OF_BOXES_TO_SEND": str, 
                            "HAS_RETURN": bool, 
                            "RETURN_TO_CARRIER_DEPOT": bool, "TYPE_OF_RETURN": str, 
                            "AMOUNT_OF_BOXES_TO_RETURN": str, 
                            "TRACKING_NUMBER": str, 
                            "RETURN_TRACKING_NUMBER": str, "PRINT_RETURN_DOCUMENT": bool, 
                            "CONTACTS": str, "CARRIER_ID": str}
        return columns_names, columns_types

    def apply_team_specific_changes_for_orders_tables(self, ordersDataFrame: pd.DataFrame) -> pd.DataFrame:
        temperatures = {"Ambiente": "Ambient", "Ambiente Controlado": "Controlled Ambient", "Refrigerado": "Refrigerated"}
        ordersDataFrame["TEMPERATURE"] = ordersDataFrame["TEMPERATURE"].replace(temperatures)
        
        ordersDataFrame["PRINT_RETURN_DOCUMENT"] = False
        
        ordersDataFrame["HAS_RETURN"] = (ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & (ordersDataFrame["TEMPERATURE"] != "Ambient")
        ordersDataFrame.loc[ordersDataFrame["HAS_RETURN"], "TYPE_OF_RETURN"] = "CREDO"
        ordersDataFrame["RETURN_TO_CARRIER_DEPOT"] = ordersDataFrame["HAS_RETURN"]

        return ordersDataFrame

    def sendEmailWithOrdersToTeam(self, folder_path_with_orders_files: str, date: str):
        self.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "")

    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            ordersDataFrame = pd.read_excel(path_from_get_data, sheet_name=orders_sheet, dtype=columns_types, header=0, skiprows=7)
        except Exception as e:
            raise e
        return ordersDataFrame
    
    def readSitesExcel(self, path_from_get_data: str, contacts_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            contactsDataFrame = pd.read_excel(path_from_get_data, sheet_name=contacts_sheet, dtype=columns_types, header=0)
        except Exception as e:
            raise e
        return contactsDataFrame

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
