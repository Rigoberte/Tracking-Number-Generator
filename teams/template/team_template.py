import os

import pandas as pd
import datetime as dt
from typing import Tuple

from .team import Team

class NAMETeam(Team):
    def __init__(self, folder_path_to_download: str = ""):
        raise NotImplementedError("Method not implemented")
        self.carrierWebpage = self.__build_carrier_Webpage__("CARRIER NAME", folder_path_to_download)

    def getTeamName(self) -> str:
        raise NotImplementedError("Method not implemented")
        return "TEAM NAME"

    def getTeamEmail(self) -> str:
        raise NotImplementedError("Method not implemented")
        return "TEAM_EMAIL@MAIL.com"

    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        raise NotImplementedError("Method not implemented")
        try:
            ordersDataFrame = pd.read_excel(path_from_get_data, sheet_name=orders_sheet, dtype=columns_types, header=0)
        except Exception as e:
            raise e
        return ordersDataFrame
    
    def readSitesExcel(self, path_from_get_data: str, contacts_sheet: str, columns_types: dict) -> pd.DataFrame:
        
        raise NotImplementedError("Method not implemented")
        try:
            contactsDataFrame = pd.read_excel(path_from_get_data, sheet_name=contacts_sheet, dtype=columns_types, header=0)
        except Exception as e:
            raise e
        return contactsDataFrame

    def get_column_rename_type_config_for_contacts_table(self) -> Tuple[dict, dict]:
        """
        Columns names must turn columns into DataRecolector.columns_for_contacts 
        """
        raise NotImplementedError("Method not implemented")
        columns_names = {}
        columns_types = {}
    
        return columns_names, columns_types
    
    def apply_team_specific_changes_for_contacts_table(self, contactsDataFrame: pd.DataFrame) -> pd.DataFrame:
        raise NotImplementedError("Method not implemented")
        return contactsDataFrame
    
    def get_column_rename_type_config_for_orders_tables(self) -> Tuple[dict, dict]:
        """
        Columns names must turn columns into DataRecolector.columns_for_orders
        """
        raise NotImplementedError("Method not implemented")
        columns_names = {}
        columns_types = {}
        return columns_names, columns_types
    
    def apply_team_specific_changes_for_orders_tables(self, ordersDataFrame: pd.DataFrame) -> pd.DataFrame:
        raise NotImplementedError("Method not implemented")
        return ordersDataFrame

    def sendEmailWithOrdersToTeam(self, folder_path_with_orders_files: str, date: str) -> None:
        self.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "inaki.costa@thermofisher")

    def build_driver(self) -> None:
        self.__build_driver__(self.carrierWebpage)

    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        return self.__check_if_user_and_password_are_correct__(self.carrierWebpage, username, password)
    
    def quit_driver(self) -> None:
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
    
    def printWayBillDocument(self, tracking_number: str, amount_of_copies: int) -> None:
        self.__printWayBillDocument__(self.carrierWebpage, tracking_number, amount_of_copies)

    def printLabelDocument(self, tracking_number: str = "") -> None:
        self.__printLabelDocument__(self.carrierWebpage, tracking_number)

    def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int) -> None:
        self.__printReturnWayBillDocument__(self.carrierWebpage, return_tracking_number, amount_of_copies)

    def get_column_rename_type_config_for_not_working_days_table(self) -> Tuple[dict, dict]:
        raise NotImplementedError("Method not implemented")
        columns_names = {}
        columns_types = {}
        return columns_names, columns_types
    
    def readNotWorkingDaysExcel(self, path_from_get_data: str, not_working_days_sheet: str, columns_types: dict) -> pd.DataFrame:
        raise NotImplementedError("Method not implemented")
        try:
            notWorkingDaysDataFrame = pd.read_excel(path_from_get_data, sheet_name=not_working_days_sheet, dtype=columns_types, header=0)
        except Exception as e:
            raise e
        return notWorkingDaysDataFrame