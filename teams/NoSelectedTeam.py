import pandas as pd
import datetime as dt
from typing import Tuple

from .team import Team
from logClass.log import Log
from emailSender.emailSender import EmailSender

class NoSelectedTeam(Team):
    def __init__(self, folder_path_to_download: str, log : Log):
        super().__init__(log)
        self.carrierWebpage = self.__build_carrier_Webpage__("", folder_path_to_download)

    def get_team_name(self) -> str:
        return "No Selected Team"
    
    def get_column_rename_type_config_for_contacts_table(self) -> Tuple[dict, dict]:
        return {}, {}
    
    def apply_team_specific_changes_for_contacts_table(self, contactsDataframe: pd.DataFrame) -> pd.DataFrame:
        contactsDataframe["CAN_RECEIVE_MEDICINES"] = False
        contactsDataframe["CAN_RECEIVE_ANCILLARIES"] = False
        return contactsDataframe
    
    def get_column_rename_type_config_for_orders_tables(self) -> Tuple[dict, dict]:
        return {}, {}
    
    def apply_team_specific_changes_for_orders_tables(self, ordersDataframe: pd.DataFrame) -> pd.DataFrame:
        return ordersDataframe

    def send_email_to_team_with_orders(self, folder_path_with_orders_files: str, date: str,
                totalAmountOfOrders: int, amountOfOrdersProcessed: int, amountOfOrdersReadyToBeProcessed: int, emailSender: EmailSender) -> None:
        self.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "inaki.costa@thermofisher",
                    totalAmountOfOrders, amountOfOrdersProcessed, amountOfOrdersReadyToBeProcessed, emailSender)
    
    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        return self.__getEmptyOrdersDataFrame__()
    
    def readContactsExcel(self, path_from_get_data: str, sites_sheet: str, columns_types: dict) -> pd.DataFrame:
        return self.__getEmptyContactsDataFrame__()

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
    
    def print_wayBill_document(self, tracking_number: str, amount_of_copies: int):
        self.__printWayBillDocument__(self.carrierWebpage, tracking_number, amount_of_copies)

    def print_label_document(self, tracking_number: str):
        self.__printLabelDocument__(self.carrierWebpage, tracking_number)

    def print_return_wayBill_document(self, return_tracking_number: str, amount_of_copies: int):
        self.__printReturnWayBillDocument__(self.carrierWebpage, return_tracking_number, amount_of_copies)

    def get_column_rename_type_config_for_not_working_days_table(self) -> Tuple[dict, dict]:
        return {}, {}
    
    def readNotWorkingDaysExcel(self, path_from_get_data: str, not_working_days_sheet: str, columns_types: dict) -> pd.DataFrame:
        return pd.DataFrame(columns=["DATE"])