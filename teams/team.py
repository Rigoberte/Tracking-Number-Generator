import os
import win32com.client as win32
import pandas as pd
from abc import ABC, abstractmethod
from typing import Tuple
import time

from carriersWebpage.carrierWebPage import CarrierWebpage
from carriersWebpage.carrierWebPage_factory import CarrierWebPageFactory
from .team_factory import TeamFactory
from logClass.log import Log
from utils.zip_folder import zip_folder
from dataPathController.dataPathController import DataPathController

class Team(ABC):
    """
    Abstract class for teams
    """
    def __init__(self, log: Log):
        """
        Class constructor for teams
        """
        self.log = log
        pass

    def get_team_names(self) -> list:
        """
        Gets teams names
        """
        return TeamFactory.get_team_names()

    def get_data_path(self, vars_to_returns: list = []) -> list:
        """
        Loads data path

        Returns:
            str: excel file path
            str: excel sheet name
            str: excel sheet name with sites info
            str: excel sheet name with not working days
        """
        possible_vars = ["team_excel_path", 
                        "team_orders_sheet",
                        "team_contacts_sheet",
                        "team_not_working_days_sheet",
                        "team_send_email_to_medical_centers",
                        "team_email"]

        result = []
        config = DataPathController().get_config_of_a_team(self.get_team_name())
        for var in vars_to_returns:
            if var in possible_vars:
                result.append(config[var])
        
        return result
    
    def getTeamEmail(self) -> str:
        """
        Gets team email
        """
        return self.get_data_path(["team_email"])[0]
    
    # Methods to be implemented by each sub class
    @abstractmethod
    def get_team_name(self) -> str:
        """
        Gets team name
        """
        pass

    @abstractmethod
    def get_column_rename_type_config_for_contacts_table(self) -> Tuple[dict, dict]:
        """
        Loads columns names and types for the sites info table

        Returns:
            dict: columns names
            dict: columns types
        """
        pass
    
    @abstractmethod
    def apply_team_specific_changes_for_contacts_table(self, contactsDataframe: pd.DataFrame) -> pd.DataFrame:
        """
        Applies team specific changes to contacts table

        Args:
            contactsDataframe (DataFrame): contacts table

        Returns:
            DataFrame: contacts table with team specific changes
        """
        pass
    
    @abstractmethod
    def get_column_rename_type_config_for_orders_tables(self) -> Tuple[dict, dict]:
        """
        Loads columns names and types for the orders table

        Returns:
            dict: columns names
            dict: columns types
        """
        pass
    
    @abstractmethod
    def apply_team_specific_changes_for_orders_tables(self, ordersDataframe: pd.DataFrame) -> pd.DataFrame:
        """
        Applies team specific changes to orders table

        Args:
            ordersDataframe (DataFrame): orders table

        Returns:
            DataFrame: orders table with team specific changes
        """
        try:
            pass
        except Exception as e:
            self.log.add_error_log(f"Error applying team specific changes for orders tables: {e}")
            return ordersDataframe
    
    @abstractmethod
    def send_email_to_team_with_orders(self, folder_path_with_orders_files: str, date: str,
                totalAmountOfOrders: int, amountOfOrdersProcessed: int, amountOfOrdersReadyToBeProcessed: int):
        """
        Sends an email with the orders table to the team

        Args:
            folder_path_with_orders_files (str): folder path with orders files
            date (str): date
            totalAmountOfOrders (int): total amount of orders
            amountOfOrdersProcessed (int): amount of orders processed
            amountOfOrdersReadyToBeProcessed (int): amount of orders ready to be processed
        """
        pass

    @abstractmethod
    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        """
        Reads orders excel

        Args:
            path_from_get_data (str): path from get data
            orders_sheet (str): orders sheet
            columns_types (dict): columns types

        Returns:
            DataFrame: orders table
        """
        try:
            pass
        except Exception as e:
            self.log.add_error_log(f"Error reading orders excel: {e}")
            return self.__getEmptyOrdersDataFrame__()
    
    @abstractmethod
    def readContactsExcel(self, path_from_get_data: str, contacts_sheet: str, columns_types: dict) -> pd.DataFrame:
        """
        Reads contacts excel

        Args:
            path_from_get_data (str): path from get data
            sites_sheet (str): contacts sheet
            columns_types (dict): columns types

        Returns:
            DataFrame: contacts table
        """
        try:
            pass
        except Exception as e:
            self.log.add_error_log(f"Error reading contacts excel: {e}")
            return self.__getEmptyContactsDataFrame__()

    @abstractmethod
    def build_driver(self) -> None:
        """
        Builds driver
        """
        pass

    @abstractmethod
    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        """
        Checks if user and password are correct
        """
        pass

    @abstractmethod
    def quit_driver(self) -> None:
        """
        Quits driver
        """
        pass

    @abstractmethod
    def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                    ship_date: str, ship_time_from: str, ship_time_to: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        """
        Completes shipping order form
        """
        pass

    @abstractmethod
    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                            return_delivery_date: str, return_delivery_hour_from: str,
                                            return_delivery_hour_to: str, type_of_return: str,
                                            contacts: str, amount_of_boxes_to_return: int,
                                            return_to_TA: bool, tracking_number: str) -> str:
        """
        Completes shipping order return form
        """
        pass

    @abstractmethod
    def print_wayBill_document(self, tracking_number: str, amount_of_copies: int) -> None:
        """
        Prints way bill document
        """
        pass

    @abstractmethod
    def print_label_document(self, tracking_number: str) -> None:
        """
        Prints label document
        """
        pass

    @abstractmethod
    def print_return_wayBill_document(self, return_tracking_number: str, amount_of_copies: int) -> None:
        """
        Prints return way bill document
        """
        pass

    @abstractmethod
    def get_column_rename_type_config_for_not_working_days_table(self) -> Tuple[dict, dict]:
        """
        Loads columns names and types for the not working days table

        Returns:
            dict: columns names
            dict: columns types
        """
        pass

    @abstractmethod
    def readNotWorkingDaysExcel(self, path_from_get_data: str, not_working_days_sheet: str, columns_types: dict) -> pd.DataFrame:
        """
        Reads not working days excel

        Args:
            path_from_get_data (str): path from get data
            not_working_days_sheet (str): not working days sheet
            columns_types (dict): columns types

        Returns:
            DataFrame: not working days table
        """
        try:
            pass
        except Exception as e:
            self.log.add_error_log(f"Error reading not working days excel: {e}")
            return pd.DataFrame(columns=["DATE"])

    # Private methods
    def __sendEmailWithOrdersToTeam__(self, 
            folder_path_with_orders_files: str, date: str, 
            emails_of_team: str, emails_of_admin: str,
            totalAmountOfOrders: int, amountOfOrdersProcessed: int, amountOfOrdersReadyToBeProcessed: int):
        
        try:
            # Obtener la carpeta padre
            parent_folder = os.path.dirname(folder_path_with_orders_files)
            
            # Nombre del archivo ZIP sin extensión
            zip_filename = 'orders_' + self.get_team_name() + '_' + date
            
            # Ruta completa del archivo ZIP
            zip_path = os.path.join(parent_folder, zip_filename)

            # Crear el archivo ZIP
            zip_folder(folder_path_with_orders_files, zip_path)

            # Ruta del archivo ZIP con extensión
            zip_file_with_extension = zip_path + '.zip'

            time.sleep(1)

            outlook = win32.Dispatch('outlook.application')

            mail = outlook.CreateItem(0)
            mail.To = emails_of_team
            mail.Cc = emails_of_admin
            mail.Subject = f"Shipping orders with dispatch date {date} - {self.get_team_name()}"
            
            emailSource = self.__get_email_source_from_TXT_file__("media/email_to_team.txt")
            emailSource = self.__replace_email_values__(emailSource, self.get_team_name(), date, 
                        totalAmountOfOrders, amountOfOrdersProcessed, amountOfOrdersReadyToBeProcessed)
            mail.HTMLBody = emailSource

            # Adjuntar el archivo ZIP
            mail.Attachments.Add(zip_file_with_extension)

            time.sleep(1)

            mail.Send()
        except Exception as e:
            self.log.add_error_log(f"Error sending email with orders to team: {e}")

    def __get_email_source_from_TXT_file__(self, file) -> str:
            with open(file, 'r') as file:
                return file.read()
            
    def __replace_email_values__(self, emailSource: str, 
        selected_team_name: str, ship_date: str, total_amount_of_orders: int, 
        amount_of_orders_processed: int, amount_of_orders_not_processed: int) -> str:
        
        emailSource = emailSource.replace("|VAR_SELECTED_TEAM|", selected_team_name)
        emailSource = emailSource.replace("|VAR_SHIP_DATE|", ship_date)
        emailSource = emailSource.replace("|VAR_TOTAL_AMOUNT_OF_ORDERS|", str(total_amount_of_orders))
        emailSource = emailSource.replace("|VAR_AMOUNT_OF_ORDERS_PROCESSED|", str(amount_of_orders_processed))
        emailSource = emailSource.replace("|VAR_AMOUNT_OF_ORDERS_NOT_PROCESSED|", str(amount_of_orders_not_processed))
        emailSource = emailSource.replace("|VAR_TMO_LOGO|", os.getcwd() + "\\media\\TMO_logo_email.jpg")

        return emailSource

    def __check_if_user_and_password_are_correct__(self, carrierWebpage: CarrierWebpage, username: str, password: str) -> bool:
        return carrierWebpage.check_if_user_and_password_are_correct(username, password)
    
    def __build_driver__(self, carrierWebpage: CarrierWebpage) -> None:
        carrierWebpage.build_driver()

    def __quit_driver__(self, carrierWebpage: CarrierWebpage) -> None:
        carrierWebpage.quit_driver()

    def __build_carrier_Webpage__(self, carrierWebpage_name: str, folder_path_to_download: str) -> CarrierWebpage:
        return CarrierWebPageFactory().create_carrier_webpage(carrierWebpage_name, folder_path_to_download, self.log)

    def __complete_shipping_order_form__(self, carrierWebpage: CarrierWebpage, carrier_id: str, reference: str,
                                    ship_date: str, ship_time_from: str, ship_time_to: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        return carrierWebpage.complete_shipping_order_form(carrier_id, reference,
                                    ship_date, ship_time_from, ship_time_to,
                                    delivery_date, delivery_time_from, delivery_time_to,
                                    type_of_material, temperature,
                                    contacts, amount_of_boxes)

    def __complete_shipping_order_return_form__(self, carrierWebpage: CarrierWebpage, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
        return carrierWebpage.complete_shipping_order_return_form(carrier_id, reference_return,
                                                delivery_date, return_time_from,
                                                return_time_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)
    
    def __printWayBillDocument__(self, carrierWebpage: CarrierWebpage, tracking_number: str, amount_of_copies: int) -> None:
        carrierWebpage.print_wayBill_document(tracking_number, amount_of_copies)

    def __printLabelDocument__(self, carrierWebpage: CarrierWebpage, tracking_number: str) -> None:
        carrierWebpage.print_label_document(tracking_number)

    def __printReturnWayBillDocument__(self, carrierWebpage: CarrierWebpage, return_tracking_number: str, amount_of_copies: int) -> None:
        carrierWebpage.print_return_wayBill_document(return_tracking_number, amount_of_copies)

    def __getEmptyOrdersDataFrame__(self) -> pd.DataFrame:
        from dataRecolector.dataRecolector import DataRecolector
        
        noSelectedTeam = CarrierWebPageFactory().create_carrier_webpage("NoCarrier", "", self.log)
        return DataRecolector(noSelectedTeam, log= self.log).get_empty_orders_dataFrame()
    
    def __getEmptyContactsDataFrame__(self) -> pd.DataFrame:
        from dataRecolector.dataRecolector import DataRecolector

        noSelectedTeam = CarrierWebPageFactory().create_carrier_webpage("NoCarrier", "", self.log)
        return DataRecolector(noSelectedTeam, log= self.log).get_empty_contacts_dataFrame()