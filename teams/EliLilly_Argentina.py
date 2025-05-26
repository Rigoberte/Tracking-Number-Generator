import pandas as pd
import datetime as dt
import numpy as np
from typing import Tuple

from .team import Team
from logClass.log import Log
from emailSender.emailSender import EmailSender

class EliLillyArgentinaTeam(Team):
    def __init__(self, folder_path_to_download: str, log : Log):
        super().__init__(log)
        self.carrierWebpage = self.__build_carrier_Webpage__("Transportes Ambientales HTTP", folder_path_to_download)

    def get_team_name(self) -> str:
        return "Eli Lilly Argentina"

    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            ordersDataFrame = pd.read_excel(path_from_get_data, sheet_name=orders_sheet, dtype=columns_types, header=0)
        except Exception as e:
            raise e
        return ordersDataFrame
    
    def readContactsExcel(self, path_from_get_data: str, contacts_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            contactsDataFrame = pd.read_excel(path_from_get_data, sheet_name=contacts_sheet, dtype=str, header=0)

            # Función para convertir valores según el tipo deseado
            def convertir_valor(valor, tipo):
                if pd.isna(valor):
                    return np.nan
                try:
                    if tipo == 'int':
                        return int(valor)
                    elif tipo == 'float':
                        return float(valor)
                    elif tipo == 'str':
                        return str(valor)
                    elif tipo == 'bool':
                        return bool(int(valor))
                    else:
                        return valor
                except ValueError:
                    return np.nan

            # Aplicar la conversión de tipos según el diccionario
            for column, _type in columns_types.items():
                if column in contactsDataFrame.columns:
                    contactsDataFrame[column] = contactsDataFrame[column].apply(lambda x: convertir_valor(x, _type))

        except Exception as e:
            raise e
        
        return contactsDataFrame

    def get_column_rename_type_config_for_contacts_table(self) -> Tuple[dict, dict]:
        columns_names = {"Protocolo": "STUDY", "Codigo": "CARRIER_ID", "Site": "SITE#",
                        "Horario inicio": "DELIVERY_TIME_FROM", "Horario fin": "DELIVERY_TIME_TO"}
        columns_types = {"Protocolo": str, "Site": str, "Codigo": str, "Horario inicio": str, "Horario fin": str}
        return columns_names, columns_types
    
    def apply_team_specific_changes_for_contacts_table(self, contactsDataFrame: pd.DataFrame) -> pd.DataFrame:
        if not "CONTACTS" in contactsDataFrame.columns:
            contactsDataFrame["CONTACTS"] = "No contact"
        
        contactsDataFrame["CAN_RECEIVE_MEDICINES"] = True
        contactsDataFrame["CAN_RECEIVE_ANCILLARIES_TYPE1"] = False
        contactsDataFrame["CAN_RECEIVE_ANCILLARIES_TYPE2"] = False
        contactsDataFrame["CAN_RECEIVE_EQUIPMENTS"] = False
        contactsDataFrame["MEDICAL_CENTER_EMAILS"] = ""
        contactsDataFrame["CUSTOMER_EMAIL"] = ""
        contactsDataFrame["CRA_EMAILS"] = ""
        return contactsDataFrame
    
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
                        "Order received": str, 
                        "Horario de Despacho": str,
                        "Destination": str,
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

    def send_email_to_team_with_orders(self, folder_path_with_orders_files: str, date: str,
                totalAmountOfOrders: int, amountOfOrdersProcessed: int, amountOfOrdersReadyToBeProcessed: int, emailSender: EmailSender) -> None:
        self.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "inaki.costa@thermofisher",
                    totalAmountOfOrders, amountOfOrdersProcessed, amountOfOrdersReadyToBeProcessed, emailSender)

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
        columns_names = {"date" : "DATE"}
        columns_types = {"date": 'datetime64[ns]'}
        return columns_names, columns_types
    
    def readNotWorkingDaysExcel(self, path_from_get_data: str, not_working_days_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            notWorkingDaysDataFrame = pd.read_excel(path_from_get_data, sheet_name=not_working_days_sheet, dtype=columns_types, header=0)
        except Exception as e:
            raise e
        return notWorkingDaysDataFrame