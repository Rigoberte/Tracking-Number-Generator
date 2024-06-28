import pandas as pd
import datetime as dt
from typing import Tuple

from .team import Team

next_weekday = dt.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) + dt.timedelta(days= 7 - dt.datetime.today().weekday())
RELATIVE_DATE = next_weekday

class TestTeam(Team):
    def __init__(self, folder_path_to_download: str = ""):
        self.carrierWebpage = self.__build_carrier_Webpage__("Carrier Webpage For Testing", folder_path_to_download)

    def getTeamName(self) -> str:
        return "Test_5_ordenes"
    
    def getTeamEmail(self) -> str:
        return "inaki.costa@thermofisher.com"

    def get_column_rename_type_config_for_contacts_table(self) -> Tuple[dict, dict]:
        columns_names = {}
        columns_types = {"STUDY": str, "SITE#": str, "CARRIER_ID": str, "DELIVERY_TIME_FROM": dt.datetime, "DELIVERY_TIME_TO": dt.datetime}
        return columns_names, columns_types
    
    def apply_team_specific_changes_for_contacts_table(self, contactsDataFrame: pd.DataFrame) -> pd.DataFrame:
        contactsDataFrame["TYPE_OF_MATERIAL_CAN_RECEIVE"] = "Can receive medicines"
        contactsDataFrame["MEDICAL_CENTER_EMAILS"] = "inaki.costa@thermofisher.com"
        contactsDataFrame["CUSTOMER_EMAIL"] = ""
        contactsDataFrame["CRA_EMAILS"] = ""
        contactsDataFrame["CAN_RECEIVE_MEDICINES"] = contactsDataFrame["CAN_RECEIVE_MEDICINES"] != ""
        contactsDataFrame["CAN_RECEIVE_ANCILLARIES_TYPE1"] = contactsDataFrame["CAN_RECEIVE_ANCILLARIES_TYPE1"] != ""
        contactsDataFrame["CAN_RECEIVE_ANCILLARIES_TYPE2"] = contactsDataFrame["CAN_RECEIVE_ANCILLARIES_TYPE2"] != ""
        contactsDataFrame["CAN_RECEIVE_EQUIPMENTS"] = contactsDataFrame["CAN_RECEIVE_EQUIPMENTS"] != ""
        return contactsDataFrame
    
    def get_column_rename_type_config_for_orders_tables(self) -> Tuple[dict, dict]:
        columns_names = {}
        columns_types = {"SITE#": str, "RETURN_TO_CARRIER_DEPOT": bool, "SHIP_TIME_TO": str}
        return columns_names, columns_types
    
    def apply_team_specific_changes_for_orders_tables(self, ordersDataFrame: pd.DataFrame) -> pd.DataFrame:
        ordersDataFrame["PRINT_RETURN_DOCUMENT"] = True

        #ordersDataFrame["HAS_RETURN"] = (ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & (ordersDataFrame["TEMPERATURE"] != "Ambient")
        #ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] = ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"].apply(lambda x: x if x > 0 else 0)
        #ordersDataFrame["TYPE_OF_RETURN"] = ordersDataFrame["HAS_RETURN"].apply(lambda x: "CREDO" if x else "NA")

        return ordersDataFrame
    
    def sendEmailWithOrdersToTeam(self, folder_path_with_orders_files: str, date: str):
        self.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "")

    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            ordersDataFrame = self.__get_orders_for_testing__()
            
            return ordersDataFrame
        except Exception as e:
            raise e
    
    def readContactsExcel(self, path_from_get_data: str, contacts_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            contacts_df = pd.DataFrame(columns=["STUDY", "SITE#", "CARRIER_ID", "DELIVERY_TIME_FROM", "DELIVERY_TIME_TO", "CONTACTS", "CAN_RECEIVE_MEDICINES", "CAN_RECEIVE_ANCILLARIES_TYPE1", "CAN_RECEIVE_ANCILLARIES_TYPE2", "CAN_RECEIVE_EQUIPMENTS"])

            contact1 = {
                "STUDY": "TEST",
                "SITE#": "01",
                "CARRIER_ID": "5616",
                "DELIVERY_TIME_FROM": "10:00:00",
                "DELIVERY_TIME_TO": "12:00:00",
                "CONTACTS": "Contact 1",
                "CAN_RECEIVE_MEDICINES": "x",
                "CAN_RECEIVE_ANCILLARIES_TYPE1": "x",
                "CAN_RECEIVE_ANCILLARIES_TYPE2": "",
                "CAN_RECEIVE_EQUIPMENTS": ""
            }
            contacts_df = self.__append_a_line_to_dataframe__(contacts_df, contact1)

            contact2 = {
                "STUDY": "TEST",
                "SITE#": "02",
                "CARRIER_ID": "5617",
                "DELIVERY_TIME_FROM": "14:00:00",
                "DELIVERY_TIME_TO": "16:00:00",
                "CONTACTS": "Contact 2",
                "CAN_RECEIVE_MEDICINES": "x",
                "CAN_RECEIVE_ANCILLARIES_TYPE1": "",
                "CAN_RECEIVE_ANCILLARIES_TYPE2": "x",
                "CAN_RECEIVE_EQUIPMENTS": ""
            }
            contacts_df = self.__append_a_line_to_dataframe__(contacts_df, contact2)

            contact3 = {
                "STUDY": "TEST",
                "SITE#": "03",
                "CARRIER_ID": "5618",
                "DELIVERY_TIME_FROM": "08:00:00",
                "DELIVERY_TIME_TO": "10:00:00",
                "CONTACTS": "Contact 3",
                "CAN_RECEIVE_MEDICINES": "x",
                "CAN_RECEIVE_ANCILLARIES_TYPE1": "",
                "CAN_RECEIVE_ANCILLARIES_TYPE2": "",
                "CAN_RECEIVE_EQUIPMENTS": "x"
            }
            contacts_df = self.__append_a_line_to_dataframe__(contacts_df, contact3)

            return contacts_df
        except Exception as e:
            raise e

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

    def get_column_rename_type_config_for_not_working_days_table(self) -> Tuple[dict, dict]:
        columns_names = {}
        columns_types = {"DATE": dt.datetime}
        return columns_names, columns_types
    
    def readNotWorkingDaysExcel(self, path_from_get_data: str, not_working_days_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            notWorkingDaysDataFrame = pd.DataFrame(columns=["DATE"])
        except Exception as e:
            raise e
        return notWorkingDaysDataFrame
    
    def __append_a_line_to_dataframe__(self, dataFrame: pd.DataFrame, line: dict) -> pd.DataFrame:
        order_df = pd.DataFrame([line])
        ordersDataFrame = pd.concat([dataFrame, order_df], ignore_index=True)
        return ordersDataFrame
    
    def __get_order_template__(self, 
                        system_number = "1",
                        ivrs_number = "IVRS_NUMBER__01",
                        customer = "CUSTOMER__01",
                        study = "TEST",
                        site = "01",
                        ship_date = RELATIVE_DATE,
                        ship_time_from = "19:00:00",
                        delivery_date = (RELATIVE_DATE + dt.timedelta(days=1)),
                        type_of_material = "Medicine",
                        temperature = "Refrigerated",
                        amount_of_boxes_to_send = 1,
                        has_return = "TRUE",
                        return_to_carrier_depot = "TRUE",
                        type_of_return = "CREDO",
                        amount_of_boxes_to_return = 1,
                        tracking_number = "",
                        return_tracking_number = ""):

        order = {
                    "SYSTEM_NUMBER": system_number,
                    "IVRS_NUMBER": ivrs_number,
                    "CUSTOMER": customer,
                    "STUDY": study,
                    "SITE#": site,
                    "SHIP_DATE" : ship_date,
                    "SHIP_TIME_FROM" : ship_time_from,
                    "DELIVERY_DATE" : delivery_date,
                    "TYPE_OF_MATERIAL" : type_of_material,
                    "TEMPERATURE" : temperature,
                    "AMOUNT_OF_BOXES_TO_SEND" : amount_of_boxes_to_send,
                    "HAS_RETURN": has_return,
                    "RETURN_TO_CARRIER_DEPOT": return_to_carrier_depot,
                    "TYPE_OF_RETURN": type_of_return,
                    "AMOUNT_OF_BOXES_TO_RETURN": amount_of_boxes_to_return,
                    "TRACKING_NUMBER": tracking_number,
                    "RETURN_TRACKING_NUMBER": return_tracking_number
                }
    
        return order
    
    def __get_orders_for_testing__(self):
        ordersDataFrame = pd.DataFrame(columns=["SYSTEM_NUMBER", "IVRS_NUMBER", "STUDY", "SITE#", 
                                                "SHIP_DATE", "SHIP_TIME_FROM", "DELIVERY_DATE", "TYPE_OF_MATERIAL", 
                                                "TEMPERATURE", "AMOUNT_OF_BOXES_TO_SEND", "HAS_RETURN", 
                                                "RETURN_TO_CARRIER_DEPOT", "TYPE_OF_RETURN", 
                                                "AMOUNT_OF_BOXES_TO_RETURN", "TRACKING_NUMBER", "RETURN_TRACKING_NUMBER"])
        orders = []
            
        orderWithOutSystemNumber = self.__get_order_template__(system_number="",
                                                        ivrs_number="ORDER WITHOUT SYSTEM NUMBER")
        orders.append(orderWithOutSystemNumber)

        orderWithOutStudy = self.__get_order_template__(study="",
                                                        ivrs_number="ORDER WITHOUT STUDY")
        orders.append(orderWithOutStudy)
        
        orderWithOutCustomer = self.__get_order_template__(customer="",
                                                        ivrs_number="ORDER WITHOUT CUSTOMER")
        orders.append(orderWithOutCustomer)

        orderWithOutSite = self.__get_order_template__(site="",
                                                        ivrs_number="ORDER WITHOUT SITE")
        orders.append(orderWithOutSite)
        
        orderWithOutShipTime = self.__get_order_template__(ship_time_from="",
                                                        ivrs_number="ORDER WITHOUT SHIP TIME")
        orders.append(orderWithOutShipTime)

        orderWithOutDeliveryDate = self.__get_order_template__(delivery_date="", 
                                                        ivrs_number="ORDER WITHOUT DELIVERY DATE")
        orders.append(orderWithOutDeliveryDate)

        orderWithOutMaterial = self.__get_order_template__(type_of_material="",
                                                        ivrs_number="ORDER WITHOUT MATERIAL")
        
        orders.append(orderWithOutMaterial)

        """orderWithOutDeliveryTimeFrom = self.__get_order_template__(delivery_time_from="",
                                                        ivrs_number="ORDER WITHOUT DELIVERY TIME FROM")
        orders.append(orderWithOutDeliveryTimeFrom)

        orderWithOutDeliveryTimeTo = self.__get_order_template__(delivery_time_to="",
                                                        ivrs_number="ORDER WITHOUT DELIVERY TIME TO")
        orders.append(orderWithOutDeliveryTimeTo)"""

        orderWithOutTemperature = self.__get_order_template__(temperature="",
                                                        ivrs_number="ORDER WITHOUT TEMPERATURE")
        orders.append(orderWithOutTemperature)

        orderWithOutBoxes = self.__get_order_template__(amount_of_boxes_to_send=0,
                                                        ivrs_number="ORDER WITHOUT BOXES")
        orders.append(orderWithOutBoxes)

        orderWithOutReturn = self.__get_order_template__(has_return="",
                                                        ivrs_number="ORDER WITHOUT RETURN")
        orders.append(orderWithOutReturn)

        orderWithOutReturnToCarrierDepot = self.__get_order_template__(return_to_carrier_depot="",
                                                        ivrs_number="ORDER WITHOUT RETURN TO CARRIER DEPOT")
        orders.append(orderWithOutReturnToCarrierDepot)

        orderWithOutReturnType = self.__get_order_template__(type_of_return="",
                                                        ivrs_number="ORDER WITHOUT RETURN TYPE - MAY HAVE NO ERROR")
        orders.append(orderWithOutReturnType)

        
        
        orderWithDeliveryDateEarlierThanShipDate = self.__get_order_template__(delivery_date=(RELATIVE_DATE - dt.timedelta(days=1)),
                                                        ivrs_number="ORDER WITH DELIVERY DATE EARLIER THAN SHIP DATE")
        orders.append(orderWithDeliveryDateEarlierThanShipDate)

        """orderWithDeliveryTimeToEarlierThanDeliveryTimeFrom = self.__get_order_template__(delivery_time_to="08:00:00",
                                                                                        delivery_time_from="10:00:00",
                                                        ivrs_number="ORDER WITH DELIVERY TIME TO EARLIER THAN DELIVERY TIME FROM")
        orders.append(orderWithDeliveryTimeToEarlierThanDeliveryTimeFrom)"""

        orderWithAnotherTypeOfMaterial = self.__get_order_template__(type_of_material="TEST",
                                                        ivrs_number="ORDER WITH ANOTHER TYPE OF MATERIAL")
        orders.append(orderWithAnotherTypeOfMaterial)

        orderWithAnotherTemperature = self.__get_order_template__(temperature="TEST",
                                                        ivrs_number="ORDER WITH ANOTHER TEMPERATURE")
        orders.append(orderWithAnotherTemperature)

        """orderWithAmountOfBoxesToSendNotInteger = self.__get_order_template__(amount_of_boxes_to_send="TEST",
                                                        ivrs_number="ORDER WITH AMOUNT OF BOXES TO SEND NOT INTEGER")
        orders.append(orderWithAmountOfBoxesToSendNotInteger)"""

        orderWithAmountOfBoxesToSendNegative = self.__get_order_template__(amount_of_boxes_to_send=-1,
                                                        ivrs_number="ORDER WITH AMOUNT OF BOXES TO SEND NEGATIVE")
        orders.append(orderWithAmountOfBoxesToSendNegative)

        """orderWithAmountOfBoxesToReturnNotInteger = self.__get_order_template__(amount_of_boxes_to_return="TEST",
                                                        ivrs_number="ORDER WITH AMOUNT OF BOXES TO RETURN NOT INTEGER")
        orders.append(orderWithAmountOfBoxesToReturnNotInteger)"""

        orderWithAmountOfBoxesToReturnNegative = self.__get_order_template__(amount_of_boxes_to_return=-1,
                                                        ivrs_number="ORDER WITH AMOUNT OF BOXES TO RETURN NEGATIVE")
        orders.append(orderWithAmountOfBoxesToReturnNegative)

        orderWithAmountOfBoxesToReturnGreaterThanAmountOfBoxesToSend = self.__get_order_template__(amount_of_boxes_to_send=1,
                                                                                                amount_of_boxes_to_return=2,
                                                        ivrs_number="ORDER WITH AMOUNT OF BOXES TO RETURN GREATER THAN AMOUNT OF BOXES TO SEND")
        orders.append(orderWithAmountOfBoxesToReturnGreaterThanAmountOfBoxesToSend)

        orderWithAnotherTypeOfReturn = self.__get_order_template__(type_of_return="TEST",
                                                        ivrs_number="ORDER WITH ANOTHER TYPE OF RETURN  - MAY HAVE NO ERROR")
        orders.append(orderWithAnotherTypeOfReturn)

        validOrderNotDone = self.__get_order_template__(ivrs_number="VALID NOT PROCESSED ORDER")
        orders.append(validOrderNotDone)

        processedOrderWithTrackingNumber = self.__get_order_template__(ivrs_number="VALID PROCESSED ORDER WITH TRACKING NUMBER",
                                                        tracking_number="TRACKING_NUMBER__01")
        orders.append(processedOrderWithTrackingNumber)

        processedOrderWithTrackingNumberAndReturnNumber = self.__get_order_template__(ivrs_number="VALID PROCESSED ORDER WITH TRACKING NUMBER AND RETURN NUMBER",
                                                        tracking_number="TRACKING_NUMBER__01",
                                                        return_tracking_number="RETURN_TRACKING_NUMBER__01")
        orders.append(processedOrderWithTrackingNumberAndReturnNumber)

        orderWithOnlyReturnNumber = self.__get_order_template__(ivrs_number="ORDER WITH ONLY RETURN NUMBER",
                                                        return_tracking_number="RETURN_TRACKING_NUMBER__01")
        orders.append(orderWithOnlyReturnNumber)

        ordersDataFrame = pd.DataFrame(orders)

        return ordersDataFrame