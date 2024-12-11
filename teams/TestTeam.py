import pandas as pd
import datetime as dt
from typing import Tuple

from .team import Team
from logClass.log import Log

next_weekday = dt.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) + dt.timedelta(days= 1 if dt.datetime.today().weekday() < 4 else 7 - dt.datetime.today().weekday())
RELATIVE_DATE = next_weekday

class TestTeam(Team):
    def __init__(self, folder_path_to_download: str, log : Log):
        super().__init__(log)
        self.carrierWebpage = self.__build_carrier_Webpage__("Transportes Ambientales HTTP", folder_path_to_download)

    def get_team_name(self) -> str:
        return "Team_for_testings"
    
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

        ordersDataFrame["HAS_RETURN"] = (ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & (ordersDataFrame["TEMPERATURE"] != "Ambient")
        #ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] = ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"].apply(lambda x: x if x > 0 else 0)
        ordersDataFrame["TYPE_OF_RETURN"] = ordersDataFrame["HAS_RETURN"].apply(lambda x: "CREDO" if x else "NA")

        return ordersDataFrame
    
    def send_email_to_team_with_orders(self, folder_path_with_orders_files: str, date: str,
                totalAmountOfOrders: int, amountOfOrdersProcessed: int, amountOfOrdersReadyToBeProcessed: int) -> None:
        self.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "inaki.costa@thermofisher",
                    totalAmountOfOrders, amountOfOrdersProcessed, amountOfOrdersReadyToBeProcessed)

    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            ordersDataFrame = self.__get_orders_for_testing__()
            
            return ordersDataFrame
        except Exception as e:
            raise e
    
    def readContactsExcel(self, path_from_get_data: str, contacts_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            contacts_df = self.__get_contacts_for_testing__()

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
    
    def print_wayBill_document(self, tracking_number: str, amount_of_copies: int):
        self.__printWayBillDocument__(self.carrierWebpage, tracking_number, amount_of_copies)

    def print_label_document(self, tracking_number: str):
        self.__printLabelDocument__(self.carrierWebpage, tracking_number)

    def print_return_wayBill_document(self, return_tracking_number: str, amount_of_copies: int):
        self.__printReturnWayBillDocument__(self.carrierWebpage, return_tracking_number, amount_of_copies)

    def get_column_rename_type_config_for_not_working_days_table(self) -> Tuple[dict, dict]:
        columns_names = {}
        columns_types = {"DATE": dt.datetime}
        return columns_names, columns_types
    
    def readNotWorkingDaysExcel(self, path_from_get_data: str, not_working_days_sheet: str, columns_types: dict) -> pd.DataFrame:
        try:
            notWorkingDaysData = {"DATE": [RELATIVE_DATE + dt.timedelta(days=4)]}
            notWorkingDaysDataFrame = pd.DataFrame(notWorkingDaysData)
        except Exception as e:
            raise e
        return notWorkingDaysDataFrame
    
    def __get_order_template__(self, 
                        system_number = "1",
                        ivrs_number = "IVRS_NUMBER__01",
                        customer = "CUSTOMER__01",
                        study = "TEST",
                        site = "Only Medicines",
                        ship_date = RELATIVE_DATE,
                        ship_time_from = "19:00:00",
                        delivery_date = (RELATIVE_DATE + dt.timedelta(days=1)),
                        type_of_material = "Medicine",
                        temperature = "Refrigerated",
                        amount_of_boxes_to_send = 1,
                        has_return = True,
                        return_to_carrier_depot = True,
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

        orderWithAnotherTypeOfMaterial = self.__get_order_template__(type_of_material="TEST",
                                                        ivrs_number="ORDER WITH ANOTHER TYPE OF MATERIAL")
        orders.append(orderWithAnotherTypeOfMaterial)

        orderWithAnotherTemperature = self.__get_order_template__(temperature="TEST",
                                                        ivrs_number="ORDER WITH ANOTHER TEMPERATURE")
        orders.append(orderWithAnotherTemperature)

        orderWithAmountOfBoxesToSendNegative = self.__get_order_template__(amount_of_boxes_to_send=-1,
                                                        ivrs_number="ORDER WITH AMOUNT OF BOXES TO SEND NEGATIVE")
        orders.append(orderWithAmountOfBoxesToSendNegative)

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

        orderWithOnlyMedicines = self.__get_order_template__(ivrs_number="ORDER WITH ONLY MEDICINES",
                                                        site="Only Medicines",
                                                        type_of_material="Medicine")
        orders.append(orderWithOnlyMedicines)

        orderWithOnlyAncillariesType1 = self.__get_order_template__(ivrs_number="ORDER WITH ONLY ANCILLARIES TYPE 1",
                                                        site="Only Ancillaries Type 1",
                                                        type_of_material="Ancillary Type 1")
        
        orders.append(orderWithOnlyAncillariesType1)
        
        orderWithOnlyAncillariesType2 = self.__get_order_template__(ivrs_number="ORDER WITH ONLY ANCILLARIES TYPE 2",
                                                        site="Only Ancillaries Type 2",
                                                        type_of_material="Ancillary Type 2")
        orders.append(orderWithOnlyAncillariesType2)

        orderWithOnlyEquipments = self.__get_order_template__(ivrs_number="ORDER WITH ONLY EQUIPMENTS",
                                                        site="Only Equipments",
                                                        type_of_material="Equipment")
        orders.append(orderWithOnlyEquipments)

        orderWithAncillariesType1AndGeneralContactsForAncillaries = self.__get_order_template__(ivrs_number="ORDER WITH ANCILLARIES TYPE AND GENERAL CONTACTS FOR ANCILLARIES",
                                                        site="Receive Ancillaries",
                                                        type_of_material="Ancillary Type 1")
        orders.append(orderWithAncillariesType1AndGeneralContactsForAncillaries)

        orderWithAncillariesType2AndGeneralContactsForAncillaries = self.__get_order_template__(ivrs_number="ORDER WITH ANCILLARIES TYPE AND GENERAL CONTACTS FOR ANCILLARIES",
                                                        site="Receive Ancillaries",
                                                        type_of_material="Ancillary Type 2")
        orders.append(orderWithAncillariesType2AndGeneralContactsForAncillaries)

        orderWithAllMaterialsAndGeneralContactsForAll = self.__get_order_template__(ivrs_number="ORDER WITH ALL MATERIALS AND GENERAL CONTACTS FOR ALL",
                                                        site="Receive All",
                                                        type_of_material="Medicine")
        orders.append(orderWithAllMaterialsAndGeneralContactsForAll)

        orderWithOutDeliveryTimeFrom = self.__get_order_template__(site = "Without Delivery Time From",
                                                        ivrs_number="ORDER WITHOUT DELIVERY TIME FROM")
        orders.append(orderWithOutDeliveryTimeFrom)

        orderWithOutDeliveryTimeTo = self.__get_order_template__(site = "Without Delivery Time To",
                                                        ivrs_number="ORDER WITHOUT DELIVERY TIME TO")
        orders.append(orderWithOutDeliveryTimeTo)

        orderWithDeliveryTimeToEarlierThanDeliveryTimeFrom = self.__get_order_template__(site = "With Delivery Time To Earlier Than Delivery Time From",
                                                        ivrs_number="ORDER WITH DELIVERY TIME TO EARLIER THAN DELIVERY TIME FROM")
        orders.append(orderWithDeliveryTimeToEarlierThanDeliveryTimeFrom)

        orderWithNotWorkingDay = self.__get_order_template__(ivrs_number="ORDER WITH NOT WORKING DAY, SHOULD RETURN ON 5TH DAY or more",
                                                        delivery_date=(RELATIVE_DATE + dt.timedelta(days=2)))
        orders.append(orderWithNotWorkingDay)
        
        validOrderAmbientWithNoReturn = self.__get_order_template__(ivrs_number="VALID ORDER AMBIENT WITH NO RETURN",
                                                        temperature="Ambient",
                                                        has_return=False,
                                                        return_to_carrier_depot=False,
                                                        type_of_return="NA",
                                                        amount_of_boxes_to_return=0)
        orders.append(validOrderAmbientWithNoReturn)

        validOrderControlledAmbientWithReturnToCarrierDepot = self.__get_order_template__(ivrs_number="VALID ORDER CONTROLLED AMBIENT WITH RETURN TO CARRIER DEPOT",
                                                        temperature="Controlled Ambient")
        orders.append(validOrderControlledAmbientWithReturnToCarrierDepot)

        validOrderControlledAmbientWithReturnToFCS = self.__get_order_template__(ivrs_number="VALID ORDER CONTROLLED AMBIENT WITH RETURN TO FCS",
                                                        temperature="Controlled Ambient",
                                                        return_to_carrier_depot=False)
        orders.append(validOrderControlledAmbientWithReturnToFCS)

        validOrderRefrigeratedWith2Boxes = self.__get_order_template__(ivrs_number="VALID ORDER REFRIGERATED WITH 2 BOXES",
                                                        temperature="Refrigerated",
                                                        amount_of_boxes_to_send=2)
        orders.append(validOrderRefrigeratedWith2Boxes)

        validOrderFrozenWithReturn = self.__get_order_template__(ivrs_number="VALID ORDER FROZEN WITH RETURN",
                                                        temperature="Frozen")
        orders.append(validOrderFrozenWithReturn)

        validOrderWithReturnOfDatalloger = self.__get_order_template__(ivrs_number="VALID ORDER WITH RETURN OF DATALOGGER",
                                                        type_of_return="DATALOGGER")
        orders.append(validOrderWithReturnOfDatalloger)

        validOrderWithReturnOfBoxAndDatalogger = self.__get_order_template__(ivrs_number="VALID ORDER WITH RETURN OF BOX AND DATALOGGER",
                                                        type_of_return="CREDO AND DATALOGGER")
        orders.append(validOrderWithReturnOfBoxAndDatalogger)

        validOrderWithAncillariesType = self.__get_order_template__(ivrs_number="VALID ORDER WITH ANCILLARIES",
                                                        site="Receive Ancillaries",
                                                        type_of_material="Ancillary Type 1")
        orders.append(validOrderWithAncillariesType)

        orderWithNotCommonLetters = self.__get_order_template__(ivrs_number="Letra ñ Ñ y tilde á Á é É í Í ó Ó ú Ú ü Ü")
        orders.append(orderWithNotCommonLetters)

        orderWithIncorrectContact = self.__get_order_template__( site="Incorrect Contacts",
                                                        ivrs_number="ORDER WITH INCORRECT CONTACT")

        orders.append(orderWithIncorrectContact)

        ordersDataFrame = pd.DataFrame(orders)

        return ordersDataFrame
    
    def __get_contact_template__(self, 
                        study = "TEST",
                        site = "01",
                        carrier_id = "83",
                        delivery_time_from = "10:00:00",
                        delivery_time_to = "12:00:00",
                        contacts = "Contact 1",
                        can_receive_medicines = "",
                        can_receive_ancillaries_type1 = "",
                        can_receive_ancillaries_type2 = "",
                        can_receive_equipments = "",
                        medical_center_emails = "",
                        customer_email = "inaki.costa@thermofisher.com",
                        cra_emails = "inaki.costa@thermofisher.com",
                        team_emails = "inaki.costa@thermofisher.com"):


        order = {
                    "STUDY": study,
                    "SITE#": site,
                    "CARRIER_ID": carrier_id,
                    "DELIVERY_TIME_FROM" : delivery_time_from,
                    "DELIVERY_TIME_TO" : delivery_time_to,
                    "CONTACTS" : contacts,
                    "CAN_RECEIVE_MEDICINES" : can_receive_medicines,
                    "CAN_RECEIVE_ANCILLARIES_TYPE1" : can_receive_ancillaries_type1,
                    "CAN_RECEIVE_ANCILLARIES_TYPE2" : can_receive_ancillaries_type2,
                    "CAN_RECEIVE_EQUIPMENTS" : can_receive_equipments,
                    "MEDICAL_CENTER_EMAILS": medical_center_emails, 
                    "CUSTOMER_EMAIL": customer_email,
                    "CRA_EMAILS": cra_emails,
                    "TEAM_EMAILS": team_emails
                }
    
        return order
    
    def __get_contacts_for_testing__(self):
        contacts = []

        contactToReceiveOnlyMedicines = self.__get_contact_template__(site="Only Medicines",
                                                contacts="Contact to Receive Only Medicines",
                                                can_receive_medicines="x")
        contacts.append(contactToReceiveOnlyMedicines)

        contactToReceiveOnlyAncillariesType1 = self.__get_contact_template__(site="Only Ancillaries Type 1",
                                                contacts="Contact to Receive Only Ancillaries Type 1",
                                                can_receive_ancillaries_type1="x")
        contacts.append(contactToReceiveOnlyAncillariesType1)

        contactToReceiveOnlyAncillariesType2 = self.__get_contact_template__(site="Only Ancillaries Type 2",
                                                contacts="Contact to Receive Only Ancillaries Type 2",
                                                can_receive_ancillaries_type2="x")
        contacts.append(contactToReceiveOnlyAncillariesType2)

        contactToReceiveOnlyEquipments = self.__get_contact_template__(site="Only Equipments",
                                                contacts="Contact to Receive Only Equipments",
                                                can_receive_equipments="x")
        
        contacts.append(contactToReceiveOnlyEquipments)

        contactToReceiveAll = self.__get_contact_template__(site="Receive All",
                                                contacts="Contact to Receive All",
                                                can_receive_medicines="x",
                                                can_receive_ancillaries_type1="x",
                                                can_receive_ancillaries_type2="x",
                                                can_receive_equipments="x")
        contacts.append(contactToReceiveAll)

        contactToReceiveAncillaries = self.__get_contact_template__(site="Receive Ancillaries",
                                                contacts="Contact to Receive Ancillaries",
                                                can_receive_ancillaries_type1="x",
                                                can_receive_ancillaries_type2="x")
        contacts.append(contactToReceiveAncillaries)

        contactsWithOutDeliveryTimeFrom = self.__get_contact_template__( site = "Without Delivery Time From",
                                                        delivery_time_from="",
                                                        contacts="Contact Without Delivery Time From")
        contacts.append(contactsWithOutDeliveryTimeFrom)

        contactsWithOutDeliveryTimeTo = self.__get_contact_template__( site = "Without Delivery Time To",
                                                        delivery_time_to="",
                                                        contacts="Contact Without Delivery Time To")
        contacts.append(contactsWithOutDeliveryTimeTo)

        contactsWithDeliveryTimeToEarlierThanDeliveryTimeFrom = self.__get_contact_template__( site = "With Delivery Time To Earlier Than Delivery Time From",
                                                                                        delivery_time_from="10:00:00",
                                                                                        delivery_time_to="08:00:00",
                                                        contacts="Contact With Delivery Time To Earlier Than Delivery Time From")
        contacts.append(contactsWithDeliveryTimeToEarlierThanDeliveryTimeFrom)

        incorrectContacts = self.__get_contact_template__( site = "Incorrect Contacts",
                                                        delivery_time_from="10:00:00",
                                                        delivery_time_to="12:00:00",
                                                        contacts="IncorrectContact1; Incorrect/ Contact. 3")
        contacts.append(incorrectContacts)

        contacts_df = pd.DataFrame(contacts)

        return contacts_df