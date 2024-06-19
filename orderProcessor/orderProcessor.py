import pandas as pd
import datetime as dt
import win32com.client as win32
import time
import os
from typing import Tuple
import queue

from teams.team import Team
from logClass.log import Log
from utils.utils import renameAllReturnFiles, export_to_excel

class OrderProcessor:
    def __init__(self, folder_path_to_download : str, selectedTeam: Team, queue: queue.Queue):
        """
        Class constructor

        Args:
            self.driver (webdriver): selenium self.driver

        Attributes:
            self.driver (webdriver): selenium self.driver
            wait (WebDriverWait): selenium wait
        """
        self.folder_path_to_download = folder_path_to_download
        self.selectedTeam = selectedTeam
        self.queue = queue
    
    def send_email_to_medical_center(self, study: str, site: str, ivrs_number: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str, amount_of_boxes: int,
                                    hasReturn: bool, type_of_return: str, amount_of_boxes_to_return: int,
                                    tracking_number: str, return_tracking_number: str,
                                    contacts: str, medical_center_emails: str, customer_emails: str, CRAs_emails: str, team_emails: str) -> None:
        """
        Sends an email to the medical center

        Args:
            study (str): study
            site (str): site
            ivrs_number (str): ivrs number
            delivery_date (str): delivery date
            delivery_time_from (str): delivery time (from)
            delivery_time_to (str): delivery time (to)
            type_of_material (str): type of material
            temperature (str): temperature
            amount_of_boxes (int): number of boxes
            hasReturn (bool): if True, creates a return order
            type_of_return (str): type of return
            amount_of_boxes_to_return (int): number of boxes to return
            tracking_number (str): tracking number
            return_tracking_number (str): return tracking number
            contacts (str): contacts
        """
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.Subject = f"Orden de envío || Estudio: {study} || Site#: {site} || Número de IVRS: {ivrs_number}"
            mail.To = medical_center_emails
            
            customer_emails = customer_emails + "; " if customer_emails != "" else ""
            CRAs_emails = CRAs_emails + "; " if CRAs_emails != "" else ""
            team_emails = team_emails if team_emails != "" else ""
            mail.CC = customer_emails + CRAs_emails + team_emails

            emailSource = self.__getEmailSourceFromTxtFile__("email.txt")
            emailSource = self.__replaceEmailValues__(emailSource, study, site, ivrs_number, 
                                                    delivery_date, delivery_time_from, delivery_time_to, 
                                                    type_of_material, temperature, amount_of_boxes, 
                                                    hasReturn, type_of_return, amount_of_boxes_to_return, 
                                                    tracking_number, return_tracking_number, contacts, team_emails)
            mail.HTMLBody = emailSource

            # Enviar el correo
            mail.Send()
            
        except Exception as e:
            Log().add_log(f"Error sending email to medical center: {e}")

    def processOrdersAndContactsTable(self, ordersDataFrame:pd.DataFrame) -> pd.DataFrame:
        """
        Process all orders in the table
        
        Variables:
            wait (WebDriverWait): selenium wait
            team (str): team to process
            df_path (str): excel file path
            sheet (str): excel sheet name
            df (DataFrame): orders table

        Args:
            ordersDataFrame (DataFrame): orders table
            user (str): user of carrier website
            password (str): password of carrier website

        Returns:
            DataFrame: orders table with tracking numbers

        Preconditions:
            The orders table must be standardized as DataRecollector() class returns

            Website must be built and logged in
        """
        try:
            self.__process_all_shipping_orders__(ordersDataFrame)

            time.sleep(2) # Wait for the download to finish

            self.__renameAllReturnFiles__(ordersDataFrame)

            self.__export_to_excel__(ordersDataFrame)

            self.queue.put("processOrdersAndContactsTable finished")

        except Exception as e:
            Log().add_log(f"Error generating shipping report: {e}")

        finally:
            try:
                self.selectedTeam.quit_driver()
            except Exception as e:
                Log().add_log(f"Error quitting the webpage: {e}")
            finally:
                return ordersDataFrame

    # Private methods
    def __export_to_excel__(self, dataFrame: pd.DataFrame):
        export_to_excel(dataFrame, self.folder_path_to_download)

    def __get_shipping_tracking_number__(self, carrier_id: int, system_number: str, ivrs_number: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        """
        Gets the shipping tracking number

        Args:
            carrier_id (int): Site ID on carrier website
            system_number (int): order system number
            ivrs_number (str): order ivrs number
            ship_date (str): ship date
            ship_time_from (str): ship time (from)
            ship_time_to (str): ship time (to)
            delivery_date (str): delivery date
            delivery_time_from (str): delivery time (from)
            delivery_time_to (str): delivery time (to)
            type_of_material (str): type of material
            temperature (str): temperature
            contacts (str): contacts
            amount_of_boxes (int): number of boxes

        Returns:
            str: tracking number
        """

        reference = f"{system_number} {ivrs_number}"[:50]
        delivery_date = dt.datetime.strptime(delivery_date, '%d/%m/%Y').strftime('%d/%m/%Y')

        tracking_number = self.selectedTeam.complete_shipping_order_form(
            carrier_id, reference, 
            ship_date, ship_time_from, ship_time_to, 
            delivery_date, delivery_time_from, delivery_time_to, 
            type_of_material, temperature, 
            contacts, amount_of_boxes
        )
        
        return tracking_number
    
    def __get_return_tracking_number__(self, carrier_id: int, system_number: str, ivrs_number: str,
                                    delivery_date: str,  tracking_number: str, hasReturn: bool,
                                    return_delivery_date: str, return_delivery_hour_from: str, return_delivery_hour_to: str,
                                    return_to_TA: bool, type_of_return: str, amount_of_boxes_to_return: int) -> str:
            """
            Gets the return tracking number
    
            Args:
                carrier_id (int): Site ID on carrier website
                system_number (int): order system number
                ivrs_number (str): order ivrs number
                delivery_date (str): delivery date
                tracking_number (str): tracking number
                hasReturn (bool): if True, creates a return order
                return_to_TA (bool): if True, the return order is sent to carrier depot
                type_of_return (str): type of return
                amount_of_boxes_to_return (int): number of boxes to return
    
            Returns:
                str: tracking number
            """
            if not hasReturn or tracking_number == "":
                return  ""

            reference_return = f"{system_number} {ivrs_number} RET {tracking_number}"[:50]
            
            return_tracking_number = self.selectedTeam.complete_shipping_order_return_form(
                carrier_id, reference_return, 
                return_delivery_date, return_delivery_hour_from, return_delivery_hour_to, 
                type_of_return, "", amount_of_boxes_to_return, 
                return_to_TA, tracking_number
            )
        
            return return_tracking_number
    
    def __renameAllReturnFiles__(self, ordersAndContactsDataframe: pd.DataFrame) -> None:
        renameAllReturnFiles(ordersAndContactsDataframe, self.folder_path_to_download)

    def __process_all_shipping_orders__(self, ordersAndContactsDataframe: pd.DataFrame) -> None:
        """
        Process all orders in the table

        Args:
            ordersAndContactsDataframe (DataFrame): orders table
        """
        email_must_be_sent = self.selectedTeam.get_data_path(["team_send_email_to_medical_centers"])

        for index, row in ordersAndContactsDataframe.iterrows():
            if row['TRACKING_NUMBER'] != "":
                # Skip orders that have already been processed
                continue

            if row['HAS_AN_ERROR'] != "No error":
                # Skip orders with errors
                continue
            
            tracking_number, return_tracking_number = self.__get_tracking_numbers_from_carrier__(
                row["CARRIER_ID"], row["SYSTEM_NUMBER"], row["IVRS_NUMBER"],
                row["SHIP_DATE"], row["SHIP_TIME_FROM"], row["SHIP_TIME_TO"],
                row["DELIVERY_DATE"], row["DELIVERY_TIME_FROM"], row["DELIVERY_TIME_TO"],
                row["TYPE_OF_MATERIAL"], row["TEMPERATURE"],
                row["CONTACTS"], row["AMOUNT_OF_BOXES_TO_SEND"],
                row["HAS_RETURN"], row["RETURN_DELIVERY_DATE"], row["RETURN_DELIVERY_HOUR_FROM"], row["RETURN_DELIVERY_HOUR_TO"], 
                row["RETURN_TO_CARRIER_DEPOT"], row["TYPE_OF_RETURN"], row["AMOUNT_OF_BOXES_TO_RETURN"]
            )

            if tracking_number == "":
                # Skip orders with errors (no tracking number generated)
                continue

            ordersAndContactsDataframe.loc[index, "TRACKING_NUMBER"] = tracking_number
            ordersAndContactsDataframe.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number
            
            self.__printOrderDocuments__(tracking_number, return_tracking_number, row["PRINT_RETURN_DOCUMENT"])

            if email_must_be_sent:
                try:
                    self.send_email_to_medical_center( row["STUDY"], row["SITE#"], row["IVRS_NUMBER"], 
                    row["DELIVERY_DATE"], row["DELIVERY_TIME_FROM"], row["DELIVERY_TIME_TO"], 
                    row["TYPE_OF_MATERIAL"], row["TEMPERATURE"], row["AMOUNT_OF_BOXES_TO_SEND"],
                    row["HAS_RETURN"], row["TYPE_OF_RETURN"], row["AMOUNT_OF_BOXES_TO_RETURN"],
                    tracking_number, return_tracking_number,
                    row["CONTACTS"], row["MEDICAL_CENTER_EMAILS"], row["CUSTOMER_EMAIL"], row["CRA_EMAILS"], row["TEAM_EMAILS"])
                except Exception as e:
                    Log().add_log(f"Error sending email to medical center: {e}")
                    Log().add_log(f"Order: {row['SYSTEM_NUMBER']} {row['IVRS_NUMBER']}")
            
            self.__updateTreeviewLine__(index, tracking_number, return_tracking_number)

    def __printOrderDocuments__(self, tracking_number: str, return_tracking_number: str, print_return_document: bool) -> None:
        """
        Prints the order documents

        Args:
            tracking_number (str): tracking number
            return_tracking_number (str): return tracking number
            print_return_document (bool): if True, prints the return document
        """
        if tracking_number != "":
            self.selectedTeam.printWayBillDocument(tracking_number, 4)
            self.selectedTeam.printLabelDocument(tracking_number)
        else:
            return
        
        if print_return_document and return_tracking_number != "" and return_tracking_number != "Error":
            self.selectedTeam.printReturnWayBillDocument(return_tracking_number, 1)

    def __get_tracking_numbers_from_carrier__(self, carrier_id: int, system_number: str, ivrs_number: str, 
        ship_date: str, ship_time_from: str, ship_time_to: str, 
        delivery_date: str, delivery_time_from: str, delivery_time_to: str,
        type_of_material: str, temperature: str, 
        contacts: str, amount_of_boxes: int, 
        hasReturn: bool, return_delivery_date: str, return_delivery_hour_from: str, return_delivery_hour_to: str, 
        return_to_TA: bool, type_of_return: str, amount_of_boxes_to_return: int) -> Tuple[str, str]:
        """
        - Process an order by completing the carrier form
        - Creates a return order if necessary
        - Prints the order documents

        Args:
            carrier_id (int): Site ID on carrier website
            system_number (int): order system number
            ivrs_number (str): order ivrs number
            ship_date (str): ship date
            ship_time_from (str): ship time (from)
            ship_time_to (str): ship time (to)
            delivery_date (str): delivery date
            delivery_time_from (str): delivery time (from)
            delivery_time_to (str): delivery time (to)
            type_of_material (str): type of material
            temperature (str): temperature
            contacts (str): contacts
            amount_of_boxes (int): number of boxes
            hasReturn (bool): if True, creates a return order
            return_to_TA (bool): if True, the return order is sent to carrier depot
            amount_of_boxes_to_return (int): number of boxes to return

        Returns:
            str: tracking number
            str: return tracking number
        """
        
        tracking_number, return_tracking_number = "", ""

        try:
            tracking_number = self.__get_shipping_tracking_number__(
                carrier_id, system_number, ivrs_number,
                ship_date, ship_time_from, ship_time_to,
                delivery_date, delivery_time_from, delivery_time_to,
                type_of_material, temperature,
                contacts, amount_of_boxes
            )

            if tracking_number == "":
                return "", ""
            
            return_tracking_number = self.__get_return_tracking_number__(
                carrier_id, system_number, ivrs_number,
                delivery_date, tracking_number, hasReturn, 
                return_delivery_date, return_delivery_hour_from, return_delivery_hour_to,
                return_to_TA, type_of_return, amount_of_boxes_to_return
            )

        except Exception as e:
            Log().add_log(f"Error processing order: {e}")
            Log().add_log(f"Order: {system_number} {ivrs_number}")

        finally:
            return tracking_number, return_tracking_number

    def __updateTreeviewLine__(self, index: int, tracking_number: str, return_tracking_number: str) -> None:
        """
        Updates a line in the treeview

        Args:
            index (int): row index
        """
        try:
            self.queue.put(f"update_tag_color_of_a_treeview_line({index}, {tracking_number}, {return_tracking_number})")
        except Exception as e:
            Log().add_log(f"Error updating treeview line: {e}")

    def __getEmailSourceFromTxtFile__(self, file) -> str:
            with open(file, 'r') as file:
                return file.read()
            
    def __replaceEmailValues__(self, emailSource: str, study: str, site: str, ivrs_number: str,
        delivery_date: str, delivery_time_from: str, delivery_time_to: str,
        type_of_material: str, temperature: str, amount_of_boxes: int,
        hasReturn: bool, type_of_return: str, amount_of_boxes_to_return: int,
        tracking_number: str, return_tracking_number: str,
        contacts: str, team_emails: str) -> str:
        
        emailSource = emailSource.replace("|VAR_STUDY|", study)
        emailSource = emailSource.replace("|VAR_SITE#|", site)
        emailSource = emailSource.replace("|VAR_IVRS_NUMBER|", ivrs_number)
        emailSource = emailSource.replace("|VAR_DELIVERY_DATE|", delivery_date)
        emailSource = emailSource.replace("|VAR_DELIVERY_TIME|", delivery_time_from + " to " + delivery_time_to)
        emailSource = emailSource.replace("|VAR_TYPE_OF_MATERIAL|", type_of_material)
        emailSource = emailSource.replace("|VAR_TEMPERATURE|", temperature)
        emailSource = emailSource.replace("|VAR_AMOUNT_OF_BOXES|", str(amount_of_boxes))
        emailSource = emailSource.replace("|VAR_TRACKING_NUMBER|", tracking_number)
        emailSource = emailSource.replace("|VAR_CONTACTS|", contacts)
        emailSource = emailSource.replace("|VAR_TEAM_EMAIL|", team_emails)
        emailSource = emailSource.replace("|VAR_TMO_LOGO|", os.getcwd() + "\\media\\TMO_logo_email.jpg")

        if hasReturn:
            emailSource = emailSource.replace("|VAR_TYPE_OF_RETURN|", type_of_return)
            emailSource = emailSource.replace("|VAR_AMOUNT_OF_BOXES_TO_RETURN|", str(amount_of_boxes_to_return))
            emailSource = emailSource.replace("|VAR_RETURN_TRACKING_NUMBER|", return_tracking_number)
        else:
            emailSource = emailSource.replace("|VAR_TYPE_OF_RETURN|", "NA")
            emailSource = emailSource.replace("|VAR_AMOUNT_OF_BOXES_TO_RETURN|", "0")
            emailSource = emailSource.replace("|VAR_RETURN_TRACKING_NUMBER|", "NA")

        return emailSource