import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC

from carriersWebpage.carrierWebPage import CarrierWebpage
from logClass.log import Log

class CarrierWebpageForTesting(CarrierWebpage):
    def __init__(self, folder_path_to_download: str, log: Log):
        """
        Class constructor for Transportes Ambientales

        Args:
            driver (webdriver): selenium driver
        """
        super().__init__(log)
        self.folder_path_to_download = folder_path_to_download

    def build_driver(self) -> None:
        pass

    def quit_driver(self) -> None:
        pass

    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        """
        Checks if user can log in webpage

        Args:
            driver (webdriver): selenium driver
            username (str): username
            password (str): password
        """
        self.complete_login_form(username, password)

        try:
            return username == "username" and password == "password"
        
        except Exception as e:
            raise Exception(e)
        
        finally:
            return False
    
    def complete_login_form(self, username: str, password: str) -> None:
        """
        self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/input[1]")))
        Completes login form

        Args:
            driver (webdriver): selenium driver
            username (str): username
            password (str): password
        """
        pass

    def complete_shipping_order_form(self, carrier_id: str, reference: str, 
                                ship_date: str, ship_time_from: str, ship_time_to: str, 
                                delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                type_of_material: str, temperature: str,
                                contacts: str, amount_of_boxes: int) -> str:
        tracking_number = ""

        try:
            if (ship_date == ""):
                raise Exception("Ship date is empty")
            
            if (delivery_date == ""):
                raise Exception("Delivery date is empty")
            
            if (type_of_material == ""):
                raise Exception("Type of material is empty")
            
            if (temperature == ""):
                raise Exception("Temperature is empty")
            
            time.sleep(0.5) # Simulate time to complete form
            
            tracking_number = f"sucessful tracking number for order {reference}"
            
        except Exception as e:
            raise Exception(e)

        finally:
            return tracking_number

    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                            delivery_date: str, return_time_from: str,
                                            return_time_to: str, type_of_return: str,
                                            contacts: str, amount_of_boxes_to_return: int,
                                            return_to_carrier_depot: bool, tracking_number: str) -> str:
        return_tracking_number = ""

        try:
            
            if (delivery_date == ""):
                raise Exception("Delivery date is empty")
            
            if (type_of_return == ""):
                raise Exception("Type of return is empty")
            
            time.sleep(0.5) # Simulate time to complete form
            
            return_tracking_number = f"sucessful return tracking number for order {reference_return}"

        except Exception as e:
            raise Exception(e)
            
        finally:
            return return_tracking_number
    
    def printWayBillDocument(self, tracking_number: str, amount_of_copies: int) -> None:
        pass

    def printLabelDocument(self, tracking_number: str) -> None:
        pass

    def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int) -> None:
        pass