import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC

from carriersWebpage.carrierWebPage import CarrierWebpage
from logClass.log import Log

class CARRIER_NAME(CarrierWebpage):
    def __init__(self, folder_path_to_download: str, log: Log):
        """
        Class constructor for Transportes Ambientales

        Args:
            driver (webdriver): selenium driver
        """
        super().__init__(log)
        self.folder_path_to_download = folder_path_to_download

    def build_driver(self) -> None:
        self.driver, self.wait = self.__build_driver__(self.folder_path_to_download)

    def quit_driver(self) -> None:
        self.__quit_driver__(self.driver)

    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        """
        Checks if user can log in webpage

        Args:
            driver (webdriver): selenium driver
            username (str): username
            password (str): password
        """
        self.driver.get("www.google.com")
        
        self.complete_login_form(username, password)

        try:
            # here you should check if the login was successful 
            raise NotImplementedError("Method not implemented")

            return True
        except Exception as e:
            raise Exception(e)
        
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
        # here you should complete the login form
        raise NotImplementedError("Method not implemented")

    def complete_shipping_order_form(self, carrier_id: str, reference: str, 
                                ship_date: str, ship_time_from: str, ship_time_to: str, 
                                delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                type_of_material: str, temperature: str,
                                contacts: str, amount_of_boxes: int) -> str:
        tracking_number = ""
        url = "www.google.com"

        try:
            self.driver.get(url)

            # Here you should complete the shipping order form
            raise NotImplementedError("Method not implemented")
            
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
        url_return = "www.google.com"

        try:
            
            # Here you should complete the shipping order return form
            raise NotImplementedError("Method not implemented")

        except Exception as e:
            raise Exception(e)
            
        finally:
            return return_tracking_number
    
    def printWayBillDocument(self, tracking_number: str, amount_of_copies: int) -> None:
        url_guias = "www.google.com"

        # Here you should print way bill document
        raise NotImplementedError("Method not implemented")
    
        self.__print_webpage__(url_guias)

    def printLabelDocument(self, tracking_number: str) -> None:
        url_rotulo = "www.google.com"

        # Here you should print label document
        raise NotImplementedError("Method not implemented")
    
        self.__print_webpage__(url_rotulo)

    def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int) -> None:
        url_guias_return = "www.google.com"

        # Here you should print return way bill document
        raise NotImplementedError("Method not implemented")
    
        self.__print_webpage__(url_guias_return)