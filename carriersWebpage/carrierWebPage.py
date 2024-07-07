from selenium import webdriver

from selenium.webdriver.support.ui import WebDriverWait
from abc import ABC, abstractmethod

from .Browser import Browser
from logClass.log import Log
from .carrierWebPage_factory import CarrierWebPageFactory

class CarrierWebpage(ABC):
    """
    Abstract class for carriers webpages
    """
    def __init__(self, log: Log):
        """
        Class constructor
        """
        self.log = log
        pass

    # Methods of Super class
    def getCarriersWebpagesNames(self):
        return CarrierWebPageFactory.get_carrier_webpage_names()

    # Methods to be implemented by each sub class
    @abstractmethod
    def build_driver(self) -> None:
        """
        Builds the driver
        """
        pass
    
    @abstractmethod
    def quit_driver(self) -> None:
        """
        Quits the browser
        """
        pass

    @abstractmethod
    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        """
        Logs in webpage

        Args:
            driver (webdriver): selenium driver
            username (str): username
            password (str): password
        """
        try:
            pass
        except Exception as e:
            self.log.add_error_log(f"Error logging in webpage: {e}")

    @abstractmethod
    def complete_shipping_order_form(self, carrier_id: str, reference: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        """
        Completes carrier form

        Args:
            self.driver (webdriver): selenium self.driver
            carrier_id (int): Site ID on carrier website
            reference (str): order reference
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
        try:
            pass
        except Exception as e:
            self.log.add_error_log(f"Error completing shipping order form: {e}")
            self.log.add_error_log(f"Order: {reference}")

    @abstractmethod
    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_carrier_depot: bool, tracking_number: str) -> str:
        """
        Completes carrier return form

        Args:
            self.driver (webdriver): selenium self.driver
            carrier_id (int): Site ID on carrier website
            reference_return (str): return reference
            delivery_date (str): delivery date
            return_time_from (str): return time (from)
            return_time_to (str): return time (to)
            type_of_return (str): type of return
            contacts (str): contacts
            amount_of_boxes_to_return (int): number of boxes to return
            return_to_TA (bool): if True, the return order is sent to carrier depot
            tracking_number (str): tracking number

        Returns:
            str: return tracking number
        """
        try:
            pass
        except Exception as e:
            self.log.add_error_log(f"Error completing shipping order return form: {e}")
            self.log.add_error_log(f"Order: {reference_return}")
            return "ERROR"
    
    @abstractmethod
    def print_wayBill_document(self, tracking_number: str, amount_of_copies: int) -> None:
        """
        Prints waybill documents

        Args:
            tracking_number (str): tracking number
            amount_of_copies (int): amount of copies to print
        """
        pass

    @abstractmethod
    def print_label_document(self, tracking_number: str) -> None:
        """
        Prints label documents

        Args:
            tracking_number (str): tracking number
        """
        pass

    @abstractmethod
    def print_return_wayBill_document(self, return_tracking_number: str, amount_of_copies: int) -> None:
        """
        Prints return waybill documents

        Args:
            return_tracking_number (str): return tracking number
            amount_of_copies (int): amount of copies to print
        """
        pass

    # Private methods
    def __build_driver__(self, folder_path_to_download: str):
        browser = Browser(folder_path_to_download)
        driver = browser.driver
        wait = WebDriverWait(driver, 10)

        return driver, wait

    def __print_webpage__(self, driver: webdriver, url: str) -> None:
        try:
            driver.get(url) # this print since chrome options are set to print automatically
            driver.implicitly_wait(5)

        except Exception as e:
            self.log.add_error_log(f"Error printing documents: {e}")
    
    def __quit_driver__(self, driver) -> None:
        try:
            driver.quit()
        except Exception as e:
            self.log.add_error_log(f"Error quitting driver: {e}")