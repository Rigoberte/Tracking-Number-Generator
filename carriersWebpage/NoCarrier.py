import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC

from .carrierWebPage import CarrierWebpage

class NoCarrier(CarrierWebpage):
    def __init__(self, folder_path_to_download: str = ""):
        """
        Class constructor for NoCarrier

        Args:
            driver (webdriver): selenium driver
        """
        self.folder_path_to_download = folder_path_to_download

    def build_driver(self) -> None:
        self.driver, self.wait = self.__build_driver__(self.folder_path_to_download)

    def quit_driver(self) -> None:
        self.__quit_driver__(self.driver)
    
    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        return False

    def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                ship_date: str, ship_time_from: str, ship_time_to: str,
                                delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                type_of_material: str, temperature: str,
                                contacts: str, amount_of_boxes: int) -> str:
        return ""

    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                            delivery_date: str, return_time_from: str,
                                            return_time_to: str, type_of_return: str,
                                            contacts: str, amount_of_boxes_to_return: int,
                                            return_to_carrier_depot: bool, tracking_number: str) -> str:
        return ""
    
    def printWayBillDocument(self, tracking_number: str, amount_of_copies: int) -> None:
        return ""
    
    def printLabelDocument(self, tracking_number: str) -> None:
        return ""
    
    def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int) -> None:
        return ""