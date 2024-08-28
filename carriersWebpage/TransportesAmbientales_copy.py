import requests
import json

from carriersWebpage.carrierWebPage import CarrierWebpage
from logClass.log import Log

class TransportesAmbientales(CarrierWebpage):
    def __init__(self, folder_path_to_download: str, log: Log):
        """
        Class constructor for Transportes Ambientales

        Args:
            folder_path_to_download (str): folder path to download the documents
            log (Log): log object
        """
        super().__init__(log)
        self.folder_path_to_download = folder_path_to_download
        
        self.url_base = "https://sgi.tanet.com.ar/webservice/"

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
        url = f"{self.url_base}/login"
        payload = {
            "username": username,
            "password": password
        }

        response = requests.post(url, data=json.dumps(payload), headers={"Content-Type": "application/json"})
        response_data = response.json()

        if response_data.get('result') == 'OK':
            self.rsid = response_data.get('rsid')
            return True
        else:
            for response_error in response_data.get('errors'):
                self.log.add_error_log(f"Error: {response_error}")
            return False
        
    def complete_shipping_order_form(self, carrier_id: str, reference: str, 
                                ship_date: str, ship_time_from: str, ship_time_to: str, 
                                delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                type_of_material: str, temperature: str,
                                contacts: str, amount_of_boxes: int) -> str:
        
        url = f"{self.url_base}/createShippingOrder"
        contacts = self.__standarize_contacts__(contacts)
        payload = {
            "rsid": self.rsid,
            "carrier_id": carrier_id,
            "reference": reference,
            "ship_date": ship_date,
            "ship_time_from": ship_time_from,
            "ship_time_to": ship_time_to,
            "delivery_date": delivery_date,
            "delivery_time_from": delivery_time_from,
            "delivery_time_to": delivery_time_to,
            "type_of_material": type_of_material,
            "temperature": temperature,
            "contacts": contacts,
            "amount_of_boxes": amount_of_boxes
        }

        response = requests.post(url, data=json.dumps(payload), headers={"Content-Type": "application/json"})
        response_data = response.json()

        if response_data.get('result') == 'OK':
            return response_data.get('tracking_number')
        else:
            raise Exception(f"Shipping order creation failed: {response_data.get('errors')}")

    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                            delivery_date: str, return_time_from: str,
                                            return_time_to: str, type_of_return: str,
                                            contacts: str, amount_of_boxes_to_return: int,
                                            return_to_carrier_depot: bool, tracking_number: str) -> str:
        
        url_return = f"{self.url_base}/createReturn"
    
        # Mapeo del tipo de retorno
        if type_of_return == "CREDO":
            it_tipo_retorno = 1
        elif type_of_return == "DATALOGGER":
            it_tipo_retorno = 2
        elif type_of_return == "CREDO AND DATALOGGER":
            it_tipo_retorno = 3
        else:
            raise ValueError("Type of return not valid")

        payload = {
            "rsid": self.rsid,
            "carrier_id": carrier_id,
            "reference_return": reference_return,
            "delivery_date": delivery_date,
            "return_time_from": return_time_from,
            "return_time_to": return_time_to,
            "type_of_return": it_tipo_retorno,
            "contacts": contacts,
            "amount_of_boxes_to_return": amount_of_boxes_to_return,
            "return_to_carrier_depot": return_to_carrier_depot,
            "tracking_number": tracking_number
        }

        response = requests.post(url_return, data=json.dumps(payload), headers={"Content-Type": "application/json"})
        response_data = response.json()

        if response_data.get('result') == 'OK':
            return response_data.get('return_tracking_number')
        else:
            raise Exception(f"Return creation failed: {response_data.get('errors')}")
    
    def print_wayBill_document(self, tracking_number: str, amount_of_copies: int) -> None:
        url_guias = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id={tracking_number[:7]}&idservicio={tracking_number[:7]}&copies={amount_of_copies}"
        self.__print_webpage__(self.driver, url_guias)

    def print_label_document(self, tracking_number: str) -> None:
        url_rotulo = f"https://sgi.tanet.com.ar/sgi/srv.RotuloFCSPdf.emitir+id={tracking_number[:7]}&idservicio={tracking_number[:7]}"
        self.__print_webpage__(self.driver, url_rotulo)

    def print_return_wayBill_document(self, return_tracking_number: str, amount_of_copies: int) -> None:
        url_guias_return = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id={return_tracking_number[:7]}&idservicio={return_tracking_number[:7]}&copies={amount_of_copies}"
        self.__print_webpage__(self.driver, url_guias_return)

    def __standarize_contacts__(self, contacts: str) -> str:
        replacements = [" / ", "/ ", " /", "/", 
                        " ; ", "; ", " ;", ";", 
                        " , ", ", ", " ,", ",",
                        " - ", "- ", " -", "-"]
        for replacement in replacements:
            contacts = contacts.replace(replacement, ", ")

        contacts = contacts.title() # Capitalize
        contacts = contacts.strip() # Trim

        if contacts[-1:] == ",":
            contacts = contacts[:-1]

        return contacts