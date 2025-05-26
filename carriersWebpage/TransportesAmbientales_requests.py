import requests

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

from carriersWebpage.carrierWebPage import CarrierWebpage
from logClass.log import Log

class TransportesAmbientales_requests(CarrierWebpage):
    def __init__(self, folder_path_to_download: str, log: Log):
        """
        Class constructor for Transportes Ambientales

        Args:
            folder_path_to_download (str): folder path to download the documents
            log (Log): log object
        """
        super().__init__(log)
        self.folder_path_to_download = folder_path_to_download
        
        self.url_base = "http://sgi.tanet.com.ar/sgi"
        self.sitesDataFrame = pd.DataFrame()

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
        login_url = f"{self.url_base}/usr.usrJSON.login+login_username={username}&login_password={password}"
        
        response = self.__do_a_http_request__(login_url)

        if response.get('result') == 'FAIL':
            return False
        
        self.rsid = response['data']['data'].get('RSID')
        self.__complete_login_form__(username, password)
        return True
        
    def complete_shipping_order_form(self, carrier_id: str, reference: str, 
                                ship_date: str, ship_time_from: str, ship_time_to: str, 
                                delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                type_of_material: str, temperature: str,
                                contacts: str, amount_of_boxes: int) -> str:
        
        # Build the dates and times (dd/mm/yyyy hh:mm)
        retiradde = f"{ship_date} {ship_time_from}"
        retirahta = f"{ship_date} {ship_time_to}"
        entregadde = f"{delivery_date} {delivery_time_from}"
        entregahta = f"{delivery_date} {delivery_time_to}"

        self.shipment_entregadde = entregadde
        self.shipment_entregahta = entregahta

        # Comments
        ubicacion = self.__get_site_info__(carrier_id)

        if len(ubicacion) > 0:
            ubicacion_sector = str(ubicacion.get('sector', 'NA'))
            ubicacion_site = str(ubicacion.get('site', 'NA'))
            telContacto = str(ubicacion.get('telefono_contacto', 'NA'))
            telContacto = telContacto.replace("+", "")
        else:
            ubicacion_sector = "NA"
            ubicacion_site = "NA"
            telContacto = "NA"

        comments = f"Sector de destino: {ubicacion_sector}\n"
        comments += f"Site: {ubicacion_site}\n"
        if temperature != "Ambient":
            comments += "***ENVÍO CON CAJA CREDO. EL COURIER AGUARDARÁ QUE EL CENTRO ALMACENE LA MEDICACIÓN Y RETORNE EL EMBALAJE.***"
        
        # Selects the type of material
        it_type_of_material = self.__get_it_type_of_material__(type_of_material)

        # Selects the temperature
        it_temperature = self.__get_it_temperature__(temperature)

        # Standarize contacts
        if contacts == "" or contacts == "No contact":
                contacts = str(ubicacion['contacto'])
        contacts = self.__standarize_contacts__(contacts)
        self.contacts = contacts

        # Builds URL
        create_shipment_url = f"{self.url_base}/srv.SrvClienteJSON.crearEnvio+RSID={self.rsid}&idubicacion={carrier_id}"
        create_shipment_url += f"&referencia={reference}&retiradde={retiradde}&retirahta={retirahta}&entregadde={entregadde}&entregahta={entregahta}"
        create_shipment_url += f"&obsOper={comments}&tipomaterial={it_type_of_material}&temperatura={it_temperature}&autRecibe={contacts}&telContacto={telContacto}&cajas={amount_of_boxes}"
        
        response = self.__do_a_http_request__(create_shipment_url)

        if response.get('result') == 'FAIL':
            return ""
        
        return response.get('data').get('data').get('strJob')

    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                            delivery_date: str, return_time_from: str,
                                            return_time_to: str, type_of_return: str,
                                            contacts: str, amount_of_boxes_to_return: int,
                                            return_to_carrier_depot: bool, tracking_number: str) -> str:
        
    
        # Build the dates and times (dd/mm/yyyy hh:mm)
        retiradde = self.shipment_entregadde
        retirahta = self.shipment_entregahta
        entregadde = f"{delivery_date} {return_time_from}"
        entregahta = f"{delivery_date} {return_time_to}"

        # Standarize contacts
        contacts = self.contacts

        ubicacion = self.__get_site_info__(carrier_id)

        if len(ubicacion) > 0:
            ubicacion_sector = str(ubicacion.get('sector', 'NA'))
            ubicacion_site = str(ubicacion.get('site', 'NA'))
            ubicacion_visitas = str(ubicacion.get('visitas', 'NA'))
            ubicacion_telefono_contacto = str(ubicacion.get('telefono_contacto', 'NA'))
            ubicacion_telefono_contacto = ubicacion_telefono_contacto.replace("+", "")
        else:
            ubicacion_sector = "NA"
            ubicacion_site = "NA"
            ubicacion_visitas = "NA"
            ubicacion_telefono_contacto = "NA"

        # Comments
        obsOper = f"Sector de recolección: {ubicacion_sector}\n"
        obsOper += f"Site: {ubicacion_site}\n"
        obsOper += f"Dias y horarios de recolección: {ubicacion_visitas}\n"
        obsOper += f"Persona de contacto: {contacts}\n"
        obsOper += f"Telefono: {ubicacion_telefono_contacto}"


        tipomaterial = self.__get_it_type_of_material__("Other")
        temperatura = self.__get_it_temperature__("Ambient")
        idserviciopadre = tracking_number[:7]
        returnTo = "TA" if return_to_carrier_depot else "FCS"
        it_tipo_retorno = self.__get_it_tipo_retorno__(type_of_return)
        
        url_return = f"{self.url_base}/srv.SrvClienteJSON.crearRetorno+RSID={self.rsid}&idubicacion={carrier_id}"
        url_return += f"&referencia={reference_return}&retiradde={retiradde}&retirahta={retirahta}&entregadde={entregadde}&entregahta={entregahta}"
        url_return += f"&obsOper={obsOper}&tipomaterial={tipomaterial}&temperatura={temperatura}&cajas={amount_of_boxes_to_return}"
        url_return += f"&idserviciopadre={idserviciopadre}&returnTo={returnTo}&tiporetorno={it_tipo_retorno}"

        response = self.__do_a_http_request__(url_return)

        if response.get('result') == 'FAIL':
            return "ERROR"
        
        return response.get('data').get('data').get('strJob')
    
    def print_wayBill_document(self, tracking_number: str, amount_of_copies: int) -> None:
        url_guias = f"{self.url_base}/srv.SrvPdf.emitirOde+id={tracking_number[:7]}&idservicio={tracking_number[:7]}&copies={amount_of_copies}"
        self.__print_webpage__(self.driver, url_guias)

    def print_label_document(self, tracking_number: str) -> None:
        url_rotulo = f"{self.url_base}/srv.RotuloFCSPdf.emitir+id={tracking_number[:7]}&idservicio={tracking_number[:7]}"
        self.__print_webpage__(self.driver, url_rotulo)

    def print_return_wayBill_document(self, return_tracking_number: str, amount_of_copies: int) -> None:
        url_guias_return = f"{self.url_base}/srv.SrvPdf.emitirOde+id={return_tracking_number[:7]}&idservicio={return_tracking_number[:7]}&copies={amount_of_copies}"
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

        contacts = f"[{contacts}]"

        return contacts

    def __complete_login_form__(self, username: str, password: str) -> None:
        """
        Completes login form

        Args:
            driver (webdriver): selenium driver
            username (str): username
            password (str): password
        """
        self.driver.get(f"{self.url_base}/index.php")
        
        self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/input[1]")))
        self.driver.find_element(By.XPATH, "/html/body/form/input[1]").send_keys(username)
        self.driver.find_element(By.XPATH, "/html/body/form/input[2]").send_keys(password)
        self.driver.find_element(By.XPATH, "/html/body/form/button").click()

    def __do_a_http_request__(self, url: str) -> dict:
        response = requests.get(url)
        fail_dict = {"result": "FAIL", "data": "N/A"}

        if response.status_code != 200:
            self.log.add_error_log(f"Error HTTP: {response.status_code}. El servidor no respondió correctamente.")
            return fail_dict

        try:
            response_data = response.json()
        except ValueError:
            self.log.add_error_log(f"Contenido de la respuesta: {response.text}")
            self.log.add_error_log(f"La respuesta no es un JSON válido. Verifica el servidor o los parámetros de la solicitud.")
            return fail_dict
            

        if 'err' in response_data and response_data['err']:
            self.log.add_error_log(f"Error: {response_data['err']}")
            return fail_dict
        
        success_dict = {"result": "OK", "data": response_data}
        return success_dict
    
    def __get_it_type_of_material__(self, type_of_material: str) -> int:
        if type_of_material == "Medicine": return 3
        elif type_of_material == "Other": return 4
        elif type_of_material == "Ancillary": return 6
        elif type_of_material == "Ancillary Type 1": return 6
        elif type_of_material == "Ancillary Type 2": return 6
        elif type_of_material == "Equipment": return 8

    def __get_it_temperature__(self, temperature: str) -> int:
        if temperature == "Ambient": return 1
        elif temperature == "Refrigerated": return 2
        elif temperature == "Frozen": return 2
        elif temperature == "Refrigerated with Dry Ice": return 3
        elif temperature == "Frozen with Liquid Nitrogen": return 4
        elif temperature == "Controlled Ambient": return 5

    def __get_it_tipo_retorno__(self, type_of_return: str) -> str:
        if type_of_return == "CREDO": return 'E'
        elif type_of_return == "DATALOGGER": return 'R'
        elif type_of_return == "CREDO AND DATALOGGER": return 'ER'

    def __load_sites__(self, rsid: str) -> None:
        if len(self.sitesDataFrame) > 0:
            return

        partial_study_name = ""
        partial_site_name = ""
        site_id = ""
        
        search_site_url = f"{self.url_base}/srv.SrvClienteJSON.buscarUbicacion"
        search_site_url += f"+RSID={rsid}"
        search_site_url += f"&nomlinea={partial_study_name}"
        search_site_url += f"&site={partial_site_name}"
        search_site_url += f"&idubicacion={site_id}"
        
        response = self.__do_a_http_request__(search_site_url)

        if response.get('result') == 'FAIL':
            return

        sitesDataFrame = pd.DataFrame(response.get('data').get('data')).T

        self.sitesDataFrame = sitesDataFrame

    def __get_site_info__(self, carrier_id: str) -> dict:
        try:
            self.__load_sites__(self.rsid)

            site_info = self.sitesDataFrame.loc[self.sitesDataFrame['idubicacion'] == carrier_id]

            if len(site_info) > 0:
                site_info = site_info.iloc[0].to_dict()
            else:
                site_info = None
        except:
            sites_info = None
        finally:
            return site_info