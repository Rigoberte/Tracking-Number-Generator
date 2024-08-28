import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC

from carriersWebpage.carrierWebPage import CarrierWebpage
from logClass.log import Log

class TransportesAmbientales(CarrierWebpage):
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
        self.driver.get("https://sgi.tanet.com.ar/sgi/index.php")
        
        self.complete_login_form(username, password)

        try:
            time.sleep(0.5)
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[2]/h1")))
            return True
        except Exception as e:
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
        self.driver.find_element(By.XPATH, "/html/body/form/input[1]").send_keys(username)
        self.driver.find_element(By.XPATH, "/html/body/form/input[2]").send_keys(password)
        self.driver.find_element(By.XPATH, "/html/body/form/button").click()

    def complete_shipping_order_form(self, carrier_id: str, reference: str, 
                                ship_date: str, ship_time_from: str, ship_time_to: str, 
                                delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                type_of_material: str, temperature: str,
                                contacts: str, amount_of_boxes: int) -> str:
        tracking_number = ""
        url = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion={carrier_id}"

        try:
            self.driver.get(url)

            # Wait for the webpage to load
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/input")))

            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").send_keys(reference)

            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[1]").send_keys(ship_date)
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[2]").send_keys(ship_time_from)
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[3]").send_keys(ship_time_to)

            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[1]").send_keys(delivery_date)
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[2]").send_keys(delivery_time_from)
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[3]").send_keys(delivery_time_to)

            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/input").send_keys("Fisher Clinical Services FCS")
            self.driver.implicitly_wait(5)

            suggestions_container = self.wait.until(EC.presence_of_element_located((By.ID, "suggest_nomDomOri_list")))
            self.driver.implicitly_wait(5)
            time.sleep(1)

            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/div/div/div[1]/table/tbody")))
            self.driver.implicitly_wait(5)
            time.sleep(1)

            # Finds all suggested buttons inside the suggestions container
            suggested_buttons = suggestions_container.find_elements(By.CLASS_NAME, "suggest")
            # Loops through the suggested buttons and finds the one that contains the text "Fisher Clinical Services FCS"
            for button in suggested_buttons:
                button_text = button.text.strip()
                if "Fisher Clinical Services FCS" in button_text:
                    button.click()
                    break

            # Selects the material type
            if type_of_material == "Medicine": it_type_of_material = 3
            elif type_of_material == "Ancillary": it_type_of_material = 5
            elif type_of_material == "Equipment": it_type_of_material = 7
            else: raise ValueError("Type of material not valid")

            for i in range(0, it_type_of_material):
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[1]").send_keys(Keys.DOWN)

            # Selects the temperature
            if temperature == "Ambient": it_temperature = 0
            elif temperature == "Controlled Ambient": it_temperature = 1
            elif temperature == "Refrigerated": it_temperature = 2
            elif temperature == "Frozen": it_temperature = 2
            elif temperature == "Refrigerated with Dry Ice": it_temperature = 3
            elif temperature == "Frozen with Liquid Nitrogen": it_temperature = 4
            else: raise ValueError("Temperature not valid")

            for i in range(0, it_temperature):
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[2]").send_keys(Keys.DOWN)

            observaciones_textbox = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/textarea")
            comments = ""
            for line in observaciones_textbox.text.split("\n"):
                if "Dias y horarios de entrega:" != line[:27]:
                    comments += line + "\n"

            if temperature != "Ambient":
                comments += "***ENVÍO CON CAJA CREDO. EL COURIER AGUARDARÁ QUE EL CENTRO ALMACENE LA MEDICACIÓN Y RETORNE EL EMBALAJE***"

            observaciones_textbox.clear()
            observaciones_textbox.send_keys(comments)

            # Get value if contacts is empty
            if contacts == "" or contacts == "No contact":
                contacts = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input").get_attribute("value")
            
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input").clear()

            contacts = self.__standarize_contacts__(contacts)

            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input").send_keys(contacts)


            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").send_keys(amount_of_boxes)

            time.sleep(0.5)

            try:
                # Sometimes the button is not found by ID, so we try different XPATHs
                buttonToConfirmOrder = self.driver.find_element(By.XPATH, '//*[@id="btn_Grabar_F8_"]')
            except:
                try:
                    buttonToConfirmOrder = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[14]/td/button")
                except:
                    try:
                        buttonToConfirmOrder = self.driver.find_element(By.ID, "btn_Grabar_F8_")
                    except:
                        raise Exception("'Grabar' Button not found")

            self.driver.implicitly_wait(5)                    
            time.sleep(0.5)

            buttonToConfirmOrder.click()

            # Wait for the webpage to load and get the tracking number
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))

            self.driver.implicitly_wait(5)

            tracking_number = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:14]


            """
            # To avoid using selenium and avoid web scraping, we will change the code with something similar to this:
            # (this code is not tested and the web service is not implemented on carrier webpage yet. It will be implemented around August 15, 2024)

            url = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion={carrier_id}"

            headers = {
                "Authorization": "Bearer TOKEN",
                "Content-Type": "application/json",
                "Accept": "application/json"
            }

            order_parameters = {
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

            response = requests.post(url, headers=headers, data=json.dumps(data))
            numero_de_servicio = ""

            if response.status_code == 201:
                respuesta_json = response.json()
                numero_de_servicio = respuesta_json.get("TRACKING_NUMBER")
                print(f"Servicio creado exitosamente. Número de servicio: {numero_de_servicio}")
                
            else:
                raise Exception("{response.status_code} - {response.text}")

            return numero_de_servicio
            """
            
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
        url_return = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarRetorno+idsp={tracking_number[:7]}&idubicacion={carrier_id}"
        url_return += "&returnToTa=true" if return_to_carrier_depot else ""

        try:
            self.driver.get(url_return)
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select")))
            
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").clear()
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").send_keys(reference_return)
            
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[1]").send_keys(delivery_date)
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[2]").send_keys(return_time_from)
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[3]").send_keys(return_time_to)

            # Selects the return type
            if type_of_return == "CREDO": it_tipo_retorno = 1
            elif type_of_return == "DATALOGGER": it_tipo_retorno = 2
            elif it_tipo_retorno == "CREDO AND DATALOGGER": it_tipo_retorno = 3
            else: raise ValueError("Type of return not valid")

            for i in range(0, it_tipo_retorno):
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select").send_keys(Keys.DOWN)

            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input").send_keys(contacts)
            
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[14]/td/input").clear()
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[14]/td/input").send_keys(amount_of_boxes_to_return)

            try:
                buttonToConfirmOrder = self.driver.find_element(By.XPATH, '//*[@id="btn_Grabar_F8_"]')
            except:
                try:
                    buttonToConfirmOrder = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[15]/td/button")
                except:
                    try:
                        buttonToConfirmOrder = self.driver.find_element(By.ID, "btn_Grabar_F8_")
                    except:
                        raise Exception("'Grabar' Button not found")
            
            self.driver.implicitly_wait(5)
            time.sleep(0.5)

            buttonToConfirmOrder.click()
            
            # Wait for the webpage to load and get the return tracking number
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))
            
            self.driver.implicitly_wait(5)
            
            return_tracking_number = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:14]
        
        except Exception as e:
            return_tracking_number = "ERROR"
            raise Exception(e)
            
        finally:
            return return_tracking_number
    
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

        contacts = contacts.title()
        contacts = contacts.strip()
        if contacts[-1:] == ",":
            contacts = contacts[:-1]

        return contacts