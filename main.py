from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import datetime as dt
import tkinter as tk
import customtkinter as ctk
from tkcalendar import Calendar
from PIL import Image, ImageTk
import os, time, copy
import win32com.client as win32

class Browser(object):
    def __init__(self, folder_path: str):
        """
        Class constructor for Browser 

        Args:
            folder_path (str): folder path to download files

        Attributes:
            self.driver (webdriver): selenium self.driver
        """
        chrome_options = webdriver.ChromeOptions()
        #chrome_options.add_argument('--headless') # Disable headless mode
        chrome_options.add_argument('--disable-gpu') # Disable GPU acceleration
        chrome_options.add_argument('--disable-software-rasterizer')  # Disable software rasterizer
        chrome_options.add_argument('--disable-dev-shm-usage') # Disable shared memory usage
        chrome_options.add_argument('--no-sandbox') # Disable sandbox
        chrome_options.add_argument('--disable-extensions') # Disable extensions
        chrome_options.add_argument('--disable-sync') # Disable syncing to a Google account
        chrome_options.add_argument('--disable-webgl') # Disable WebGL 
        chrome_options.add_argument('--disable-gl-extensions') # Disable WebGL extensions
        chrome_options.add_argument('--disable-in-process-stack-traces') # Disable stack traces
        chrome_options.add_argument('--disable-logging') # Disable logging
        chrome_options.add_argument('--disable-cache') # Disable cache
        chrome_options.add_argument('--disable-application-cache') # Disable application cache
        chrome_options.add_argument('--disk-cache-size=1') # Set disk cache size to 1
        chrome_options.add_argument('--media-cache-size=1') # Set media cache size to 1
        chrome_options.add_argument('--kiosk-printing') # Enable kiosk printing
        chrome_options.add_argument('--kiosk-pdf-printing')

        chrome_prefs = {
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "download.open_pdf_in_system_reader": False,
            "profile.default_content_settings.popups": 0,
            #"download.prompt_for_download": False, #To auto download the file
            "printing.print_to_pdf": True,
            "download.default_directory": folder_path,
            "savefile.default_directory": folder_path
        }

        chrome_options.add_experimental_option("prefs", chrome_prefs)

        self.driver = webdriver.Chrome(options=chrome_options)

    def changeToHeadless(self):
        """
        Changes to headless mode
        """
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')

        self.driver = webdriver.Chrome(options=chrome_options)

class CarriersWebpage:
    def __init__(self, carrier: str = "", folder_path: str = ""):
        """
        Class constructor
        """

        self.carrierWebpageNames = ["Transportes Ambientales"]

        match carrier:
            case "Transportes Ambientales":
                self.carrierWebpage = self.TransportesAmbientales(folder_path)
            
            case _:
                self.carrierWebpage = self.NoCarrier(folder_path)

    def getCarrierWebpage(self):
        return self.carrierWebpage
    
    def getCarriersNames(self):
        return self.carrierWebpageNames
    
    def build_driver(self):
        """
        Builds the driver
        """
        self.carrierWebpage.build_driver()
    
    def quit(self):
        """
        Quits the browser
        """
        self.getCarrierWebpage().quit()

    def log_in_website(self) -> bool:
        """
        Logs in website

        Args:
            driver (webdriver): selenium driver
        """
        return self.getCarrierWebpage().log_in_website()
    
    def complete_shipping_order_form(self, TA_ID: str, reference: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        """
        Completes carrier form

        Args:
            self.driver (webdriver): selenium self.driver
            TA_ID (int): Site ID on carrier website
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
        return self.getCarrierWebpage().complete_shipping_order_form(TA_ID, reference,
                                    ship_date, ship_time_from, ship_time_to,
                                    delivery_date, delivery_time_from, delivery_time_to,
                                    type_of_material, temperature,
                                    contacts, amount_of_boxes)
    
    def complete_shipping_order_return_form(self, TA_ID: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
        """
        Completes carrier return form

        Args:
            self.driver (webdriver): selenium self.driver
            TA_ID (int): Site ID on carrier website
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
        return self.getCarrierWebpage().complete_shipping_order_return_form(TA_ID, reference_return,
                                                delivery_date, return_time_from,
                                                return_time_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)
    
    def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
        """
        Prints waybill documents

        Args:
        """
        self.getCarrierWebpage().printWayBillDocument(tracking_number, amount_of_copies)

    def printLabelDocument(self, tracking_number: str):
        """
        Prints label documents

        Args:
        """
        self.getCarrierWebpage().printLabelDocument(tracking_number)

    def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
        """
        Prints return waybill documents

        Args:
        """
        self.getCarrierWebpage().printReturnWayBillDocument(return_tracking_number, amount_of_copies)

    def print_webpage(self, url: str):
        """
        Prints webpage

        Args:
            self.driver (webdriver): selenium self.driver
            url (str): webpage url
        """
        self.getCarrierWebpage().print_webpage(url)

    def __print_webpage(self, url: str):
        """
        Prints webpage

        Args:
            self.driver (webdriver): selenium self.driver
            url (str): webpage url
        """
        try:
            self.driver.get(url) # this print since chrome options are set to print automatically
            self.driver.implicitly_wait(5)

        except Exception as e:
            print(f"Error printing documents: {e}")

    class NoCarrier:
        def __init__(self, folder_path: str = ""):
            """
            Class constructor for NoCarrier

            Args:
                driver (webdriver): selenium driver
            """
            self.folder_path = folder_path

        def build_driver(self):
            self.browser = Browser(self.folder_path)
            self.driver = self.browser.driver
            self.wait = WebDriverWait(self.driver, 10)

        def quit(self):
            self.browser.quit()

        def log_in_website(self) -> bool:
            return False

        def complete_shipping_order_form(self, TA_ID: str, reference: str,
                                    ship_date: str, ship_time_from: str, ship_time_to: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
            return ""

        def complete_shipping_order_return_form(self, TA_ID: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
            return ""
        
        def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
            return ""
        
        def printLabelDocument(self, tracking_number: str):
            return ""
        
        def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
            return ""
        
        def print_webpage(self, url: str):
            return ""
    
    class TransportesAmbientales:
        def __init__(self, folder_path: str = ""):
            """
            Class constructor for Transportes Ambientales

            Args:
                driver (webdriver): selenium driver
            """
            self.folder_path = folder_path

        def build_driver(self):
            self.browser = Browser(self.folder_path)
            self.driver = self.browser.driver
            self.wait = WebDriverWait(self.driver, 10)

        def quit(self):
            self.browser.quit()

        def log_in_website(self) -> bool:
            self.driver.get("https://sgi.tanet.com.ar/sgi/index.php")
            # Wait for the login webpage to load
            self.wait = WebDriverWait(self.driver, 30)

            try:
                # Wait user input their credentials 
                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/div[2]/div[1]/table/tbody/tr[1]/td")))
                print("Logged in")
                return True
            except:
                print("Not logged in")
                return False
            
            finally:
                self.wait = WebDriverWait(self.driver, 10)

        def complete_shipping_order_form(self, TA_ID: str, reference: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
            tracking_number = ""
            url = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion={TA_ID}"

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
                else: it_type_of_material = 8

                for i in range(0, it_type_of_material):
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[1]").send_keys(Keys.DOWN)

                # Selects the temperature
                if temperature == "Ambient": it_temperature = 0
                elif temperature == "Controlled Ambient": it_temperature = 1
                elif temperature == "Refrigerated": it_temperature = 2
                elif temperature == "Congelado": it_temperature = 3
                else: it_temperature = 4

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

                if contacts != "" and contacts != "No Contact":
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").clear()
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").send_keys(contacts)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input").send_keys(amount_of_boxes)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/button").click()

                # Wait for the webpage to load and get the tracking number
                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))

                self.driver.implicitly_wait(5)

                tracking_number = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:14]
                
            except Exception as e:
                print(f"Error completing shipping order form: {e}")
                print(f"Order: {reference}")

            finally:
                return tracking_number

        def complete_shipping_order_return_form(self, TA_ID: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
            return_tracking_number = ""
            url_return = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarRetorno+idsp={tracking_number[:7]}&idubicacion={TA_ID}"
            url_return += "&returnToTa=true" if return_to_TA else ""

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
                else: it_tipo_retorno = 3
                for i in range(0, it_tipo_retorno):
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select").send_keys(Keys.DOWN)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input[1]").send_keys(contacts)
                
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").clear()
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").send_keys(amount_of_boxes_to_return)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[14]/td/button").click()
                
                # Wait for the webpage to load and get the return tracking number
                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))
                
                self.driver.implicitly_wait(5)
                
                return_tracking_number = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:14]
            except Exception as e:
                print(f"Error completing shipping order return form: {e}")
                print(f"Order: {reference_return}")
                
            finally:
                return return_tracking_number
        
        def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
            url_guias = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id={tracking_number[:7]}&idservicio={tracking_number[:7]}&copies={amount_of_copies}"
            self.print_webpage(url_guias)

        def printLabelDocument(self, tracking_number: str):
            url_rotulo = f"https://sgi.tanet.com.ar/sgi/srv.RotuloPdf.emitir+id={tracking_number[:7]}&idservicio={tracking_number[:7]}"
            self.print_webpage(url_rotulo)

        def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
            url_guias_return = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id={return_tracking_number[:7]}&idservicio={return_tracking_number[:7]}&copies={amount_of_copies}"
            self.print_webpage(url_guias_return)
        
        def print_webpage(self, url: str):
            #self.CarriersWebpage().__print_webpage(url)
            try:
                self.driver.get(url) # this print since chrome options are set to print automatically
                self.driver.implicitly_wait(5)

            except Exception as e:
                print(f"Error printing documents: {e}")

class OrderProcessor:
    def __init__(self, folder_path : str, carrierWebpage):
        """
        Class constructor

        Args:
            self.driver (webdriver): selenium self.driver

        Attributes:
            self.driver (webdriver): selenium self.driver
            wait (WebDriverWait): selenium wait
        """
        self.folder_path = folder_path
        self.carrierWebpage = carrierWebpage
    
    def renameAllReturnFiles(self, df: pd.DataFrame):
        """
        Renames all return files

        Args:
            df (DataFrame): return tracking numbers
        """
        dataFrameWithReturnTrackingNumbers = df[(df["RETURN_TRACKING_NUMBER"] != "") & (df["PRINT_RETURN_DOCUMENT"])][["RETURN_TRACKING_NUMBER"]]
        
        for index, row in dataFrameWithReturnTrackingNumbers.iterrows():
            self.renameReturnPDFFile(row["RETURN_TRACKING_NUMBER"])

    def renameReturnPDFFile(self, return_tracking_number: str):
        """
        Renames the return file

        Args:
            return_tracking_number (str): return tracking number
        """
        try:
            pdf_path = self.folder_path + "\\JOB " + return_tracking_number + ".pdf"
            new_pdf_path = self.folder_path + "\\JOB " + return_tracking_number + " RETORNO DE CREDO.pdf"
            os.rename(pdf_path, new_pdf_path)
        except Exception as e:
            print(f"Error renaming return file: {e}")

    def setUserForm(self, userForm):
        # I know this breaks encapsulation
        self.userForm = userForm

    def updateTreeviewLine(self, index: int, tracking_number: str, return_tracking_number: str):
        """
        Updates a line in the treeview

        Args:
            index (int): row index
        """
        try:
            self.userForm.update_tag_color_of_a_treeview_line(index, self.userForm.getTreeview(), tracking_number, return_tracking_number)
        except Exception as e:
            print(f"Error updating treeview line: {e}")

    def send_email_to_medical_center(self, study: str, site: str, ivrs_number: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str, amount_of_boxes: int,
                                    return_: bool, type_of_return: str, amount_of_boxes_to_return: int,
                                    tracking_number: str, return_tracking_number: str,
                                    contacts: str):
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
            return_ (bool): if True, creates a return order
            type_of_return (str): type of return
            amount_of_boxes_to_return (int): number of boxes to return
            tracking_number (str): tracking number
            return_tracking_number (str): return tracking number
            contacts (str): contacts
        """
        try:
            def getEmailSourceFromTxtFile(file):
                """
                Gets the body from a txt file

                Args:
                    file (str): txt file

                Returns:
                    str: body
                """
                with open(file, 'r') as file:
                    return file.read()
            
            # Crear una instancia de Outlook
            outlook = win32.Dispatch('outlook.application')

            # Crear un nuevo mensaje de correo
            mail = outlook.CreateItem(0)

            emailSource = getEmailSourceFromTxtFile("email.txt")

            # Configurar los campos del correo
            mail.Subject = 'Asunto del Correo'
            mail.To = 'inaki.costa@thermofisher.com'
            mail.HTMLBody = emailSource  # Configurar el cuerpo del correo como HTML

            # Enviar el correo
            mail.Send()
            
            #send_email(subject, body)
        except Exception as e:
            print(f"Error sending email to medical center: {e}")

    def process_all_shipping_orders(self, df: pd.DataFrame):
        """
        Process all orders in the table

        Args:
            df (DataFrame): orders table
        """
        for index, row in df.iterrows():
            if df.loc[index, "TRACKING_NUMBER"] != "":
                continue

            error = row['HAS_AN_ERROR'] != "No error"

            if error:
                continue
            
            tracking_number, return_tracking_number = self.get_tracking_numbers_from_carrier(
                row["TA_ID"], row["SYSTEM_NUMBER"], row["IVRS_NUMBER"],
                row["SHIP_DATE"], row["SHIP_TIME_FROM"], row["SHIP_TIME_TO"],
                row["DELIVERY_DATE"], row["DELIVERY_TIME_FROM"], row["DELIVERY_TIME_TO"],
                row["TYPE_OF_MATERIAL"], row["TEMPERATURE"],
                row["CONTACTS"], row["AMOUNT_OF_BOXES"],
                row["RETURN"], row["RETURN_TO_TA"], row["RETURN_TYPE"], row["RETURN_CANTIDAD"], row["PRINT_RETURN_DOCUMENT"]
            )

            df.loc[index, "TRACKING_NUMBER"] = tracking_number
            df.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number

            """
            self.send_email_to_medical_center( row["STUDY"], row["SITE#"], row["IVRS_NUMBER"], 
            row["DELIVERY_DATE"], row["DELIVERY_TIME_FROM"], row["DELIVERY_TIME_TO"], 
            row["TYPE_OF_MATERIAL"], row["TEMPERATURE"], row["AMOUNT_OF_BOXES"],
            row["RETURN"], row["RETURN_TYPE"], row["RETURN_CANTIDAD"],
            row["TRACKING_NUMBER"], row["RETURN_TRACKING_NUMBER"],
            row["CONTACTS"] ) """

            self.updateTreeviewLine(index, tracking_number, return_tracking_number)

    def printOrderDocuments(self, tracking_number: str, return_tracking_number: str, print_return_document: bool):
        """
        Prints the order documents

        Args:
            tracking_number (str): tracking number
            return_tracking_number (str): return tracking number
            print_return_document (bool): if True, prints the return document
        """
        self.carrierWebpage.printWayBillDocument(tracking_number, 4)

        self.carrierWebpage.printLabelDocument(tracking_number)
        
        if print_return_document and return_tracking_number != "":
            self.carrierWebpage.printReturnWayBillDocument(return_tracking_number, 1)

    def get_tracking_numbers_from_carrier(self, TA_ID: int, system_number: str, ivrs_number: str, 
                               ship_date: str, ship_time_from: str, ship_time_to: str, 
                               delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                               type_of_material: str, temperature: str, 
                               contacts: str, amount_of_boxes: int, 
                               return_: bool, return_to_TA: bool, type_of_return: str, amount_of_boxes_to_return: int, print_return_document: bool) -> (str, str):
        """
        - Process an order by completing the carrier form
        - Creates a return order if necessary
        - Prints the order documents

        Args:
            TA_ID (int): Site ID on carrier website
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
            return_ (bool): if True, creates a return order
            return_to_TA (bool): if True, the return order is sent to carrier depot
            amount_of_boxes_to_return (int): number of boxes to return

        Returns:
            str: tracking number
            str: return tracking number
        """
        
        tracking_number, return_tracking_number = "", ""

        try:
            tracking_number = self.get_shipping_tracking_number(
                TA_ID, system_number, ivrs_number,
                ship_date, ship_time_from, ship_time_to,
                delivery_date, delivery_time_from, delivery_time_to,
                type_of_material, temperature,
                contacts, amount_of_boxes
            )

            if tracking_number == "":
                return "", ""
            
            return_tracking_number = self.get_return_tracking_number(
                TA_ID, system_number, ivrs_number,
                delivery_date, tracking_number, return_, return_to_TA, type_of_return, amount_of_boxes_to_return
            )

            self.printOrderDocuments(tracking_number, return_tracking_number, print_return_document)

        except Exception as e:
            print(f"Error processing order: {e}")
            print(f"Order: {system_number} {ivrs_number}")

        finally:
            return tracking_number, return_tracking_number

    def get_shipping_tracking_number(self, TA_ID: int, system_number: str, ivrs_number: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        """
        Gets the shipping tracking number

        Args:
            TA_ID (int): Site ID on carrier website
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

        tracking_number = self.carrierWebpage.complete_shipping_order_form(
            TA_ID, reference, 
            ship_date, ship_time_from, ship_time_to, 
            delivery_date, delivery_time_from, delivery_time_to, 
            type_of_material, temperature, 
            contacts, amount_of_boxes
        )
        
        return tracking_number
    
    def get_return_tracking_number(self, TA_ID: int, system_number: str, ivrs_number: str,
                                    delivery_date: str,  tracking_number: str, return_: bool,
                                    return_to_TA: bool, type_of_return: str, amount_of_boxes_to_return: int) -> str:
            """
            Gets the return tracking number
    
            Args:
                TA_ID (int): Site ID on carrier website
                system_number (int): order system number
                ivrs_number (str): order ivrs number
                delivery_date (str): delivery date
                tracking_number (str): tracking number
                return_ (bool): if True, creates a return order
                return_to_TA (bool): if True, the return order is sent to carrier depot
                type_of_return (str): type of return
                amount_of_boxes_to_return (int): number of boxes to return
    
            Returns:
                str: tracking number
            """
            if not return_ or tracking_number == "":
                return  ""

            reference_return = f"{system_number} {ivrs_number} RET {tracking_number}"[:50]
            transit_days = dt.timedelta(days=1)
            #transit_days = max(dt.datetime.strptime(delivery_date, '%d/%m/%Y') - dt.datetime.strptime(ship_date, '%d/%m/%Y'), dt.timedelta(days=1))
            return_delivery_date = dt.datetime.strptime(delivery_date, '%d/%m/%Y') + transit_days
            return_delivery_date += dt.timedelta(days=2) if return_delivery_date.weekday() >= 5 else dt.timedelta(days=0) # Add 2 days if the return delivery date is on a weekend
            return_delivery_date = return_delivery_date.strftime('%d/%m/%Y')

            return_tracking_number = self.carrierWebpage.complete_shipping_order_return_form(
                TA_ID, reference_return, 
                return_delivery_date, "9", "17", 
                type_of_return, "", amount_of_boxes_to_return, 
                return_to_TA, tracking_number
            )
        
            return return_tracking_number
    
    def export_to_excel(self, df: pd.DataFrame):
        """
        Exports the orders table to an excel file

        Args:
            df (DataFrame): orders table
        """
        if not df.empty:
            dataframe_name = "orders_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            df.to_excel(self.folder_path + "\\" + dataframe_name, index=False)
        else:
            print("Empty DataFrame")
    
    def generate_shipping_report(self, df:pd.DataFrame) -> pd.DataFrame:
        """
        Process all orders in the table
        
        Variables:
            wait (WebDriverWait): selenium wait
            team (str): team to process
            df_path (str): excel file path
            sheet (str): excel sheet name
            df (DataFrame): orders table

        Args:
            df (DataFrame): orders table

        Returns:
            DataFrame: orders table with tracking numbers
        """
        try:
            self.carrierWebpage.build_driver()

            if not self.carrierWebpage.log_in_website():
                return

            self.process_all_shipping_orders(df)

            time.sleep(2) # Wait for the download to finish

            self.renameAllReturnFiles(df)

            self.export_to_excel(df)

        except Exception as e:
            print(f"Error generating shipping report: {e}")
              
        finally:
            try:
                self.carrierWebpage.quit()
            finally:
                return df

class MyUserForm(tk.Tk):
    class Chroma:
        def __init__(self):
            """
            Class constructor for Chroma

            Attributes:
                self.dark (bool): dark mode
                self.body_color (str): body color
                self.sidebar_color (str): sidebar color
                self.primary_color (str): primary color
                self.primary_color_light (str): primary color light
                self.toggle_color (str): toggle color
                self.text_color (str): text color
            """
            self.dark = True
            self.body_color = '#18191A'
            self.sidebar_color = '#242526'
            self.primary_color = '#3A3B3C'
            self.primary_color_light = '#3A3B3C'
            self.toggle_color = '#FFF'
            self.text_color = '#CCC'

        def toggle(self):
            """
            Toggles between dark and light mode
            """
            self.dark = not self.dark
            if self.dark:
                self.body_color = '#18191A'
                self.sidebar_color = '#242526'
                self.primary_color = '#3A3B3C'
                self.primary_color_light = '#3A3B3C'
                self.toggle_color = '#FFF'
                self.text_color = '#CCC'
            else:
                self.body_color = '#E4E9F7'
                self.sidebar_color = '#FFF'
                self.primary_color = '#695CFE'
                self.primary_color_light = '#F6F5FF'
                self.toggle_color = '#DDD'
                self.text_color = '#707070'

    def __init__(self):
        """
        Class constructor for UserForm

        Attributes:
            self.colors (Chroma): color palette
            self.canvas (tk.Canvas): canvas
            self.cal (Calendar): calendar
            self.team_combobox (tk.ttk.Combobox): teams combobox
            self.treeview (tk.ttk.Treeview): self.treeview to show orders table
        """
        super().__init__()

        # UserForm
        self.title("Carrier Form AutoLoad")
        self.state("zoomed")

        self.colors = self.Chroma()

        self.columns_df = ['SYSTEM_NUMBER', 'IVRS_NUMBER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO', 
                'TYPE_OF_MATERIAL', 'TEMPERATURE', 'AMOUNT_OF_BOXES',
                'RETURN', 'RETURN_TO_TA', 'RETURN_TYPE', 'RETURN_CANTIDAD',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT', 'CONTACTS', 'TA_ID', 'HAS_AN_ERROR']

        self.selected_team = Teams()
        
        self.load_userform()
        
    def load_userform(self):
        """
        Loads the UserForm and their widgets
        """
        def create_frames(self):
            # Top Frame
            frame_top = ctk.CTkFrame(self)
            frame_top.pack(side=tk.TOP, fill=tk.X, pady=0)

            # Bottom Frame
            frame_bottom = ctk.CTkFrame(self)
            frame_bottom.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, pady=0)

            return frame_top, frame_bottom

        def load_top_logo_image(self, frame):
            # Image banner
            imagen = Image.open("TMO_logo.png")
            imagen = imagen.resize((284, 67))
            imagen_tk = ImageTk.PhotoImage(imagen)

            label_banner = tk.Label(frame, image=imagen_tk, bg=ctk.ThemeManager.theme["CTkFrame"]["fg_color"][1])
            label_banner.image = imagen_tk
            label_banner.pack(side=tk.LEFT, padx=10)

        def load_calendar_datePicker(self, frame):
            self.cal = Calendar(frame, selectmode='day', locale='en_US', disabledforeground='red',
                    cursor="hand2", background=ctk.ThemeManager.theme["CTkFrame"]["fg_color"][1],
                    selectbackground=ctk.ThemeManager.theme["CTkButton"]["fg_color"][1], date_pattern='yyyy-MM-dd')

            self.cal.pack(padx=50, pady=0, side=tk.LEFT)

        def load_teams_combobox(self, frame):
            teams_options = self.selected_team.getTeamsNames()
            self.team_combobox = tk.ttk.Combobox(frame, values=teams_options, width=20, height=15, font=30)
            self.team_combobox.pack(side=tk.LEFT, padx=10)
            self.team_combobox.current(0)
        
        def load_treeview(self, frame, treeviewColumns):
            style = tk.ttk.Style()
            style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 13)) # Font of the body
            style.configure("mystyle.Treeview.Heading", font=('Calibri', 13,'bold')) # Font of the headings
            style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) # Remove the borders
            
            self.treeview = tk.ttk.Treeview(frame, columns=treeviewColumns , show='headings', style="mystyle.Treeview")
            self.treeview.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            self.treeview.tag_configure('odd', background='#E8E8E8')
            self.treeview.tag_configure('odd_done', background='#C6E0B4')
            self.treeview.tag_configure('odd_error', background='#FFC7CE')

            self.treeview.tag_configure('even', background='#DFDFDF')
            self.treeview.tag_configure('even_done', background='#A9D08E')
            self.treeview.tag_configure('even_error', background='#FFA7BB')

            # Treeview columns headings and columns width
            self.treeview.column("#0", width=0, stretch=tk.NO)  # Hide the first column
            for col in treeviewColumns:
                self.treeview.heading(col, text=col)
                self.treeview.column(col, anchor=tk.W, width=int(self.winfo_screenwidth() * 0.7 * 0.3))
            self.treeview.column("#", anchor=tk.W, width=int(self.winfo_screenwidth() * 0.7 * 0.05))
        
        def load_carrier_combobox(self, frame):
            carrier_options = CarriersWebpage().getCarriersNames()
            self.carrier_combobox = tk.ttk.Combobox(frame, values=carrier_options, width=20, height=15, font=30)
            self.carrier_combobox.pack(side=tk.LEFT, padx=10)
            self.carrier_combobox.current(0)

        def load_horizontal_scrollbar(self, frame, treeview):
            x_scrollbar = tk.ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=treeview.xview)
            x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            treeview.configure(xscrollcommand=x_scrollbar.set)

        def load_button_processOrders(self, frame):
            processOrders_btn = ctk.CTkButton(master=frame, text="Process Orders", command=self.on_processOrders_btn_click,
                                            width=150, height=50, font=('Calibri', 22, 'bold'))
            processOrders_btn.pack(side=tk.RIGHT, padx=10)

        def load_button_loadOrders(self, frame):
            loadOrders_btn = ctk.CTkButton(master=frame, text="Load Orders", command=self.on_loadOrders_btn_click,
                                            width=150, height=50, font=('Calibri', 22, 'bold'))
            loadOrders_btn.pack(side=tk.RIGHT, padx=10)
        
        frame_top, frame_bottom = create_frames(self)
        load_top_logo_image(self, frame_top)
        load_calendar_datePicker(self, frame_top)
        load_teams_combobox(self, frame_top)
        #load_carrier_combobox(self, frame_top)
        
        load_treeview(self, frame_bottom, ['#'] + self.columns_df)

        load_horizontal_scrollbar(self, frame_bottom, self.treeview)

        load_button_processOrders(self, frame_top)

        load_button_loadOrders(self, frame_top)
    
    def clear_treeview(self, treeview):
        """
        Clears the treeview
        """
        for item in treeview.get_children():
            treeview.delete(item)

    def tag_color_of_a_treeview_line(self, parity: bool, order_done: bool, error: bool):
        """
        Tags colors on each treeview line

        Args:
            row (Series): row of the orders table
            parity (bool): parity
        """
        tag_color = 'odd' if parity else 'even'
        tag_color += "_done" if order_done else ""
        tag_color += ("_error" if error else "") if not order_done else ""
        return tag_color
    
    def update_tag_color_of_a_treeview_line(self, index: int, treeview, tracking_number: str = "", return_tracking_number: str = ""):
        """
        Updates tag color on each treeview line

        Args:
            index (int): row index
            row (Series): row of the orders table
            treeview (tk.ttk.Treeview): self.treeview to show orders table
        """
        try:
            if tracking_number != "":
                self.df.loc[index, "TRACKING_NUMBER"] = tracking_number
            
            if return_tracking_number != "":
                self.df.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number
            
            row_values = self.df.loc[index]
            row_values_list = [index] + list(row_values)

            treeview.item(index, values = row_values_list)

            parity = index % 2 == 0
            is_order_done = row_values['TRACKING_NUMBER'] != ""
            has_an_error = row_values['HAS_AN_ERROR'] != "No error"
            tag_color = self.tag_color_of_a_treeview_line(parity, is_order_done, has_an_error)
            treeview.item(index, tags=tag_color)

        except Exception as e:
            print(f"Error updating tag color of a treeview line: {e}")
    
    def tag_colors_for_each_treeview_line(self, df: pd.DataFrame, treeview):
        """
        Tags colors on each treeview line

        Args:
            df (DataFrame): orders table
        """
        if df.empty:
            return

        for index, row in df.iterrows():
            row_values = [index] + list(row)
            treeview.insert("", "end", iid=index, values=row_values)
            self.update_tag_color_of_a_treeview_line(index, treeview, row['TRACKING_NUMBER'], row['RETURN_TRACKING_NUMBER'])
            
    def update_treeview(self, df, treeview):
        """
        Updates the treeview

        Args:
            df (DataFrame): orders table
            treeview (tk.ttk.Treeview): self.treeview to show orders table
        """
        self.clear_treeview(treeview)
        self.tag_colors_for_each_treeview_line(df, treeview)
        treeview.update()

    def getTreeview(self) -> tk.ttk.Treeview:
        return self.treeview
    
    def on_loadOrders_btn_click(self):
        """
        Loads orders table according to date and team

        Args:
            date (dt.datetime): date to process
            team (str): team to process
            shipdate (dt.datetime): ship date
            self.treeview (tk.ttk.Treeview): self.treeview to show orders table
        
        Returns:
            DataFrame: orders table
        """

        def create_folder(team, date):
            folder_path = os.path.expanduser("~\\Downloads") + "\\" + "TA_Form_AutoLoad" + "\\" + team
            
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            folder_path += "\\" + date

            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            return folder_path
        
        selected_team_name = self.team_combobox.get()
        
        self.selected_date = self.cal.get_date()
        
        self.folder_path = create_folder(selected_team_name, self.selected_date.replace("/", "_"))
        selected_date = dt.datetime.strptime(self.selected_date, '%Y-%m-%d')
        
        self.selected_team = Teams(selected_team_name, self.folder_path)

        self.df = self.generate_shipping_order_tables(self.selected_team, selected_date)

        self.update_treeview(self.df, self.treeview)

    def on_processOrders_btn_click(self):
        """
        Button to process all orders in the table
        """
        carriersWebpage = self.selected_team.getCarrierWebpage()
        
        orderProcessor = OrderProcessor(self.folder_path, carriersWebpage)

        orderProcessor.setUserForm(self)
        
        self.df = orderProcessor.generate_shipping_report(self.df)

        self.update_treeview(self.df, self.treeview) # useless?

    def load_shipping_order_table(self, date: dt.datetime, team,  path: str, sheet: str) -> pd.DataFrame:
        """
        Loads orders table according to date and team

        Args:
            date (dt.datetime): date to process
            team (str): team to process
            path (str): excel file path
            sheet (str): excel sheet name

        Returns:
            DataFrame: orders table
        """
        
        columns_names, columns_types = team.get_column_rename_type_config_for_orders_tables()

        df = pd.read_excel(path, sheet_name=sheet, header=0, dtype=columns_types)
        df.rename(columns=columns_names, inplace=True)
        
        df = df[df["SHIP_DATE"] == date]
        
        df["SHIP_DATE"] = df["SHIP_DATE"].astype("datetime64[ns]")
        df["SHIP_DATE"] = pd.to_datetime(df["SHIP_DATE"], format='%d/%m/%Y', errors='coerce')
        #df["SHIP_DATE"] = df["SHIP_DATE"].dt.date

        df["SHIP_DATE"] = df["SHIP_DATE"].dt.strftime('%d/%m/%Y')

        df["TEMPERATURE"] = df["TEMPERATURE"].str.strip()
        df["SITE#"] = df["SITE#"].astype(object)
        df["AMOUNT_OF_BOXES"] = df["AMOUNT_OF_BOXES"].fillna(0).astype(int)

        df = team.apply_team_specific_changes_for_orders_tables(df)

        df.fillna('', inplace=True)
        
        return df[['SYSTEM_NUMBER', 'IVRS_NUMBER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'TYPE_OF_MATERIAL', 
                'TEMPERATURE', 'AMOUNT_OF_BOXES', 'RETURN', 
                'RETURN_TO_TA', 'RETURN_TYPE', 'RETURN_CANTIDAD',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT']]

    def load_table_info_sites(self, team, path: str, sheet: str) -> pd.DataFrame:
        """
        Loads sites info table according to team

        Args:
            team (str): team to process
            path (str): excel file path
            sheet (str): excel sheet name

        Returns:
            DataFrame: sites info table
        """
        columns_names, columns_types = team.get_column_rename_type_config_for_sites_table()

        df = pd.read_excel(path, sheet_name=sheet, header=0, dtype=columns_types)
        df.rename(columns=columns_names, inplace=True)
        df.fillna('', inplace=True)

        df = team.apply_team_specific_changes_for_sites_table(df)

        df["DELIVERY_TIME_FROM"] = pd.to_datetime(df["DELIVERY_TIME_FROM"], format='%H:%M:%S', errors='coerce')
        df["DELIVERY_TIME_TO"] = pd.to_datetime(df["DELIVERY_TIME_TO"], format='%H:%M:%S', errors='coerce')
        df["DELIVERY_TIME_FROM"] = df["DELIVERY_TIME_FROM"].dt.strftime('%H:%M')
        df["DELIVERY_TIME_TO"] = df["DELIVERY_TIME_TO"].dt.strftime('%H:%M')

        df.fillna('', inplace=True)

        return df[["STUDY", "SITE#", "TA_ID", "DELIVERY_TIME_FROM", "DELIVERY_TIME_TO", "CONTACTS"]]
    
    def checkErrorsOnEachOrder(self, row: pd.Series) -> str:
        """
        Checks errors on each order

        Args:
            row (Series): order row
        """
        def assertIfIsValidShipDate(row: pd.Series) -> bool:
            """
            Asserts if the ship date is valid

            Args:
                row (Series): order row
            """
            return row['SHIP_DATE'] != ""
        
        def assertIfIsValidDeliveryDate(row: pd.Series) -> bool:
            """
            Asserts if the delivery date is valid

            Args:
                row (Series): order row
            """
            return row['DELIVERY_DATE'] != ""
        
        def assertIfIsValidShipTime(row: pd.Series) -> bool:
            """
            Asserts if the ship time is valid

            Args:
                row (Series): order row
            """
            return row['SHIP_TIME_FROM'] != "" and row['SHIP_TIME_TO'] != ""
        
        def assertIfIsValidDeliveryTime(row: pd.Series) -> bool:
            """
            Asserts if the delivery time is valid

            Args:
                row (Series): order row
            """
            return row['DELIVERY_TIME_FROM'] != "" and row['DELIVERY_TIME_TO'] != ""
        
        def assertIfIsValidAmountOfBoxes(row: pd.Series) -> bool:
            """
            Asserts if the amount of boxes is valid

            Args:
                row (Series): order row
            """
            return row['AMOUNT_OF_BOXES'] > 0
        
        def assertIfIsValidCarrier_ID(row: pd.Series) -> bool:
            """
            Asserts if the carrier ID is valid

            Args:
                row (Series): order row
            """
            return row['TA_ID'] != ""
        
        def assertIfTemperaturesAreValid(row: pd.Series) -> bool:
            """
            Asserts if the temperatures are valid

            Args:
                row (Series): order row
            """
            return row['TEMPERATURE'] in ["Ambient", "Controlled Ambient", "Refrigerated"]
        
        errors = ""
        if not assertIfIsValidShipDate(row):
            errors += "No ship date; "
        
        if not assertIfIsValidDeliveryDate(row):
            errors += "No delivery date; "
        
        if not assertIfIsValidShipTime(row):
            errors += "No ship time; "
        
        if not assertIfIsValidDeliveryTime(row):
            errors += "No delivery time; "
        
        if not assertIfIsValidAmountOfBoxes(row):
            errors += "No amount of boxes; "
        
        if not assertIfIsValidCarrier_ID(row):
            errors += "No carrier ID; "

        if not assertIfTemperaturesAreValid(row):
            errors += "Invalid temperature; "

        return "No error" if errors == "" else errors

    def generate_shipping_order_tables(self, team, shipdate: dt.datetime) -> pd.DataFrame:
        """
        Process all orders in the table

        Args:
            team (str): team to process
            shipdate (dt.datetime): date to process

        Returns:
            DataFrame: orders table with standarisized data
        """
        df_path, sheet, info_sites_sheet = team.get_data_path()
        
        df_orders = self.load_shipping_order_table(shipdate, team, df_path, sheet)

        df_info_sites = self.load_table_info_sites(team, df_path, info_sites_sheet)
        
        df = pd.merge(df_orders, df_info_sites, on=["STUDY", "SITE#"], how="inner")

        df["HAS_AN_ERROR"] = df.apply(self.checkErrorsOnEachOrder, axis=1)
        
        return df[self.columns_df]

class Teams():
    def __init__(self, team: str = "", folder_path: str = ""):
        """
        Class constructor for teams
        """

        self.teamsNames = ["Eli Lilly Argentina", "GPM Argentina", "Test", "Test_5_ordenes"]
        
        match team:
            case "Eli Lilly Argentina":
                self.selectedTeam = self.EliLillyArgentinaTeam(folder_path)
            case "GPM Argentina":
                self.selectedTeam = self.GPMArgentinaTeam(folder_path)
            case "Test":
                self.selectedTeam = self.TestTeam(folder_path)
            case "Test_5_ordenes":
                self.selectedTeam = self.TestTeam(folder_path)

            # ---------------------
            case _:
                self.selectedTeam = self.NoSelectedTeam("")

    def getCarrierWebpage(self):
        """
        Gets the carrier webpage
        """
        return self.getTeam().getCarrierWebpage()

    def getTeam(self):
        """
        Gets the team
        """
        return copy.copy(self.selectedTeam)
    
    def getTeamsNames(self) -> list:
        """
        Gets teams names
        """
        return self.teamsNames
    
    def getTeamName(self) -> str:
        """
        Gets team name
        """
        return self.getTeam().getTeamName()
    
    def get_column_rename_type_config_for_sites_table(self) -> (dict, dict):
        """
        Loads columns names and types for the sites info table

        Returns:
            dict: columns names
            dict: columns types
        """
        return self.getTeam().get_column_rename_type_config_for_sites_table()
    
    def apply_team_specific_changes_for_sites_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Applies team specific changes to sites info table

        Args:
            df (DataFrame): sites info table

        Returns:
            DataFrame: sites info table
        """
        return self.getTeam().apply_team_specific_changes_for_sites_table(df)
    
    def get_data_path(self) -> (str, str, str):
        """
        Loads data path

        Returns:
            str: excel file path
            str: excel sheet name
            str: excel sheet name with sites info
        """
        return self.getTeam().get_data_path()
    
    def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
        """
        Loads columns names and types for the orders table

        Returns:
            dict: columns names
            dict: columns types
        """
        return self.getTeam().get_column_rename_type_config_for_orders_tables()
    
    def apply_team_specific_changes_for_orders_tables(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Applies team specific changes to orders table

        Args:
            df (DataFrame): orders table

        Returns:
            DataFrame: orders table
        """
        return self.getTeam().apply_team_specific_changes_for_orders_tables(df)
    
    class NoSelectedTeam:
        def __init__(self, folder_path: str):
            self.carrierWebpage = CarriersWebpage("", folder_path)

        def getCarrierWebpage(self):
            return self.carrierWebpage

        def getTeamName(self) -> str:
            return "No Selected Team"
        
        def get_column_rename_type_config_for_sites_table(self) -> (dict, dict):
            return {}, {}
        
        def apply_team_specific_changes_for_sites_table(self, df: pd.DataFrame) -> pd.DataFrame:
            return df
        
        def get_data_path(self) -> (str, str, str):
            return ""
        
        def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
            return {}, {}
        
        def apply_team_specific_changes_for_orders_tables(self, df: pd.DataFrame) -> pd.DataFrame:
            return df
    
    class EliLillyArgentinaTeam:
        def __init__(self, folder_path: str):
            self.carrierWebpage = CarriersWebpage("Transportes Ambientales", folder_path)

        def getCarrierWebpage(self):
            return self.carrierWebpage

        def getTeamName(self) -> str:
            return "Eli Lilly Argentina"

        def get_column_rename_type_config_for_sites_table(self) -> (dict, dict):
            columns_names = {"Protocolo": "STUDY", "Codigo": "TA_ID", "Site": "SITE#",
                            "Horario inicio": "DELIVERY_TIME_FROM", "Horario fin": "DELIVERY_TIME_TO"}
            columns_types = {"Protocolo": str, "Site": str, "Codigo": str, "Horario inicio": str, "Horario fin": str}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_sites_table(self, df: pd.DataFrame) -> pd.DataFrame:
            df["CONTACTS"] = "No contact"
            return df
        
        def get_data_path(self) -> (str, str, str):
            path = os.path.expanduser("~\\Thermo Fisher Scientific\Power BI Lilly Argentina - General\Share Point Lilly Argentina.xlsx")
            sheet = "Shipments"
            info_sites_sheet = "Dias y Destinos"
            return path, sheet, info_sites_sheet
        
        def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
            columns_names = {"CT-WIN": "SYSTEM_NUMBER", "IVRS": "IVRS_NUMBER",
                            "Trial Alias": "STUDY", "Site ": "SITE#",
                            "Order received": "ENTER DATE", "Ship date": "SHIP_DATE",
                            "Horario de Despacho": "SHIP_TIME_FROM",  
                            "Dia de entrega": "DELIVERY_DATE", "Destination": "DESTINATION",
                            "CONDICION": "TEMPERATURE", "TT4": "AMOUNT_OF_BOXES",  
                            "AWB": "TRACKING_NUMBER", "Shipper return AWB": "RETURN_TRACKING_NUMBER"}
            columns_types = {"CT-WIN": str, "IVRS": str, 
                            "Trial Alias": str, "Site ": str, 
                            "Order received": str, "Ship date": dt.datetime,
                            "Horario de Despacho": str,
                            "Dia de entrega": dt.datetime, "Destination": str,
                            "CONDICION": str, "TT4": str, 
                            "AWB": str, "Shipper return AWB": str}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_orders_tables(self, df: pd.DataFrame) -> pd.DataFrame:
            df["TYPE_OF_MATERIAL"] = "Medicine"
            
            temperatures = {"L": "Ambient",
                            "M": "Controlled Ambient", "M + L": "Controlled Ambient",
                            "H": "Controlled Ambient", "H + M": "Controlled Ambient", "H + L": "Controlled Ambient", "H + M + L": "Controlled Ambient",
                            "REF": "Refrigerated", "REF + H": "Refrigerated", "REF + M": "Refrigerated", "REF + L": "Refrigerated",
                            "REF + H + M": "Refrigerated", "REF + H + L": "Refrigerated", "REF + M + L": "Refrigerated",
                            "REF + H + M + L": "Refrigerated"}
            df["TEMPERATURE"] = df["TEMPERATURE"].str.strip()
            df["TEMPERATURE"] = df["TEMPERATURE"].replace(temperatures)
            df.loc[(df["TEMPERATURE"] == "Ambient") & (df["RETURN_TRACKING_NUMBER"] != "N"), "TEMPERATURE"] = "Controlled Ambient"
            df["Cajas (Carton)"] = df["Cajas (Carton)"].fillna(0).astype(int)
            df["RETURN_CANTIDAD"] = df["AMOUNT_OF_BOXES"] - df["Cajas (Carton)"]
            df["RETURN"] = (df["RETURN_CANTIDAD"] > 0) & (df["TEMPERATURE"] != "Ambient")
            df["RETURN_TO_TA"] = False
            df["RETURN_TYPE"] = "NA"
            df.loc[df["RETURN"], "RETURN_TYPE"] = "CREDO"

            shipSchedules = {"8": "08:00:00", "16.3": "16:30:00", "19": "19:00:00"} 
            df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].replace(shipSchedules)
            df["SHIP_TIME_FROM"] = pd.to_datetime(df["SHIP_TIME_FROM"], format='%H:%M:%S', errors='coerce')
            df["SHIP_TIME_TO"] = df["SHIP_TIME_FROM"] + dt.timedelta(minutes=30)
            df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].dt.strftime('%H:%M')
            df["SHIP_TIME_TO"] = df["SHIP_TIME_TO"].dt.strftime('%H:%M')

            df["RETURN_CANTIDAD"] = df["RETURN_CANTIDAD"].astype(int)
            df["DELIVERY_DATE"] = df["DELIVERY_DATE"].astype("datetime64[ns]")
            df["DELIVERY_DATE"] = pd.to_datetime(df["DELIVERY_DATE"], format='%d/%m/%Y', errors='coerce')
            df["DELIVERY_DATE"] = df["DELIVERY_DATE"].dt.strftime('%d/%m/%Y')

            df["PRINT_RETURN_DOCUMENT"] = df["RETURN"]

            return df

    class GPMArgentinaTeam:
        def __init__(self, folder_path: str):
            self.carrierWebpage = CarriersWebpage("Transportes Ambientales", folder_path)

        def getCarrierWebpage(self):
            return self.carrierWebpage

        def getTeamName(self) -> str:
            return "GPM"

        def get_column_rename_type_config_for_sites_table(self) -> (dict, dict):
            columns_names = {} #TODO
            columns_types = {} #TODO
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_sites_table(self, df: pd.DataFrame) -> pd.DataFrame:
            return df #TODO
        
        def get_data_path(self) -> (str, str, str):
            path = os.path.expanduser("~\\Desktop\Automatizacion_Ordenes.xlsx")
            sheet = "Vacio"
            info_sites_sheet = ""
            return path, sheet, info_sites_sheet
        
        def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
            return {}, {}

        def apply_team_specific_changes_for_orders_tables(self, df: pd.DataFrame) -> pd.DataFrame:
            return df #TODO

    class TestTeam:
        def __init__(self, folder_path: str):
            self.carrierWebpage = CarriersWebpage("Transportes Ambientales", folder_path)

        def getCarrierWebpage(self):
            return self.carrierWebpage

        def getTeamName(self) -> str:
            return "Test"

        def get_column_rename_type_config_for_sites_table(self) -> (dict, dict):
            columns_names = {}
            columns_types = {"STUDY": str, "SITE#": str, "TA_ID": str, "DELIVERY_TIME_FROM": dt.datetime, "DELIVERY_TIME_TO": dt.datetime}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_sites_table(self, df: pd.DataFrame) -> pd.DataFrame:
            return df #TODO
        
        def get_data_path(self) -> (str, str, str):
            path = os.path.expanduser("~\\OneDrive - Thermo Fisher Scientific\Desktop\Automatizacion_Ordenes.xlsx")
            sheet = "Test" # "Test_5_ordenes"
            info_sites_sheet = "SiteInfo"
            return path, sheet, info_sites_sheet
        
        def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
            columns_names = {}
            columns_types = {"SITE#": str, "RETURN_TO_TA": bool, "SHIP_TIME_TO": str}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_orders_tables(self, df: pd.DataFrame) -> pd.DataFrame:
            df["SHIP_TIME_TO"] = ""
            df["SHIP_TIME_FROM"] = pd.to_datetime(df["SHIP_TIME_FROM"], format='%H:%M:%S', errors='coerce')
            df["SHIP_TIME_TO"] = df["SHIP_TIME_FROM"] + dt.timedelta(minutes=30)

            df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].dt.strftime('%H:%M')
            df["SHIP_TIME_TO"] = df["SHIP_TIME_TO"].dt.strftime('%H:%M')

            df["RETURN_CANTIDAD"] = df["AMOUNT_OF_BOXES"]

            df["PRINT_RETURN_DOCUMENT"] = True

            df["DELIVERY_DATE"] = df["DELIVERY_DATE"].astype("datetime64[ns]")
            df["DELIVERY_DATE"] = pd.to_datetime(df["DELIVERY_DATE"], format='%d/%m/%Y', errors='coerce')
            df["DELIVERY_DATE"] = df["DELIVERY_DATE"].dt.strftime('%d/%m/%Y')

            return df

def main():
    ctk.set_appearance_mode("dark")
    app = MyUserForm()
    app.mainloop()

if __name__ == "__main__":
    main()