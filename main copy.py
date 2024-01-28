from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import datetime as dt
import tkinter as tk
import customtkinter as ctk
from tkcalendar import Calendar
from PIL import Image, ImageTk
import win32print
import os, time
from PyPDF2 import PdfReader

class Browser(object):
    def __init__(self, folder_path: str):
        """
        Class constructor for Browser 

        Args:
            folder_path (str): folder path to download files

        Attributes:
            self.driver (webdriver): selenium self.driver
        """
        chrome_options = webdriver.ChromeOptions() #Options()
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

class OrderProcessor:
    def __init__(self, folder_path : str):
        """
        Class constructor

        Args:
            self.driver (webdriver): selenium self.driver

        Attributes:
            self.driver (webdriver): selenium self.driver
            wait (WebDriverWait): selenium wait
        """
        self.folder_path = folder_path
        self.browser = Browser(folder_path)
        self.driver = self.browser.driver
        self.wait = WebDriverWait(self.driver, 10)

    def log_in_TA_website(self) -> bool:
        """
        Logs in TA website

        Args:
            driver (webdriver): selenium driver
        """
        self.driver.get("https://sgi.tanet.com.ar/sgi/index.php")
        # Wait for the login webpage to load
        self.wait = WebDriverWait(self.driver, 30)

        try:
            # Wait user input their credentials 
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/div[2]/div[1]/table/tbody/tr[1]/td")))
            self.wait = WebDriverWait(self.driver, 10)
            return True
        except:
            self.wait = WebDriverWait(self.driver, 10)
            return False

    def process_all_shipping_orders(self, df: pd.DataFrame):
        """
        Process all orders in the table

        Args:
            self.driver (webdriver): selenium self.driver
            df (DataFrame): orders table
        """
        for index, row in df.iterrows():
            if df.loc[index, "TRACKING_NUMBER"] != "":
                continue

            error = row['SHIP_DATE'] == "" or row['SHIP_TIME_FROM'] == "" or row['SHIP_TIME_TO'] == ""
            error = error or row['DELIVERY_DATE'] == "" or row['DELIVERY_TIME_FROM'] == "" or row['DELIVERY_TIME_TO'] == ""
            error = error or row['AMOUNT_OF_BOXES'] == 0 or row['TA_ID'] == ""

            if error:
                continue
            
            tracking_number, return_tracking_number = self.process_shipping_order(
                row["TA_ID"], row["SYSTEM_NUMBER"], row["IVRS_NUMBER"],
                row["SHIP_DATE"], row["SHIP_TIME_FROM"], row["SHIP_TIME_TO"],
                row["DELIVERY_DATE"], row["DELIVERY_TIME_FROM"], row["DELIVERY_TIME_TO"],
                row["TYPE_OF_MATERIAL"], row["TEMPERATURE"],
                row["CONTACTS"], row["AMOUNT_OF_BOXES"],
                row["RETURN"], row["RETURN_TO_TA"], row["RETURN_TYPE"], row["RETURN_CANTIDAD"], row["PRINT_RETURN_DOCUMENT"]
            )

            df.loc[index, "TRACKING_NUMBER"] = tracking_number
            df.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number

    def process_shipping_order(self, TA_ID: int, system_number: str, ivrs_number: str, 
                               ship_date: str, ship_time_from: str, ship_time_to: str, 
                               delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                               type_of_material: str, temperature: str, 
                               contacts: str, amount_of_boxes: int, 
                               return_: bool, return_to_TA: bool, type_of_return: str, amount_of_boxes_to_return: int, print_return_document: bool) -> (str, str):
        """
        - Process an order by completing the TA form
        - Creates a return order if necessary
        - Prints the order documents

        Args:
            self.driver (webdriver): selenium self.driver
            TA_ID (int): Site ID on TA website
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
            return_to_TA (bool): if True, the return order is sent to TA depot
            amount_of_boxes_to_return (int): number of boxes to return

        Returns:
            str: tracking number
            str: return tracking number
        """
        
        tracking_number, return_tracking_number = "", ""

        try:
            url = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion={TA_ID}"
            reference = f"{system_number} {ivrs_number}"[:50]

            transit_days = max(dt.datetime.strptime(delivery_date, '%d/%m/%Y') - dt.datetime.strptime(ship_date, '%d/%m/%Y'), dt.timedelta(days=1))
            return_delivery_date = dt.datetime.strptime(delivery_date, '%d/%m/%Y') + transit_days
            return_delivery_date += dt.timedelta(days=2) if return_delivery_date.weekday() >= 5 else dt.timedelta(days=0) # Add 2 days if the return delivery date is on a weekend

            delivery_date = dt.datetime.strptime(delivery_date, '%d/%m/%Y').strftime('%d/%m/%Y')
            return_delivery_date = return_delivery_date.strftime('%d/%m/%Y')

            tracking_number = self.complete_shipping_order_form(
                url, reference, 
                ship_date, ship_time_from, ship_time_to,
                delivery_date, delivery_time_from, delivery_time_to,
                type_of_material, temperature, contacts, amount_of_boxes)

            if tracking_number == "":
                return  "", ""

            if return_:
                url_return = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarRetorno+idsp={tracking_number[:7]}&idubicacion={TA_ID}"
                url_return += "&returnToTa=true" if return_to_TA else ""

                reference_return = f"{reference} RET {tracking_number}"[:50]

                return_tracking_number = self.complete_shipping_order_return_form(
                    url_return, reference_return, 
                    return_delivery_date, "9", "17",
                    type_of_return, "Personal de FCS", amount_of_boxes_to_return
                )
            else:
                return_tracking_number = ""

            url_guias = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id={tracking_number[:7]}&idservicio={tracking_number[:7]}&copies=4"
            self.print_webpage(url_guias)

            self.driver.implicitly_wait(5)

            url_rotulo = f"https://sgi.tanet.com.ar/sgi/srv.RotuloPdf.emitir+id={tracking_number[:7]}&idservicio={tracking_number[:7]}"
            self.print_webpage(url_rotulo)

            if print_return_document:
                self.driver.implicitly_wait(5)
                url_guias_return = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id={return_tracking_number[:7]}&idservicio={return_tracking_number[:7]}&copies=1"
                self.print_webpage(url_guias_return)

            self.driver.implicitly_wait(5)

        except Exception as e:
            print(f"Error processing order: {e}")
            print(f"Order: {system_number} {ivrs_number}")
            return "", ""

        return tracking_number, return_tracking_number

    def complete_shipping_order_form(self, url: str, reference: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        """
        Completes the TA form

        Args:
            self.driver (webdriver): selenium self.driver
            url (str): TA url
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
        
        tracking_number = ""

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
            if temperature == "Ambiente": it_temperature = 0
            elif temperature == "Ambiente Controlado": it_temperature = 1
            elif temperature == "Refrigerado": it_temperature = 2
            elif temperature == "Congelado": it_temperature = 3
            else: it_temperature = 4

            for i in range(0, it_temperature):
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[2]").send_keys(Keys.DOWN)

            observaciones_textbox = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/textarea")
            comments = ""
            for line in observaciones_textbox.text.split("\n"):
                if "Dias y horarios de entrega:" != line[:27]:
                    comments += line + "\n"
            
            if temperature != "Ambiente":
                comments += "***ENVÍO CON CAJA CREDO. EL COURIER AGUARDARÁ QUE EL CENTRO ALMACENE LA MEDICACIÓN Y RETORNE EL EMBALAJE***"
            
            observaciones_textbox.clear()
            observaciones_textbox.send_keys(comments)
            
            if contacts != "" and contacts != "NA":
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").clear()
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").send_keys(contacts)
            
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input").send_keys(amount_of_boxes)
            
            self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/button").click()

            # Wait for the webpage to load and get the tracking number
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))
            
            self.driver.implicitly_wait(5)

            return self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:15]
            
        except Exception as e:
            print(f"Error completing shipping order form: {e}")
            print(f"Order: {reference}")

        return tracking_number

    def complete_shipping_order_return_form(self, url_return: str, reference_return: str,
                                            delivery_date: str, return_time_from: str,
                                            return_time_to: str, type_of_return: str,
                                            contacts: str, amount_of_boxes_to_return: int) -> str:
        """
        Completes the TA return form

        Args:
            self.driver (webdriver): selenium self.driver
            url_return (str): TA return url
            reference_return (str): return order reference
            return_delivery_date (str): return delivery date
            return_time_from (str): return delivery time (from)
            return_time_to (str): return delivery time (to)
            type_of_return (str): type of return
            contact_person (str): contact person
            amount_of_boxes_to_return (int): number of boxes to return

        Returns:
            str: return tracking number
        """

        return_tracking_number = ""

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
            
            return_tracking_number = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:15]
        except Exception as e:
            print(f"Error completing shipping order return form: {e}")
            print(f"Order: {reference_return}")
            return return_tracking_number

        return return_tracking_number

    def print_webpage(self, url: str):
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
            print(f"Error printing TA documents: {e}")

    def update_shipping_order_table(self, df: pd.DataFrame, path: str, sheet: str):
        """
        Updates orders table

        Args:
            df (DataFrame): orders table
            path (str): excel file path
            sheet (str): excel sheet name
        """

        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet, index=False)

    def generate_shipping_report(self, df:pd.DataFrame) -> pd.DataFrame:
        """
        Process all orders in the table
        
        Variables:
            self.driver (webdriver): selenium self.driver
            wait (WebDriverWait): selenium wait
            team (str): team to process
            df_path (str): excel file path
            sheet (str): excel sheet name
            df (DataFrame): orders table

        Args:
            self.driver (webdriver): selenium self.driver
            df (DataFrame): orders table

        Returns:
            DataFrame: orders table with tracking numbers
        """
        if self.log_in_TA_website():
            print("Logged in")
        else:
            print("Not logged in")
            return
            
        self.process_all_shipping_orders(df)

        if not df.empty:
            dataframe_name = "orders_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            df.to_excel(self.folder_path + "\\" + dataframe_name, index=False)
        else:
            print("Empty DataFrame")
        
        time.sleep(2) # Wait for the download to finish

        self.driver.quit()

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

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        self.colors = self.Chroma()

        self.configure(bg=self.colors.body_color)

        """self.canvas = tk.Canvas(
            self, #tk.Tk(),
            bg = self.colors.body_color,
            height = screen_height,
            width = screen_width,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )
        
        self.canvas.place(x = 0, y = 0)
        
        self.canvas.create_rectangle( # Topbar
            0.0,
            0.0,
            screen_width,
            screen_height * 0.3,
            fill=self.colors.primary_color,
            outline="")
        
        self.canvas.create_rectangle( # Body
            0.0,
            screen_height * 0.3,
            screen_width,
            screen_height * 0.7,
            fill=self.colors.body_color,
            outline="")
        
        self.canvas.create_rectangle( # Top
            screen_width * 0.8,
            0.0,
            screen_width * 0.2,
            screen_height * 0.3,
            fill="red",
            outline="")
        
        self.canvas.create_rectangle(
            screen_width * 0.6,
            0.0,
            screen_width * 0.2,
            screen_height * 0.3,
            fill="green",
            outline="")
        
        self.canvas.create_rectangle(
            screen_width * 0.4,
            0.0,
            screen_width * 0.2,
            screen_height * 0.3,
            fill="blue",
            outline="")
        
        self.canvas.create_rectangle(
            screen_width * 0.2,
            0.0,
            screen_width * 0.2,
            screen_height * 0.3,
            fill="yellow",
            outline="")
        
        self.canvas.create_rectangle(
            screen_width*0.0,
            0.0,
            screen_width * 0.2,
            screen_height * 0.3,
            fill="pink",
            outline="")"""
        
        # Top Frame
        frame_top = ctk.CTkFrame(self)
        frame_top.pack(pady=0, side = tk.TOP, fill=tk.BOTH)

        frame_top_2 = ctk.CTkFrame(self)
        frame_top_2.pack(pady=0, after = frame_top, side = tk.LEFT)

        frame_top_3 = ctk.CTkFrame(self)
        frame_top_3.pack(pady=0, after = frame_top_2, side = tk.LEFT)

        frame_top_4 = ctk.CTkFrame(self)
        frame_top_4.pack(pady=0, after = frame_top_3, side = tk.LEFT)

        frame_top_5 = ctk.CTkFrame(self)
        frame_top_5.pack(pady=0, after = frame_top_4, side = tk.LEFT)

        # Bottom Frame
        frame_bottom = ctk.CTkFrame(self)
        frame_bottom.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, pady=0, anchor = tk.S, after = frame_top)

        # Image banner
        imagen = Image.open("fcs.jpg")
        imagen = imagen.resize((284, 67))
        imagen = ImageTk.PhotoImage(imagen)

        label_banner = tk.Label(frame_top, image=imagen)
        label_banner.image = imagen
        label_banner.pack(padx=10)

        # DatePicker
        self.cal = Calendar(frame_top_2, selectmode='day', locale='en_US', disabledforeground='red',
                cursor="hand2", background=ctk.ThemeManager.theme["CTkFrame"]["fg_color"][1],
                selectbackground=ctk.ThemeManager.theme["CTkButton"]["fg_color"][1], date_pattern='yyyy-MM-dd')

        self.cal.pack(padx=50, pady=0)
        
        # Teams Combobox
        teams_options = ["Lilly", "GPM", "Test", "Test_5_ordenes"]
        self.team_combobox = tk.ttk.Combobox(frame_top_3, values=teams_options, width=20, height=15, font=30)
        self.team_combobox.pack(padx=10)
        self.team_combobox.current(0)

        carrier_options = ["Transportes Ambientales", "NA"]
        self.carrier_combobox = tk.ttk.Combobox(frame_top_3, values=carrier_options, width=20, height=15, font=30)
        self.carrier_combobox.pack(padx=10)
        self.carrier_combobox.current(0)

        # Buttons ###################
        # Load orders button
        loadOrders_btn = ctk.CTkButton(master=frame_top_4, text="Load Orders", command=self.on_loadOrders_btn_click,
                                       width=150, height=50, font=('Calibri', 22, 'bold'))
        loadOrders_btn.pack(padx=10)

        # Process orders button
        processOrders_btn = ctk.CTkButton(master=frame_top_4, text="Process Orders", command=self.on_processOrders_btn_click,
                                        width=150, height=50, font=('Calibri', 22, 'bold'))
        processOrders_btn.pack(padx=10)

        exporttable_btn = ctk.CTkButton(master=frame_top_5, text="Export Table", command=print("Export Table"),
                                       width=150, height=50, font=('Calibri', 22, 'bold'))
        exporttable_btn.pack(padx=10)

        # Process orders button
        configuration_btn = ctk.CTkButton(master=frame_top_5, text="Configuration", command=print("Configuration"),
                                        width=150, height=50, font=('Calibri', 22, 'bold'))
        configuration_btn.pack(padx=10)

        # Orders self.treeview (table)
        style = tk.ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 13)) # Font of the body
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 13,'bold')) # Font of the headings
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) # Remove the borders
        
        columns_df = ['#', 'SYSTEM_NUMBER', 'IVRS_NUMBER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO', 
                'TYPE_OF_MATERIAL', 'TEMPERATURE', 'AMOUNT_OF_BOXES',
                'RETURN', 'RETURN_TO_TA', 'RETURN_TYPE', 'RETURN_CANTIDAD',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT', 'CONTACTS', 'TA_ID']
        self.treeview = tk.ttk.Treeview(frame_bottom, columns=columns_df, show='headings', style="mystyle.Treeview")
        self.treeview.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.treeview.tag_configure('odd', background='#E8E8E8')
        self.treeview.tag_configure('odd_done', background='#C6E0B4')
        self.treeview.tag_configure('odd_error', background='#FFC7CE')

        self.treeview.tag_configure('even', background='#DFDFDF')
        self.treeview.tag_configure('even_done', background='#A9D08E')
        self.treeview.tag_configure('even_error', background='#FFA7BB')

        # Treeview columns headings and columns width
        self.treeview.column("#0", width=0, stretch=tk.NO)  # Hide the first column
        for col in columns_df:
            self.treeview.heading(col, text=col)
            self.treeview.column(col, anchor=tk.W, width=int(self.winfo_screenwidth() * 0.7 * 0.3))
        self.treeview.column("#", anchor=tk.W, width=int(self.winfo_screenwidth() * 0.7 * 0.05))

        # Scrollbars ###################
        # Vertical scrollbar
        y_scrollbar = tk.ttk.Scrollbar(frame_bottom, orient=tk.VERTICAL, command=self.treeview.yview)
        y_scrollbar.pack(after = self.treeview, side=tk.RIGHT, fill=tk.Y)
        self.treeview.configure(yscrollcommand=y_scrollbar.set)
        #self.treeview.configure(command=self.treeview.yview)
        #y_scrollbar.grid(row=0, column=1, rowspan=10, sticky=NS)
        #y_scrollbar.place(x=self.treeview.winfo_x() + self.treeview.winfo_width(), y=self.treeview.winfo_y(), height=self.treeview.winfo_height())

        # Horizontal scrollbar
        x_scrollbar = tk.ttk.Scrollbar(frame_bottom, orient=tk.HORIZONTAL, command=self.treeview.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.treeview.configure(xscrollcommand=x_scrollbar.set)

        
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
        selected_date = dt.datetime.strptime(self.cal.get_date(), '%Y-%m-%d')
        selected_team = self.team_combobox.get()

        self.df = self.generate_shipping_order_tables(selected_team, selected_date)

        for item in self.treeview.get_children():
            self.treeview.delete(item)
        
        parity = False
        i = 1
        for index, row in self.df.iterrows():
            r = [i] + list(row)

            error = row['SHIP_DATE'] == "" or row['SHIP_TIME_FROM'] == "" or row['SHIP_TIME_TO'] == ""
            error = error or row['DELIVERY_DATE'] == "" or row['DELIVERY_TIME_FROM'] == "" or row['DELIVERY_TIME_TO'] == ""
            error = error or row['AMOUNT_OF_BOXES'] == 0 or row['TA_ID'] == ""

            order_done = row['TRACKING_NUMBER'] != ""
            tag_color = 'odd' if parity else 'even'
            tag_color += "_done" if order_done else ""
            tag_color += ("_error" if error else "") if not order_done else ""
            self.treeview.insert("", "end", values=r, tags=(tag_color))
            parity = not parity
            i += 1
        
        self.treeview.update()

    def on_processOrders_btn_click(self):
        """
        Button to process all orders in the table
        """
        selected_date = self.cal.get_date()
        selected_date = selected_date.replace("/", "_")
        selected_team = self.team_combobox.get()

        folder_path = os.path.expanduser("~\\Downloads") + "\\" + "TA_Form_AutoLoad" + "\\" + selected_team

        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        folder_path += "\\" + selected_date

        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        self.df = OrderProcessor(folder_path).generate_shipping_report(self.df)

        for item in self.treeview.get_children():
            self.treeview.delete(item)
        
        parity = False
        i = 1
        for index, row in self.df.iterrows():
            r = [i] + list(row)
            order_done = row['TRACKING_NUMBER'] != ""
            tag_color = 'odd' if parity else 'even'
            tag_color += "_done" if order_done else ""
            self.treeview.insert("", "end", values=r, tags=(tag_color))
            parity = not parity
            i += 1
        
        self.treeview.update()

    def load_paths(self, team: str) -> (str, str, str):
        """
        Loads excel files paths

        Args:
            team (str): team to process

        Returns:
            str: excel file path
            str: excel sheet name
            str: excel sheet name with sites info
        """
        
        path, sheet, info_sites_sheet = "", "", ""

        if team == "Lilly": # Specific cases for Lilly team
            path = os.path.expanduser("~\\Thermo Fisher Scientific\Power BI Lilly Argentina - General\Share Point Lilly Argentina.xlsx")
            sheet = "Shipments"
            info_sites_sheet = "Dias y Destinos"
        elif team == "GPM": # Specific cases for GPM team
            path = os.path.expanduser("~\\Desktop\Automatizacion_Ordenes.xlsx")
            sheet = "Vacio"
            info_sites_sheet = ""
        elif team == "Test": # Specific cases for tests
            path = os.path.expanduser("~\\OneDrive - Thermo Fisher Scientific\Desktop\Automatizacion_Ordenes.xlsx")
            sheet = "Test"
            info_sites_sheet = "SiteInfo"
        elif team == "Test_5_ordenes": # Specific cases for tests_5_ordenes
            path = os.path.expanduser("~\\OneDrive - Thermo Fisher Scientific\Desktop\Automatizacion_Ordenes.xlsx")
            sheet = "Test_5_ordenes"
            info_sites_sheet = "SiteInfo"
        
        return path, sheet, info_sites_sheet

    def load_shipping_order_table(self, date: dt.datetime, team: str,  path: str, sheet: str) -> pd.DataFrame:
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
        if team == "Lilly": # Specific cases for Lilly team
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

        elif team == "GPM": # Specific cases for GPM team
            columns_names = {}
            columns_types = {}

        elif team == "Test" or team == "Test_5_ordenes": # Specific cases for tests
            columns_names = {}
            columns_types = {"SITE#": str, "RETURN_TO_TA": bool, "SHIP_TIME_TO": str}

        df = pd.read_excel(path, sheet_name=sheet, header=0, dtype=columns_types)
        df.rename(columns=columns_names, inplace=True)
        
        df = df[df["SHIP_DATE"] == date]
        
        df["SHIP_DATE"] = df["SHIP_DATE"].astype("datetime64[ns]")
        df["SHIP_DATE"] = pd.to_datetime(df["SHIP_DATE"], format='%d/%m/%Y', errors='coerce')
        #df["SHIP_DATE"] = df["SHIP_DATE"].dt.date

        df["SHIP_DATE"] = df["SHIP_DATE"].dt.strftime('%d/%m/%Y')

        df["SITE#"] = df["SITE#"].astype(object)
        df["AMOUNT_OF_BOXES"] = df["AMOUNT_OF_BOXES"].fillna(0).astype(int)

        if team == "Lilly": # Specific cases for Lilly team
            df["TYPE_OF_MATERIAL"] = "Medicine"
            
            temperatures = {"L": "Ambiente",
                            "M": "Ambiente Controlado", "M + L": "Ambiente Controlado",
                            "H": "Ambiente Controlado", "H + M": "Ambiente Controlado", "H + L": "Ambiente Controlado", "H + M + L": "Ambiente Controlado",
                            "REF": "Refrigerado", "REF + H": "Refrigerado", "REF + M": "Refrigerado", "REF + L": "Refrigerado",
                            "REF + H + M": "Refrigerado", "REF + H + L": "Refrigerado", "REF + M + L": "Refrigerado",
                            "REF + H + M + L": "Refrigerado"}
            df["TEMPERATURE"] = df["TEMPERATURE"].str.strip()
            df["TEMPERATURE"] = df["TEMPERATURE"].replace(temperatures)
            df.loc[(df["TEMPERATURE"] == "Ambiente") & (df["RETURN_TRACKING_NUMBER"] != "NA"), "TEMPERATURE"] = "Ambiente Controlado"
            df["Cajas (Carton)"] = df["Cajas (Carton)"].fillna(0).astype(int)
            df["RETURN_CANTIDAD"] = df["AMOUNT_OF_BOXES"] - df["Cajas (Carton)"]
            df["RETURN"] = (df["RETURN_CANTIDAD"] > 0) & (df["TEMPERATURE"] != "Ambiente")
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
    
        elif team == "GPM": # Specific cases for GPM team
            df = df #TODO

        elif team == "Test" or team == "Test_5_ordenes": # Specific cases for tests
            df["SHIP_TIME_TO"] = ""
            df["SHIP_TIME_FROM"] = pd.to_datetime(df["SHIP_TIME_FROM"], format='%H:%M:%S', errors='coerce')
            df["SHIP_TIME_TO"] = df["SHIP_TIME_FROM"] + dt.timedelta(minutes=30)

            df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].dt.strftime('%H:%M')
            df["SHIP_TIME_TO"] = df["SHIP_TIME_TO"].dt.strftime('%H:%M')

            df["RETURN_CANTIDAD"] = df["AMOUNT_OF_BOXES"]

            df["PRINT_RETURN_DOCUMENT"] = False

        df.fillna('', inplace=True)
        
        return df[['SYSTEM_NUMBER', 'IVRS_NUMBER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'TYPE_OF_MATERIAL', 
                'TEMPERATURE', 'AMOUNT_OF_BOXES', 'RETURN', 
                'RETURN_TO_TA', 'RETURN_TYPE', 'RETURN_CANTIDAD',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT']]

    def load_table_info_sites(self, team: str, path: str, sheet: str) -> pd.DataFrame:
        """
        Loads sites info table according to team

        Args:
            team (str): team to process
            path (str): excel file path
            sheet (str): excel sheet name

        Returns:
            DataFrame: sites info table
        """
        if team == "Lilly": # Specific cases for Lilly team
            columns_names = {"Protocolo": "STUDY", "Codigo": "TA_ID", "Site": "SITE#",
                            "Horario inicio": "DELIVERY_TIME_FROM", "Horario fin": "DELIVERY_TIME_TO"}
            columns_types = {"Protocolo": str, "Site": str, "Codigo": str, "Horario inicio": str, "Horario fin": str}
        elif team == "GPM": # Specific cases for GPM team
            columns_names = {} #TODO
            columns_types = {} #TODO
        elif team == "Test" or team == "Test_5_ordenes": # Specific cases for tests
            columns_names = {}
            columns_types = {"STUDY": str, "SITE#": str, "TA_ID": str, "DELIVERY_TIME_FROM": dt.datetime, "DELIVERY_TIME_TO": dt.datetime}

        df = pd.read_excel(path, sheet_name=sheet, header=0, dtype=columns_types)
        df.rename(columns=columns_names, inplace=True)
        df.fillna('', inplace=True)

        if team == "Lilly": # Specific cases for Lilly team
            df["CONTACTS"] = "NA"
        elif team == "GPM": # Specific cases for GPM team
            df = df #TODO
        elif team == "Test" or team == "Test_5_ordenes": # Specific cases for tests
            df = df #TODO

        df["DELIVERY_TIME_FROM"] = pd.to_datetime(df["DELIVERY_TIME_FROM"], format='%H:%M:%S', errors='coerce')
        df["DELIVERY_TIME_TO"] = pd.to_datetime(df["DELIVERY_TIME_TO"], format='%H:%M:%S', errors='coerce')
        df["DELIVERY_TIME_FROM"] = df["DELIVERY_TIME_FROM"].dt.strftime('%H:%M')
        df["DELIVERY_TIME_TO"] = df["DELIVERY_TIME_TO"].dt.strftime('%H:%M')

        return df[["STUDY", "SITE#", "TA_ID", "DELIVERY_TIME_FROM", "DELIVERY_TIME_TO", "CONTACTS"]]

    def generate_shipping_order_tables(self, team: str, shipdate: dt.datetime) -> pd.DataFrame:
        """
        Process all orders in the table

        Args:
            team (str): team to process
            shipdate (dt.datetime): date to process

        Returns:
            DataFrame: orders table with standarisized data
        """
        df_path, sheet, info_sites_sheet = self.load_paths(team)
        
        df_orders = self.load_shipping_order_table(shipdate, team, df_path, sheet)

        df_info_sites = self.load_table_info_sites(team, df_path, info_sites_sheet)
        
        df = pd.merge(df_orders, df_info_sites, on=["STUDY", "SITE#"], how="inner")
        
        return df[['SYSTEM_NUMBER', 'IVRS_NUMBER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO', 
                'TYPE_OF_MATERIAL', 'TEMPERATURE', 'AMOUNT_OF_BOXES',
                'RETURN', 'RETURN_TO_TA', 'RETURN_TYPE', 'RETURN_CANTIDAD',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT', 'CONTACTS', 'TA_ID']]

def main():
    ctk.set_appearance_mode("dark")
    app = MyUserForm()
    app.mainloop()

if __name__ == "__main__":
    main()