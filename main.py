from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
import datetime as dt
import os
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkcalendar import Calendar
import customtkinter as ctk
import sys
from PIL import Image, ImageTk

class PrintingBrowser(object):
    def __init__(self):
        chrome_options = Options()
        #chrome_options.add_argument('--headless') # Disable headless mode (do not 
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

        self.driver = webdriver.Chrome(options=chrome_options)

    def get_page_and_print(self, page):
        self.driver.get(page)
        time.sleep(1)
        self.driver.execute_script("window.print = function(msg) {return false;}")

class DatePicker(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self._last_click_event = None
        
        frame_banner = tk.Frame(master)
        frame_banner.pack(side=tk.TOP, fill=tk.X)

        # Cargar la imagen y mostrarla en un Label
        imagen = Image.open("tu_imagen.jpg")  # Reemplaza "tu_imagen.jpg" con la ruta de tu propia imagen
        imagen = imagen.resize((master.winfo_reqwidth(), 100))
        imagen = ImageTk.PhotoImage(imagen)

        label_banner = tk.Label(frame_banner, image=imagen)
        label_banner.image = imagen  # Se requiere mantener una reference a la imagen para que no sea eliminada por el recolector de basura
        label_banner.pack(fill=tk.X)

        frame = ctk.CTkFrame(master)
        frame.pack(fill="both", padx=10, pady=10, expand=True)
        frame_calendario = tk.Frame(frame)
        frame_calendario.pack(side=tk.LEFT, padx=10, pady=10, expand=True, fill=tk.BOTH)

        style = ttk.Style(master)
        style.theme_use("default")

        self.selected_date = dt.datetime.now()

        cal = Calendar(frame_calendario, selectmode='day', locale='en_US', disabledforeground='red',
                cursor="hand2", background=ctk.ThemeManager.theme["CTkFrame"]["fg_color"][1],
                selectbackground=ctk.ThemeManager.theme["CTkButton"]["fg_color"][1])
        cal.pack(fill="both", expand=True, padx=10, pady=10)
        cal.bind("<<DateEntrySelected>>", self._on_click_handler)
        
        self.opcion_var = tk.StringVar(None, "Lilly")
        radio1 = tk.Radiobutton(frame, text="Team Lilly", variable=self.opcion_var, value="Lilly")
        radio2 = tk.Radiobutton(frame, text="Team GPM", variable=self.opcion_var, value="GPM")
        radio3 = tk.Radiobutton(frame, text="Test", variable=self.opcion_var, value="Test")
        radio4 = tk.Radiobutton(frame, text="Test_5_ordenes", variable=self.opcion_var, value="Test_5_ordenes")

        radio1.pack(pady=5, anchor="n")
        radio2.pack(pady=5, anchor="n")
        radio3.pack(pady=5, anchor="n")
        radio4.pack(pady=5, anchor="n")

        def print_sel():
            date.config(text = "Click here to process orders for " + cal.get_date() + " for " + self.opcion_var.get() + " team" )
            self.selected_date = dt.datetime.strptime(cal.get_date(), "%m/%d/%y")

        Button(master, text = "Process orders",
                command = print_sel, font=30).pack(pady = 10, padx=10, side=tk.BOTTOM)
        
        date = Label(master, text = "", font=30)
        date.pack(pady = 20)
        date.bind("<Button-1>", self._on_click_handler)

    def _on_click_handler(self, event):
        if event is None and self._last_click_event is not None:
            self._last_click_event = None
            #self.call_label_top(self._last_click_event)
        elif self._last_click_event is None and event is not None:
            self._last_click_event = event
            self.after(300, self._on_click_handler, None)
        elif event is not None:
            self.call_label_top_double(event)
            self._last_click_event = None

    def call_label_top_double(self, event):
        print("Double click in label")
        generate_shipping_report(self.selected_date, self.opcion_var.get())

def init_driver() -> (webdriver, WebDriverWait):
    """
    Starts selenium driver and selenium wait

    Returns:
        webdriver: selenium driver
        WebDriverWait: selenium wait
    """
    global browser_that_prints
    browser_that_prints = PrintingBrowser()
    driver = browser_that_prints.driver
    wait = WebDriverWait(driver, 10)
    return driver, wait

def load_paths(team: str) -> (str, str, str):
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

    if team == "Lilly":
        #path = r"C:/Users/inaki.costa/Thermo Fisher Scientific/Power BI Lilly Argentina - General/Share Point Lilly Argentina.xlsx"
        path = os.path.expanduser("~\\Thermo Fisher Scientific\Power BI Lilly Argentina - General\Share Point Lilly Argentina.xlsx")
        sheet = "Shipments"
        info_sites_sheet = "Dias y Destinos"
    elif team == "GPM":
        path = r"C:/Users/inaki.costa/OneDrive - Thermo Fisher Scientific/Desktop/Automatizacion_Ordenes.xlsx"
        sheet = "Vacio"
        info_sites_sheet = ""
    elif team == "Test":
        path = r"C:/Users/inaki.costa/OneDrive - Thermo Fisher Scientific/Desktop/Automatizacion_Ordenes.xlsx"
        sheet = "Test"
        info_sites_sheet = "SiteInfo"
    elif team == "Test_5_ordenes":
        path = r"C:/Users/inaki.costa/OneDrive - Thermo Fisher Scientific/Desktop/Automatizacion_Ordenes.xlsx"
        sheet = "Test_5_ordenes"
        info_sites_sheet = "SiteInfo"
    
    return path, sheet, info_sites_sheet

def log_in_TA_website(driver) -> bool:
    """
    Logs in TA website

    Args:
        driver (webdriver): selenium driver
    """
    driver.get("https://sgi.tanet.com.ar/sgi/index.php")
    # Wait for the login webpage to load
    wait = WebDriverWait(driver, 30)

    try:
        # Wait user input their credentials 
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/div[2]/div[1]/table/tbody/tr[1]/td")))
        wait = WebDriverWait(driver, 10)
        return True
    except:
        return False

def load_shipping_order_table(date : dt.datetime, team: str,  path: str, sheet: str ) -> pd.DataFrame:
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
    if team == "Lilly":
        columns_names = {"CT-WIN": "SYSTEM_NUMBER", "IVRS": "IVRS_NUMBER",
                        "Trial Alias": "STUDY", "Site ": "SITE#",
                        "Order received": "ENTER DATE", "Ship date": "SHIP_DATE",
                        "Horario de Despacho": "SHIP_TIME_FROM",  
                        "Dia de entrega": "DELIVERY_DATE", "Destination": "DESTINATION",
                        "CONDICION": "TEMPERATURE", "TT4": "AMOUNT_OF_BOXES",  
                        "AWB": "TRACKING_NUMBER", "Shipper return AWB": "RETURN_TRACKING_NUMBER"}
        
        columns_types = {"CT-WIN": int, "IVRS": str, 
                        "Trial Alias": str, "Site ": str, 
                        "Order received": str, "Ship date": dt.datetime,
                        "Horario de Despacho": str,
                        "Dia de entrega": dt.datetime, "Destination": str,
                        "CONDICION": str, "TT4": str, 
                        "AWB": str, "Shipper return AWB": str}

    elif team == "GPM":
        columns_names = {}
        columns_types = {}

    elif team == "Test" or team == "Test_5_ordenes":
        columns_names = {}
        columns_types = {"SITE#": str, "RETURN_TO_TA": bool, "SHIP_TIME_TO": str}

    df = pd.read_excel(path, sheet_name=sheet, header=0, dtype=columns_types)
    df.rename(columns=columns_names, inplace=True)
    df = df[df["SHIP_DATE"] == date]
    
    df["SHIP_DATE"] = df["SHIP_DATE"].dt.strftime('%d%m%y')
    df["SITE#"] = df["SITE#"].astype(object)
    df["AMOUNT_OF_BOXES"] = df["AMOUNT_OF_BOXES"].fillna(0).astype(int)

    if team == "Lilly": # Specific cases for Lilly team
        df["TYPE_OF_MATERIAL"] = "Medicacion"
        
        temperaturas = {"L": "Ambiente",
                        "M": "Ambiente Controlado", "M + L": "Ambiente Controlado",
                        "H": "Ambiente Controlado", "H + M": "Ambiente Controlado", "H + L": "Ambiente Controlado", "H + M + L": "Ambiente Controlado",
                        "REF": "Refrigerado", "REF + H": "Refrigerado", "REF + M": "Refrigerado", "REF + L": "Refrigerado",
                        "REF + H + M": "Refrigerado", "REF + H + L": "Refrigerado", "REF + M + L": "Refrigerado",
                        "REF + H + M + L": "Refrigerado"}
        df["TEMPERATURE"] = df["TEMPERATURE"].replace(temperaturas)
        df[(df["TEMPERATURE"] == "Ambiente") & (df["RETURN_TRACKING_NUMBER"] != "NA")]["TEMPERATURE"] = "Ambiente Controlado"
        df["Cajas (Carton)"] = df["Cajas (Carton)"].fillna(0).astype(int)
        df["RETURN_CANTIDAD"] = df["AMOUNT_OF_BOXES"] - df["Cajas (Carton)"]
        df["RETURN"] = (df["RETURN_CANTIDAD"] > 0) & (df["TEMPERATURE"] != "Ambiente")
        df["RETURN_TO_TA"] = False
        df["RETURN_TYPE"] = "CREDO"

        horariosDeDespacho = {"8": "08:00:00", "16.3": "16:30:00", "19": "19:00:00"} 
        df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].replace(horariosDeDespacho)
        df["SHIP_TIME_FROM"] = pd.to_datetime(df["SHIP_TIME_FROM"], format='%H:%M:%S', errors='coerce')
        df["SHIP_TIME_TO"] = df["SHIP_TIME_FROM"] + dt.timedelta(minutes=30)
        df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].dt.strftime('%H:%M')
        df["SHIP_TIME_TO"] = df["SHIP_TIME_TO"].dt.strftime('%H:%M')

        df["RETURN_CANTIDAD"] = df["RETURN_CANTIDAD"].astype(int)
        df["DELIVERY_DATE"] = pd.to_datetime(df["DELIVERY_DATE"], format='%d%m%Y', errors='coerce')
 
    elif team == "GPM": # Specific cases for GPM team
        df = df #TODO
    
    elif team == "Test" or team == "Test_5_ordenes": # Specific cases for tests
        df["SHIP_TIME_TO"] = ""
        df["SHIP_TIME_FROM"] = pd.to_datetime(df["SHIP_TIME_FROM"], format='%H:%M:%S', errors='coerce')
        df["SHIP_TIME_TO"] = df["SHIP_TIME_FROM"] + dt.timedelta(minutes=30)

        df["SHIP_TIME_FROM"] = df["SHIP_TIME_FROM"].dt.strftime('%H:%M')
        df["SHIP_TIME_TO"] = df["SHIP_TIME_TO"].dt.strftime('%H:%M')

        df["RETURN_CANTIDAD"] = df["AMOUNT_OF_BOXES"]

    df.fillna('', inplace=True)
    
    return df[['SYSTEM_NUMBER', 'IVRS_NUMBER', 'STUDY', 'SITE#',
               'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
               'DELIVERY_DATE', 'TYPE_OF_MATERIAL', 
               'TEMPERATURE', 'AMOUNT_OF_BOXES', 'RETURN', 
               'RETURN_TO_TA', 'RETURN_TYPE', 'RETURN_CANTIDAD',
               'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER']]

def load_table_info_sites(team: str, path: str, sheet: str) -> pd.DataFrame:
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

def complete_shipping_order_form(driver, url: str, reference: str,
                                    ship_date: str, ship_time_from: str, ship_time_to: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str :
    """
    Completes the order form

    Args:
        driver (webdriver): selenium driver
        url (str): page url
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
        str: tracking number // waybill number
    """
    try:
        driver.get(url)
        
        # Wait for the webpage to load
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/input")))
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").send_keys(reference)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[1]").send_keys(ship_date)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[2]").send_keys(ship_time_from)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[3]").send_keys(ship_time_to)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[1]").send_keys(delivery_date)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[2]").send_keys(delivery_time_from)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[3]").send_keys(delivery_time_to)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/input").send_keys("FCS")
        suggestions_container = wait.until(EC.presence_of_element_located((By.ID, "suggest_nomDomOri_list")))

        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/div/div/div[1]/table/tbody")))

        # Finds all suggested buttons inside the suggestions container
        suggested_buttons = suggestions_container.find_elements(By.CLASS_NAME, "suggest")
        # Loops through the suggested buttons and finds the one that contains the text "Fisher Clinical Services FCS"
        for button in suggested_buttons:
            button_text = button.text.strip()
            if "Fisher Clinical Services FCS" in button_text:
                button.click()
                break
        
        # Selects the material type
        if type_of_material == "Medicacion": it_tipo_material = 3
        elif type_of_material == "Material": it_tipo_material = 5
        elif type_of_material == "Equipo": it_tipo_material = 7
        else: it_tipo_material = 8

        for i in range(0, it_tipo_material):
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[1]").send_keys(Keys.DOWN)
        

        # Selects the temperature
        if temperature == "Ambiente": it_temperatura = 0
        elif temperature == "Ambiente Controlado": it_temperatura = 1
        elif temperature == "Refrigerado": it_temperatura = 2
        elif temperature == "Congelado": it_temperatura = 3
        else: it_temperatura = 4

        for i in range(0, it_temperatura):
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[2]").send_keys(Keys.DOWN)

        if temperature != "Ambiente":
            leyenda = "\n ***ENVÍO CON CAJA CREDO. EL COURIER AGUARDARÁ QUE EL CENTRO ALMACENE LA MEDICACIÓN Y RETORNE EL EMBALAJE***"
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/textarea").send_keys(leyenda)
        
        if contacts != "" and contacts != "NA":
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").clear()
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").send_keys(contacts)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input").send_keys(amount_of_boxes)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/button").click()

        # Wait for the webpage to load and get the tracking number
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))
    
        return driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:15]
    except:
        return ""

def complete_shipping_order_return_form(driver, url: str, reference: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_return: str, contacts: str, amount_of_boxes: int) -> str:
    """
    Completes the return form

    Args:
        driver (webdriver): selenium driver
        url (str): page url
        reference (str): order reference
        delivery_date (str): delivery date
        delivery_time_from (str): delivery time (from)
        delivery_time_to (str): delivery time (to)
        type_of_return (str): return type
        contacts (str): contacts
        amount_of_boxes (int): number of boxes

    Returns:
        str: return tracking number // return waybill number
    """
    try:
        # Load the webpage
        driver.get(url)

        # Wait for the webpage to load
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select")))
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").clear()
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").send_keys(reference)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[1]").send_keys(delivery_date)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[2]").send_keys(delivery_time_from)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[3]").send_keys(delivery_time_to)

        # Selects the return type
        if type_of_return == "CREDO": it_tipo_retorno = 1
        elif type_of_return == "DATALOGGER": it_tipo_retorno = 2
        else: it_tipo_retorno = 3
        for i in range(0, it_tipo_retorno):
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select").send_keys(Keys.DOWN)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input[1]").send_keys(contacts)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").clear()
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").send_keys(amount_of_boxes)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[14]/td/button").click()
        
        # Wait for the webpage to load and get the return tracking number
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))
        
        return driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:15]
    except:
        return ""

def print_TA_documents(driver, url: str):
    """
    Prints order documents

    Args:
        driver (webdriver): selenium driver
        url (str): page url
    """
    
    time.sleep(2)
    browser_that_prints.get_page_and_print(url)
    #driver.execute_script("window.print = function(msg) {return false;};")
    
    #driver.execute_script("setTimeout(function() { window.print(); }, 0);");
    #base64code = driver.print_page(print_options)
    
    #driver.execute_script("window.print = function(msg) {return false;};")
    
    #time.sleep(2)
    #driver.executeScript("window.confirm = function(msg){return false;};");
    #driver.find_element(By.XPATH, "/html/body/div/div[1]/div[1]/div/div[3]/div/button[1]").click()

def process_shipping_order(driver, TA_ID:int, system_number:int, ivrs_number:str,
                        ship_date: str, ship_time_from: str, ship_time_to: str,
                        delivery_date: dt.datetime, delivery_time_from: str, delivery_time_to: str,
                        type_of_material: str, temperature: str,
                        contacts: str, amount_of_boxes: int,
                        return_: bool, return_to_TA: bool, type_of_return: str , amount_of_boxes_to_return: int) -> str:
    """
    - Process an order by completing the TA form
    - Creates a return order if necessary
    - Prints the order documents

    Args:
        driver (webdriver): selenium driver
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
        url = "https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion=" + str(TA_ID)
        reference = str(system_number) + " " + ivrs_number
        reference = reference[:50]

        driver.get(url)
        return_delivery_date = delivery_date + dt.timedelta(days=1)
        
        delivery_date = delivery_date.strftime('%d%m%Y')
        return_delivery_date = return_delivery_date.strftime('%d%m%Y')

        tracking_number = complete_shipping_order_form(driver,
                                    url, reference, 
                                    ship_date, ship_time_from, ship_time_to,
                                    str(delivery_date) , delivery_time_from, delivery_time_to,
                                    type_of_material, temperature,
                                    contacts, amount_of_boxes)
        
        if tracking_number == "": return "", ""

        if return_:
            url_return = "https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarRetorno+idsp=" + tracking_number[:7] + "&idubicacion=" + str(TA_ID)
            url_return += "&returnToTa=true" if return_to_TA else ""

            reference_return = reference + " RET " + tracking_number
            reference_return = reference_return[:50]

            driver.get(url_return)
            return_tracking_number = complete_shipping_order_return_form(driver, url_return, 
                                                                            reference_return,
                                                                            str(return_delivery_date), "9", "16",
                                                                            type_of_return, "Personal de FCS", 
                                                                            amount_of_boxes_to_return)
        else:
            return_tracking_number = ""

        url_guias = "https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id=" + str(tracking_number)[:7] + "&idservicio=" + str(tracking_number)[:7] + "&copies=4"
        driver.get(url_guias)
        print_TA_documents(driver, url_guias)

        url_rotulo = "https://sgi.tanet.com.ar/sgi/srv.RotuloPdf.emitir+id=" + str(tracking_number)[:7] + "&idservicio=" + str(tracking_number)[:7] + "&copies=1"
        driver.get(url_rotulo)
        print_TA_documents(driver, url_rotulo)
    
    except:
        return "", ""

    return tracking_number, return_tracking_number

def process_all_shipping_orders(driver, df: pd.DataFrame):
    """
    Process all orders in the table

    Args:
        driver (webdriver): selenium driver
        df (DataFrame): orders table
    """
    
    for index, row in df.iterrows():
        if df.loc[index, "TRACKING_NUMBER"] != "": continue

        tracking_number, return_tracking_number = process_shipping_order(driver, 
                                                        row["TA_ID"], row["SYSTEM_NUMBER"], row["IVRS_NUMBER"],
                                                        row["SHIP_DATE"], row["SHIP_TIME_FROM"], row["SHIP_TIME_TO"],
                                                        row["DELIVERY_DATE"], row["DELIVERY_TIME_FROM"], row["DELIVERY_TIME_TO"],
                                                        row["TYPE_OF_MATERIAL"], row["TEMPERATURE"],
                                                        row["CONTACTS"], row["AMOUNT_OF_BOXES"],
                                                        row["RETURN"], row["RETURN_TO_TA"], row["RETURN_TYPE"], row["RETURN_CANTIDAD"])
        
        df.loc[index, "TRACKING_NUMBER"] = tracking_number
        df.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number

def update_shipping_order_table(df: pd.DataFrame, path: str, sheet: str):
    """
    Updates orders table

    Args:
        df (DataFrame): orders table
        path (str): excel file path
        sheet (str): excel sheet name
    """

    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

def generate_shipping_report(shipdate, team):
    """
    Process all orders in the table
    
    Variables:
        driver (webdriver): selenium driver
        wait (WebDriverWait): selenium wait
        team (str): team to process
        df_path (str): excel file path
        sheet (str): excel sheet name
        df (DataFrame): orders table
    """
    global wait

    driver, wait = init_driver()

    df_path, sheet, info_sites_sheet = load_paths(team)

    if log_in_TA_website(driver):
        print("Logged in")
    else:
        print("Not logged in")
        return
        
    df_orders = load_shipping_order_table(shipdate, team, df_path, sheet)
    
    df_info_sites = load_table_info_sites(team, df_path, info_sites_sheet)
    
    df = pd.merge(df_orders, df_info_sites, on=["STUDY", "SITE#"], how="inner")

    process_all_shipping_orders(driver, df)
    
    if not df.empty:
        dataframe_name = "orders_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
        df.to_excel(os.path.expanduser("~\\Downloads") + "\\" + dataframe_name, index=False)
        
        print(df[["SYSTEM_NUMBER", "STUDY", "SITE#", "IVRS_NUMBER", "SHIP_DATE", "TRACKING_NUMBER", "RETURN_TRACKING_NUMBER"]])
        print("Total: " + str(len(df.index)))
    else:
        print("Empty DataFrame")

    time.sleep(1)
    driver.quit()

if __name__ == "__main__":
    print(sys.version)
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("green")
    root = ctk.CTk()
    root.geometry("550x400")
    app = DatePicker(root)
    root.title('App Title')
    root.mainloop()