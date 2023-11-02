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

def init_driver() -> (webdriver, WebDriverWait):
    """
    Inicia el driver de selenium y la espera de selenium

    Returns:
        webdriver: driver de selenium
        WebDriverWait: espera de selenium
    """
    global browser_that_prints
    browser_that_prints = PrintingBrowser()
    driver = browser_that_prints.driver
    wait = WebDriverWait(driver, 10)
    return driver, wait

def cargar_rutas(team: str) -> (str, str, str):
    """
    Carga las rutas de los archivos de excel

    Args:
        team (str): equipo a procesar

    Returns:
        str: ruta del archivo de excel
        str: nombre de la hoja de excel
        str: nombre de la hoja de excel de información de los sitios
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

def iniciar_sesion_TA(driver, username: str, password: str):
    """
    Inicia sesión en la página de TA

    Args:
        driver (webdriver): driver de selenium
        username (str): nombre de usuario
        password (str): contraseña
    """
    driver.get("https://sgi.tanet.com.ar/sgi/index.php")

    wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/button")))

    driver.find_element(By.XPATH, "/html/body/form/input[1]").send_keys(username)
    driver.find_element(By.XPATH, "/html/body/form/input[2]").send_keys(password)
    driver.find_element(By.XPATH, "/html/body/form/button").click()

def cargar_tabla_ordenes_envio(date : dt.datetime, team: str,  path: str, sheet: str ) -> pd.DataFrame:
    """
    Carga la tabla de ordenes de envío segun la fecha y el equipo

    Args:
        date (dt.datetime): fecha de las ordenes de envío
        team (str): equipo a procesar
        path (str): ruta del archivo de excel
        sheet (str): nombre de la hoja de excel

    Returns:
        DataFrame: tabla de ordenes de envío
    """
    if team == "Lilly":
        columns_names = {"CT-WIN": "SYSTEM_NUMBER", "IVRS": "IVRS_NUMBER",
                        "Trial Alias": "STUDY", "Site ": "SITE#",
                        "Order received": "ENTER DATE", "Ship date": "SHIP DATE",
                        "Horario de Despacho": "RECOLECCION_HORA_DESDE",  
                        "Dia de entrega": "ENTREGA_FECHA", "Destination": "DESTINATION",
                        "CONDICION": "TEMPERATURA", "TT4": "CANTIDAD_BULTOS",  
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
        columns_types = {"SITE#": str, "RETURN_TO_TA": bool, "RECOLECCION_HORA_HASTA": str}

    df = pd.read_excel(path, sheet_name=sheet, header=0, dtype=columns_types)
    df.rename(columns=columns_names, inplace=True)
    df = df[df["SHIP DATE"] == date]
    
    df["SHIP DATE"] = df["SHIP DATE"].dt.strftime('%d%m%y')
    df["SITE#"] = df["SITE#"].astype(object)
    df["CANTIDAD_BULTOS"] = df["CANTIDAD_BULTOS"].fillna(0).astype(int)

    if team == "Lilly": # Casos especiales del equipo Lilly
        df["TIPO_MATERIAL"] = "Medicacion"
        
        temperaturas = {"L": "Ambiente",
                        "M": "Ambiente Controlado", "M + L": "Ambiente Controlado",
                        "H": "Ambiente Controlado", "H + M": "Ambiente Controlado", "H + L": "Ambiente Controlado", "H + M + L": "Ambiente Controlado",
                        "REF": "Refrigerado", "REF + H": "Refrigerado", "REF + M": "Refrigerado", "REF + L": "Refrigerado",
                        "REF + H + M": "Refrigerado", "REF + H + L": "Refrigerado", "REF + M + L": "Refrigerado",
                        "REF + H + M + L": "Refrigerado"}
        df["TEMPERATURA"] = df["TEMPERATURA"].replace(temperaturas)
        df[(df["TEMPERATURA"] == "Ambiente") & (df["RETURN_TRACKING_NUMBER"] != "NA")]["TEMPERATURA"] = "Ambiente Controlado"
        df["Cajas (Carton)"] = df["Cajas (Carton)"].fillna(0).astype(int)
        df["RETURN_CANTIDAD"] = df["CANTIDAD_BULTOS"] - df["Cajas (Carton)"]
        df["RETURN"] = (df["RETURN_CANTIDAD"] > 0) & (df["TEMPERATURA"] != "Ambiente")
        df["RETURN_TO_TA"] = False
        df["TIPO_RETORNO"] = "CREDO"

        horariosDeDespacho = {"8": "08:00:00", "16.3": "16:30:00", "19": "19:00:00"} 
        df["RECOLECCION_HORA_DESDE"] = df["RECOLECCION_HORA_DESDE"].replace(horariosDeDespacho)
        df["RECOLECCION_HORA_DESDE"] = pd.to_datetime(df["RECOLECCION_HORA_DESDE"], format='%H:%M:%S', errors='coerce')
        df["RECOLECCION_HORA_HASTA"] = df["RECOLECCION_HORA_DESDE"] + dt.timedelta(minutes=30)
        df["RECOLECCION_HORA_DESDE"] = df["RECOLECCION_HORA_DESDE"].dt.strftime('%H:%M')
        df["RECOLECCION_HORA_HASTA"] = df["RECOLECCION_HORA_HASTA"].dt.strftime('%H:%M')

        df["RETURN_CANTIDAD"] = df["RETURN_CANTIDAD"].astype(int)
        df["ENTREGA_FECHA"] = pd.to_datetime(df["ENTREGA_FECHA"], format='%d%m%Y', errors='coerce')
 
    elif team == "GPM": # Casos especiales del equipo GPM
        df = df #TODO
    
    elif team == "Test" or team == "Test_5_ordenes": # Casos especiales de tests
        df["RECOLECCION_HORA_HASTA"] = ""
        df["RECOLECCION_HORA_DESDE"] = pd.to_datetime(df["RECOLECCION_HORA_DESDE"], format='%H:%M:%S', errors='coerce')
        df["RECOLECCION_HORA_HASTA"] = df["RECOLECCION_HORA_DESDE"] + dt.timedelta(minutes=30)

        df["RECOLECCION_HORA_DESDE"] = df["RECOLECCION_HORA_DESDE"].dt.strftime('%H:%M')
        df["RECOLECCION_HORA_HASTA"] = df["RECOLECCION_HORA_HASTA"].dt.strftime('%H:%M')

        df["RETURN_CANTIDAD"] = df["CANTIDAD_BULTOS"]

    df.fillna('', inplace=True)
    
    return df[['SYSTEM_NUMBER', 'IVRS_NUMBER', 'STUDY', 'SITE#',
               'SHIP DATE', 'RECOLECCION_HORA_DESDE', 'RECOLECCION_HORA_HASTA', 
               'ENTREGA_FECHA', 'TIPO_MATERIAL', 
               'TEMPERATURA', 'CANTIDAD_BULTOS', 'RETURN', 
               'RETURN_TO_TA', 'TIPO_RETORNO', 'RETURN_CANTIDAD',
               'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER']]

def cargar_tabla_info_sites(team: str, path: str, sheet: str) -> pd.DataFrame:
    """
    Carga la tabla de información de los sitios segun el equipo

    Args:
        team (str): equipo a procesar
        path (str): ruta del archivo de excel
        sheet (str): nombre de la hoja de excel

    Returns:
        DataFrame: tabla de información de los sitios
    """
    if team == "Lilly":
        columns_names = {"Protocolo": "STUDY", "Codigo": "TA_ID", "Site": "SITE#",
                        "Horario inicio": "ENTREGA_HORA_DESDE", "Horario fin": "ENTREGA_HORA_HASTA"}
        columns_types = {"Protocolo": str, "Site": str, "Codigo": str, "Horario inicio": str, "Horario fin": str}
    elif team == "Test" or team == "Test_5_ordenes":
        columns_names = {}
        columns_types = {"STUDY": str, "SITE#": str, "TA_ID": str, "ENTREGA_HORA_DESDE": dt.datetime, "ENTREGA_HORA_HASTA": dt.datetime}

    df = pd.read_excel(path, sheet_name=sheet, header=0, dtype=columns_types)
    df.rename(columns=columns_names, inplace=True)
    df.fillna('', inplace=True)

    df["CONTACTOS"] = ""
    df["ENTREGA_HORA_DESDE"] = pd.to_datetime(df["ENTREGA_HORA_DESDE"], format='%H:%M:%S', errors='coerce')
    df["ENTREGA_HORA_HASTA"] = pd.to_datetime(df["ENTREGA_HORA_HASTA"], format='%H:%M:%S', errors='coerce')
    df["ENTREGA_HORA_DESDE"] = df["ENTREGA_HORA_DESDE"].dt.strftime('%H:%M')
    df["ENTREGA_HORA_HASTA"] = df["ENTREGA_HORA_HASTA"].dt.strftime('%H:%M')

    return df[["STUDY", "SITE#", "TA_ID", "ENTREGA_HORA_DESDE", "ENTREGA_HORA_HASTA", "CONTACTOS"]]

def completar_formulario_orden_envio(driver, url: str, referencia: str,
                                    recoleccion_fecha: str, recoleccion_hora_desde: str, recoleccion_hora_hasta: str,
                                    entrega_fecha: str, entrega_hora_desde: str, entrega_hora_hasta: str,
                                    tipo_material: str, temperatura: str,
                                    contactos: str, cantidad_bultos: int) -> str :
    """
    Completa el formulario de orden de envío

    Args:
        driver (webdriver): driver de selenium
        url (str): url de la página
        referencia (str): referencia de la orden de envío
        recoleccion_fecha (str): fecha de recolección
        recoleccion_hora_desde (str): hora de recolección desde
        recoleccion_hora_hasta (str): hora de recolección hasta
        entrega_fecha (str): fecha de entrega
        entrega_hora_desde (str): hora de entrega desde
        entrega_hora_hasta (str): hora de entrega hasta
        tipo_material (str): tipo de material
        temperatura (str): temperatura
        contactos (str): contactos
        cantidad_bultos (int): cantidad de bultos
    
    Returns:
        str: tracking number
    """
    try:
        # Carga la página
        driver.get(url)
        
        # Espera a que se cargue el formulario
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/input")))
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").send_keys(referencia)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[1]").send_keys(recoleccion_fecha)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[2]").send_keys(recoleccion_hora_desde)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[3]").send_keys(recoleccion_hora_hasta)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[1]").send_keys(entrega_fecha)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[2]").send_keys(entrega_hora_desde)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[3]").send_keys(entrega_hora_hasta)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/input").send_keys("FCS")
        suggestions_container = wait.until(EC.presence_of_element_located((By.ID, "suggest_nomDomOri_list")))

        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/div/div/div[1]/table/tbody")))

        # Encuentra todos los botones sugeridos dentro del contenedor de sugerencias
        suggested_buttons = suggestions_container.find_elements(By.CLASS_NAME, "suggest")
        # Itera a través de los botones sugeridos y encuentra el que contiene el texto "Fisher Clinical Services FCS"
        for button in suggested_buttons:
            button_text = button.text.strip()
            if "Fisher Clinical Services FCS" in button_text:
                button.click()
                break
        
        # Selecciona el tipo de material
        if tipo_material == "Medicacion": it_tipo_material = 3
        elif tipo_material == "Material": it_tipo_material = 5
        elif tipo_material == "Equipo": it_tipo_material = 7
        else: it_tipo_material = 8

        for i in range(0, it_tipo_material):
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[1]").send_keys(Keys.DOWN)
        

        # Selecciona la temperatura
        if temperatura == "Ambiente": it_temperatura = 0
        elif temperatura == "Ambiente Controlado": it_temperatura = 1
        elif temperatura == "Refrigerado": it_temperatura = 2
        elif temperatura == "Congelado": it_temperatura = 3
        else: it_temperatura = 4

        for i in range(0, it_temperatura):
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[2]").send_keys(Keys.DOWN)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").clear()
        if contactos != "" or contactos != "NA":
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").clear()
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").send_keys(contactos)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input").send_keys(cantidad_bultos)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/button").click()

        # Espera a que se cargue la página con el tracking number
        time.sleep(1)
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))
    
        return driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:15]
    except:
        return ""

def completar_formulario_retorno(driver, url: str, referencia: str,
                                    entrega_fecha: str, entrega_hora_desde: str, entrega_hora_hasta: str,
                                    tipo_retorno: str, contactos: str, cantidad_bultos: int) -> str:
    """
    Completa el formulario de retorno de credo

    Args:
        driver (webdriver): driver de selenium
        url (str): url de la página
        referencia (str): referencia de la orden de envío
        entrega_fecha (str): fecha de entrega
        entrega_hora_desde (str): hora de entrega desde
        entrega_hora_hasta (str): hora de entrega hasta
        tipo_retorno (str): tipo de retorno
        contactos (str): contactos
        cantidad_bultos (int): cantidad de bultos

    Returns:
        str: tracking number de retorno de credo
    """
    try:
        # Carga la página
        driver.get(url)

        # Espera a que se cargue el formulario
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select")))
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").clear()
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").send_keys(referencia)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[1]").send_keys(entrega_fecha)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[2]").send_keys(entrega_hora_desde)
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[3]").send_keys(entrega_hora_hasta)

        # Selecciona el tipo de retorno
        if tipo_retorno == "CREDO": it_tipo_retorno = 1
        elif tipo_retorno == "DATALOGGER": it_tipo_retorno = 2
        else: it_tipo_retorno = 3
        for i in range(0, it_tipo_retorno):
            driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select").send_keys(Keys.DOWN)

        leyenda = "\n ***ENVÍO CON CAJA CREDO. EL COURIER AGUARDARÁ QUE EL CENTRO ALMACENE LA MEDICACIÓN Y RETORNE EL EMBALAJE***"
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/textarea").send_keys(leyenda)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input[1]").send_keys(contactos)
        
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").clear()
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").send_keys(cantidad_bultos)

        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[14]/td/button").click()
        
        # Espera a que se cargue la página con el tracking number
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))
        
        return driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:15]
    except:
        return ""

def imprimir_documentos_TA(driver, url: str):
    """
    Imprime los documentos de la página

    Args:
        driver (webdriver): driver de selenium
        url (str): url de la página
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

def procesar_orden_envio(driver, TA_ID:int, system_number:int, ivrs_number:str,
                        recoleccion_fecha: str, recoleccion_hora_desde: str, recoleccion_hora_hasta: str,
                        entrega_fecha: dt.datetime, entrega_hora_desde: str, entrega_hora_hasta: str,
                        tipo_material: str, temperatura: str,
                        contactos: str, cantidad_bultos: int,
                        return_: bool, return_to_TA: bool, tipo_retorno: str , return_cantidad: int) -> str:
    """
    -Procesa una orden de envio completando el formulario de TA
    -Le crea una orden de retorno de credo si es necesario
    -Imprime los documentos de la orden de envío

    Args:
        driver (webdriver): driver de selenium
        TA_ID (int): id de ubicación de TA
        system_number (int): número de sistema
        ivrs_number (str): número de ivrs
        recoleccion_fecha (str): fecha de recolección
        recoleccion_hora_desde (str): hora de recolección desde
        recoleccion_hora_hasta (str): hora de recolección hasta
        entrega_fecha (str): fecha de entrega
        entrega_hora_desde (str): hora de entrega desde
        entrega_hora_hasta (str): hora de entrega hasta
        tipo_material (str): tipo de material
        temperatura (str): temperatura
        contactos (str): contactos
        cantidad_bultos (int): cantidad de bultos
        return_ (bool): si es necesario crear una orden de retorno de credo
        return_to_TA (bool): si es necesario, decide si la orden de retorno va a TA
        return_cantidad (int): cantidad de cajas de la orden de retorno de credo
    
    Returns:
        str: tracking number
        str: tracking number de retorno de credo
    """

    url = "https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion=" + str(TA_ID)
    referencia = str(system_number) + " " + ivrs_number
    referencia = referencia[:50]

    driver.get(url)
    entrega_fecha = entrega_fecha.strftime('%d%m%Y')
    tracking_number = completar_formulario_orden_envio(driver,
                                url, referencia, 
                                recoleccion_fecha, recoleccion_hora_desde, recoleccion_hora_hasta,
                                str(entrega_fecha) , entrega_hora_desde, entrega_hora_hasta,
                                tipo_material, temperatura,
                                contactos, cantidad_bultos)
    
    if tracking_number == "": return "", ""

    if return_:
        url_return = "https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarRetorno+idsp=" + tracking_number[:7] + "&idubicacion=" + str(TA_ID)
        url_return += "&returnToTa=true" if return_to_TA else ""

        referencia_return = referencia + " RET " + tracking_number
        referencia_return = referencia_return[:50]

        driver.get(url_return)
        return_tracking_number = completar_formulario_retorno(driver, url_return, 
                                                                        referencia_return,
                                                                        str(entrega_fecha + dt.timedelta(days=1)), "9", "16",
                                                                        tipo_retorno, "Personal de FCS", 
                                                                        return_cantidad)
    else:
        return_tracking_number = ""

    url_guias = "https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id=" + str(tracking_number)[:7] + "&idservicio=" + str(tracking_number)[:7] + "&copies=4"
    driver.get(url_guias)
    imprimir_documentos_TA(driver, url_guias)

    url_rotulo = "https://sgi.tanet.com.ar/sgi/srv.RotuloPdf.emitir+id=" + str(tracking_number)[:7] + "&idservicio=" + str(tracking_number)[:7] + "&copies=1"
    driver.get(url_rotulo)
    imprimir_documentos_TA(driver, url_rotulo)

    return tracking_number, return_tracking_number

def procesar_ordenes_envios(driver, df: pd.DataFrame):
    """
    Procesa todas las ordenes de envío de la tabla

    Args:
        driver (webdriver): driver de selenium
        df (DataFrame): tabla de ordenes de envío
    """
    
    for index, row in df.iterrows():
        if df.loc[index, "TRACKING_NUMBER"] != "": continue

        tracking_number, return_tracking_number = procesar_orden_envio(driver, 
                                                        row["TA_ID"], row["SYSTEM_NUMBER"], row["IVRS_NUMBER"],
                                                        row["SHIP DATE"], row["RECOLECCION_HORA_DESDE"], row["RECOLECCION_HORA_HASTA"],
                                                        row["ENTREGA_FECHA"], row["ENTREGA_HORA_DESDE"], row["ENTREGA_HORA_HASTA"],
                                                        row["TIPO_MATERIAL"], row["TEMPERATURA"],
                                                        row["CONTACTOS"], row["CANTIDAD_BULTOS"],
                                                        row["RETURN"], row["RETURN_TO_TA"], row["TIPO_RETORNO"], row["RETURN_CANTIDAD"])
        
        df.loc[index, "TRACKING_NUMBER"] = tracking_number
        df.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number

def actualizar_tabla_ordenes_envio(df: pd.DataFrame, path: str, sheet: str):
    """
    Actualiza la tabla de ordenes de envío

    Args:
        df (DataFrame): tabla de ordenes de envío
        path (str): ruta del archivo de excel
        sheet (str): nombre de la hoja de excel
    """

    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

def exportar_a_excel(df: pd.DataFrame, path: str):
    """
    Exporta la tabla de ordenes de envío a un archivo de excel

    Args:
        df (DataFrame): tabla de ordenes de envío
        path (str): ruta del archivo de excel
    """
    df.to_excel(path, index=False)

def main():
    """
    Procesa todas las ordenes de envío de la tabla
    
    Variables:
        driver (webdriver): driver de selenium
        wait (WebDriverWait): espera de selenium
        team (str): equipo a procesar
        df_path (str): ruta del archivo de excel
        sheet (str): nombre de la hoja de excel
        df (DataFrame): tabla de ordenes de envío
    """
    global wait

    team = "Lilly"
    shipdate = dt.datetime(2023, 11, 4) # Año, mes, día

    driver, wait = init_driver()

    df_path, sheet, info_sites_sheet = cargar_rutas(team)

    if True: # Para ocultar el usuario y contraseña :)
        username = "inaki.costa"
        password = "Alejan1961"
    
    iniciar_sesion_TA(driver, username, password)

    df_ordenes = cargar_tabla_ordenes_envio(shipdate, team, df_path, sheet)
    
    df_info_sites = cargar_tabla_info_sites(team, df_path, info_sites_sheet)
    
    df = pd.merge(df_ordenes, df_info_sites, on=["STUDY", "SITE#"], how="inner")

    procesar_ordenes_envios(driver, df)
    
    if not df.empty:
        dataframe_name = "orders_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
        exportar_a_excel(df, os.path.expanduser("~\\Downloads") + "\\" + dataframe_name)
        
        print(df[["SYSTEM_NUMBER", "STUDY", "SITE#", "IVRS_NUMBER", "SHIP DATE", "RECOLECCION_HORA_DESDE", "RECOLECCION_HORA_HASTA"]]) #, "TRACKING_NUMBER", "RETURN_TRACKING_NUMBER"]])
        print("Total: " + str(len(df.index)))
    else:
        print("Empty DataFrame")

    time.sleep(1)
    driver.quit()

main()