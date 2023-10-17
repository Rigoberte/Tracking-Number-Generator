from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

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

def completar_formulario_orden_envio(driver, url: str, referencia: str,
                                    recoleccion_fecha: str, recoleccion_hora_desde: str, recoleccion_hora_hasta: str,
                                    entrega_fecha: str, entrega_hora_desde: str, entrega_hora_hasta: str,
                                    tipo_material: str, temperatura: str,
                                    contactos: str, cantidad_cajas: int) -> str :
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
        cantidad_cajas (int): cantidad de cajas
    
    Returns:
        str: tracking number
    """
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
    driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").send_keys(contactos)
    driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input").send_keys(cantidad_cajas)
    
    #driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/button").click()

    return "" #driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption/text()").text[-9:]

def completar_formulario_retorno_credo(driver, url: str, referencia: str,
                                    entrega_fecha: str, entrega_hora_desde: str, entrega_hora_hasta: str,
                                    tipo_retorno: str, contactos: str, cantidad_cajas: int) -> str:
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
        cantidad_cajas (int): cantidad de cajas

    Returns:
        str: tracking number de retorno de credo
    """
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
    if tipo_retorno == "CREDO": it_tipo_retorno = 0
    elif tipo_retorno == "DATALOGGER": it_tipo_retorno = 1
    else: it_tipo_retorno = 2
    for i in range(0, it_tipo_retorno):
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select").send_keys(Keys.DOWN)

    driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input[1]").send_keys(contactos)
    
    driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").clear()
    driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").send_keys(cantidad_cajas)

    #driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[14]/td/button").click()
    
    return "" #driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption/text()").text[-9:]

def imprimir_documentos_TA(driver, tracking_number: str):
    """
    Imprime los documentos (guias y rotulo) de la orden de envío

    Args:
        driver (webdriver): driver de selenium
        tracking_number (str): tracking number de la orden de envío
    """
    # Carga la página
    url_guias = "https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id=" + str(tracking_number)[:7] + "&idservicio=" + str(tracking_number)[:7]
    #driver.get(url_guias)

    url_rotulo = "https://sgi.tanet.com.ar/sgi/srv.RotuloPdf.emitir+id=" + str(tracking_number)[:7] + "&idservicio=" + str(tracking_number)[:7]
    #driver.get(url_rotulo)

def procesar_orden_envio(driver, TA_ID:int, system_number:int, ivrs_number:str,
                        recoleccion_fecha: str, recoleccion_hora_desde: str, recoleccion_hora_hasta: str,
                        entrega_fecha: str, entrega_hora_desde: str, entrega_hora_hasta: str,
                        tipo_material: str, temperatura: str,
                        contactos: str, cantidad_cajas: int,
                        credo_return: bool, credo_return_TA: bool, credo_return_cant: int) -> str:
    """
    Procesa una orden de envio completando el formulario de TA y le crea una orden de retorno de credo si es necesario

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
        cantidad_cajas (int): cantidad de cajas
        credo_return (bool): si es necesario crear una orden de retorno de credo
        credo_return_TA (bool): si es necesario, decide si la orden de retorno va a TA
        credo_return_cant (int): cantidad de cajas de la orden de retorno de credo
    
    Returns:
        str: tracking number
        str: tracking number de retorno de credo
    """

    url = "https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion=" + str(TA_ID)
    referencia = str(system_number) + " " + ivrs_number

    driver.get(url)

    tracking_number = completar_formulario_orden_envio(driver,
                                url, referencia, 
                                recoleccion_fecha, recoleccion_hora_desde, recoleccion_hora_hasta,
                                entrega_fecha, entrega_hora_desde, entrega_hora_hasta,
                                tipo_material, temperatura,
                                contactos, cantidad_cajas)
    
    if credo_return:
        url_credo_return = "https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarRetorno+idsp=" + tracking_number[:7] + "&idubicacion=" + str(TA_ID)
        url_credo_return += "&returnToTa=true" if credo_return_TA else ""

        referencia_credo_return = referencia + " RET " + tracking_number

        driver.get(url_credo_return)

        credo_return_tracking_number = completar_formulario_retorno_credo(driver, url_credo_return, 
                                                                        referencia_credo_return,
                                                                        "2512", "9", "16",
                                                                        "CREDO", "Personal de FCS", 
                                                                        credo_return_cant)
    else:
        credo_return_tracking_number = ""


    #imprimir_documentos_TA(driver, tracking_number)

    return tracking_number, credo_return_tracking_number

def cargar_tabla_ordenes_envio(date) -> pd.DataFrame:
    """
    Carga la tabla de ordenes de envío

    Args:
        driver (webdriver): driver de selenium
        date (str): fecha de las ordenes a procesar

    Returns:
        DataFrame: tabla de ordenes de envío
    """

    df = pd.DataFrame({ "SYSTEM_NUMBER": [12345, 12346, 12347, 12348, 12349],
                        "IVRS_NUMBER": ["IVRS_NUMBER_01", "IVRS_NUMBER_02", "IVRS_NUMBER_03", "IVRS_NUMBER_04", "IVRS_NUMBER_05"],
                        "TA_ID": [5616, 5616, 5616, 5616, 5616],
                        "RECOLECCION_FECHA": ["2012", "2012", "2012", "2012", "2012"],
                        "RECOLECCION_HORA_DESDE": ["19", "19", "19", "19", "19"],
                        "RECOLECCION_HORA_HASTA": ["1930", "1930", "1930", "1930", "1930"],
                        "ENTREGA_FECHA": ["2112", "2112", "2112", "2112", "2112"],
                        "ENTREGA_HORA_DESDE": ["10", "10", "10", "10", "10"],
                        "ENTREGA_HORA_HASTA": ["12", "12", "12", "12", "12"],
                        "TIPO_MATERIAL": ["Medicacion", "Materiales", "Medicacion", "Materiales", "Medicacion"],
                        "TEMPERATURA": ["Refrigerado", "Ambiente", "Ambiente Controlado", "Ambiente", "Ambiente"],
                        "CONTACTOS": ["Contactos_01", "Contactos_02", "Contactos_03", "Contactos_04", "Contactos_05"],
                        "CANTIDAD_CAJAS": [1, 3, 2, 4, 1],
                        "CREDO_RETURN": [False, False, False, False, False],
                        "TRACKING_NUMBER": ["", "", "", "", ""],
                        "CREDO_RETURN_TRACKING_NUMBER": ["", "", "", "", ""]})
    
    return df

def procesar_ordenes_envios(driver, df):
    """
    Procesa todas las ordenes de envío de la tabla

    Args:
        driver (webdriver): driver de selenium
        df (DataFrame): tabla de ordenes de envío
    """
    
    #for index, row in df.iterrows():
    #    tracking_number, credo_return_tracking_number = procesar_orden_envio(driver, 
    #                                                row["TA_ID"], row["SYSTEM_NUMBER"], row["IVRS_NUMBER"],
    #                                                row["RECOLECCION_FECHA"], row["RECOLECCION_HORA_DESDE"], row["RECOLECCION_HORA_HASTA"],
    #                                                row["ENTREGA_FECHA"], row["ENTREGA_HORA_DESDE"], row["ENTREGA_HORA_HASTA"],
    #                                                row["TIPO_MATERIAL"], row["TEMPERATURA"],
    #                                                row["CONTACTOS"], row["CANTIDAD_CAJAS"], row["CREDO_RETURN"])
    #    
    #    df.loc[index, "TRACKING_NUMBER"] = tracking_number
    #    df.loc[index, "CREDO_RETURN_TRACKING_NUMBER"] = credo_return_tracking_number

    tracking_number, credo_return_tracking_number = procesar_orden_envio(driver, 
                                                    5616, 12345, "IVRS_NUMBER",
                                                    "2012", "19", "1930",
                                                    "2112", "10", "12",
                                                    "Medicacion", "Refrigerado",
                                                    "Contactos", 2, False, False, 2)
    
    print(df)
    

def main():
    global wait
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)

    iniciar_sesion_TA(driver, "inaki.costa", "Alejan1961")

    df = cargar_tabla_ordenes_envio("20dic23")

    procesar_ordenes_envios(driver, df)

    time.sleep(20)
    driver.quit()

main()