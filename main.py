from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd


def iniciar_sesion_TA(driver, username, password):
    driver.get("https://sgi.tanet.com.ar/sgi/index.php")

    wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/button")))

    driver.find_element(By.XPATH, "/html/body/form/input[1]").send_keys(username)
    driver.find_element(By.XPATH, "/html/body/form/input[2]").send_keys(password)
    driver.find_element(By.XPATH, "/html/body/form/button").click()

def completar_formulario_orden_envio(driver, url, referencia,
                                    recoleccion_fecha, recoleccion_hora_desde, recoleccion_hora_hasta,
                                    entrega_fecha, entrega_hora_desde, entrega_hora_hasta,
                                    tipo_material, temperatura,
                                    contactos, cantidad_cajas):
    driver.get(url)
    
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

    #Encuentra todos los botones sugeridos dentro del contenedor de sugerencias
    suggested_buttons = suggestions_container.find_elements(By.CLASS_NAME, "suggest")

    # Itera a través de los botones sugeridos y encuentra el que deseas seleccionar
    for button in suggested_buttons:
        button_text = button.text.strip()
        if "Fisher Clinical Services FCS" in button_text:
            button.click()
            break
    
    if tipo_material == "Medicacion": it_tipo_material = 3
    elif tipo_material == "Material": it_tipo_material = 5
    elif tipo_material == "Equipo": it_tipo_material = 7
    else: it_tipo_material = 8

    for i in range(0, it_tipo_material):
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[1]").send_keys(Keys.DOWN)
    
    if temperatura == "Ambiente": it_temperatura = 0
    elif temperatura == "Ambiente Controlado": it_temperatura = 1
    elif temperatura == "Refrigerado": it_temperatura = 2
    elif temperatura == "Congelado": it_temperatura = 3
    else: it_temperatura = 4

    for i in range(0, it_temperatura):
        driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[2]").send_keys(Keys.DOWN)


    driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").send_keys(contactos)
    driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input").send_keys(cantidad_cajas)
    
    #driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/button").click()

    #return driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption/text()").text[-9:]

def cargar_tabla_ordenes_envio(driver, date):
    df = pd.DataFrame({})
    return df

def procesar_ordenes_envios(driver, df):

    tracking_number = completar_formulario_orden_envio(driver,
                                "https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion=5616", "Referencia", 
                                "2012", "19", "1930",
                                "2112", "10", "12",
                                "Medicacion", "Ambiente Controlado",
                                "Contactos", "2")

def main():
    global wait
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)

    iniciar_sesion_TA(driver, "inaki.costa", "Alejan1961")

    driver.get("https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion=5616")

    df = cargar_tabla_ordenes_envio(driver, "20dic23")

    procesar_ordenes_envios(driver, df)

    time.sleep(20)
    driver.quit()
main()