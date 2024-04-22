from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import datetime as dt
import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk
import os, time

class Browser(object):
    def __init__(self, folder_path: str):
        """
        Class constructor for Browser 

        Args:
            folder_path (str): folder path to download files

        Attributes:
            self.driver (webdriver): selenium self.driver
        """
        self.folder_path = folder_path
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

        self.chrome_prefs = {
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "download.open_pdf_in_system_reader": False,
            "profile.default_content_settings.popups": 0,
            "download.prompt_for_download": False, #To auto download the file
            "printing.print_to_pdf": True,
            "download.default_directory": folder_path,
            "savefile.default_directory": folder_path
        }

        chrome_options.add_experimental_option("prefs", self.chrome_prefs)

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
            self.driver.find_element(By.XPATH, "/html/body/form/input[1]").send_keys("inaki.costa")
            self.driver.find_element(By.XPATH, "/html/body/form/input[2]").send_keys("Chavi1961")

            self.driver.find_element(By.XPATH, "/html/body/form/button").click()
            # Wait user input their credentials 
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/div[2]/div[1]/table/tbody/tr[1]/td")))
            self.wait = WebDriverWait(self.driver, 10)
            return True
        except:
            self.wait = WebDriverWait(self.driver, 10)
            return False

    def imprimir_todos_los_registros_de_servicio(self, df: pd.DataFrame):
        """
        Process all orders in the table

        Args:
            self.driver (webdriver): selenium self.driver
            df (DataFrame): orders table
        """
        
        if self.log_in_TA_website():
            print("Logged in")
        else:
            print("Not logged in")
            return

        if not os.path.exists(self.browser.folder_path):
            os.makedirs(self.browser.folder_path)

        for index, row in df.iterrows():
            if df.loc[index, "TRACKING_NUMBER"] == "":
                continue
            
            tracking_number = df.loc[index, "TRACKING_NUMBER"][:7]
            url = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirCopia+id={tracking_number}&idservicio={tracking_number}"
            self.print_webpage(url)
            #self.open_new_tab(url)

        time.sleep(2)

        self.driver.quit()

        return df

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

    def open_new_tab(self, url: str):
        """
        Opens new tab

        Args:
            self.driver (webdriver): selenium self.driver
            url (str): webpage url
        """
        self.driver.execute_script(f"window.open('{url}', '_blank')")

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
        self.title("Service Records Printer")
        self.state("zoomed")

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        self.colors = self.Chroma()

        # Top Frame
        frame_top = ctk.CTkFrame(self)
        frame_top.pack(side=tk.TOP, fill=tk.X, pady=0)

        # Image banner
        imagen = Image.open("fcs.jpg")
        imagen = imagen.resize((284, 67))
        imagen = ImageTk.PhotoImage(imagen)

        label_banner = tk.Label(frame_top, image=imagen)
        label_banner.image = imagen
        label_banner.pack(side=tk.LEFT, padx=10)

        # Teams Combobox
        teams_options = ["Lilly"]
        self.team_combobox = tk.ttk.Combobox(frame_top, values=teams_options, width=20, height=15, font=30)
        self.team_combobox.pack(side=tk.LEFT, padx=10)
        self.team_combobox.current(0)

        # Bottom Frame
        frame_bottom = ctk.CTkFrame(self)
        frame_bottom.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, pady=0)

        # Orders self.treeview (table)
        style = tk.ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 13)) # Font of the body
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 13,'bold')) # Font of the headings
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) # Remove the borders
        
        columns_df = ['#', 'SYSTEM_NUMBER', 'IVRS_NUMBER', 'TRACKING_NUMBER', 'BOX']
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
        #y_scrollbar = tk.ttk.Scrollbar(frame_bottom, orient=tk.VERTICAL, command=self.treeview.yview)
        #y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        #self.treeview.configure(command=self.treeview.yview)
        #y_scrollbar.grid(row=0, column=1, rowspan=10, sticky=NS)
        #y_scrollbar.place(x=self.treeview.winfo_x() + self.treeview.winfo_width(), y=self.treeview.winfo_y(), height=self.treeview.winfo_height())

        # Horizontal scrollbar
        x_scrollbar = tk.ttk.Scrollbar(frame_bottom, orient=tk.HORIZONTAL, command=self.treeview.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.treeview.configure(xscrollcommand=x_scrollbar.set)

        # Buttons ###################
        # Process orders button
        processOrders_btn = ctk.CTkButton(master=frame_top, text="Save PDFs", command=self.on_processOrders_btn_click,
                                        width=150, height=50, font=('Calibri', 22, 'bold'))
        processOrders_btn.pack(side=tk.RIGHT, padx=10)

        # Load orders button
        loadOrders_btn = ctk.CTkButton(master=frame_top, text="Load Orders", command=self.on_loadOrders_btn_click,
                                       width=150, height=50, font=('Calibri', 22, 'bold'))
        loadOrders_btn.pack(side=tk.RIGHT, padx=10)

        imprimir_btn = ctk.CTkButton(master=frame_top, text="Imprimir", command=self.on_imprimir_btn_click,
                                        width=150, height=50, font=('Calibri', 22, 'bold'))
        imprimir_btn.pack(side=tk.RIGHT, padx=10)
        
    def on_imprimir_btn_click(self):
        # Ruta de la carpeta que contiene los PDFs
        carpeta = os.path.expanduser("~\\Downloads") + "\\" + "Registros_de_Servicio"

        """# Obtener la lista de archivos en la carpeta
        archivos = sorted(os.listdir(carpeta), key=lambda x: os.path.getctime(os.path.join(carpeta, x)))
        
        # Obtener la impresora predeterminada
        printer_name = win32print.GetDefaultPrinter()

        # Iterar sobre cada archivo en la carpeta
        for archivo in archivos:
            # Comprobar si el archivo es un PDF
            if archivo.endswith('.pdf'):
                # Obtener la ruta completa del archivo PDF
                ruta_pdf = os.path.join(carpeta, archivo)
                
                # Imprimir el archivo PDF usando win32print
                hprinter = win32print.OpenPrinter(printer_name)
                try:
                    win32print.StartDocPrinter(hprinter, 1, (archivo, None, "RAW"))
                    win32print.StartPagePrinter(hprinter)
                    with open(ruta_pdf, "rb") as f:
                        win32api.SetJob(hprinter, win32print.StartDocPrinter(hprinter, 1, (archivo, None, "RAW")), 1, f.read(), 0)
                    win32print.EndPagePrinter(hprinter)
                    win32print.EndDocPrinter(hprinter)
                finally:
                    win32print.ClosePrinter(hprinter)"""
        
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
        selected_date = dt.datetime(2021, 1, 1)
        selected_team = "Lilly"

        self.df = self.generate_shipping_order_tables(selected_team, selected_date)

        for item in self.treeview.get_children():
            self.treeview.delete(item)
        
        parity = False
        i = 1
        for index, row in self.df.iterrows():
            r = [i] + list(row)

            tag_color = 'odd' if parity else 'even'
            self.treeview.insert("", "end", values=r, tags=(tag_color))
            parity = not parity
            i += 1
        
        self.treeview.update()

    def on_processOrders_btn_click(self):
        """
        Button to process all orders in the table
        """
        boxes = list(set(self.df["BOX"]))
        self.df = self.generate_shipping_order_tables("Lilly", dt.datetime.now())
        
        for box in boxes:
            self.df_box = self.df[self.df["BOX"] == box]

            folder_path = os.path.expanduser("~\\Downloads") + "\\" + "Registros_de_Servicio" + "\\" + box

            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            order_processor = OrderProcessor(folder_path)

            self.df_box = order_processor.imprimir_todos_los_registros_de_servicio(self.df_box)

            # Obtener la lista de archivos en el folder_path
            archivos = os.listdir(folder_path)

            # Filtrar solo los archivos (excluir directorios)
            archivos = [archivo for archivo in archivos if os.path.isfile(os.path.join(folder_path, archivo))]

            # Obtener la fecha de creación de cada archivo y almacenarla en un diccionario
            fechas_creacion = {}
            for archivo in archivos:
                ruta_archivo = os.path.join(folder_path, archivo)
                fecha_creacion = dt.datetime.fromtimestamp(os.path.getctime(ruta_archivo))
                fechas_creacion[archivo] = fecha_creacion

            # Ordenar los archivos por fecha de creación
            archivos_ordenados = sorted(archivos, key=lambda x: fechas_creacion[x])

            # Renombrar los archivos con un índice al principio del nombre
            for indice, archivo in enumerate(archivos_ordenados, start=1):
                nuevo_nombre = f"{indice}_{archivo}"
                os.rename(os.path.join(folder_path, archivo), os.path.join(folder_path, nuevo_nombre))

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
        
        df.fillna('', inplace=True)
        
        return df[['SYSTEM_NUMBER', 'IVRS_NUMBER', 'TRACKING_NUMBER']]

    def cargar_ordenes(self) -> pd.DataFrame:
        path = os.path.expanduser("~\\Downloads\escaneo.xlsx")
        columns_types = {"SYSTEM_NUMBER": str, "BOX": str}
        df_excel = pd.read_excel(path, header=0, dtype=columns_types)
        return df_excel

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

        df_excel = self.cargar_ordenes()

        df = pd.merge(df_excel, df_orders, on=["SYSTEM_NUMBER"], how="left")

        df.fillna('', inplace=True)

        return df[['SYSTEM_NUMBER', 'IVRS_NUMBER', 'TRACKING_NUMBER', 'BOX']]

def main():
    ctk.set_appearance_mode("dark")
    app = MyUserForm()
    app.mainloop()

if __name__ == "__main__":
    main()