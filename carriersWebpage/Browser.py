from selenium import webdriver

class Browser(object):
    def __init__(self, folder_path_to_download: str):
        """
        Class constructor for Browser 

        Args:
            folder_path_to_download (str): folder path to download files

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
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
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
            "printing.print_to_pdf": True,
            "download.default_directory": folder_path_to_download,
            "savefile.default_directory": folder_path_to_download
        }

        chrome_options.add_experimental_option("prefs", chrome_prefs)

        self.driver = webdriver.Chrome(options=chrome_options)