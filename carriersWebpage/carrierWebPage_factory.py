class CarrierWebPageFactory:
    """
    Factory class to create carrier webpages
    """
    def __init__(self):
        """
        Class constructor
        """
        pass

    def create_carrier_webpage(self, carrier_name: str, folder_path_to_download: str):
        """
        Creates a carrier webpage

        Args:
            carrier_name (str): carrier name
            folder_path_to_download (str): folder path to download

        Returns:
            CarrierWebpage: carrier webpage
        """
        match carrier_name:
            case "Transportes Ambientales":
                from .TransportesAmbientales import TransportesAmbientales
                return TransportesAmbientales(folder_path_to_download)
            
            case "Carrier Webpage For Testing":
                from .CarrierWebpageForTesting import CarrierWebpageForTesting
                return CarrierWebpageForTesting(folder_path_to_download)
            
            case _:
                from .NoCarrier import NoCarrier
                return NoCarrier(folder_path_to_download)
            
    def get_carrier_webpage_names(self) -> list:
        """
        Returns the carrier webpage names

        Returns:
            list: carrier webpage names
        """
        return ["Transportes Ambientales", "Carrier Webpage For Testing"]