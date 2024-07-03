
import os

def renameReturnPDFFile(return_tracking_number: str, folder_path_to_download: str):
    """
    Renames the return file

    Args:
        return_tracking_number (str): return tracking number
    """
    try:
        pdf_path = folder_path_to_download + "\\JOB " + return_tracking_number + ".pdf"
        new_pdf_path = folder_path_to_download + "\\JOB " + return_tracking_number + " RETORNO DE CREDO.pdf"
        os.rename(pdf_path, new_pdf_path)
    except Exception as e:
        raise Exception(f"Error renaming return file: {e}")