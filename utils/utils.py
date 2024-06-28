import os
import shutil

import pandas as pd
import datetime as dt

from logClass.log import Log

def create_folder(folder_path_to_download: str):
    """
    Creates a folder

    Args:
        folder_path_to_download (str): folder path to download
    """
    try:
        os.makedirs(folder_path_to_download, exist_ok=True)
    except FileExistsError:
        pass

def getFolderPathToDownload(team: str, date: str) -> str:
    """
    Gets the folder path to download

    Args:
        team (str): team to process
        date (str): date to process

    Returns:
        str: folder path to download
    """
    downloads_path = os.path.expanduser("~\\Downloads")

    folder_path_to_download = os.path.join(downloads_path, "Tracking_Number_Generator", team, date)
    
    return folder_path_to_download

def renameAllReturnFiles(ordersAndContactsDataframe: pd.DataFrame, folder_path_to_download: str):
    """
    Renames all return files

    Args:
        ordersAndContactsDataframe (DataFrame): return tracking numbers
    """
    dataFrameWithReturnTrackingNumbers = ordersAndContactsDataframe[(ordersAndContactsDataframe["RETURN_TRACKING_NUMBER"] != "") & (ordersAndContactsDataframe["PRINT_RETURN_DOCUMENT"])][["RETURN_TRACKING_NUMBER"]]
    
    if dataFrameWithReturnTrackingNumbers.empty:
        return

    for index, row in dataFrameWithReturnTrackingNumbers.iterrows():
        __renameReturnPDFFile__(row["RETURN_TRACKING_NUMBER"], folder_path_to_download)

def export_to_excel(dataFrame: pd.DataFrame, folder_path_to_download: str, filename: str):
    """
    Exports the orders table to an excel file

    Args:
        dataFrame (DataFrame): orders table
    """
    if not dataFrame.empty:
        dataframe_name = filename + "_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
        dataFrame.to_excel(folder_path_to_download + "\\" + dataframe_name, index=False)
    else:
        Log().add_warning_log("Empty DataFrame")

def zip_folder(folder_path: str, zip_path: str):
        shutil.make_archive(zip_path, 'zip', folder_path)

def __renameReturnPDFFile__(return_tracking_number: str, folder_path_to_download: str):
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
        Log().add_error_log(f"Error renaming return file: {e}")