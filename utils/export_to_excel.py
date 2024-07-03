import pandas as pd
import datetime as dt

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
        raise Exception("Empty DataFrame")