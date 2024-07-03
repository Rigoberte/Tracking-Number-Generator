import os

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