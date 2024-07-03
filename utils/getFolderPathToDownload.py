import os

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