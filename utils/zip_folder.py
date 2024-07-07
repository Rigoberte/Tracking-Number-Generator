import shutil

def zip_folder(folder_path: str, zip_path: str) -> None:
        """
        Zips a folder

        Args:
                folder_path (str): folder path to zip
                zip_path (str): zip path
        """
        shutil.make_archive(zip_path, 'zip', folder_path)