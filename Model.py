import pandas as pd
import datetime as dt
import time
import os
import queue
import threading

from dataRecolector.dataRecolector import DataRecolector
from orderProcessor.orderProcessor import OrderProcessor
from dataPathController.dataPathController import DataPathController
from teams.team import Team
from teams.team_factory import TeamFactory
from logClass.log import Log
from utils.utils import getFolderPathToDownload, create_folder

class Model:
    def __init__(self):
        self.selected_team = TeamFactory().create_team("No Selected Team", "")
        self.ordersAndContactsDataframe = self.getEmptyOrdersAndContactsData()
        
        self.queue = queue.Queue()

        self.log = Log()
        self.add_info_log("Application started")

    def on_loadOrders_btn_click(self, selected_team_name, selected_date) -> None:
        thread = threading.Thread(target=self.__loadOrdersAndCalculateTime__, args=(selected_team_name, selected_date), daemon=True)
        thread.start()

    def on_clearOrders_btn_click(self) -> None:
        self.ordersAndContactsDataframe = self.getEmptyOrdersAndContactsData()
    
        self.queue.put(self.ordersAndContactsDataframe)

        self.log.add_info_log("Orders table cleaned")

    def getEmptyOrdersAndContactsData(self) -> pd.DataFrame:
        no_selected_team = TeamFactory().create_team("No Selected Team", "")
        return DataRecolector(no_selected_team).getEmptyOrdersAndContactsData()

    def getTeamsNames(self) -> list:
        return TeamFactory().get_team_names()

    def validate_login(self, username: str, password: str) -> bool:
        self.selected_team.build_driver()
        return self.selected_team.check_if_user_and_password_are_correct(username, password)

    def on_login_successful(self) -> None:
        thread = threading.Thread( target= self.__processOrdersAndCalculateTime__, args=(self.selected_team, self.selected_date, self.folder_path_to_download, self.ordersAndContactsDataframe), daemon=True)
        thread.start()

    def on_login_failed(self) -> None:
        self.selected_team.quit_driver()

    def add_error_log(self, text: str) -> None:
        self.log.add_error_log(text)

    def add_warning_log(self, text: str) -> None:
        self.log.add_warning_log(text)

    def add_info_log(self, text: str) -> None:
        self.log.add_info_log(text)

    def print_logs(self) -> pd.DataFrame:
        return self.log.print_logs()

    def print_last_n_logs(self, n: int) -> pd.DataFrame:
        return self.log.print_last_n_logs(n)

    def on_export_logs_to_csv(self) -> None:
        self.log.export_to_csv(os.path.expanduser("~\\Downloads"))

    def get_config_of_a_team(self, teamName: str) -> str:
        return DataPathController().get_config_of_a_team(teamName)
    
    def on_click_save_config_button(self, 
                                    team_name : str, 
                                    team_excel_path : str, 
                                    team_orders_sheet : str, 
                                    team_contacts_sheet : str, 
                                    team_not_working_days_sheet : str, 
                                    team_send_email_to_medical_centers : bool) -> None:
        
        DataPathController().redefine_a_config_of_a_team(team_name, 
            {
            "team_excel_path": team_excel_path,
            "team_orders_sheet": team_orders_sheet,
            "team_contacts_sheet": team_contacts_sheet,
            "team_not_working_days_sheet": team_not_working_days_sheet,
            "team_send_email_to_medical_centers": team_send_email_to_medical_centers
            }
        )

    def on_open_excel_double_btn_click(self, temporal_selected_team_name: str) -> None:
        temporal_selected_team = TeamFactory().create_team(temporal_selected_team_name, "")

        excel_dataPath = temporal_selected_team.get_data_path(["team_excel_path"])

        try:
            if excel_dataPath[0] == "" or not os.path.exists(excel_dataPath[0]) or not excel_dataPath[0].endswith(".xlsx"):
                self.add_warning_log("Excel file path not found")
                return
                
            os.startfile(excel_dataPath[0])
        except Exception as e:
            self.add_error_log(f"Error opening Excel file: {e}")

    def __loadOrders__(self, selected_team_name: str, selected_date: str) -> None:
        self.folder_path_to_download = getFolderPathToDownload(selected_team_name, selected_date.replace("/", "_"))
        self.selected_team = TeamFactory().create_team(selected_team_name, self.folder_path_to_download)

        self.selected_date = dt.datetime.strptime(selected_date, '%Y-%m-%d')
        self.ordersAndContactsDataframe = DataRecolector(self.selected_team, self.queue).recolectOrdersAndContactsData(self.selected_date)

    def __loadOrdersAndCalculateTime__(self, selected_team_name: str, selected_date: str) -> None:
        self.queue.put("BLOCK MAIN USERFORM WIDGETS")
        
        self.log.add_separator()
        self.add_info_log(f"Start loading orders for {selected_team_name} team")
        
        time0 = time.time()
        self.__loadOrders__(selected_team_name, selected_date)
        time1 = time.time()
        total_time = time1 - time0
        total_time_with_2_decimals = round(total_time, 2)

        self.add_info_log(f"End loading orders")
        self.add_info_log(f"Total loading time: {total_time_with_2_decimals} s")

        self.queue.put("UNBLOCK MAIN USERFORM WIDGETS")

    def __processOrders__(self, selected_team: Team, selected_date: dt.datetime, folder_path_to_download: str, ordersAndContactsDataframe: pd.DataFrame) -> pd.DataFrame:
        create_folder(folder_path_to_download)

        ordersAndContactsDataframe = OrderProcessor(folder_path_to_download, selected_team, self.queue).processOrdersAndContactsTable(ordersAndContactsDataframe)
        
        selected_date_str = selected_date.strftime("%Y-%m-%d")
        
        selected_team.sendEmailWithOrdersToTeam(folder_path_to_download, selected_date_str)

        return ordersAndContactsDataframe

    def __processOrdersAndCalculateTime__(self, selected_team: Team, selected_date: dt.datetime, folder_path_to_download: str, ordersAndContactsDataframe: pd.DataFrame) -> None:
        self.queue.put("BLOCK MAIN USERFORM WIDGETS")
        
        amountOfOrdersToProcess = len(ordersAndContactsDataframe[(ordersAndContactsDataframe["TRACKING_NUMBER"] == "") & (ordersAndContactsDataframe["HAS_AN_ERROR"] == "No error")])
        
        self.add_info_log(f"Start processing orders")
        
        time0 = time.time()
        try:
            self.ordersAndContactsDataframe = self.__processOrders__(
                                            selected_team, selected_date,
                                            folder_path_to_download, ordersAndContactsDataframe)
        finally:
            time1 = time.time()
            total_time = time1 - time0
            total_time_with_2_decimals = round(total_time, 2)
            
            self.add_info_log(f"End processing orders")
            self.add_info_log(f"Total processing time: {total_time_with_2_decimals} s for {amountOfOrdersToProcess} orders")

            if amountOfOrdersToProcess == 0:
                avarage_time_with_2_decimals = 0.00
            else:
                avarage_time_with_2_decimals = round(total_time / amountOfOrdersToProcess, 2)

            self.add_info_log(f"Average processing time: {avarage_time_with_2_decimals} s")

            self.queue.put(self.ordersAndContactsDataframe)

            self.queue.put("UNBLOCK MAIN USERFORM WIDGETS")