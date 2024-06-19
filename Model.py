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
        self.addToLog("Application started")

    def on_loadOrders_btn_click(self, selected_team_name, selected_date) -> None:
        thread = threading.Thread(target=self.__loadOrdersAndCalculateTime__(selected_team_name, selected_date))
        thread.start()
        # self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)

    def on_clearOrders_btn_click(self) -> None:
        self.ordersAndContactsDataframe = self.getEmptyOrdersAndContactsData()
        self.addToLog("Orders table cleaned")
        # self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)

    def on_login_successful(self, selected_team, selected_date, folder_path_to_download) -> None:
        #self.destroyLogInUserForm()
        thread = threading.Thread( target= self.__processOrdersAndCalculateTime__( selected_team, selected_date, folder_path_to_download, self.ordersAndContactsDataframe) )
        thread.start()
        #self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)

    def getEmptyOrdersAndContactsData(self) -> pd.DataFrame:
        no_selected_team = TeamFactory().create_team("No Selected Team", "")
        return DataRecolector(no_selected_team).getEmptyOrdersAndContactsData()

    def getTeamsNames(self) -> list:
        return TeamFactory().getTeamsNames()

    def addToLog(self, text: str) -> None:
        self.log.add_log(text)

    def get_last_n_logs(self, n: int) -> list:
        return self.log.print_last_n_logs(n)

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
                self.addToLog("Excel file path not found")
                return
                
            os.startfile(excel_dataPath[0])
        except Exception as e:
            self.addToLog(f"Error opening Excel file: {e}")

    def __loadOrders__(self, selected_team_name: str, selected_date: str) -> pd.DataFrame:
        self.folder_path_to_download = getFolderPathToDownload(selected_team_name, selected_date.replace("/", "_"))
        self.selected_team = TeamFactory().create_team(selected_team_name, self.folder_path_to_download)

        self.selected_date = dt.datetime.strptime(selected_date, '%Y-%m-%d')
        self.ordersAndContactsDataframe = DataRecolector(self.selected_team).recolectOrdersAndContactsData(self.selected_date)

        return self.ordersAndContactsDataframe

    def __loadOrdersAndCalculateTime__(self, selected_team_name: str, selected_date: str) -> pd.DataFrame:
        self.log.add_separator()
        self.addToLog(f"Start loading orders for {selected_team_name} team")
        
        time0 = time.time()
        ordersAndContactsDataframe = self.__loadOrders__(selected_team_name, selected_date)
        time1 = time.time()
        total_time = time1 - time0
        total_time_with_2_decimals = round(total_time, 2)

        self.addToLog(f"End loading orders")
        self.addToLog(f"Total loading time: {total_time_with_2_decimals} s")

        self.queue.put("loadOrdersAndCalculateTime finished")

        return ordersAndContactsDataframe

    def __processOrders__(self, selected_team: Team, selected_date: dt.datetime, folder_path_to_download: str, ordersAndContactsDataframe: pd.DataFrame) -> pd.DataFrame:
        create_folder(folder_path_to_download)

        ordersAndContactsDataframe = OrderProcessor(folder_path_to_download, selected_team, self.queue).processOrdersAndContactsTable(ordersAndContactsDataframe)
        
        selected_date_str = selected_date.strftime("%Y-%m-%d")
        
        selected_team.sendEmailWithOrdersToTeam(folder_path_to_download, selected_date_str)

        return ordersAndContactsDataframe

    def __processOrdersAndCalculateTime__(self, selected_team: Team, selected_date: dt.datetime, folder_path_to_download: str, ordersAndContactsDataframe: pd.DataFrame) -> None:
        amountOfOrdersToProcess = len(ordersAndContactsDataframe[(ordersAndContactsDataframe["TRACKING_NUMBER"] == "") & (ordersAndContactsDataframe["HAS_AN_ERROR"] == "No error")])
        
        self.addToLog(f"Start processing orders")
        
        time0 = time.time()
        self.ordersAndContactsDataframe = self.__processOrders__(
                                        selected_team, selected_date,
                                        folder_path_to_download, ordersAndContactsDataframe)
        time1 = time.time()
        total_time = time1 - time0
        total_time_with_2_decimals = round(total_time, 2)
        
        self.addToLog(f"End processing orders")
        self.addToLog(f"Total processing time: {total_time_with_2_decimals} s for {amountOfOrdersToProcess} orders")

        if amountOfOrdersToProcess == 0:
            avarage_time_with_2_decimals = 0.00
        else:
            avarage_time_with_2_decimals = round(total_time / amountOfOrdersToProcess, 2)

        self.addToLog(f"Average processing time: {avarage_time_with_2_decimals} s")

        self.queue.put("processOrdersAndCalculateTime finished")