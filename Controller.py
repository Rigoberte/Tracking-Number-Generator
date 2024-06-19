import pandas as pd
import datetime as dt
import time
import os

from userForms.mains.mainUserForm import MyUserForm
from userForms.logins.logInUserForm import LogInUserForm
from userForms.mains.logConsole import LogConsole
from dataRecolector.dataRecolector import DataRecolector
from orderProcessor.orderProcessor import OrderProcessor
from teams.team import Team
from teams.team_factory import TeamFactory
from logClass.log import Log
from dataPathController.dataPathController import DataPathController
from userForms.mains.configUserForm import ConfigUserForm
from utils.utils import getFolderPathToDownload, create_folder

class Controller:
    def __init__(self):
        self.selected_team = TeamFactory().create_team("No Selected Team", "")
        self.ordersAndContactsDataframe = self.getEmptyOrdersAndContactsData()
        Log().add_log("Application started")

    def showMainUserForm(self) -> None:
        self.mainUserForm = MyUserForm(self)
        self.mainUserForm.show_userform()

    def destroyMainUserForm(self) -> None:
        self.mainUserForm.hide_userform()

    def showLogInUserForm(self) -> None:
        self.logInUserForm = LogInUserForm(self)
        self.logInUserForm.show_userform()

    def destroyLogInUserForm(self) -> None:
        self.logInUserForm.hide_userform()

    def on_log_btn_click(self) -> None:
        LogConsole(self).show_userform()

    def on_loadOrders_btn_click(self) -> None:
        selected_team_name = self.mainUserForm.getSelectedTeamName()
        selected_date = self.mainUserForm.getSelectedDate()

        self.__loadOrdersAndCalculateTime__(selected_team_name, selected_date)
        self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)

    def on_clearOrders_btn_click(self) -> None:
        self.ordersAndContactsDataframe = self.getEmptyOrdersAndContactsData()
        self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)
        Log().add_log("Orders table cleaned")

    def on_processOrders_btn_click(self) -> None:
        self.showLogInUserForm()

    def validate_login(self) -> None:
        username = self.logInUserForm.get_username()
        password = self.logInUserForm.get_password()

        self.selected_team.build_driver()
        if self.selected_team.check_if_user_and_password_are_correct(username, password):
            # WebDriver keeps built
            self.on_login_successful()
        else:
            self.logInUserForm.clear_password_entry()
            self.selected_team.quit_driver()
            self.on_login_failed()

    def on_login_failed(self) -> None:
        self.logInUserForm.show_login_failed()

    def on_login_successful(self) -> None:
        self.destroyLogInUserForm()
        
        self.ordersAndContactsDataframe = self.__processOrdersAndCalculateTime__(
                                        self.selected_team, self.selected_date,
                                        self.folder_path_to_download, self.ordersAndContactsDataframe)

        self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)

    def update_tag_color_of_a_treeview_line(self, index: int, tracking_number: str, return_tracking_number: str) -> None:
        self.mainUserForm.update_tag_color_of_a_treeview_line(self.ordersAndContactsDataframe, index, tracking_number, return_tracking_number)
        self.mainUserForm.increase_amount_of_orders_processed()

    def getEmptyOrdersAndContactsData(self) -> pd.DataFrame:
        no_selected_team = TeamFactory().create_team("No Selected Team", "")
        return DataRecolector(no_selected_team).getEmptyOrdersAndContactsData()

    def getTeamsNames(self) -> list:
        return ["Eli Lilly Argentina", "GPM Argentina", "Test_5_ordenes"]

    def addToLog(self, text: str) -> None:
        Log().add_log(text)

    def get_last_n_logs(self, n: int) -> list:
        return Log().print_last_n_logs(n)

    def config_button_on_click(self) -> None:
        self.configUserForm = ConfigUserForm(self)
        self.configUserForm.show_userform()

    def get_config_of_a_team(self, teamName: str) -> str:
        return DataPathController().get_config_of_a_team(teamName)
    
    def on_click_save_config_button(self) -> None:
        teamName = self.configUserForm.getSelectedTeamName()
        team_excel_path = self.configUserForm.getTeamExcelPath()
        team_orders_sheet = self.configUserForm.getTeamOrdersSheet()
        team_contacts_sheet = self.configUserForm.getTeamContactsSheet()
        team_not_working_days_sheet = self.configUserForm.getTeamNotWorkingDaysSheet()
        team_send_email_to_medical_centers = self.configUserForm.getSendEmailVar()

        DataPathController().redefine_a_config_of_a_team(teamName, 
            {
            "team_excel_path": team_excel_path,
            "team_orders_sheet": team_orders_sheet,
            "team_contacts_sheet": team_contacts_sheet,
            "team_not_working_days_sheet": team_not_working_days_sheet,
            "team_send_email_to_medical_centers": team_send_email_to_medical_centers
            }
        )

    def on_open_excel_double_btn_click(self) -> None:
        temporal_selected_team_name = self.mainUserForm.getSelectedTeamName()
        temporal_selected_team = TeamFactory().create_team(temporal_selected_team_name, "")

        excel_dataPath = temporal_selected_team.get_data_path(["team_excel_path"])

        try:
            if excel_dataPath[0] == "" or not os.path.exists(excel_dataPath[0]) or not excel_dataPath[0].endswith(".xlsx"):
                Log().add_log("Excel file path not found")
                return
                
            os.startfile(excel_dataPath[0])
        except Exception as e:
            Log().add_log(f"Error opening Excel file: {e}")

    def __loadOrders__(self, selected_team_name: str, selected_date: str) -> pd.DataFrame:
        self.folder_path_to_download = getFolderPathToDownload(selected_team_name, selected_date.replace("/", "_"))
        self.selected_team = TeamFactory().create_team(selected_team_name, self.folder_path_to_download)

        self.selected_date = dt.datetime.strptime(selected_date, '%Y-%m-%d')
        self.ordersAndContactsDataframe = DataRecolector(self.selected_team).recolectOrdersAndContactsData(self.selected_date)

        return self.ordersAndContactsDataframe

    def __loadOrdersAndCalculateTime__(self, selected_team_name: str, selected_date: str) -> pd.DataFrame:
        Log().add_separator()
        Log().add_log(f"Start loading orders for {selected_team_name} team")
        
        time0 = time.time()
        ordersAndContactsDataframe = self.__loadOrders__(selected_team_name, selected_date)
        time1 = time.time()
        total_time = time1 - time0
        total_time_with_2_decimals = round(total_time, 2)

        Log().add_log(f"End loading orders")
        Log().add_log(f"Total loading time: {total_time_with_2_decimals} s")

        return ordersAndContactsDataframe

    def __processOrders__(self, selected_team: Team, selected_date: dt.datetime, folder_path_to_download: str, ordersAndContactsDataframe: pd.DataFrame) -> pd.DataFrame:
        create_folder(folder_path_to_download)

        ordersAndContactsDataframe = OrderProcessor(folder_path_to_download, selected_team, self).processOrdersAndContactsTable(ordersAndContactsDataframe)
        
        selected_date_str = selected_date.strftime("%Y-%m-%d")
        
        selected_team.sendEmailWithOrdersToTeam(folder_path_to_download, selected_date_str)

        return ordersAndContactsDataframe

    def __processOrdersAndCalculateTime__(self, selected_team: Team, selected_date: dt.datetime, folder_path_to_download: str, ordersAndContactsDataframe: pd.DataFrame) -> pd.DataFrame:
        amountOfOrdersToProcess = len(ordersAndContactsDataframe[(ordersAndContactsDataframe["TRACKING_NUMBER"] == "") & (ordersAndContactsDataframe["HAS_AN_ERROR"] == "No error")])
        
        Log().add_log(f"Start processing orders")
        
        time0 = time.time()
        self.ordersAndContactsDataframe = self.__processOrders__(
                                        selected_team, selected_date,
                                        folder_path_to_download, ordersAndContactsDataframe)
        time1 = time.time()
        total_time = time1 - time0
        total_time_with_2_decimals = round(total_time, 2)
        
        Log().add_log(f"End processing orders")
        Log().add_log(f"Total processing time: {total_time_with_2_decimals} s for {amountOfOrdersToProcess} orders")

        if amountOfOrdersToProcess == 0:
            avarage_time_with_2_decimals = 0.00
        else:
            avarage_time_with_2_decimals = round(total_time / amountOfOrdersToProcess, 2)

        Log().add_log(f"Average processing time: {avarage_time_with_2_decimals} s")

        return ordersAndContactsDataframe
    
# Initializer
if __name__ == "__main__":
    Controller().showMainUserForm()