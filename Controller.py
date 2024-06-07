import pandas as pd
import datetime as dt
import time

from userForms.mains.mainUserForm import MyUserForm
from userForms.logins.logInUserForm import LogInUserForm
from userForms.mains.logConsole import LogConsole
from dataRecolector.dataRecolector import DataRecolector
from orderProcessor.orderProcessor import OrderProcessor
from teams.team import Team
from teams.team_factory import TeamFactory
from logClass.log import Log
from utils.utils import getFolderPathToDownload, create_folder

class Controller:
    def __init__(self):
        self.NO_SELECTED_TEAM = TeamFactory().create_team("No Selected Team", "")
        
        self.mainUserForm = MyUserForm(self)
        self.logInUserForm = LogInUserForm(self).hide_userform() # TODO I dont know why this userform is shown by default
        self.selected_team = self.NO_SELECTED_TEAM
        self.ordersAndContactsDataframe = DataRecolector(self.NO_SELECTED_TEAM).getEmptyOrdersAndContactsData()
        Log().add_log("Application started")

    def showMainUserForm(self) -> None:
        self.mainUserForm.show_userform()

    def destroyMainUserForm(self) -> None:
        self.mainUserForm.hide_userform()

    def showLogInUserForm(self) -> None:
        self.logInUserForm = LogInUserForm(self)
        self.logInUserForm.show_userform() # TODO here i should only show the userform

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
        self.ordersAndContactsDataframe = DataRecolector(self.NO_SELECTED_TEAM).getEmptyOrdersAndContactsData()
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
        return DataRecolector(self.NO_SELECTED_TEAM).getEmptyOrdersAndContactsData()

    def getTeamsNames(self) -> list:
        return ["Eli Lilly Argentina", "GPM Argentina", "Test_5_ordenes"]

    def addToLog(self, text: str) -> None:
        Log().add_log(text)

    def get_last_n_logs(self, n: int) -> list:
        return Log().print_last_n_logs(n)

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