import pandas as pd
import queue

class Controller:
    def __init__(self, model, view):
        self.model = model
        self.view = view

    def showMainUserForm(self) -> None:
        self.view.get_main_userform_root().after(100, self.check_queue)
        self.view.show_mainUserForm_and_connect_with_controller(controller=self)

    def destroyMainUserForm(self) -> None:
        self.view.destroyMainUserForm()

    def destroyLogInUserForm(self) -> None:
        self.view.hide_userform()

    def on_log_btn_click(self) -> None:
        self.view.on_log_btn_click(controller=self)

    def on_loadOrders_btn_click(self) -> None:
        selected_team_name = self.view.get_selected_team_name()
        selected_date = self.view.get_selected_date()

        self.model.on_loadOrders_btn_click(selected_team_name, selected_date)

    def on_clearOrders_btn_click(self) -> None:
        self.model.on_clearOrders_btn_click()
        self.view.on_clearOrders_btn_click()

    def on_processOrders_btn_click(self) -> None:
        self.view.on_processOrders_btn_click(self)

    def validate_login(self) -> None:
        username = self.view.get_username_from_logInUserForm()
        password = self.view.get_password_from_logInUserForm()

        if self.model.validate_login(username, password):
            # WebDriver keeps built
            self.on_login_successful()
        else:
            self.on_login_failed()

    def on_login_failed(self) -> None:
        self.model.on_login_failed()
        self.view.on_login_failed()

    def on_login_successful(self) -> None:
        self.view.destroyLogInUserForm()
        self.model.on_login_successful()

    def update_a_line_to_processed_of_represented_ordersAndContactsDataframe(self, index: int, tracking_number: str, return_tracking_number: str) -> None:
        self.view.update_a_line_to_processed_of_represented_ordersAndContactsDataframe(index, tracking_number, return_tracking_number)

    def getEmptyOrdersAndContactsData(self) -> pd.DataFrame:
        return self.model.getEmptyOrdersAndContactsData()

    def getTeamsNames(self) -> list:
        return self.model.getTeamsNames()

    def add_error_log(self, text: str) -> None:
        self.model.add_error_log(text)

    def add_warning_log(self, text: str) -> None:
        self.model.add_warning_log(text)

    def add_info_log(self, text: str) -> None:
        self.model.add_info_log(text)

    def print_logs(self) -> pd.DataFrame:
        return self.model.print_logs()

    def print_last_n_logs(self, n: int) -> pd.DataFrame:
        return self.model.print_last_n_logs(n)

    def config_button_on_click(self) -> None:
        self.view.config_button_on_click(controller = self)
        self.update_widgets_from_configUserForm()

    def update_widgets_from_configUserForm(self) -> None:
        team_name = self.view.get_selected_team_name_on_config()
        config = self.get_config_of_a_team(team_name)
        self.view.update_widgets_from_configUserForm(config)

    def get_config_of_a_team(self, teamName: str) -> str:
        return self.model.get_config_of_a_team(teamName)
    
    def on_click_save_config_button(self) -> None:
        teamName = self.view.get_selected_team_name_on_config()
        team_excel_path = self.view.get_team_excel_path_from_configUserForm()
        team_orders_sheet = self.view.get_team_orders_sheet_from_configUserForm()
        team_contacts_sheet = self.view.get_team_contacts_sheet_from_configUserForm()
        team_not_working_days_sheet = self.view.get_team_not_working_days_sheet_from_configUserForm()
        team_send_email_to_medical_centers = self.view.get_team_send_email_to_medical_centers_from_configUserForm()

        self.model.on_click_save_config_button(
            teamName,
            team_excel_path,
            team_orders_sheet,
            team_contacts_sheet,
            team_not_working_days_sheet,
            team_send_email_to_medical_centers)

    def on_open_excel_double_btn_click(self) -> None:
        temporal_selected_team_name = self.view.get_selected_team_name()
        self.model.on_open_excel_double_btn_click(temporal_selected_team_name)

    def on_export_logs_to_csv(self) -> None:
        try:
            self.model.on_export_logs_to_csv()
            self.view.show_success_export_to_csv()
        except Exception as e:
            self.add_error_log(f"Error exporting logs to csv: {e}")
            self.view.show_failure_export_to_csv()

    def check_queue(self) -> None:
        try:
            while True:
                task = self.model.queue.get_nowait()
                self.view.queue_action(task)
        except queue.Empty:
            pass
        self.view.get_main_userform_root().after(100, self.check_queue)