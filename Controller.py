import pandas as pd
import queue

class Controller:
    def __init__(self, model, view):
        self.model = model
        self.view = view

    # Show and destroy userforms
    def show_mainUserForm(self) -> None:
        self.view.get_main_userform_root().after(100, self.check_queue)
        self.view.show_mainUserForm(controller=self)

    def destroy_mainUserForm(self) -> None:
        self.view.destroy_mainUserForm()

    # MainUserForm methods
    def on_log_btn_click(self) -> None:
        self.view.on_log_btn_click(controller=self)

    def on_loadOrders_btn_click(self) -> None:
        selected_team_name = self.view.get_selected_team_name_from_mainUserForm()
        selected_date = self.view.get_selected_date_from_mainUserForm()

        self.model.on_loadOrders_btn_click(selected_team_name, selected_date)

    def on_clearOrders_btn_click(self) -> None:
        self.model.on_clearOrders_btn_click()
        self.view.on_clearOrders_btn_click()

    def on_processOrders_btn_click(self) -> None:
        self.view.on_processOrders_btn_click(self)

    def config_button_on_click(self) -> None:
        self.view.config_button_on_click(controller = self)
        self.update_widgets_from_configUserForm()
    
    def on_open_excel_double_btn_click(self) -> None:
        temporal_selected_team_name = self.view.get_selected_team_name_from_mainUserForm()
        self.model.on_open_excel_double_btn_click(temporal_selected_team_name)

    def update_a_line_to_processed_of_represented_ordersAndContactsDataframe(self, index: int, tracking_number: str, return_tracking_number: str) -> None:
        self.view.update_a_line_to_processed_of_represented_ordersAndContactsDataframe(index, tracking_number, return_tracking_number)

    def get_selected_team_name_on_mainUserForm(self) -> str:
        return self.view.get_selected_team_name_from_mainUserForm()
    
    # ConfigUserForm methods
    def on_click_save_config_button(self) -> None:
        teamName = self.view.get_selected_team_name_from_configUserForm()
        team_excel_path = self.view.get_team_excel_path_from_configUserForm()
        team_orders_sheet = self.view.get_team_orders_sheet_from_configUserForm()
        team_contacts_sheet = self.view.get_team_contacts_sheet_from_configUserForm()
        team_not_working_days_sheet = self.view.get_team_not_working_days_sheet_from_configUserForm()
        team_send_email_to_medical_centers = self.view.get_team_send_email_to_medical_centers_from_configUserForm()
        team_email = self.view.get_team_email_from_configUserForm()

        self.model.on_click_save_config_button(
            teamName,
            team_excel_path,
            team_orders_sheet,
            team_contacts_sheet,
            team_not_working_days_sheet,
            team_send_email_to_medical_centers,
            team_email)

    def update_widgets_from_configUserForm(self) -> None:
        team_name = self.view.get_selected_team_name_from_configUserForm()
        config = self.model.get_config_of_a_team(team_name)
        self.view.update_widgets_from_configUserForm(config)

    # ConfigUserForm buttons methods
    def on_export_logs_to_csv(self) -> None:
        try:
            self.model.on_export_logs_to_csv()
            self.view.show_success_export_to_csv()
        except Exception as e:
            self.add_error_log(f"Error exporting logs to csv: {e}")
            self.view.show_failure_export_to_csv()

    # LogInUserForm buttons methods
    def validate_login(self) -> None:
        username = self.view.get_username_from_logInUserForm()
        password = self.view.get_password_from_logInUserForm()

        if self.model.validate_login(username, password):
            # WebDriver keeps built
            self.model.on_login_successful()
            self.view.destroy_logInUserForm()
        else:
            self.model.on_login_failed()
            self.view.on_login_failed()

    # Getters from Model
    def get_empty_ordersAndContactsData(self) -> pd.DataFrame:
        return self.model.get_empty_ordersAndContactsData()

    def get_team_names(self) -> list:
        return self.model.get_team_names()

    # Log methods
    def print_logs(self) -> pd.DataFrame:
        return self.model.print_logs()

    def print_last_n_logs(self, n: int) -> pd.DataFrame:
        return self.model.print_last_n_logs(n)

    # Check if there are tasks in the queue
    def check_queue(self) -> None:
        try:
            while True:
                if self.model.queue.empty():
                    break
                
                task = self.model.queue.get_nowait()
                self.view.queue_action(task)
        except queue.Empty:
            pass
        self.view.get_main_userform_root().after(100, self.check_queue)