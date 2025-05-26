import pandas as pd

from userForms.mains.mainUserForm import MyUserForm
from userForms.logins.logInUserForm import LogInUserForm
from userForms.mains.logConsole import LogConsole
from userForms.mains.configUserForm import ConfigUserForm
from userForms.mains.emailDataUserForm import EmailDataUserForm

class View:
    def __init__(self):
        self.mainUserForm = MyUserForm()

    def queue_action(self, objectOrInstruction) -> None:
        if type(objectOrInstruction) == pd.DataFrame:
            self.update_ordersAndContactsDataframe_and_widgets(objectOrInstruction)
        elif type(objectOrInstruction) == dict:
            self.update_a_line_to_processed_of_represented_ordersAndContactsDataframe(objectOrInstruction["INDEX"],
                                                                                    objectOrInstruction["TRACKING_NUMBER"], 
                                                                                    objectOrInstruction["RETURN_TRACKING_NUMBER"])
        elif type(objectOrInstruction) == str:
            if objectOrInstruction == "BLOCK MAIN USERFORM WIDGETS":
                self.mainUserForm.block_widgets()
            elif objectOrInstruction == "UNBLOCK MAIN USERFORM WIDGETS":
                self.mainUserForm.unblock_widgets()

    # Show and hide userForms
    def show_mainUserForm(self, *, controller) -> None:
        self.mainUserForm.connect_commands_with_controller(controller)
        self.mainUserForm.show_userform()

    def destroy_mainUserForm(self) -> None:
        self.mainUserForm.hide_userform()

    def show_logInUserForm(self, controller) -> None:
        self.logInUserForm = LogInUserForm()
        self.logInUserForm.connect_with_controller(controller)
        self.logInUserForm.show_userform()

    def destroy_logInUserForm(self) -> None:
        self.logInUserForm.hide_userform()

    def show_logConsole(self, controller) -> None:
        self.logconsole = LogConsole()
        self.logconsole.connect_with_controller(controller = controller)
        self.logconsole.show_userform()

    def show_configUserForm(self, controller) -> None:
        self.configUserForm = ConfigUserForm()
        self.configUserForm.connect_with_controller(controller = controller)
        self.configUserForm.show_userform()

    def show_emailDataUserForm(self, controller) -> None:
        self.emailDataUserForm = EmailDataUserForm()
        self.emailDataUserForm.connect_with_controller(controller = controller)
        self.emailDataUserForm.show_userform()

    def destroy_emailDataUserForm(self) -> None:
        self.emailDataUserForm.hide_userform()

    def update_widgets_from_emailDataUserForm(self, config: dict) -> None:
        self.emailDataUserForm.update_widgets(config)

    # Getters from MainUserForm
    def get_selected_team_name_from_mainUserForm(self) -> str:
        return self.mainUserForm.get_selected_team_name()

    def get_selected_date_from_mainUserForm(self) -> str:
        return self.mainUserForm.get_selected_date()

    # Getters from LogInUserForm
    def get_username_from_logInUserForm(self) -> str:
        return self.logInUserForm.get_username()
    
    def get_password_from_logInUserForm(self) -> str:
        return self.logInUserForm.get_password()
    
    def get_main_userform_root(self):
        return self.mainUserForm.get_root()
    
    # Getters from ConfigUserForm
    def get_selected_team_name_from_configUserForm(self) -> str:
        return self.configUserForm.get_selected_team_name()
    
    def get_team_excel_path_from_configUserForm(self) -> str:
        return self.configUserForm.get_team_excel_path()
    
    def get_team_orders_sheet_from_configUserForm(self) -> str:
        return self.configUserForm.get_team_orders_sheet()
    
    def get_team_contacts_sheet_from_configUserForm(self) -> str:
        return self.configUserForm.get_team_contacts_sheet()
    
    def get_team_not_working_days_sheet_from_configUserForm(self) -> str:
        return self.configUserForm.get_team_not_working_days_sheet()
    
    def get_team_send_email_to_medical_centers_from_configUserForm(self) -> bool:
        return self.configUserForm.get_team_send_email_to_medical_centers()
    
    def get_team_email_from_configUserForm(self) -> str:
        return self.configUserForm.get_team_email()

    # MainUserForm buttons methods
    def on_log_btn_click(self, controller) -> None:
        self.show_logConsole(controller = controller)

    def on_loadOrders_btn_click(self, ordersAndContactsDataframe: pd.DataFrame) -> None:
        self.update_ordersAndContactsDataframe_and_widgets(ordersAndContactsDataframe)
        self.mainUserForm.unblock_processOrders_btn()

    def on_clearOrders_btn_click(self) -> None:
        self.mainUserForm.block_processOrders_btn()
        
    def on_processOrders_btn_click(self, controller) -> None:
        self.show_logInUserForm(controller = controller)

    def config_button_on_click(self, controller) -> None:
        self.show_configUserForm(controller = controller)

    # Getters from EmailDataUserForm
    def get_full_name_from_emailDataUserForm(self) -> str:
        return self.emailDataUserForm.get_full_name()
    
    def get_job_position_from_emailDataUserForm(self) -> str:
        return self.emailDataUserForm.get_job_position()
    
    def get_address_from_emailDataUserForm(self) -> str:
        return self.emailDataUserForm.get_address()
    
    def get_phone_number_from_emailDataUserForm(self) -> str:
        return self.emailDataUserForm.get_phone_number()
    
    def get_email_address_from_emailDataUserForm(self) -> str:
        return self.emailDataUserForm.get_email_address()

    # LogInUserForm methods
    def on_login_failed(self) -> None:
        self.logInUserForm.clear_password_entry()
        self.logInUserForm.show_login_failed()

    def on_login_successful(self, ordersAndContactsDataframe: pd.DataFrame) -> None:
        self.update_ordersAndContactsDataframe_and_widgets(ordersAndContactsDataframe)

    # Update widgets from MainUserForm
    def update_ordersAndContactsDataframe_and_widgets(self, ordersAndContactsDataframe: pd.DataFrame) -> None:
        self.mainUserForm.update_whole_represented_ordersAndContactsDataframe(ordersAndContactsDataframe)

    def update_a_line_to_processed_of_represented_ordersAndContactsDataframe(self, index: int, tracking_number: str, return_tracking_number: str) -> None:
        self.mainUserForm.update_a_line_to_processed_of_represented_ordersAndContactsDataframe(index, tracking_number, return_tracking_number)

    # Update widgets from ConfigUserForm
    def update_widgets_from_configUserForm(self, config: dict) -> None:
        self.configUserForm.update_widgets(config)

    # ConfigUserForm methods
    def show_success_export_to_csv(self) -> None:
        self.logconsole.show_success_export_to_csv()

    def show_failure_export_to_csv(self) -> None:
        self.logconsole.show_failure_export_to_csv()