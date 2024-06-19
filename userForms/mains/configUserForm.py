import customtkinter as ctk
import tkinter as tk
import tkinter.ttk as ttk

class ConfigUserForm(tk.Tk):
    def __init__(self, controller=None):
        super().__init__()
        self.controller = controller

        self.title("Config")
        self.geometry("800x300")

        self.send_email_var = tk.BooleanVar()

    def getSelectedTeamName(self) -> str:
        return self.teams_combobox.get()
    
    def getTeamExcelPath(self) -> str:
        return self.team_excel_path.get()
    
    def getTeamOrdersSheet(self) -> str:
        return self.team_orders_sheet.get()
    
    def getTeamContactsSheet(self) -> str:
        return self.team_contacts_sheet.get()
    
    def getTeamNotWorkingDaysSheet(self) -> str:
        return self.team_not_working_days_sheet.get()
    
    def getSendEmailVar(self) -> bool:
        return self.send_email_var.get()

    def show_userform(self) -> None:
        self.__create_widgets__()
        self.__update_widgets__()
        self.mainloop()

    def hide_userform(self) -> None:
        self.destroy()

    def __set_configs__(self, team_name: str) -> None:
        config = self.__get_config_of_a_team__(team_name)
        
        self.team_excel_path.delete(0, tk.END)
        self.team_excel_path.insert(0, config["team_excel_path"])

        self.team_orders_sheet.delete(0, tk.END)
        self.team_orders_sheet.insert(0, config["team_orders_sheet"])

        self.team_contacts_sheet.delete(0, tk.END)
        self.team_contacts_sheet.insert(0, config["team_contacts_sheet"])

        self.team_not_working_days_sheet.delete(0, tk.END)
        self.team_not_working_days_sheet.insert(0, config["team_not_working_days_sheet"])

        self.send_email_var.set(config["team_send_email_to_medical_centers"])

    def __update_widgets__(self) -> None:
        team_name = self.teams_combobox.get()
        self.__set_configs__(team_name)

    def __on_click_teams_combobox__(self, event) -> None:
        self.__update_widgets__()

    def __get_config_of_a_team__(self, team_name: str) -> None:
        return self.controller.get_config_of_a_team(team_name)

    def __create_frames__(self) -> dict:
        frames = {}

        top_frame = ctk.CTkFrame(self)
        top_frame.pack(side=tk.TOP, fill=tk.X, expand=True)
        frames["top_frame"] = top_frame

        bottom_frame = ctk.CTkFrame(self)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

        excel_path_frame = ctk.CTkFrame(bottom_frame)
        excel_path_frame.pack(side=tk.TOP, expand=True)
        frames["excel_path_frame"] = excel_path_frame

        orders_sheet_frame = ctk.CTkFrame(bottom_frame)
        orders_sheet_frame.pack(side=tk.TOP, expand=True)
        frames["orders_sheet_frame"] = orders_sheet_frame

        contacts_sheet_frame = ctk.CTkFrame(bottom_frame)
        contacts_sheet_frame.pack(side=tk.TOP, expand=True)
        frames["contacts_sheet_frame"] = contacts_sheet_frame

        not_working_days_sheet_frame = ctk.CTkFrame(bottom_frame)
        not_working_days_sheet_frame.pack(side=tk.TOP, expand=True)
        frames["not_working_days_sheet_frame"] = not_working_days_sheet_frame

        send_email_button_frame = ctk.CTkFrame(bottom_frame)
        send_email_button_frame.pack(side=tk.TOP, expand=True)
        frames["send_email_button_frame"] = send_email_button_frame

        bottom_bottom_frame = ctk.CTkFrame(bottom_frame)
        bottom_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, expand=True)
        frames["bottom_bottom_frame"] = bottom_bottom_frame

        return frames

    def __create_entry_with_label__(self, frame, label_text, entry_text, and_button_to_open_file_dialog=False) -> ttk.Entry:
        label = ttk.Label(frame, text=label_text, width=30)
        label.pack(side=tk.LEFT)

        if and_button_to_open_file_dialog:
            button = ttk.Button(frame, text="...", command=lambda: self.__open_file_dialog__(entry), width=2)
            button.pack(side=tk.RIGHT)

        entry = ttk.Entry(frame, width=70, takefocus=False)
        entry.pack(side=tk.RIGHT)
        entry.insert(0, entry_text)

        return entry

    def __open_file_dialog__(self, entry) -> None:
        file_path = tk.filedialog.askopenfilename()
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

    def __create_widgets__(self) -> None:
        frames = self.__create_frames__()
        
        teams = self.controller.getTeamsNames()

        self.teams_combobox = ttk.Combobox(frames["top_frame"], values=teams)
        self.teams_combobox.pack(fill=tk.X, expand=True, side=tk.RIGHT)
        self.teams_combobox.set(teams[0])

        self.team_excel_path = self.__create_entry_with_label__(frames["excel_path_frame"], "Excel path", "", and_button_to_open_file_dialog=True)

        self.team_orders_sheet = self.__create_entry_with_label__(frames["orders_sheet_frame"], "Orders sheet", "")

        self.team_contacts_sheet = self.__create_entry_with_label__(frames["contacts_sheet_frame"], "Contacts sheet", "")

        self.team_not_working_days_sheet = self.__create_entry_with_label__(frames["not_working_days_sheet_frame"], "Not working days sheet", "")

        self.teams_combobox.bind("<<ComboboxSelected>>", self.__on_click_teams_combobox__)

        self.send_email_toggle_button = ttk.Checkbutton(frames["send_email_button_frame"], text="Send email", variable=self.send_email_var)
        self.send_email_toggle_button.pack(fill=tk.X)

        save_button = ctk.CTkButton(frames["bottom_bottom_frame"], text="Save", command=self.__on_click_save__)
        save_button.pack(fill=tk.X)

    def __on_click_save__(self) -> None:
        self.controller.on_click_save_config_button()
    
if __name__ == "__main__":
    ConfigUserForm().show_userform()