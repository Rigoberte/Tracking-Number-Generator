import os

import customtkinter as ctk
import tkinter as tk
import tkinter.ttk as ttk

class ConfigUserForm(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Config")
        self.geometry("800x300")
        self.iconbitmap(os.getcwd() + "\\media\\icon.ico")

        self.send_email_var = tk.IntVar()
        self.__create_widgets__()

    def get_selected_team_name(self) -> str:
        return self.teams_combobox.get()
    
    def get_team_excel_path(self) -> str:
        return self.team_excel_path.get()
    
    def get_team_orders_sheet(self) -> str:
        return self.team_orders_sheet.get()
    
    def get_team_contacts_sheet(self) -> str:
        return self.team_contacts_sheet.get()
    
    def get_team_not_working_days_sheet(self) -> str:
        return self.team_not_working_days_sheet.get()
    
    def get_team_send_email_to_medical_centers(self) -> bool:
        return self.send_email_var.get() == 1
    
    def get_team_email(self) -> str:
        return self.team_email.get()

    def show_userform(self) -> None:
        self.mainloop()

    def hide_userform(self) -> None:
        self.destroy()

    def connect_with_controller(self, controller) -> None:
        self.teams_combobox['values'] = controller.get_team_names()
        
        selected_team_name_on_mainUserForm = controller.get_selected_team_name_on_mainUserForm()
        self.teams_combobox.current(0)
        for i, team_name in enumerate(self.teams_combobox['values']):
            if team_name == selected_team_name_on_mainUserForm:
                self.teams_combobox.current(i)
                break

        def __on_click_teams_combobox__(event):
            controller.update_widgets_from_configUserForm()

        self.teams_combobox.bind("<<ComboboxSelected>>", __on_click_teams_combobox__)
        
        self.save_button.configure(command=controller.on_click_save_config_button)

        self.send_email_toggle_button.invoke()
        self.send_email_toggle_button.invoke()

        controller.update_widgets_from_configUserForm()

    def update_widgets(self, config: dict) -> None:
        self.__set_configs__(config)

    def __set_configs__(self, config: dict) -> None:
        self.team_excel_path.delete(0, tk.END)
        self.team_excel_path.insert(0, config["team_excel_path"])

        self.team_orders_sheet.delete(0, tk.END)
        self.team_orders_sheet.insert(0, config["team_orders_sheet"])

        self.team_contacts_sheet.delete(0, tk.END)
        self.team_contacts_sheet.insert(0, config["team_contacts_sheet"])

        self.team_not_working_days_sheet.delete(0, tk.END)
        self.team_not_working_days_sheet.insert(0, config["team_not_working_days_sheet"])

        self.team_email.delete(0, tk.END)
        self.team_email.insert(0, config["team_email"])

        if config["team_send_email_to_medical_centers"]:
            self.send_email_var.set(1)
        else:
            self.send_email_var.set(0)
        
        if config['team_send_email_to_medical_centers'] != (self.send_email_toggle_button.instate(['selected'])):
            self.send_email_toggle_button.invoke()

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

        team_email_frame = ctk.CTkFrame(bottom_frame)
        team_email_frame.pack(side=tk.TOP, expand=True)
        frames["team_email_frame"] = team_email_frame

        send_email_button_frame = ctk.CTkFrame(bottom_frame)
        send_email_button_frame.pack(side=tk.TOP, expand=True)
        frames["send_email_button_frame"] = send_email_button_frame

        bottom_bottom_frame = ctk.CTkFrame(bottom_frame)
        bottom_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, expand=True)
        frames["bottom_bottom_frame"] = bottom_bottom_frame

        return frames

    def __create_entry_with_label__(self, frame, label_text, and_button_to_open_file_dialog=False) -> ttk.Entry:
        label = ttk.Label(frame, text=label_text, width=30)
        label.pack(side=tk.LEFT)

        if and_button_to_open_file_dialog:
            button = ttk.Button(frame, text="...", command=lambda: self.__open_file_dialog__(entry), width=2)
            button.pack(side=tk.RIGHT)

        entry = ttk.Entry(frame, width=70, takefocus=False)
        entry.pack(side=tk.RIGHT)
        entry.insert(0, "")

        return entry

    def __open_file_dialog__(self, entry) -> None:
        file_path = tk.filedialog.askopenfilename()
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

    def __create_widgets__(self) -> None:
        frames = self.__create_frames__()
        
        self.teams_combobox = ttk.Combobox(frames["top_frame"], values=[""], font=30, state='readonly')
        self.teams_combobox.pack(fill=tk.X, expand=True, side=tk.RIGHT)
        self.teams_combobox.set("")

        self.team_excel_path = self.__create_entry_with_label__(frames["excel_path_frame"], "Excel path", and_button_to_open_file_dialog=True)

        self.team_orders_sheet = self.__create_entry_with_label__(frames["orders_sheet_frame"], "Orders sheet")

        self.team_contacts_sheet = self.__create_entry_with_label__(frames["contacts_sheet_frame"], "Contacts sheet")

        self.team_not_working_days_sheet = self.__create_entry_with_label__(frames["not_working_days_sheet_frame"], "Not working days sheet")

        self.team_email = self.__create_entry_with_label__(frames["team_email_frame"], "Team email")

        self.send_email_toggle_button = ttk.Checkbutton(frames["send_email_button_frame"], text="Send email", variable=self.send_email_var)
        self.send_email_toggle_button.pack(fill=tk.X)

        self.save_button = ctk.CTkButton(frames["bottom_bottom_frame"], text="Save")
        self.save_button.pack(fill=tk.X)