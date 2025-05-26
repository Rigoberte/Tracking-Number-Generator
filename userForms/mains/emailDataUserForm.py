import os

import customtkinter as ctk
import tkinter as tk
import tkinter.ttk as ttk

class EmailDataUserForm(tk.Toplevel):
    def __init__(self):
        super().__init__()
        
        self.title("Config")
        self.geometry("800x300")
        self.iconbitmap(os.getcwd() + "\\media\\icon.ico")

        self.__create_widgets__()

        self.focus_force()

    def get_full_name(self) -> str:
        return self.full_name.get()
    
    def get_job_position(self) -> str:
        return self.job_position.get()
    
    def get_address(self) -> str:
        return self.address.get()
    
    def get_phone_number(self) -> str:
        return self.phone_number.get()
    
    def get_email_address(self) -> str:
        return self.email_address.get()

    def show_userform(self) -> None:
        self.mainloop()

    def hide_userform(self) -> None:
        self.destroy()

    def connect_with_controller(self, controller) -> None:
        self.ok_button.configure(command=controller.confirm_email_sender)

        # get user data from a JSON
        controller.update_widgets_from_emailDataForm()

    def update_widgets(self, config: dict) -> None:
        self.__set_configs__(config)

    def __set_configs__(self, config: dict) -> None:
        self.full_name.delete(0, tk.END)
        self.full_name.insert(0, config["full_name"])

        self.job_position.delete(0, tk.END)
        self.job_position.insert(0, config["job_position"])

        self.address.delete(0, tk.END)
        self.address.insert(0, config["site_address"])

        self.phone_number.delete(0, tk.END)
        self.phone_number.insert(0, config["phone_number"])

        self.email_address.delete(0, tk.END)
        self.email_address.insert(0, config["email_address"])

    def __create_frames__(self) -> dict:
        frames = {}

        bottom_frame = ctk.CTkFrame(self)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

        full_name_frame = ctk.CTkFrame(bottom_frame)
        full_name_frame.pack(side=tk.TOP, expand=True)
        frames["full_name_frame"] = full_name_frame

        job_position_frame = ctk.CTkFrame(bottom_frame)
        job_position_frame.pack(side=tk.TOP, expand=True)
        frames["job_position_frame"] = job_position_frame

        site_address_frame = ctk.CTkFrame(bottom_frame)
        site_address_frame.pack(side=tk.TOP, expand=True)
        frames["site_address_frame"] = site_address_frame

        phone_number_frame = ctk.CTkFrame(bottom_frame)
        phone_number_frame.pack(side=tk.TOP, expand=True)
        frames["phone_number_frame"] = phone_number_frame

        email_address_frame = ctk.CTkFrame(bottom_frame)
        email_address_frame.pack(side=tk.TOP, expand=True)
        frames["email_address_frame"] = email_address_frame

        bottom_bottom_frame = ctk.CTkFrame(bottom_frame)
        bottom_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, expand=True)
        frames["bottom_bottom_frame"] = bottom_bottom_frame

        return frames

    def __create_entry_with_label__(self, frame, label_text) -> ttk.Entry:
        label = ttk.Label(frame, text=label_text, width=30)
        label.pack(side=tk.LEFT)

        entry = ttk.Entry(frame, width=70, takefocus=False)
        entry.pack(side=tk.RIGHT)
        entry.insert(0, "")

        return entry

    def __create_widgets__(self) -> None:
        frames = self.__create_frames__()
        
        self.full_name = self.__create_entry_with_label__(frames["full_name_frame"], "Full name")

        self.job_position = self.__create_entry_with_label__(frames["job_position_frame"], "Job Position")

        self.address = self.__create_entry_with_label__(frames["site_address_frame"], "Site Address")

        self.phone_number = self.__create_entry_with_label__(frames["phone_number_frame"], "Phone Number")

        self.email_address = self.__create_entry_with_label__(frames["email_address_frame"], "Email Adress")

        self.ok_button = ctk.CTkButton(frames["bottom_bottom_frame"], text="Ok")
        self.ok_button.pack(fill=tk.X)