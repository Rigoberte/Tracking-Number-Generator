import os

import customtkinter as ctk
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog

from ..chroma import Chroma, UI


class ConfigUserForm(ctk.CTkToplevel):
    """Modern configuration form with customtkinter styling."""
    
    def __init__(self):
        super().__init__()
        
        self.colors = Chroma()
        if self.colors.getDarkMode():
            self.colors.toggle()
        
        self.title("Configuration")
        self.geometry("700x500")
        self.minsize(600, 450)
        
        # Center window
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.winfo_screenheight() // 2) - (500 // 2)
        self.geometry(f"700x500+{x}+{y}")
        
        try:
            self.iconbitmap(os.getcwd() + "\\media\\icon.ico")
        except:
            pass
        
        self.configure(fg_color=self.colors.theme.surface)
        
        self.send_email_var = tk.BooleanVar(value=False)
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
        return self.send_email_var.get()
    
    def get_team_email(self) -> str:
        return self.team_email.get()

    def show_userform(self) -> None:
        self.grab_set()
        self.focus_force()
        self.wait_window()

    def hide_userform(self) -> None:
        self.grab_release()
        self.destroy()

    def connect_with_controller(self, controller) -> None:
        self.teams_combobox.configure(values=controller.get_team_names())
        
        selected_team_name_on_mainUserForm = controller.get_selected_team_name_on_mainUserForm()
        team_names = controller.get_team_names()
        
        # Set current team
        if selected_team_name_on_mainUserForm in team_names:
            self.teams_combobox.set(selected_team_name_on_mainUserForm)
        elif team_names:
            self.teams_combobox.set(team_names[0])

        def on_team_change(choice):
            controller.update_widgets_from_configUserForm()

        self.teams_combobox.configure(command=on_team_change)
        self.save_button.configure(command=controller.on_click_save_config_button)
        
        controller.update_widgets_from_configUserForm()

    def update_widgets(self, config: dict) -> None:
        self.__set_configs__(config)

    def __set_configs__(self, config: dict) -> None:
        self.team_excel_path.delete(0, tk.END)
        self.team_excel_path.insert(0, config.get("team_excel_path", ""))

        self.team_orders_sheet.delete(0, tk.END)
        self.team_orders_sheet.insert(0, config.get("team_orders_sheet", ""))

        self.team_contacts_sheet.delete(0, tk.END)
        self.team_contacts_sheet.insert(0, config.get("team_contacts_sheet", ""))

        self.team_not_working_days_sheet.delete(0, tk.END)
        self.team_not_working_days_sheet.insert(0, config.get("team_not_working_days_sheet", ""))

        self.team_email.delete(0, tk.END)
        self.team_email.insert(0, config.get("team_email", ""))

        self.send_email_var.set(config.get("team_send_email_to_medical_centers", False))

    def __create_entry_with_label__(self, frame, label_text, with_file_dialog=False) -> ctk.CTkEntry:
        """Create a labeled entry field with optional file dialog button."""
        row_frame = ctk.CTkFrame(frame, fg_color="transparent")
        row_frame.pack(fill=tk.X, pady=UI.spacing_xs)
        
        label = ctk.CTkLabel(
            row_frame,
            text=label_text,
            font=(UI.font_family, UI.font_sm, "bold"),
            text_color=self.colors.theme.text_secondary,
            width=180,
            anchor="w",
        )
        label.pack(side=tk.LEFT, padx=(0, UI.spacing_sm))
        
        entry = ctk.CTkEntry(
            row_frame,
            height=UI.btn_height_sm,
            font=(UI.font_family, UI.font_sm),
            border_width=1,
            border_color=self.colors.theme.input_border,
            fg_color=self.colors.theme.input_bg,
            text_color=self.colors.theme.text_primary,
            corner_radius=UI.radius_sm,
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        if with_file_dialog:
            browse_btn = ctk.CTkButton(
                row_frame,
                text="...",
                width=40,
                height=UI.btn_height_sm,
                font=(UI.font_family, UI.font_md),
                fg_color=self.colors.theme.surface_hover,
                hover_color=self.colors.theme.input_border,
                text_color=self.colors.theme.text_primary,
                corner_radius=UI.radius_sm,
                command=lambda: self.__open_file_dialog__(entry),
            )
            browse_btn.pack(side=tk.RIGHT, padx=(UI.spacing_xs, 0))
        
        return entry

    def __open_file_dialog__(self, entry) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def __create_widgets__(self) -> None:
        # Main container
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill=tk.BOTH, expand=True, padx=UI.spacing_xl, pady=UI.spacing_lg)
        
        # Header
        header_frame = ctk.CTkFrame(container, fg_color="transparent")
        header_frame.pack(fill=tk.X, pady=(0, UI.spacing_lg))
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="Team Configuration",
            font=(UI.font_family, UI.font_2xl, "bold"),
            text_color=self.colors.theme.text_primary,
        )
        title_label.pack(side=tk.LEFT)
        
        # Team selector
        selector_frame = ctk.CTkFrame(container, fg_color="transparent")
        selector_frame.pack(fill=tk.X, pady=(0, UI.spacing_lg))
        
        team_label = ctk.CTkLabel(
            selector_frame,
            text="Select Team",
            font=(UI.font_family, UI.font_sm, "bold"),
            text_color=self.colors.theme.text_secondary,
            width=180,
            anchor="w",
        )
        team_label.pack(side=tk.LEFT, padx=(0, UI.spacing_sm))
        
        self.teams_combobox = ctk.CTkComboBox(
            selector_frame,
            values=[""],
            height=UI.btn_height_md,
            font=(UI.font_family, UI.font_md),
            dropdown_font=(UI.font_family, UI.font_sm),
            border_width=1,
            border_color=self.colors.theme.input_border,
            fg_color=self.colors.theme.input_bg,
            button_color=self.colors.theme.primary,
            button_hover_color=self.colors.theme.primary_hover,
            dropdown_fg_color=self.colors.theme.surface,
            dropdown_hover_color=self.colors.theme.surface_hover,
            text_color=self.colors.theme.text_primary,
            corner_radius=UI.radius_sm,
            state="readonly",
        )
        self.teams_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Separator
        separator = ctk.CTkFrame(container, height=1, fg_color=self.colors.theme.input_border)
        separator.pack(fill=tk.X, pady=UI.spacing_md)
        
        # Form fields container
        form_frame = ctk.CTkFrame(container, fg_color="transparent")
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Fields
        self.team_excel_path = self.__create_entry_with_label__(
            form_frame, "Excel Path", with_file_dialog=True
        )
        self.team_orders_sheet = self.__create_entry_with_label__(
            form_frame, "Orders Sheet"
        )
        self.team_contacts_sheet = self.__create_entry_with_label__(
            form_frame, "Contacts Sheet"
        )
        self.team_not_working_days_sheet = self.__create_entry_with_label__(
            form_frame, "Non-Working Days Sheet"
        )
        self.team_email = self.__create_entry_with_label__(
            form_frame, "Team Email"
        )
        
        # Checkbox
        checkbox_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        checkbox_frame.pack(fill=tk.X, pady=UI.spacing_md)
        
        self.send_email_checkbox = ctk.CTkCheckBox(
            checkbox_frame,
            text="Send email to medical centers",
            font=(UI.font_family, UI.font_sm),
            text_color=self.colors.theme.text_primary,
            fg_color=self.colors.theme.primary,
            hover_color=self.colors.theme.primary_hover,
            border_color=self.colors.theme.input_border,
            checkmark_color=self.colors.theme.text_on_primary,
            corner_radius=UI.radius_sm,
            variable=self.send_email_var,
        )
        self.send_email_checkbox.pack(side=tk.LEFT, padx=(180 + UI.spacing_sm, 0))
        
        # Bottom buttons
        bottom_frame = ctk.CTkFrame(container, fg_color="transparent")
        bottom_frame.pack(fill=tk.X, pady=(UI.spacing_lg, 0))
        
        self.save_button = ctk.CTkButton(
            bottom_frame,
            text="Save Configuration",
            height=UI.btn_height_lg,
            font=(UI.font_family, UI.font_lg, "bold"),
            fg_color=self.colors.theme.primary,
            hover_color=self.colors.theme.primary_hover,
            text_color=self.colors.theme.text_on_primary,
            corner_radius=UI.radius_md,
        )
        self.save_button.pack(fill=tk.X)