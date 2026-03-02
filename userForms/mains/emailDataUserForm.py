import os

import customtkinter as ctk
import tkinter as tk

from ..chroma import Chroma, UI


class EmailDataUserForm(ctk.CTkToplevel):
    """Modern email data form for sender information."""
    
    def __init__(self):
        super().__init__()
        
        self.colors = Chroma()
        if self.colors.getDarkMode():
            self.colors.toggle()
        
        self.title("Sender Information")
        self.geometry("550x420")
        self.resizable(False, False)
        
        # Center window
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (550 // 2)
        y = (self.winfo_screenheight() // 2) - (420 // 2)
        self.geometry(f"550x420+{x}+{y}")
        
        try:
            self.iconbitmap(os.getcwd() + "\\media\\icon.ico")
        except:
            pass
        
        self.configure(fg_color=self.colors.theme.surface)
        
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
        self.grab_set()
        self.focus_force()
        self.wait_window()

    def hide_userform(self) -> None:
        self.grab_release()
        self.destroy()

    def connect_with_controller(self, controller) -> None:
        self.ok_button.configure(command=controller.confirm_email_sender)
        controller.update_widgets_from_emailDataForm()

    def update_widgets(self, config: dict) -> None:
        self.__set_configs__(config)

    def __set_configs__(self, config: dict) -> None:
        self.full_name.delete(0, tk.END)
        self.full_name.insert(0, config.get("full_name", ""))

        self.job_position.delete(0, tk.END)
        self.job_position.insert(0, config.get("job_position", ""))

        self.address.delete(0, tk.END)
        self.address.insert(0, config.get("site_address", ""))

        self.phone_number.delete(0, tk.END)
        self.phone_number.insert(0, config.get("phone_number", ""))

        self.email_address.delete(0, tk.END)
        self.email_address.insert(0, config.get("email_address", ""))

    def __create_entry_with_label__(self, parent, label_text) -> ctk.CTkEntry:
        """Create a modern labeled entry field."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill=tk.X, pady=UI.spacing_xs)
        
        label = ctk.CTkLabel(
            frame,
            text=label_text,
            font=(UI.font_family, UI.font_sm, "bold"),
            text_color=self.colors.theme.text_secondary,
            anchor="w",
        )
        label.pack(fill=tk.X)
        
        entry = ctk.CTkEntry(
            frame,
            height=UI.btn_height_md,
            font=(UI.font_family, UI.font_md),
            border_width=1,
            border_color=self.colors.theme.input_border,
            fg_color=self.colors.theme.input_bg,
            text_color=self.colors.theme.text_primary,
            corner_radius=UI.radius_md,
        )
        entry.pack(fill=tk.X, pady=(UI.spacing_xs, 0))
        
        return entry

    def __create_widgets__(self) -> None:
        # Main container
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill=tk.BOTH, expand=True, padx=UI.spacing_xl, pady=UI.spacing_xl)
        
        # Header
        title_label = ctk.CTkLabel(
            container,
            text="Email Signature",
            font=(UI.font_family, UI.font_2xl, "bold"),
            text_color=self.colors.theme.text_primary,
        )
        title_label.pack(anchor="w", pady=(0, UI.spacing_xs))
        
        subtitle_label = ctk.CTkLabel(
            container,
            text="This information will be used in emails sent to medical centers",
            font=(UI.font_family, UI.font_sm),
            text_color=self.colors.theme.text_muted,
        )
        subtitle_label.pack(anchor="w", pady=(0, UI.spacing_lg))
        
        # Form fields
        form_frame = ctk.CTkFrame(container, fg_color="transparent")
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        self.full_name = self.__create_entry_with_label__(form_frame, "Full Name")
        self.job_position = self.__create_entry_with_label__(form_frame, "Job Position")
        self.address = self.__create_entry_with_label__(form_frame, "Site Address")
        self.phone_number = self.__create_entry_with_label__(form_frame, "Phone Number")
        self.email_address = self.__create_entry_with_label__(form_frame, "Email Address")
        
        # Button
        button_frame = ctk.CTkFrame(container, fg_color="transparent")
        button_frame.pack(fill=tk.X, pady=(UI.spacing_lg, 0))
        
        self.ok_button = ctk.CTkButton(
            button_frame,
            text="Continue",
            height=UI.btn_height_lg,
            font=(UI.font_family, UI.font_lg, "bold"),
            fg_color=self.colors.theme.primary,
            hover_color=self.colors.theme.primary_hover,
            text_color=self.colors.theme.text_on_primary,
            corner_radius=UI.radius_md,
        )
        self.ok_button.pack(fill=tk.X)
        self.ok_button.pack(fill=tk.X)