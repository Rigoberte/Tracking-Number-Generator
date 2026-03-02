import os

import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

from ..chroma import Chroma, UI


class LogInUserForm(ctk.CTkToplevel):
    """Modern login form with customtkinter styling."""
    
    def __init__(self):
        super().__init__()
        
        self.colors = Chroma()
        # Set to light mode for login form
        if self.colors.getDarkMode():
            self.colors.toggle()
        
        self.title("Login")
        self.geometry("400x320")
        self.resizable(False, False)
        
        # Center window on screen
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.winfo_screenheight() // 2) - (320 // 2)
        self.geometry(f"400x320+{x}+{y}")
        
        try:
            self.iconbitmap(os.getcwd() + "\\media\\icon.ico")
        except:
            pass
        
        self.configure(fg_color=self.colors.theme.surface)
        
        self.__load_userform__()
    
    def get_username(self) -> str:
        return self.username_entry.get()
    
    def get_password(self) -> str:
        return self.password_entry.get()

    def clear_password_entry(self) -> None:
        self.password_entry.delete(0, 'end')

    def show_userform(self) -> None:
        self.grab_set()  # Make modal
        self.focus_force()
        self.wait_window()

    def hide_userform(self) -> None:
        self.grab_release()
        self.destroy()

    def show_login_failed(self) -> None:
        messagebox.showerror("Login Failed", "Username or Password incorrect")

    def __load_userform__(self) -> None:
        # Main container with padding
        container = ctk.CTkFrame(
            self,
            fg_color="transparent",
        )
        container.pack(fill=tk.BOTH, expand=True, padx=UI.spacing_xl, pady=UI.spacing_xl)
        
        # Title
        title_label = ctk.CTkLabel(
            container,
            text="Sign In",
            font=(UI.font_family, UI.font_3xl, "bold"),
            text_color=self.colors.theme.text_primary,
        )
        title_label.pack(pady=(0, UI.spacing_lg))
        
        # Subtitle
        subtitle_label = ctk.CTkLabel(
            container,
            text="Enter your credentials to continue",
            font=(UI.font_family, UI.font_sm),
            text_color=self.colors.theme.text_muted,
        )
        subtitle_label.pack(pady=(0, UI.spacing_xl))
        
        # Username field
        username_frame = ctk.CTkFrame(container, fg_color="transparent")
        username_frame.pack(fill=tk.X, pady=(0, UI.spacing_md))
        
        username_label = ctk.CTkLabel(
            username_frame,
            text="Username",
            font=(UI.font_family, UI.font_sm, "bold"),
            text_color=self.colors.theme.text_secondary,
            anchor="w",
        )
        username_label.pack(fill=tk.X)
        
        self.username_entry = ctk.CTkEntry(
            username_frame,
            height=UI.btn_height_md,
            font=(UI.font_family, UI.font_md),
            border_width=1,
            border_color=self.colors.theme.input_border,
            fg_color=self.colors.theme.input_bg,
            text_color=self.colors.theme.text_primary,
            corner_radius=UI.radius_md,
            placeholder_text="Enter username",
        )
        self.username_entry.pack(fill=tk.X, pady=(UI.spacing_xs, 0))
        
        # Password field
        password_frame = ctk.CTkFrame(container, fg_color="transparent")
        password_frame.pack(fill=tk.X, pady=(0, UI.spacing_lg))
        
        password_label = ctk.CTkLabel(
            password_frame,
            text="Password",
            font=(UI.font_family, UI.font_sm, "bold"),
            text_color=self.colors.theme.text_secondary,
            anchor="w",
        )
        password_label.pack(fill=tk.X)
        
        self.password_entry = ctk.CTkEntry(
            password_frame,
            height=UI.btn_height_md,
            font=(UI.font_family, UI.font_md),
            border_width=1,
            border_color=self.colors.theme.input_border,
            fg_color=self.colors.theme.input_bg,
            text_color=self.colors.theme.text_primary,
            corner_radius=UI.radius_md,
            placeholder_text="Enter password",
            show="*",
        )
        self.password_entry.pack(fill=tk.X, pady=(UI.spacing_xs, 0))
        
        # Buttons frame
        buttons_frame = ctk.CTkFrame(container, fg_color="transparent")
        buttons_frame.pack(fill=tk.X, pady=(UI.spacing_md, 0))
        
        # Login button
        self.login_button = ctk.CTkButton(
            buttons_frame,
            text="Login",
            height=UI.btn_height_lg,
            font=(UI.font_family, UI.font_lg, "bold"),
            fg_color=self.colors.theme.primary,
            hover_color=self.colors.theme.primary_hover,
            text_color=self.colors.theme.text_on_primary,
            corner_radius=UI.radius_md,
        )
        self.login_button.pack(fill=tk.X, pady=(0, UI.spacing_sm))
        
        # Exit button
        self.exit_button = ctk.CTkButton(
            buttons_frame,
            text="Cancel",
            height=UI.btn_height_md,
            font=(UI.font_family, UI.font_md),
            fg_color="transparent",
            hover_color=self.colors.theme.surface_hover,
            text_color=self.colors.theme.text_secondary,
            border_width=1,
            border_color=self.colors.theme.input_border,
            corner_radius=UI.radius_md,
            command=self.hide_userform,
        )
        self.exit_button.pack(fill=tk.X)

    def connect_with_controller(self, controller) -> None:
        def on_login_btn_click(event=None):
            if self.get_username() != "" and self.get_password() != "":
                controller.validate_login()

        self.username_entry.bind("<Return>", on_login_btn_click)
        self.password_entry.bind("<Return>", on_login_btn_click)

        self.login_button.configure(command=controller.validate_login)
        
        # Focus on username entry
        self.username_entry.focus_set()