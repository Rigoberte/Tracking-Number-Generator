import os

import tkinter as tk
from tkinter import messagebox

class LogInUserForm(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Login Form")
        self.geometry("300x150")
        self.iconbitmap(os.getcwd() + "\\media\\icon.ico")
        
        self.__load_userform__(self)
    
    def get_username(self) -> str:
        return self.username_entry.get()
    
    def get_password(self) -> str:
        return self.password_entry.get()

    def clear_password_entry(self) -> None:
        self.password_entry.delete(0, 'end')

    def show_userform(self) -> None:
        self.mainloop()

    def hide_userform(self) -> None:
        self.destroy()

    def show_login_failed(self) -> None:
        messagebox.showerror("Login Failed", "Username or Password incorrect")

    def __load_userform__(self, root) -> None:
        # Username Label and Entry
        self.username_label = tk.Label(root, text="Username")
        self.username_label.pack(pady=5)
        self.username_entry = tk.Entry(root)
        self.username_entry.pack(pady=5)
        
        # Password Label and Entry
        self.password_label = tk.Label(root, text="Password")
        self.password_label.pack(pady=5)
        self.password_entry = tk.Entry(root, show="*")
        self.password_entry.pack(pady=5)

        # Login Button
        self.login_button = tk.Button(root, text="Login")
        self.login_button.pack(pady=5)
        
        # Exit Button
        self.exit_button = tk.Button(root, text="Exit", command=root.quit)
        self.exit_button.pack(pady=5)

    def connect_with_controller(self, controller) -> None:
        def on_login_btn_click(event):
            if self.get_username() != "" and self.get_password() != "":
                controller.validate_login()

        self.username_entry.bind("<Return>", on_login_btn_click)
        self.password_entry.bind("<Return>", on_login_btn_click)

        self.login_button.configure(command=controller.validate_login)