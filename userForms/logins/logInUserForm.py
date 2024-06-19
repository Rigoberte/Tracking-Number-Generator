import tkinter as tk
from tkinter import messagebox

class LogInUserForm(tk.Tk):
    def __init__(self, controller = None):
        super().__init__()
        self.title("Login Form")
        self.geometry("300x150")

        self.controller = controller
    
    def get_username(self) -> str:
        return self.username_entry.get()
    
    def get_password(self) -> str:
        return self.password_entry.get()

    def clear_password_entry(self) -> None:
        self.password_entry.delete(0, 'end')

    def show_userform(self) -> None:
        self.__load_userform__(self)
        self.mainloop()

    def hide_userform(self) -> None:
        self.destroy()

    def validate_login(self) -> None:
        self.controller.validate_login()

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

        def password_entry_on_return(event):
            self.validate_login()
        
        self.password_entry.bind("<Return>", password_entry_on_return)
        
        # Login Button
        self.login_button = tk.Button(root, text="Login", command=self.validate_login)
        self.login_button.pack(pady=5)
        
        # Exit Button
        self.exit_button = tk.Button(root, text="Exit", command=root.quit)
        self.exit_button.pack(pady=5)