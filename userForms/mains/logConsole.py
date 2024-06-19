import tkinter as tk

class LogConsole(tk.Tk):
    def __init__(self, controller = None):
        super().__init__()
        self.controller = controller
        self.title("Log Console")
        self.geometry("800x600")

    def copy(self, event):
        self.clipboard_clear()
        self.clipboard_append(self.text.selection_get())

    def show_userform(self):
        self.__create_widgets__()
        self.mainloop()

    def hide_userform(self):
        self.destroy()

    def __create_widgets__(self):
        self.text = tk.Text(self, wrap=tk.WORD)
        self.text.pack(fill=tk.BOTH, expand=True)
        text = self.controller.get_last_n_logs(100)
        self.text.insert(tk.END, text)
        self.text.config(state=tk.DISABLED)

        self.text.bind("<Key>", lambda e: "break")
        self.text.bind("<Button-1>", lambda e: "break")
        self.text.bind("<Button-2>", lambda e: "break")
        self.text.bind("<Button-3>", lambda e: "break")

        self.text.bind("<Control-c>", self.copy)

        self.text.config(bg="black", fg="white")
        self.text.config(font=("Courier", 14))