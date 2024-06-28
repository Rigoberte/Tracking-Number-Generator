import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
class LogConsole(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Log Console")
        self.geometry("800x600")

        self.logs_text = ""

        self.__create_widgets__()

    def copy(self, event):
        self.clipboard_clear()
        self.clipboard_append(self.logs_text_label.selection_get())

    def show_userform(self):
        self.mainloop()

    def hide_userform(self):
        self.destroy()

    def connect_with_controller(self, controller):
        self.logs_df = controller.print_logs()

        for index, row in self.logs_df.iterrows():
            self.logs_text += row["Date and Time"] + " - " + row["Type"] + ": " + row["Text"] + "\n"

        self.logs_text_label = tk.Text(self.frames["main"], wrap=tk.WORD)
        self.logs_text_label.pack(fill=tk.BOTH, expand=True)
        self.logs_text_label.insert(tk.END, self.logs_text)
        self.logs_text_label.config(state=tk.DISABLED)

        self.logs_text_label.config(bg="black", fg="white")
        self.logs_text_label.config(font=("Courier", 14))

        def on_export_logs_to_csv(event):
            controller.on_export_logs_to_csv()

        self.export_to_csv_label.bind("<Double-1>", on_export_logs_to_csv)

    def show_success_export_to_csv(self):
        messagebox.showinfo("Success", "Logs exported to .CSV")

    def show_failure_export_to_csv(self):
        messagebox.showerror("Error", "Logs not exported to .CSV")

    def __create_frames__(self, master) -> dict:
        frames = {}

        windows = ctk.CTkFrame(master, corner_radius=0)
        windows.pack(fill=tk.BOTH, expand=True)


        # Calcular la altura de la fuente
        font = ("Calibri Light", 11)
        dummy_label = tk.Label(self, text="Test", font=font)
        dummy_label.update_idletasks()  # Asegurarse de que la GUI está actualizada para obtener el tamaño correcto
        font_height = dummy_label.winfo_reqheight()
        dummy_label.destroy()

        frame_bottom = ctk.CTkFrame(windows, fg_color='white', bg_color='white', corner_radius=0, height=font_height)
        frame_bottom.pack(side=tk.BOTTOM, pady=0, fill=tk.X)
        frames["bottom"] = frame_bottom


        frame_main = ctk.CTkFrame(windows, fg_color='transparent', bg_color='transparent', corner_radius=0)
        frame_main.pack(side=tk.TOP, pady=0, fill=tk.BOTH, expand=True)
        frames["main"] = frame_main

        return frames

    def __create_widgets__(self):
        self.frames = self.__create_frames__(self)

        self.export_to_csv_label = tk.Label(self.frames["bottom"], text="Export logs to .CSV", font=('Calibri Light', 11, "normal"))
        self.export_to_csv_label.pack(expand=True)

        self.config(bg="black")