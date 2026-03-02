import os

import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox

from ..chroma import Chroma, UI


class LogConsole(ctk.CTkToplevel):
    """Modern log console with customtkinter styling."""
    
    # Log type colors
    LOG_COLORS = {
        "Info": "#3B82F6",     # Blue
        "Warning": "#F59E0B",  # Amber
        "Error": "#EF4444",    # Red
    }
    
    def __init__(self):
        super().__init__()
        
        self.colors = Chroma()
        if not self.colors.getDarkMode():
            self.colors.toggle()  # Use dark mode for console
        
        self.title("Log Console")
        self.geometry("900x600")
        self.minsize(600, 400)
        
        try:
            self.iconbitmap(os.getcwd() + "\\media\\icon.ico")
        except:
            pass
        
        self.configure(fg_color="#0F172A")  # Dark slate background
        
        self.logs_text = ""
        self.__create_widgets__()

    def show_userform(self):
        self.focus_force()
        self.mainloop()

    def hide_userform(self):
        self.destroy()

    def connect_with_controller(self, controller):
        self.logs_df = controller.print_logs()
        
        # Create text widget with modern styling
        self.logs_text_label = tk.Text(
            self.main_frame,
            wrap=tk.WORD,
            bg="#0F172A",
            fg="#E2E8F0",
            font=(UI.font_family_mono, 12),
            highlightthickness=0,
            borderwidth=0,
            padx=UI.spacing_md,
            pady=UI.spacing_md,
            insertbackground="#E2E8F0",
            selectbackground="#334155",
            selectforeground="#FFFFFF",
        )
        
        self.__insert_logs_into_label__(self.logs_df)
        
        # Modern scrollbar
        self.scrollbar = ctk.CTkScrollbar(
            self.main_frame,
            command=self.logs_text_label.yview,
            fg_color="#1E293B",
            button_color="#475569",
            button_hover_color="#64748B",
        )
        self.scrollbar.pack(side="right", fill="y")
        
        self.logs_text_label.pack(fill=tk.BOTH, expand=True)
        self.logs_text_label.config(yscrollcommand=self.scrollbar.set)
        
        # Go to the end of the text
        self.logs_text_label.see(tk.END)

        def on_export_logs_to_csv(event=None):
            controller.on_export_logs_to_csv()

        self.export_btn.configure(command=on_export_logs_to_csv)
        
        def on_visibility(event):
            current_quantity_of_logs = len(self.logs_df)
            logs = controller.print_logs()
            
            if len(logs) > current_quantity_of_logs:
                new_logs = logs.iloc[current_quantity_of_logs:]
                self.__insert_logs_into_label__(new_logs)
                self.logs_df = logs
        
        # Update logs when window is visible
        self.bind("<Visibility>", on_visibility)
        self.bind("<FocusIn>", on_visibility)
        self.bind("<Enter>", on_visibility)

    def show_success_export_to_csv(self):
        messagebox.showinfo("Success", "Logs exported to CSV file")

    def show_failure_export_to_csv(self):
        messagebox.showerror("Error", "Failed to export logs to CSV")

    def __create_widgets__(self):
        # Header
        header_frame = ctk.CTkFrame(self, fg_color="#1E293B", corner_radius=0, height=50)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="Log Console",
            font=(UI.font_family, UI.font_lg, "bold"),
            text_color="#F8FAFC",
        )
        title_label.pack(side=tk.LEFT, padx=UI.spacing_lg, pady=UI.spacing_sm)
        
        self.export_btn = ctk.CTkButton(
            header_frame,
            text="Export to CSV",
            width=120,
            height=32,
            font=(UI.font_family, UI.font_sm),
            fg_color="#334155",
            hover_color="#475569",
            text_color="#F8FAFC",
            corner_radius=UI.radius_sm,
        )
        self.export_btn.pack(side=tk.RIGHT, padx=UI.spacing_lg, pady=UI.spacing_sm)
        
        # Main content area
        self.main_frame = ctk.CTkFrame(self, fg_color="#0F172A", corner_radius=0)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

    def __insert_logs_into_label__(self, new_logs):
        self.logs_text_label.config(state=tk.NORMAL)
        
        # Configure tag colors
        for log_type, color in self.LOG_COLORS.items():
            self.logs_text_label.tag_config(log_type.lower(), foreground=color)
        self.logs_text_label.tag_config("timestamp", foreground="#64748B")
        self.logs_text_label.tag_config("separator", foreground="#475569")

        for index, row in new_logs.iterrows():
            if row["Text"] == "":
                continue

            if row["Type"] == "Separator":
                self.logs_text_label.insert(tk.END, "\n" + "─" * 60 + "\n\n", "separator")
                continue
            
            # Timestamp
            timestamp = row["Date and Time"]
            self.logs_text_label.insert(tk.END, f"{timestamp} ", "timestamp")
            
            # Log type badge
            log_type = row["Type"]
            tag = log_type.lower()
            self.logs_text_label.insert(tk.END, f"[{log_type.upper()}] ", tag)
            
            # Log message
            self.logs_text_label.insert(tk.END, row["Text"] + "\n\n")

        self.logs_text_label.config(state=tk.DISABLED)
        self.logs_text_label.see(tk.END)