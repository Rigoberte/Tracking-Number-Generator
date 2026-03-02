import tkinter as tk


class TreeviewToolTip:
    """Modern tooltip for treeview cells."""
    
    # Modern styling constants
    BG_COLOR = "#1F2937"  # Dark slate
    FG_COLOR = "#F9FAFB"  # Light gray
    FONT = ("Segoe UI", 9)
    PADDING = 8
    BORDER_RADIUS = 6
    
    def __init__(self, widget):
        self.widget = widget
        self.tip_window = None
        self.text = ''

    def show_tip(self, text, x, y):
        if self.tip_window or not text:
            return
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.attributes('-alpha', 0.95)  # Slight transparency
        
        # Modern styled frame
        frame = tk.Frame(
            tw,
            background=self.BG_COLOR,
            highlightbackground="#374151",
            highlightthickness=1,
        )
        frame.pack()
        
        label = tk.Label(
            frame,
            text=text,
            justify=tk.LEFT,
            background=self.BG_COLOR,
            foreground=self.FG_COLOR,
            font=self.FONT,
            padx=self.PADDING,
            pady=self.PADDING // 2,
        )
        label.pack()

        screen_x = self.widget.winfo_pointerx()
        screen_y = self.widget.winfo_pointery()
        screen_width = tw.winfo_screenwidth()
        screen_height = tw.winfo_screenheight()
        
        tw.update_idletasks()
        width, height = tw.winfo_width(), tw.winfo_height()
        
        # Position calculation
        number_of_current_screen_x = screen_x // screen_width
        if screen_x + width + 20 > screen_width * (number_of_current_screen_x + 1):
            x = screen_width * (number_of_current_screen_x + 1) - width - 20
        else:
            x = screen_x + 20

        number_of_current_screen_y = screen_y // screen_height
        if screen_y + height + 20 > screen_height * (number_of_current_screen_y + 1):
            y = screen_height * (number_of_current_screen_y + 1) - height - 20
        else:
            y = screen_y + 20

        tw.wm_geometry(f"+{x}+{y}")

    def hide_tip(self):
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None