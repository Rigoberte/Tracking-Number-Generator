import tkinter as tk

class TreeviewToolTip:
    def __init__(self, widget):
        self.widget = widget
        self.tip_window = None
        self.text = ''

    def show_tip(self, text, x, y):
        if self.tip_window or not text:
            return
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        
        label = tk.Label(tw, text=text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

        screen_x = self.widget.winfo_pointerx()
        screen_y = self.widget.winfo_pointery()
        screen_width = tw.winfo_screenwidth()
        screen_height = tw.winfo_screenheight()
        
        # Update geometry of the tooltip to ensure it doesn't go off screen
        tw.update_idletasks()
        width, height = tw.winfo_width(), tw.winfo_height()
        
        # Adjust x position if tooltip goes beyond screen width and place it in the current screen
        number_of_current_screen_x = screen_x // screen_width
        if screen_x + width + 20 > screen_width * (number_of_current_screen_x + 1):
            x = screen_width * (number_of_current_screen_x + 1) - width - 20
        else:
            x = screen_x + 20

        # Adjust y position if tooltip goes beyond screen height and place it in the current screen
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