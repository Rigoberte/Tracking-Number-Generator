import pandas as pd
import datetime as dt
import tkinter as tk
import customtkinter as ctk
from tkcalendar import Calendar
from PIL import Image, ImageTk
import os

from ..tooltips.treeviewTooltip import TreeviewToolTip
from ..tooltips.bottomBarToolTip import BottomBarToolTip
from ..chroma import Chroma

class MyUserForm(tk.Tk):
    def __init__(self):
        """
        Class constructor for UserForm

        Attributes:
            self.colors (Chroma): color palette
            self.representedOrdersAndContactsDataframe (DataFrame): orders table
        """
        super().__init__()

        self.title("Tracking Number Generator")
        self.state("zoomed")
        self.iconbitmap(os.getcwd() + "\\media\\icon.ico")

        self.colors = Chroma()
        
        # Represented model
        self.representedOrdersAndContactsDataframe = pd.DataFrame( columns=['TRACKING_NUMBER', 'HAS_AN_ERROR'] )

        self.__load_userform__()

    def get_selected_team_name(self) -> str:
        """
        Returns the selected team name from the combobox
        """

        selected_team_name = self.team_combobox.get()
        return selected_team_name

    def get_selected_date(self) -> str:
        """
        Returns the selected date from the calendar
        """
        selected_date = self.cal.get_date()
        return selected_date

    def update_a_line_to_processed_of_represented_ordersAndContactsDataframe(self, index: int, tracking_number: str, return_tracking_number: str) -> None:
        """
        Updates a line of the orders table

        Args:
            index (int): row index
            tracking_number (str): tracking number
            return_tracking_number (str): return tracking number
        """
        self.representedOrdersAndContactsDataframe.loc[index, "TRACKING_NUMBER"] = tracking_number
        self.representedOrdersAndContactsDataframe.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number

        self.__update_tag_color_of_a_treeview_line__(index)
        self.__update_bottom_dataFrame_description__()

    def update_whole_represented_ordersAndContactsDataframe(self, ordersAndContactsDataframe: pd.DataFrame) -> None:
        """
        Updates the orders table and the widgets

        Args:
            ordersAndContactsDataframe (DataFrame): orders table
        """
        self.representedOrdersAndContactsDataframe = ordersAndContactsDataframe

        self.__update_treeview__()
        self.__update_bottom_dataFrame_description__()

    def show_userform(self) -> None:
        """
        Shows the UserForm
        """
        self.mainloop()

    def hide_userform(self) -> None:
        """
        Hides the UserForm
        """
        self.destroy()

    def get_root(self) -> tk.Tk:
        """
        Returns the root

        i know this breaks the encapsulation but i need it for the controller
        """
        return self

    def block_widgets(self) -> None:
        """
        Blocks the widgets
        """
        self.loadOrders_btn.configure(state=tk.DISABLED)
        self.processOrders_btn.configure(state=tk.DISABLED)
        self.clear_treeview_btn.configure(state=tk.DISABLED)
        self.config_btn.configure(state=tk.DISABLED)
        self.cal.configure(state='disabled')
        self.team_combobox.configure(state='disabled')

    def unblock_widgets(self) -> None:
        """
        Unblocks the widgets
        """
        self.loadOrders_btn.configure(state=tk.NORMAL)
        self.processOrders_btn.configure(state=tk.NORMAL)
        self.clear_treeview_btn.configure(state=tk.NORMAL)
        self.config_btn.configure(state=tk.NORMAL)
        self.cal.configure(state='normal')
        self.team_combobox.configure(state='readonly')

    def block_processOrders_btn(self) -> None:
        """
        Blocks the processOrders_btn
        """
        self.processOrders_btn.configure(state=tk.DISABLED)

    def unblock_processOrders_btn(self) -> None:
        """
        Unblocks the processOrders_btn
        """
        self.processOrders_btn.configure(state=tk.NORMAL)

    def connect_commands_with_controller(self, controller) -> None:
        """
        Connects the commands with the controller

        Args:
            controller (Controller): controller
        """
        self.loadOrders_btn.configure(command=controller.on_loadOrders_btn_click)
        self.processOrders_btn.configure(command=controller.on_processOrders_btn_click)
        self.clear_treeview_btn.configure(command=controller.on_clearOrders_btn_click)
        self.config_btn.configure(command=controller.config_button_on_click)

        self.team_combobox['values'] = controller.get_team_names()
        self.team_combobox.current(0)

        self.__change_treeview_columns__(self.treeview, controller.get_empty_ordersAndContactsData().columns.tolist())

        def on_log_motion(event):
            self.log_tooltip.hide_tip()
            
            logs_df = controller.print_last_n_logs(35)

            text = ""
            for index, row in logs_df.iterrows():
                text += f"{row['Date and Time']} - {row['Type']}: {row['Text']}\n"

            text = text[:-1] # Remove the last '\n'

            self.log_tooltip.show_tip(text, event.x_root, event.y_root)

        def on_treeview_leave(event):
            self.log_tooltip.hide_tip()

        def on_log_btn_click(event):
            controller.on_log_btn_click()

        self.log_tooltip = BottomBarToolTip(self)
        self.log_image.bind("<Motion>", on_log_motion)
        self.log_image.bind("<Leave>", on_treeview_leave)
        self.log_image.bind("<Double-1>", on_log_btn_click)

        def on_open_excel_motion(event):
            self.open_excel_tooltip.hide_tip()
            
            text = "Open Excel"
            self.open_excel_tooltip.show_tip(text, event.x_root, event.y_root)

        def on_open_excel_leave(event):
            self.open_excel_tooltip.hide_tip()

        def on_open_excel_btn_click(event):
            controller.on_open_excel_double_btn_click()

        self.open_excel_tooltip = BottomBarToolTip(self)

        self.open_excel_image.bind("<Motion>", on_open_excel_motion)
        self.open_excel_image.bind("<Leave>", on_open_excel_leave)
        self.open_excel_image.bind("<Double-1>", on_open_excel_btn_click)

        def on_treeview_motion(event):
            item_id = self.treeview.identify_row(event.y)
            if item_id and (item_id != on_treeview_motion.last_item_id):
                on_treeview_motion.last_item_id = item_id
                item_id_column = self.treeview.identify_column(event.x)
                text = self.treeview.item(item_id, "values")[int(item_id_column[1:]) - 1]
                self.tooltip.hide_tip()
                self.tooltip.show_tip(text, event.x_root, event.y_root)
            elif not item_id:
                self.tooltip.hide_tip()
                on_treeview_motion.last_item_id = None

        def on_treeview_leave(event):
            self.tooltip.hide_tip()

        def copy_selection(event):
            if self.selected_cells:
                selected_texts = {}
                for item, column in self.selected_cells:
                    cell_value = self.treeview.set(item, column)
                    if item not in selected_texts:
                        selected_texts[item] = []
                    selected_texts[item].append(cell_value)

                final_text = []
                for item in selected_texts:
                    if len(selected_texts[item]) == 1:
                        final_text.append(selected_texts[item][0])
                    else:
                        final_text.append('\t'.join(selected_texts[item]))
                
                self.clipboard_clear()
                self.clipboard_append('\n'.join(final_text))

        def on_treeview_click(event):
            self.start_cell = (self.treeview.identify_row(event.y), self.treeview.identify_column(event.x))
            self.selected_cells = [self.start_cell]
            update_selection()
        
        def on_treeview_drag(event):
            end_cell = (self.treeview.identify_row(event.y), self.treeview.identify_column(event.x))
            self.selected_cells = get_cells_in_range(self.start_cell, end_cell)
            update_selection()
        
        def get_cells_in_range(start, end):
            start_item, start_col = start
            end_item, end_col = end
            start_row_index = self.treeview.index(start_item)
            end_row_index = self.treeview.index(end_item)
            
            if start_row_index > end_row_index:
                start_row_index, end_row_index = end_row_index, start_row_index
            if start_col > end_col:
                start_col, end_col = end_col, start_col
            
            items = self.treeview.get_children()
            selected = []
            if len(items) != 0:
                for row_index in range(start_row_index, end_row_index + 1):
                    for col in range(int(start_col[1:]), int(end_col[1:]) + 1):
                        selected.append((items[row_index], f"#{col}"))
            
            return selected
        
        def update_selection():
            try:
                for index, item in enumerate(self.treeview.get_children()):
                    self.__update_tag_color_of_a_treeview_line__(index)
                
                # Highlight selected cells
                for item, column in self.selected_cells:
                    self.treeview.selection_add(item)
            except Exception as e:
                print(f"Error updating selection: {e}")

        # TreeviewToolTip
        self.tooltip = TreeviewToolTip(self.treeview)
        on_treeview_motion.last_item_id = None
        self.selected_cells = []
        self.start_cell = None
        self.treeview.bind("<Motion>", on_treeview_motion)
        self.treeview.bind("<Leave>", on_treeview_leave)
        
        # Copy selection
        self.bind('<Control-c>', copy_selection)

        # Bind mouse click to select a cell
        self.treeview.bind('<Button-1>', on_treeview_click)

        # Bind mouse drag to select multiple cells
        self.treeview.bind('<B1-Motion>', on_treeview_drag)

    # Private methods
    def __change_treeview_columns__(self, treeview: tk.ttk.Treeview, columns_names: list) -> None:
        """
        Changes the columns of the treeview

        Args:
            columns_names (list): columns names
            treeview (tk.ttk.Treeview): self.treeview to show orders table
        """
        columns_names = ['#'] + columns_names # Please do not merge this line with the next one
        treeview["columns"] = columns_names
        treeview.delete(*treeview.get_children())
        
        for col in columns_names:
            treeview.heading(col, text=col)
            treeview.column(col, anchor=tk.W, width=int(self.winfo_screenwidth() * 0.7 * 0.3))
        treeview.column('#', anchor=tk.W, width=int(self.winfo_screenwidth() * 0.7 * 0.05))

    def __load_userform__(self):
        """
        Loads the UserForm and their widgets
        """
        def create_frame_template(master, side) -> ctk.CTkFrame:
            frame = ctk.CTkFrame(master, fg_color='transparent', bg_color='transparent', corner_radius=0)
            frame.pack(side=side, fill=tk.X, pady=0)

            return frame

        def create_all_frames(self) -> dict:
            frames = {}
            frame_top = create_frame_template(self, tk.TOP)
            frames["top"] = frame_top

            frame_for_logo = create_frame_template(frame_top, tk.LEFT)
            frames["logo"] = frame_for_logo

            frame_for_calendar_and_its_text = create_frame_template(frame_top, tk.LEFT)
            frame_for_calender_text = create_frame_template(frame_for_calendar_and_its_text, tk.TOP)
            frame_for_calendar = create_frame_template(frame_for_calendar_and_its_text, tk.TOP)
            frames["calendar_text"] = frame_for_calender_text
            frames["calendar"] = frame_for_calendar


            
            frame_for_team_picker_and_its_text = create_frame_template(frame_top, tk.LEFT)
            frame_for_team_picker_text = create_frame_template(frame_for_team_picker_and_its_text, tk.TOP)
            frame_for_team_picker = create_frame_template(frame_for_team_picker_and_its_text, tk.TOP)
            frames["team_picker_text"] = frame_for_team_picker_text
            frames["team_picker"] = frame_for_team_picker

            

            frame_for_4_buttons = create_frame_template(frame_top, tk.RIGHT)
            frame_for_top_buttons = create_frame_template(frame_for_4_buttons, tk.TOP)
            frame_for_button_bottom = create_frame_template(frame_for_4_buttons, tk.BOTTOM)
            frame_for_button_top_and_left = create_frame_template(frame_for_top_buttons, tk.LEFT)
            frame_for_button_top_and_right = create_frame_template(frame_for_top_buttons, tk.RIGHT)
            frame_for_button_bottom_and_left = create_frame_template(frame_for_button_bottom, tk.LEFT)
            frame_for_button_bottom_and_right = create_frame_template(frame_for_button_bottom, tk.RIGHT)
            frames["button_top_and_left"] = frame_for_button_top_and_left
            frames["button_top_and_right"] = frame_for_button_top_and_right
            frames["button_bottom_and_left"] = frame_for_button_bottom_and_left
            frames["button_bottom_and_right"] = frame_for_button_bottom_and_right



            frame_bottom = ctk.CTkFrame(self, corner_radius=0)
            frame_bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=0)
            frames["bottom"] = frame_bottom

            frame_mid = ctk.CTkFrame(self, fg_color='transparent', bg_color='transparent', corner_radius=0)
            frame_mid.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, pady=0)
            frame_mid.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, pady=0)

            frames["mid"] = frame_mid

            frame_for_vertical_scrollbar = ctk.CTkFrame(frame_mid, corner_radius=0)
            frame_for_vertical_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=True)
            
            frames["vertical_scrollbar"] = frame_for_vertical_scrollbar

            return frames

        def load_top_logo_image(self, frame) -> tk.Label:
            label_banner = tk.Label(frame,  bg= self.colors.getSidebarColor())
            label_banner.pack(side=tk.LEFT, padx=10)

            return label_banner

        def load_calendar_datePicker(self, frame) -> Calendar:

            def nextWorkingDay(fecha):
                while fecha.weekday() >= 5:  # 5: sábado, 6: domingo
                    fecha += dt.timedelta(days=1)
                return fecha
            
            today = dt.datetime.now()
            next_working_day = nextWorkingDay(today + dt.timedelta(days=1))

            cal = Calendar(frame, selectmode='day', locale='en_US', 
                            disabledforeground='red',
                            cursor="hand2", date_pattern='yyyy-MM-dd',
                            year=next_working_day.year, month=next_working_day.month, day=next_working_day.day)

            cal.pack(padx=50, pady=0, side=tk.LEFT)

            return cal

        def load_teams_combobox(self, frame) -> tk.ttk.Combobox:
            teams_options = ["Team"]
            team_combobox = tk.ttk.Combobox(frame, values=teams_options, width=20, height=15, font=30, state='readonly')
            team_combobox.pack(side=tk.LEFT, padx=10)
            team_combobox.current(0)

            return team_combobox
        
        def load_treeview(self, frame, treeviewColumns) -> tk.ttk.Treeview:
            style = tk.ttk.Style()
            style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 13)) # Font of the body
            style.configure("mystyle.Treeview.Heading", font=('Calibri', 13,'bold')) # Font of the headings
            style.configure("mystyle.Treeview", rowheight=25)
            style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) # Remove the borders
            
            treeview = tk.ttk.Treeview(frame, columns=treeviewColumns , show='headings', style="mystyle.Treeview")
            treeview.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            treeview.tag_configure('odd', background='#E8E8E8')
            treeview.tag_configure('odd_done', background='#C6E0B4')
            treeview.tag_configure('odd_error', background='#FFC7CE')

            treeview.tag_configure('even', background='#DFDFDF')
            treeview.tag_configure('even_done', background='#A9D08E')
            treeview.tag_configure('even_error', background='#FFA7BB')

            # Treeview columns headings and columns width
            treeview.column("#0", width=0, stretch=tk.NO)  # Hide the first column
            self.__change_treeview_columns__(treeview, treeviewColumns)

            return treeview

        def load_horizontal_scrollbar(self, frame, treeview) -> tk.ttk.Scrollbar:
            x_scrollbar = tk.ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=treeview.xview)
            x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            treeview.configure(xscrollcommand=x_scrollbar.set)

            return x_scrollbar

        def load_vertical_scrollbar(self, frame, treeview) -> tk.ttk.Scrollbar:
            y_scrollbar = tk.ttk.Scrollbar(frame, orient=tk.VERTICAL, command=treeview.yview)
            y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            treeview.configure(yscrollcommand=y_scrollbar.set)

            return y_scrollbar

        def load_button_processOrders(self, frame) -> ctk.CTkButton:
            processOrders_btn = create_button_template(self, frame, "Process Orders")

            processOrders_btn.configure(state=tk.DISABLED)

            return processOrders_btn

        def load_button_loadOrders(self, frame) -> ctk.CTkButton:
            loadOrders_btn = create_button_template(self, frame, "Load Orders")

            return loadOrders_btn
        
        def load_button_config(self, frame) -> ctk.CTkButton:
            config_btn = create_button_template(self, frame, "Config")

            return config_btn

        def load_button_clear_treeview(self, frame) -> ctk.CTkButton:
            clear_treeview_btn = create_button_template(self, frame, "Clear")

            return clear_treeview_btn

        def load_dataFrame_description(self, frame) -> tk.Label:
            bottom_dataFrame_description = create_label_template(self, frame, "", font_size = 11, is_bold=False)
            return bottom_dataFrame_description

        def create_label_template(self, frame, text, font = 'Calibri Light', font_size = 16, is_bold = True, side = tk.LEFT) -> tk.Label:
            font_bold = 'bold' if is_bold else 'normal'
            label = tk.Label(frame, text=text, font=(font, font_size, font_bold))
            label.pack(expand=True, side=side)
            
            return label

        def create_button_template(self, frame, text) -> ctk.CTkButton:
            button = ctk.CTkButton(frame, text=text, width=150, height=50, font=('Calibri', 22, 'bold'))
            button.pack(pady=10, padx=10)
            return button

        def load_dark_mode_image(self, frame) -> tk.Label:
            dark_mode_image = create_label_template(self, frame, "", font_size = 11, side=tk.RIGHT)
            dark_mode_image.bind('<Button-1>', self.__toggle_color_btn_on_click__)
            return dark_mode_image

        def load_log_image(self, frame) -> tk.Label:
            log_image = create_label_template(self, frame, "", font_size = 11, side=tk.RIGHT)
            return log_image

        def load_open_excel_image(self, frame) -> tk.Label:
            open_excel_image = create_label_template(self, frame, "", font_size = 11, side=tk.RIGHT)
            return open_excel_image

        self.frames = create_all_frames(self)

        # Top widgets
        self.logo = load_top_logo_image(self, self.frames["logo"])
        self.calendar_text = create_label_template(self, self.frames["calendar_text"], "Select a date:")
        self.cal = load_calendar_datePicker(self, self.frames["calendar"])

        self.team_picker_text = create_label_template(self, self.frames["team_picker_text"], "Select a team:")
        self.team_combobox = load_teams_combobox(self, self.frames["team_picker"])

        # Top buttons
        self.loadOrders_btn = load_button_loadOrders(self, self.frames["button_top_and_left"])
        self.processOrders_btn = load_button_processOrders(self, self.frames["button_top_and_right"])
        self.clear_treeview_btn = load_button_clear_treeview(self, self.frames["button_bottom_and_left"])
        self.config_btn = load_button_config(self, self.frames["button_bottom_and_right"])

        # Bottom widgets
        self.bottom_dataFrame_description = load_dataFrame_description(self, self.frames["bottom"])
        self.dark_mode_image = load_dark_mode_image(self, self.frames["bottom"])
        self.log_image = load_log_image(self, self.frames["bottom"])
        self.open_excel_image = load_open_excel_image(self, self.frames["bottom"])

        # Treeview
        treeviewColumns = self.representedOrdersAndContactsDataframe.columns.tolist()
        self.treeview = load_treeview(self, self.frames["mid"], treeviewColumns)

        # Scrollbars
        self.y_scrollbar = load_vertical_scrollbar(self, self.frames["vertical_scrollbar"], self.treeview)
        self.x_scrollbar = load_horizontal_scrollbar(self, self.frames["mid"], self.treeview)

        # Set colors
        self.__set_colors__()

        # Update represented model
        self.update_whole_represented_ordersAndContactsDataframe(self.representedOrdersAndContactsDataframe)
    
    def __tag_color_of_a_treeview_line__(self, parity: bool, order_done: bool, error: bool) -> str:
        """
        Tags colors on each treeview line

        Args:
            row (Series): row of the orders table
            parity (bool): parity
        """
        tag_color = 'odd' if parity else 'even'
        tag_color += "_done" if order_done else ""
        tag_color += ("_error" if error else "") if not order_done else ""
        return tag_color
    
    def __tag_colors_for_each_treeview_line__(self) -> None:
        """
        Tags colors on each treeview line

        Args:
            ordersAndContactsDataframe (DataFrame): orders table
        """
        if self.representedOrdersAndContactsDataframe.empty:
            return

        for index, row in self.representedOrdersAndContactsDataframe.iterrows():
            row_values = [index] + list(row)
            self.treeview.insert("", "end", iid=index, values=row_values)
            self.update_a_line_to_processed_of_represented_ordersAndContactsDataframe(index, row['TRACKING_NUMBER'], row['RETURN_TRACKING_NUMBER'])
    
    def __set_colors__(self) -> None:
            def changeImage(self, widget, dark_image, light_image, size, bg_color):
                if self.colors.getDarkMode():
                    logoPath = os.getcwd() + "\\media\\" + dark_image
                else:
                    logoPath = os.getcwd() + "\\media\\" + light_image

                imagen = Image.open(logoPath)
                imagen = imagen.resize(size)
                imagen_tk = ImageTk.PhotoImage(imagen)
                
                widget.configure(image=imagen_tk)
                widget.image = imagen_tk
                widget.configure(bg=bg_color)

            def changeAllImages(self):
                changeImage(self, self.logo, "TMO_logo_dark.png", "TMO_logo_light.png", (284, 61), self.colors.getSidebarColor())
                
                changeImage(self, self.dark_mode_image, "moon-regular-24.png", "sun-solid-24.png", (24, 24), self.colors.getBodyColor())
                changeImage(self, self.log_image, "message-alt-detail-regular-24.png", "message-alt-detail-solid-24.png", (24, 24), self.colors.getBodyColor())
                changeImage(self, self.open_excel_image, "data-regular-24.png", "data-solid-24.png", (24, 24), self.colors.getBodyColor())
            
            self.colors.toggle()
            self.frames["top"].configure(bg_color=self.colors.getTextColor(),
                                        fg_color=self.colors.getSidebarColor())
            
            self.frames["bottom"].configure(bg_color=self.colors.getTextColor(),
                                        fg_color=self.colors.getBodyColor())
            
            self.frames["mid"].configure(bg_color=self.colors.getTextColor(),
                                        fg_color=self.colors.getBodyColor())
            
            self.calendar_text.configure(bg=self.colors.getSidebarColor(), fg="white")
            self.team_picker_text.configure(bg=self.colors.getSidebarColor(), fg="white")
            self.bottom_dataFrame_description.configure(fg= self.colors.getTextColor(), bg = self.colors.getBodyColor()) 
            
            
            changeAllImages(self)


            buttons = [self.loadOrders_btn, self.processOrders_btn, self.clear_treeview_btn, self.config_btn]
            for button in buttons:
                button.configure(fg_color=self.colors.getPrimaryColor(),
                                hover_color = self.colors.getPrimaryColorLight(),
                                text_color= self.colors.getTextColorForButtons())

    def __update_treeview__(self) -> None:
        """
        Updates the treeview

        Args:
            ordersAndContactsDataframe (DataFrame): orders table
            treeview (tk.ttk.Treeview): self.treeview to show orders table
        """
        self.__clear_treeview__()
        self.__tag_colors_for_each_treeview_line__()
    
    def __update_bottom_dataFrame_description__(self) -> None:
        ordersAndContactsDataframe = self.representedOrdersAndContactsDataframe

        amountOfTotalOrders = len(ordersAndContactsDataframe)
        amountOfOrdersProcessed = len(ordersAndContactsDataframe[ordersAndContactsDataframe['TRACKING_NUMBER'] != ""])
        amountOfOrdersReadyToBeProcessed = len(ordersAndContactsDataframe[(ordersAndContactsDataframe['TRACKING_NUMBER'] == "") & (ordersAndContactsDataframe['HAS_AN_ERROR'] == "No error")])
        amountOfOrdersWithErrors = len(ordersAndContactsDataframe[(ordersAndContactsDataframe['HAS_AN_ERROR'] != "No error") & (ordersAndContactsDataframe['TRACKING_NUMBER'] == "")])

        text = f"Amount of total orders: {amountOfTotalOrders} | "
        text += f"Amount of orders processed: {amountOfOrdersProcessed} | "
        text += f"Amount of orders ready to be processed: {amountOfOrdersReadyToBeProcessed} | "
        text += f"Amount of orders with errors: {amountOfOrdersWithErrors}"
        
        self.bottom_dataFrame_description.config(text=text)
    
    def __clear_treeview__(self) -> None:
        """
        Clears the treeview
        """
        if len(self.treeview.get_children("")) == 0:
            return
        
        for item in self.treeview.get_children():
            self.treeview.delete(item)

    def __toggle_color_btn_on_click__(self, event):
        self.__set_colors__()

    def __update_tag_color_of_a_treeview_line__(self, index: int) -> None:
        """
        Updates tag color on each treeview line

        Args:
            index (int): row index
            row (Series): row of the orders table
            treeview (tk.ttk.Treeview): self.treeview to show orders table
        """
        try:
            row_values = self.representedOrdersAndContactsDataframe.loc[index]
            row_values_list = [index] + list(row_values)

            self.treeview.item(index, values = row_values_list)

            parity = index % 2 == 0
            is_a_processed_order = row_values['TRACKING_NUMBER'] != ""
            has_an_error = row_values['HAS_AN_ERROR'] != "No error"
            tag_color = self.__tag_color_of_a_treeview_line__(parity, is_a_processed_order, has_an_error)
            
            self.treeview.item(index, tags=tag_color)

        except Exception as e:
            print(f"Error updating tag color of a treeview line: {e}")