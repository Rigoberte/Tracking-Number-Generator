from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import datetime as dt
import tkinter as tk
import customtkinter as ctk
from tkcalendar import Calendar
from PIL import Image, ImageTk
import os, time, shutil
import win32com.client as win32

# View
class MyUserForm(tk.Tk):
    def __init__(self, controller = None):
        """
        Class constructor for UserForm

        Attributes:
            self.colors (Chroma): color palette
            self.selected_team (Teams): selected team
            self.cal (Calendar): calendar widget
            self.team_combobox (Combobox): team combobox
            self.carrier_combobox (Combobox): carrier combobox
            self.treeview (Treeview): treeview widget
            self.tooltip (TreeviewToolTip): tooltip widget
            self.ordersAndContactsDataframe (DataFrame): orders table

        Args:
            controller (Controller): controller
        """
        super().__init__()
        self.controller = controller

        self.title("Tracking Number Generator")
        self.state("zoomed")

        self.colors = Chroma()

        self.selected_team = Teams()
        self.selected_date = dt.datetime.now().strftime('%Y-%m-%d')

        self.treeview = None
        self.bottom_dataFrame_description = None
        self.team_combobox = None

        self.clear_treeview_btn = None
        self.loadOrders_btn = None
        self.processOrders_btn = None
        self.config_btn = None
        self.ordersAndContactsDataframe = DataRecolector(Teams()).getEmptyOrdersAndContactsData()

        self.amountOfTotalOrders = 0
        self.amountOfOrdersProcessed = 0
        self.amountOfOrdersReadyToBeProcessed = 0
        self.amountOfOrdersWithErrors = 0

    def getSelectedTeam(self):
        return self.selected_team
    
    def getSelectedTeamName(self) -> str:
        return self.team_combobox.get()

    def getSelectedDate(self) -> str:
        return self.cal.get_date()

    def getFolderPathToDownload(self):
        return self.folder_path_to_download

    def update_tag_color_of_a_treeview_line(self, ordersAndContactsDataframe: pd.DataFrame, index: int, tracking_number: str = "", return_tracking_number: str = ""):
        """
        Updates tag color on each treeview line

        Args:
            index (int): row index
            row (Series): row of the orders table
            treeview (tk.ttk.Treeview): self.treeview to show orders table
        """
        try:
            if tracking_number != "":
                ordersAndContactsDataframe.loc[index, "TRACKING_NUMBER"] = tracking_number
            
            if return_tracking_number != "":
                ordersAndContactsDataframe.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number
            
            row_values = ordersAndContactsDataframe.loc[index]
            row_values_list = [index] + list(row_values)

            self.treeview.item(index, values = row_values_list)

            parity = index % 2 == 0
            is_a_processed_order = row_values['TRACKING_NUMBER'] != ""
            has_an_error = row_values['HAS_AN_ERROR'] != "No error"
            tag_color = self.__tag_color_of_a_treeview_line__(parity, is_a_processed_order, has_an_error)
            self.treeview.item(index, tags=tag_color)

            # i know this is not the best way to do it, but this doesn't need too much resources
            # the correct way is to implement a multithreading
            self.treeview.update()

        except Exception as e:
            Log().add_log(f"Error updating tag color of a treeview line: {e}")
    
    def update_ordersAndContactsDataframe_and_widgets(self, ordersAndContactsDataframe):
        """
        Updates the orders table and the widgets

        Args:
            ordersAndContactsDataframe (DataFrame): orders table
        """
        self.__update_treeview__(ordersAndContactsDataframe)


        self.amountOfTotalOrders = len(ordersAndContactsDataframe)
        self.amountOfOrdersProcessed = len(ordersAndContactsDataframe[ordersAndContactsDataframe['TRACKING_NUMBER'] != ""])
        self.amountOfOrdersReadyToBeProcessed = len(ordersAndContactsDataframe[(ordersAndContactsDataframe['TRACKING_NUMBER'] == "") & (ordersAndContactsDataframe['HAS_AN_ERROR'] == "No error")])
        self.amountOfOrdersWithErrors = len(ordersAndContactsDataframe[(ordersAndContactsDataframe['HAS_AN_ERROR'] != "No error") & (ordersAndContactsDataframe['TRACKING_NUMBER'] == "")])
        
        self.__update_bottom_dataFrame_description__(self.amountOfTotalOrders, self.amountOfOrdersProcessed, self.amountOfOrdersReadyToBeProcessed, self.amountOfOrdersWithErrors)
        self.ordersAndContactsDataframe = ordersAndContactsDataframe

    def increase_amount_of_orders_processed(self):
        self.amountOfOrdersProcessed += 1
        self.amountOfOrdersReadyToBeProcessed = max(0, self.amountOfOrdersReadyToBeProcessed - 1)
        self.__update_bottom_dataFrame_description__(self.amountOfTotalOrders, self.amountOfOrdersProcessed, self.amountOfOrdersReadyToBeProcessed, self.amountOfOrdersWithErrors)

    def show_userform(self):
        self.__load_userform__()
        self.mainloop()

    def hide_userform(self):
        self.destroy()

    # Private methods
    def __getFolderPathToDownload__(self, team: str, date: str) -> str:
        return FilesInSystemController().getFolderPathToDownload(team, date)
    
    def __create_folder__(self, folder_path_to_download: str):
        FilesInSystemController().create_folder(folder_path_to_download)
    
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

            self.selected_date = cal.get_date()

            return cal

        def load_teams_combobox(self, frame) -> tk.ttk.Combobox:
            teams_options = Teams().getTeamsNames()
            team_combobox = tk.ttk.Combobox(frame, values=teams_options, width=20, height=15, font=30)
            team_combobox.pack(side=tk.LEFT, padx=10)
            team_combobox.current(0)

            return team_combobox
        
        def load_treeview(self, frame, treeviewColumns) -> tk.ttk.Treeview:
            style = tk.ttk.Style()
            style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 13)) # Font of the body
            style.configure("mystyle.Treeview.Heading", font=('Calibri', 13,'bold')) # Font of the headings
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
            for col in treeviewColumns:
                treeview.heading(col, text=col)
                treeview.column(col, anchor=tk.W, width=int(self.winfo_screenwidth() * 0.7 * 0.3))
            treeview.column("#", anchor=tk.W, width=int(self.winfo_screenwidth() * 0.7 * 0.05))

            def on_treeview_motion(event):
                item_id = treeview.identify_row(event.y)
                if item_id and (item_id != on_treeview_motion.last_item_id):
                    on_treeview_motion.last_item_id = item_id
                    item_id_column = treeview.identify_column(event.x)
                    text = treeview.item(item_id, "values")[int(item_id_column[1:]) - 1]
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
                        cell_value = treeview.set(item, column)
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
                self.start_cell = (treeview.identify_row(event.y), treeview.identify_column(event.x))
                self.selected_cells = [self.start_cell]
                update_selection()
            
            def on_treeview_drag(event):
                end_cell = (treeview.identify_row(event.y), treeview.identify_column(event.x))
                self.selected_cells = get_cells_in_range(self.start_cell, end_cell)
                update_selection()
            
            def get_cells_in_range(start, end):
                start_item, start_col = start
                end_item, end_col = end
                start_row_index = treeview.index(start_item)
                end_row_index = treeview.index(end_item)
                
                if start_row_index > end_row_index:
                    start_row_index, end_row_index = end_row_index, start_row_index
                if start_col > end_col:
                    start_col, end_col = end_col, start_col
                
                items = treeview.get_children()
                selected = []
                if len(items) != 0:
                    for row_index in range(start_row_index, end_row_index + 1):
                        for col in range(int(start_col[1:]), int(end_col[1:]) + 1):
                            selected.append((items[row_index], f"#{col}"))
                
                return selected
            
            def update_selection():
                try:
                    for index, item in enumerate(treeview.get_children()):
                        self.update_tag_color_of_a_treeview_line(self.ordersAndContactsDataframe, index)
                    
                    # Highlight selected cells
                    for item, column in self.selected_cells:
                        treeview.selection_add(item)
                except Exception as e:
                    Log().add_log(f"Error updating selection: {e}")

            # TreeviewToolTip
            self.tooltip = TreeviewToolTip(treeview)
            on_treeview_motion.last_item_id = None
            self.selected_cells = []
            self.start_cell = None
            treeview.bind("<Motion>", on_treeview_motion)
            treeview.bind("<Leave>", on_treeview_leave)
            
            # Copy selection
            self.bind('<Control-c>', copy_selection)

            # Bind mouse click to select a cell
            treeview.bind('<Button-1>', on_treeview_click)

            # Bind mouse drag to select multiple cells
            treeview.bind('<B1-Motion>', on_treeview_drag)

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

        def load_button_processOrders(self, frame):
            processOrders_btn = create_button_template(self, frame, "Process Orders", self.__on_processOrders_btn_click__)

            return processOrders_btn

        def load_button_loadOrders(self, frame):
            loadOrders_btn = create_button_template(self, frame, "Load Orders", self.__on_loadOrders_btn_click__)

            return loadOrders_btn
        
        def load_button_config(self, frame):
            config_btn = create_button_template(self, frame, "Config", self.__config_button_on_click__)

            return config_btn

        def load_button_clear_treeview(self, frame):
            clear_treeview_btn = create_button_template(self, frame, "Clear", self.__clear_treeview_on_click__)

            return clear_treeview_btn

        def load_dataFrame_description(self, frame) -> tk.Label:
            bottom_dataFrame_description = create_label_template(self, frame, "", font_size = 11, is_bold=False)
            return bottom_dataFrame_description

        def create_label_template(self, frame, text, font = 'Calibri Light', font_size = 16, is_bold = True, side = tk.LEFT) -> tk.Label:
            font_bold = 'bold' if is_bold else 'normal'
            label = tk.Label(frame, text=text, font=(font, font_size, font_bold))
            label.pack(expand=True, side=side)
            
            return label

        def create_button_template(self, frame, text, command) -> ctk.CTkButton:
            button = ctk.CTkButton(frame, text=text, command=command, width=150, height=50, font=('Calibri', 22, 'bold'))
            button.pack(pady=10, padx=10)
            return button

        def load_dark_mode_image(self, frame) -> tk.Label:
            dark_mode_image = create_label_template(self, frame, "", font_size = 11, side=tk.RIGHT)
            dark_mode_image.bind('<Button-1>', self.__toggle_color_btn_on_click__)
            return dark_mode_image

        def load_log_image(self, frame) -> tk.Label:
            
            def on_log_motion(event):
                self.log_tooltip.hide_tip()
                self.log_tooltip.show_tip(event.x_root, event.y_root)

            def on_treeview_leave(event):
                self.log_tooltip.hide_tip()

            def on_log_btn_click(event):
                self.controller.on_log_btn_click()

            log_image = create_label_template(self, frame, "", font_size = 11, side=tk.RIGHT)
            self.log_tooltip = LogToolTip(self)
            log_image.bind("<Motion>", on_log_motion)
            log_image.bind("<Leave>", on_treeview_leave)
            log_image.bind("<Double-1>", on_log_btn_click)

            return log_image

        self.frames = create_all_frames(self)

        self.logo = load_top_logo_image(self, self.frames["logo"])
        self.calendar_text = create_label_template(self, self.frames["calendar_text"], "Select a date:")
        self.cal = load_calendar_datePicker(self, self.frames["calendar"])

        self.team_picker_text = create_label_template(self, self.frames["team_picker_text"], "Select a team:")
        self.team_combobox = load_teams_combobox(self, self.frames["team_picker"])


        self.loadOrders_btn = load_button_loadOrders(self, self.frames["button_top_and_left"])
        self.processOrders_btn = load_button_processOrders(self, self.frames["button_top_and_right"])
        self.clear_treeview_btn = load_button_clear_treeview(self, self.frames["button_bottom_and_left"])
        self.config_btn = load_button_config(self, self.frames["button_bottom_and_right"])

        self.bottom_dataFrame_description = load_dataFrame_description(self, self.frames["bottom"])
        self.dark_mode_image = load_dark_mode_image(self, self.frames["bottom"])
        self.log_image = load_log_image(self, self.frames["bottom"])

        treeviewColumns = ['#'] + DataRecolector(Teams()).getColumnNames()
        self.treeview = load_treeview(self, self.frames["mid"], treeviewColumns)
        self.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)

        self.y_scrollbar = load_vertical_scrollbar(self, self.frames["vertical_scrollbar"], self.treeview)
        self.x_scrollbar = load_horizontal_scrollbar(self, self.frames["mid"], self.treeview)

        self.__set_colors__()
    
    def __tag_color_of_a_treeview_line__(self, parity: bool, order_done: bool, error: bool):
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
    
    def __tag_colors_for_each_treeview_line__(self, ordersAndContactsDataframe: pd.DataFrame):
        """
        Tags colors on each treeview line

        Args:
            ordersAndContactsDataframe (DataFrame): orders table
        """
        if ordersAndContactsDataframe.empty:
            return

        for index, row in ordersAndContactsDataframe.iterrows():
            row_values = [index] + list(row)
            self.treeview.insert("", "end", iid=index, values=row_values)
            self.update_tag_color_of_a_treeview_line(ordersAndContactsDataframe, index, row['TRACKING_NUMBER'], row['RETURN_TRACKING_NUMBER'])
    
    def __set_colors__(self):
            def changeImage(self, widget, dark_image, light_image, size):
                if self.colors.getDarkMode():
                    logoPath = os.getcwd() + "\\media\\" + dark_image
                else:
                    logoPath = os.getcwd() + "\\media\\" + light_image

                imagen = Image.open(logoPath)
                imagen = imagen.resize(size)
                imagen_tk = ImageTk.PhotoImage(imagen)
                
                widget.configure(image=imagen_tk)
                widget.image = imagen_tk

            def changeAllImages(self):
                changeImage(self, self.logo, "TMO_logo.png", "TMO_logo_light.png", (284, 61))
                changeImage(self, self.dark_mode_image, "moon-regular-24.png", "sun-solid-24.png", (24, 24))
                changeImage(self, self.log_image, "message-alt-detail-regular-24.png", "message-alt-detail-solid-24.png", (24, 24))
            
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
                                text_color= self.colors.getTextColor())

    def __update_treeview__(self, ordersAndContactsDataframe):
        """
        Updates the treeview

        Args:
            ordersAndContactsDataframe (DataFrame): orders table
            treeview (tk.ttk.Treeview): self.treeview to show orders table
        """
        self.__clear_treeview__()
        self.__tag_colors_for_each_treeview_line__(ordersAndContactsDataframe)
        self.treeview.update()
    
    def __update_bottom_dataFrame_description__(self, amount_of_total_orders, amount_of_orders_processed, amount_of_orders_ready_to_be_processed, amount_of_orders_with_errors):
        text = f"Amount of total orders: {amount_of_total_orders} | "
        text += f"Amount of orders processed: {amount_of_orders_processed} | "
        text += f"Amount of orders ready to be processed: {amount_of_orders_ready_to_be_processed} | "
        text += f"Amount of orders with errors: {amount_of_orders_with_errors}"
        
        self.bottom_dataFrame_description.config(text=text)
    
    def __clear_treeview__(self):
        """
        Clears the treeview
        """
        if self.ordersAndContactsDataframe.empty:
            return
        
        for item in self.treeview.get_children():
            self.treeview.delete(item)

    def __on_loadOrders_btn_click__(self):
        """
        Loads orders table according to date and team

        Args:
            date (dt.datetime): date to process
            team (str): team to process
            shipdate (dt.datetime): ship date
            self.treeview (tk.ttk.Treeview): self.treeview to show orders table
        
        Returns:
            DataFrame: orders table
        """

        self.controller.on_loadOrders_btn_click()

    def __on_processOrders_btn_click__(self):
        """
        Button to process all orders in the table
        """
        self.controller.on_processOrders_btn_click()

    def __toggle_color_btn_on_click__(self, event):
        self.__set_colors__()

    def __clear_treeview_on_click__(self):
        self.controller.on_clearOrders_btn_click()

    def __config_button_on_click__(self):
        return #TODO

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

class LogToolTip:
    def __init__(self, widget):
        self.widget = widget
        self.tip_window = None

    def show_tip(self, x, y):
        if self.tip_window:
            return
         
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        
        text = Log().print_last_n_logs(35)
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
        x = screen_x - width // 2

        # Adjust y position to place tip above the log image
        number_of_current_screen_y = screen_y // screen_height
        y = screen_y - height - 20

        tw.wm_geometry(f"+{x}+{y}")

    def hide_tip(self):
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None

class LogConsole(tk.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("Log Console")
        self.geometry("800x600")

        self.text = tk.Text(self, wrap=tk.WORD)
        self.text.pack(fill=tk.BOTH, expand=True)
        self.text.insert(tk.END, Log().print_last_n_logs(100))

        self.text.bind("<Key>", lambda e: "break")
        self.text.bind("<Button-1>", lambda e: "break")
        self.text.bind("<Button-2>", lambda e: "break")
        self.text.bind("<Button-3>", lambda e: "break")

        self.text.bind("<Control-c>", self.copy)

        self.text.config(bg="black", fg="white")
        self.text.config(font=("Courier", 14))

    def copy(self, event):
        self.clipboard_clear()
        self.clipboard_append(self.text.selection_get())

    def show_userform(self):
        self.mainloop()

    def hide_userform(self):
        self.destroy()

class LogInUserForm(tk.Tk):
    def __init__(self, team = None, controller = None):
        super().__init__()
        self.title("Login Form")
        self.geometry("300x150")

        if team is None:
            team = Teams("")
        self.selectedTeam = team

        self.loginResult = False

        self.controller = controller

        self.__load_userform__(self)
    
    def validate_login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        self.selectedTeam.build_driver()
        if self.selectedTeam.check_if_user_and_password_are_correct(username, password):
            # Driver keeps built
            self.loginResult = True
            self.controller.on_login_successful()
        else:
            self.loginResult = False
            self.password_entry.delete(0, 'end')
            self.selectedTeam.quit_driver()

    def getLoginFormResult(self):
        return self.loginResult

    def show_userform(self):
        self.mainloop()

    def hide_userform(self):
        self.destroy()

    def __load_userform__(self, root):
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

class Chroma:
    def __init__(self):
        """
        Class constructor for Chroma
        """
        self.dark = False

        self.body_color = None
        self.sidebar_color = None
        self.primary_color = None
        self.primary_color_light = None
        self.toggle_color = None
        self.text_color = None
        self.toggle()

    def toggle(self):
        """
        Toggles between dark and light mode
        """
        self.dark = not self.dark
        if self.dark:
            self.body_color = '#121212'
            self.sidebar_color = '#3A3B3C'
            self.primary_color = '#3700B3' 
            self.primary_color_light = '#BB86FC'
            self.toggle_color = '#3700B3'
            self.text_color = '#CCC'
        else:
            self.body_color = '#CCC'
            self.sidebar_color = '#E71316'
            self.primary_color = '#c00000'
            self.primary_color_light = '#E83713'
            self.toggle_color = '#BE2F31'
            self.text_color = '#000'

    def getBodyColor(self):
        return self.body_color
    
    def getSidebarColor(self):
        return self.sidebar_color
    
    def getPrimaryColor(self):
        return self.primary_color
    
    def getPrimaryColorLight(self):
        return self.primary_color_light
    
    def getToggleColor(self):
        return self.toggle_color
    
    def getTextColor(self):
        return self.text_color

    def getDarkMode(self):
        return self.dark

# Model
class Browser(object):
    def __init__(self, folder_path_to_download: str):
        """
        Class constructor for Browser 

        Args:
            folder_path_to_download (str): folder path to download files

        Attributes:
            self.driver (webdriver): selenium self.driver
        """
        chrome_options = webdriver.ChromeOptions()
        #chrome_options.add_argument('--headless') # Disable headless mode
        chrome_options.add_argument('--disable-gpu') # Disable GPU acceleration
        chrome_options.add_argument('--disable-software-rasterizer')  # Disable software rasterizer
        chrome_options.add_argument('--disable-dev-shm-usage') # Disable shared memory usage
        chrome_options.add_argument('--no-sandbox') # Disable sandbox
        chrome_options.add_argument('--disable-extensions') # Disable extensions
        chrome_options.add_argument('--disable-sync') # Disable syncing to a Google account
        chrome_options.add_argument('--disable-webgl') # Disable WebGL 
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--disable-gl-extensions') # Disable WebGL extensions
        chrome_options.add_argument('--disable-in-process-stack-traces') # Disable stack traces
        chrome_options.add_argument('--disable-logging') # Disable logging
        chrome_options.add_argument('--disable-cache') # Disable cache
        chrome_options.add_argument('--disable-application-cache') # Disable application cache
        chrome_options.add_argument('--disk-cache-size=1') # Set disk cache size to 1
        chrome_options.add_argument('--media-cache-size=1') # Set media cache size to 1
        chrome_options.add_argument('--kiosk-printing') # Enable kiosk printing
        chrome_options.add_argument('--kiosk-pdf-printing')

        chrome_prefs = {
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "download.open_pdf_in_system_reader": False,
            "profile.default_content_settings.popups": 0,
            #"download.prompt_for_download": False, #To auto download the file
            "printing.print_to_pdf": True,
            "download.default_directory": folder_path_to_download,
            "savefile.default_directory": folder_path_to_download
        }

        chrome_options.add_experimental_option("prefs", chrome_prefs)

        self.driver = webdriver.Chrome(options=chrome_options)

class OrderProcessor:
    def __init__(self, folder_path_to_download : str, selectedTeam, controller = None):
        """
        Class constructor

        Args:
            self.driver (webdriver): selenium self.driver

        Attributes:
            self.driver (webdriver): selenium self.driver
            wait (WebDriverWait): selenium wait
        """
        self.folder_path_to_download = folder_path_to_download
        self.selectedTeam = selectedTeam
        self.controller = controller
    
    def send_email_to_medical_center(self, study: str, site: str, ivrs_number: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str, amount_of_boxes: int,
                                    hasReturn: bool, type_of_return: str, amount_of_boxes_to_return: int,
                                    tracking_number: str, return_tracking_number: str,
                                    contacts: str, medical_center_emails: str, customer_emails: str, CRAs_emails: str, team_emails: str):
        """
        Sends an email to the medical center

        Args:
            study (str): study
            site (str): site
            ivrs_number (str): ivrs number
            delivery_date (str): delivery date
            delivery_time_from (str): delivery time (from)
            delivery_time_to (str): delivery time (to)
            type_of_material (str): type of material
            temperature (str): temperature
            amount_of_boxes (int): number of boxes
            hasReturn (bool): if True, creates a return order
            type_of_return (str): type of return
            amount_of_boxes_to_return (int): number of boxes to return
            tracking_number (str): tracking number
            return_tracking_number (str): return tracking number
            contacts (str): contacts
        """
        def getEmailSourceFromTxtFile(file):
            with open(file, 'r') as file:
                return file.read()
    
        def replaceEmailValues(emailSource: str, study: str, site: str, ivrs_number: str,
                            delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                            type_of_material: str, temperature: str, amount_of_boxes: int,
                            hasReturn: bool, type_of_return: str, amount_of_boxes_to_return: int,
                            tracking_number: str, return_tracking_number: str,
                            contacts: str, team_emails: str) -> str:
            emailSource = emailSource.replace("|VAR_STUDY|", study)
            emailSource = emailSource.replace("|VAR_SITE#|", site)
            emailSource = emailSource.replace("|VAR_IVRS_NUMBER|", ivrs_number)
            emailSource = emailSource.replace("|VAR_DELIVERY_DATE|", delivery_date)
            emailSource = emailSource.replace("|VAR_DELIVERY_TIME|", delivery_time_from + " to " + delivery_time_to)
            emailSource = emailSource.replace("|VAR_TYPE_OF_MATERIAL|", type_of_material)
            emailSource = emailSource.replace("|VAR_TEMPERATURE|", temperature)
            emailSource = emailSource.replace("|VAR_AMOUNT_OF_BOXES|", str(amount_of_boxes))
            emailSource = emailSource.replace("|VAR_TRACKING_NUMBER|", tracking_number)
            emailSource = emailSource.replace("|VAR_CONTACTS|", contacts)
            emailSource = emailSource.replace("|VAR_TEAM_EMAIL|", team_emails)
            emailSource = emailSource.replace("|VAR_TMO_LOGO|", os.getcwd() + "\\media\\TMO_logo_email.jpg")

            if hasReturn:
                emailSource = emailSource.replace("|VAR_TYPE_OF_RETURN|", type_of_return)
                emailSource = emailSource.replace("|VAR_AMOUNT_OF_BOXES_TO_RETURN|", str(amount_of_boxes_to_return))
                emailSource = emailSource.replace("|VAR_RETURN_TRACKING_NUMBER|", return_tracking_number)
            else:
                emailSource = emailSource.replace("|VAR_TYPE_OF_RETURN|", "NA")
                emailSource = emailSource.replace("|VAR_AMOUNT_OF_BOXES_TO_RETURN|", "0")
                emailSource = emailSource.replace("|VAR_RETURN_TRACKING_NUMBER|", "NA")

            return emailSource
        
        try:
            # Crear una instancia de Outlook
            outlook = win32.Dispatch('outlook.application')

            # Crear un nuevo mensaje de correo
            mail = outlook.CreateItem(0)

            emailSource = getEmailSourceFromTxtFile("email.txt")

            # Configurar los campos del correo
            mail.Subject = f"Orden de envío || Estudio: {study} || Site#: {site} || Número de IVRS: {ivrs_number}"
            mail.To = medical_center_emails
            
            customer_emails = customer_emails + "; " if customer_emails != "" else ""
            CRAs_emails = CRAs_emails + "; " if CRAs_emails != "" else ""
            team_emails = team_emails if team_emails != "" else ""
            mail.CC = customer_emails + CRAs_emails + team_emails

            emailSource = replaceEmailValues(emailSource, study, site, ivrs_number, 
                                             delivery_date, delivery_time_from, delivery_time_to, 
                                             type_of_material, temperature, amount_of_boxes, 
                                             hasReturn, type_of_return, amount_of_boxes_to_return, 
                                             tracking_number, return_tracking_number, contacts, team_emails)
            mail.HTMLBody = emailSource  # Configurar el cuerpo del correo como HTML

            # Enviar el correo
            mail.Send()
        except Exception as e:
            Log().add_log(f"Error sending email to medical center: {e}")

    def processOrdersAndContactsTable(self, ordersDataFrame:pd.DataFrame) -> pd.DataFrame:
        """
        Process all orders in the table
        
        Variables:
            wait (WebDriverWait): selenium wait
            team (str): team to process
            df_path (str): excel file path
            sheet (str): excel sheet name
            df (DataFrame): orders table

        Args:
            ordersDataFrame (DataFrame): orders table
            user (str): user of carrier website
            password (str): password of carrier website

        Returns:
            DataFrame: orders table with tracking numbers

        Preconditions:
            The orders table must be standardized as DataRecollector() class returns

            Website must be built and logged in
        """
        try:
            self.__process_all_shipping_orders__(ordersDataFrame)

            time.sleep(2) # Wait for the download to finish

            self.__renameAllReturnFiles__(ordersDataFrame)

            self.__export_to_excel__(ordersDataFrame)

        except Exception as e:
            Log().add_log(f"Error generating shipping report: {e}")
              
        finally:
            try:
                self.selectedTeam.quit_driver()
            except Exception as e:
                Log().add_log(f"Error quitting the webpage: {e}")
            finally:
                return ordersDataFrame

    # Private methods
    def __export_to_excel__(self, dataFrame: pd.DataFrame):
        FilesInSystemController().export_to_excel(dataFrame, self.folder_path_to_download)

    def __get_shipping_tracking_number__(self, carrier_id: int, system_number: str, ivrs_number: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        """
        Gets the shipping tracking number

        Args:
            carrier_id (int): Site ID on carrier website
            system_number (int): order system number
            ivrs_number (str): order ivrs number
            ship_date (str): ship date
            ship_time_from (str): ship time (from)
            ship_time_to (str): ship time (to)
            delivery_date (str): delivery date
            delivery_time_from (str): delivery time (from)
            delivery_time_to (str): delivery time (to)
            type_of_material (str): type of material
            temperature (str): temperature
            contacts (str): contacts
            amount_of_boxes (int): number of boxes

        Returns:
            str: tracking number
        """

        reference = f"{system_number} {ivrs_number}"[:50]
        delivery_date = dt.datetime.strptime(delivery_date, '%d/%m/%Y').strftime('%d/%m/%Y')

        tracking_number = self.selectedTeam.complete_shipping_order_form(
            carrier_id, reference, 
            ship_date, ship_time_from, ship_time_to, 
            delivery_date, delivery_time_from, delivery_time_to, 
            type_of_material, temperature, 
            contacts, amount_of_boxes
        )
        
        return tracking_number
    
    def __get_return_tracking_number__(self, carrier_id: int, system_number: str, ivrs_number: str,
                                    delivery_date: str,  tracking_number: str, hasReturn: bool,
                                    return_to_TA: bool, type_of_return: str, amount_of_boxes_to_return: int) -> str:
            """
            Gets the return tracking number
    
            Args:
                carrier_id (int): Site ID on carrier website
                system_number (int): order system number
                ivrs_number (str): order ivrs number
                delivery_date (str): delivery date
                tracking_number (str): tracking number
                hasReturn (bool): if True, creates a return order
                return_to_TA (bool): if True, the return order is sent to carrier depot
                type_of_return (str): type of return
                amount_of_boxes_to_return (int): number of boxes to return
    
            Returns:
                str: tracking number
            """
            if not hasReturn or tracking_number == "":
                return  ""

            reference_return = f"{system_number} {ivrs_number} RET {tracking_number}"[:50]
            transit_days = dt.timedelta(days=1)
            #transit_days = max(dt.datetime.strptime(delivery_date, '%d/%m/%Y') - dt.datetime.strptime(ship_date, '%d/%m/%Y'), dt.timedelta(days=1))
            return_delivery_date = dt.datetime.strptime(delivery_date, '%d/%m/%Y') + transit_days
            return_delivery_date += dt.timedelta(days=2) if return_delivery_date.weekday() >= 5 else dt.timedelta(days=0) # Add 2 days if the return delivery date is on a weekend
            return_delivery_date = return_delivery_date.strftime('%d/%m/%Y')
            return_delivery_hour_from = "9"
            return_delivery_hour_to = "17"
            
            return_tracking_number = self.selectedTeam.complete_shipping_order_return_form(
                carrier_id, reference_return, 
                return_delivery_date, return_delivery_hour_from, return_delivery_hour_to, 
                type_of_return, "", amount_of_boxes_to_return, 
                return_to_TA, tracking_number
            )
        
            return return_tracking_number
    
    def __renameAllReturnFiles__(self, ordersAndContactsDataframe: pd.DataFrame):
        FilesInSystemController().renameAllReturnFiles(ordersAndContactsDataframe, self.folder_path_to_download)

    def __process_all_shipping_orders__(self, ordersAndContactsDataframe: pd.DataFrame):
        """
        Process all orders in the table

        Args:
            ordersAndContactsDataframe (DataFrame): orders table
        """
        for index, row in ordersAndContactsDataframe.iterrows():
            if row['TRACKING_NUMBER'] != "":
                # Skip orders that have already been processed
                continue

            if row['HAS_AN_ERROR'] != "No error":
                # Skip orders with errors
                continue
            
            tracking_number, return_tracking_number = self.__get_tracking_numbers_from_carrier__(
                row["CARRIER_ID"], row["SYSTEM_NUMBER"], row["IVRS_NUMBER"],
                row["SHIP_DATE"], row["SHIP_TIME_FROM"], row["SHIP_TIME_TO"],
                row["DELIVERY_DATE"], row["DELIVERY_TIME_FROM"], row["DELIVERY_TIME_TO"],
                row["TYPE_OF_MATERIAL"], row["TEMPERATURE"],
                row["CONTACTS"], row["AMOUNT_OF_BOXES_TO_SEND"],
                row["HAS_RETURN"], row["RETURN_TO_CARRIER_DEPOT"], row["TYPE_OF_RETURN"], row["AMOUNT_OF_BOXES_TO_RETURN"]
            )

            if tracking_number == "":
                continue

            ordersAndContactsDataframe.loc[index, "TRACKING_NUMBER"] = tracking_number
            ordersAndContactsDataframe.loc[index, "RETURN_TRACKING_NUMBER"] = return_tracking_number
            
            self.__printOrderDocuments__(tracking_number, return_tracking_number, row["PRINT_RETURN_DOCUMENT"])

            medical_center_emails = row["MEDICAL_CENTER_EMAILS"]
            customer_email = ""
            cra_emails = ""
            team_email = ""
            delivery_date = "05/06/2024"

            try:
                self.send_email_to_medical_center( row["STUDY"], row["SITE#"], row["IVRS_NUMBER"], 
                delivery_date, row["DELIVERY_TIME_FROM"], row["DELIVERY_TIME_TO"], 
                row["TYPE_OF_MATERIAL"], row["TEMPERATURE"], row["AMOUNT_OF_BOXES_TO_SEND"],
                row["HAS_RETURN"], row["TYPE_OF_RETURN"], row["AMOUNT_OF_BOXES_TO_RETURN"],
                tracking_number, return_tracking_number,
                row["CONTACTS"], medical_center_emails, customer_email, cra_emails, team_email)
            except Exception as e:
                Log().add_log(f"Error sending email to medical center: {e}")
                Log().add_log(f"Order: {row['SYSTEM_NUMBER']} {row['IVRS_NUMBER']}")
            
            try:
                self.__updateTreeviewLine__(index, tracking_number, return_tracking_number)
            except Exception as e:
                Log().add_log(f"Error updating treeview line: {e}")

    def __printOrderDocuments__(self, tracking_number: str, return_tracking_number: str, print_return_document: bool):
        """
        Prints the order documents

        Args:
            tracking_number (str): tracking number
            return_tracking_number (str): return tracking number
            print_return_document (bool): if True, prints the return document
        """
        if tracking_number != "":
            self.selectedTeam.printWayBillDocument(tracking_number, 4)
            self.selectedTeam.printLabelDocument(tracking_number)
        else:
            return
        
        if print_return_document and return_tracking_number != "" and return_tracking_number != "Error":
            self.selectedTeam.printReturnWayBillDocument(return_tracking_number, 1)

    def __get_tracking_numbers_from_carrier__(self, carrier_id: int, system_number: str, ivrs_number: str, 
                               ship_date: str, ship_time_from: str, ship_time_to: str, 
                               delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                               type_of_material: str, temperature: str, 
                               contacts: str, amount_of_boxes: int, 
                               hasReturn: bool, return_to_TA: bool, type_of_return: str, amount_of_boxes_to_return: int) -> (str, str):
        """
        - Process an order by completing the carrier form
        - Creates a return order if necessary
        - Prints the order documents

        Args:
            carrier_id (int): Site ID on carrier website
            system_number (int): order system number
            ivrs_number (str): order ivrs number
            ship_date (str): ship date
            ship_time_from (str): ship time (from)
            ship_time_to (str): ship time (to)
            delivery_date (str): delivery date
            delivery_time_from (str): delivery time (from)
            delivery_time_to (str): delivery time (to)
            type_of_material (str): type of material
            temperature (str): temperature
            contacts (str): contacts
            amount_of_boxes (int): number of boxes
            hasReturn (bool): if True, creates a return order
            return_to_TA (bool): if True, the return order is sent to carrier depot
            amount_of_boxes_to_return (int): number of boxes to return

        Returns:
            str: tracking number
            str: return tracking number
        """
        
        tracking_number, return_tracking_number = "", ""

        try:
            tracking_number = self.__get_shipping_tracking_number__(
                carrier_id, system_number, ivrs_number,
                ship_date, ship_time_from, ship_time_to,
                delivery_date, delivery_time_from, delivery_time_to,
                type_of_material, temperature,
                contacts, amount_of_boxes
            )

            if tracking_number == "":
                return "", ""
            
            return_tracking_number = self.__get_return_tracking_number__(
                carrier_id, system_number, ivrs_number,
                delivery_date, tracking_number, hasReturn, return_to_TA, type_of_return, amount_of_boxes_to_return
            )

        except Exception as e:
            Log().add_log(f"Error processing order: {e}")
            Log().add_log(f"Order: {system_number} {ivrs_number}")

        finally:
            return tracking_number, return_tracking_number
 
    def __updateTreeviewLine__(self, index: int, tracking_number: str, return_tracking_number: str):
        """
        Updates a line in the treeview

        Args:
            index (int): row index
        """
        try:
            if self.controller is not None:
                self.controller.update_tag_color_of_a_treeview_line(index, tracking_number, return_tracking_number)
                time.sleep(1)
        except Exception as e:
            Log().add_log(f"Error updating treeview line: {e}")
    
class DataRecolector:
    def __init__(self, team):
        self.team = team
        self.columns_df = ['SYSTEM_NUMBER', 'IVRS_NUMBER', 'CUSTOMER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO', 
                'TYPE_OF_MATERIAL', 'TEMPERATURE', 'AMOUNT_OF_BOXES_TO_SEND',
                'HAS_RETURN', 'RETURN_TO_CARRIER_DEPOT', 'TYPE_OF_RETURN', 'AMOUNT_OF_BOXES_TO_RETURN',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT', 'CONTACTS', 'ROLE_OF_CONTACTS', 
                "MEDICAL_CENTER_EMAILS", "CUSTOMER_EMAIL", "CRA_EMAILS", "TEAM_EMAILS",
                'CARRIER_ID', 'HAS_AN_ERROR']
        
        self.columns_for_orders = ['SYSTEM_NUMBER', 'IVRS_NUMBER', 'CUSTOMER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'TYPE_OF_MATERIAL', 
                'TEMPERATURE', 'AMOUNT_OF_BOXES_TO_SEND', 'HAS_RETURN', 
                'RETURN_TO_CARRIER_DEPOT', 'TYPE_OF_RETURN', 'AMOUNT_OF_BOXES_TO_RETURN',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT']
        
        self.columns_for_sites = ["STUDY", "SITE#", "CARRIER_ID", 
                                  "DELIVERY_TIME_FROM", "DELIVERY_TIME_TO", 
                                  "CONTACTS", "ROLE_OF_CONTACTS", 
                                  "MEDICAL_CENTER_EMAILS", "CUSTOMER_EMAIL", "CRA_EMAILS", "TEAM_EMAILS"]
    
    def getColumnNames(self):
        return self.columns_df
    
    def recolectOrdersAndContactsData(self, shipdate: dt.datetime) -> pd.DataFrame:
        """
        Process all orders in the table

        Args:
            team (str): team to process
            shipdate (dt.datetime): date to process

        Returns:
            DataFrame: orders table with standarisized data
        """
        ordersDataframe = self.__load_shipping_order_table__(shipdate, self.team)
        contactsDataframe = self.__load_contacts_table__(self.team)
        ordersAndContactsDataframe = pd.merge(ordersDataframe, contactsDataframe, on=["STUDY", "SITE#"], how="left")

        ordersAndContactsDataframe["HAS_AN_ERROR"] = ordersAndContactsDataframe.apply(self.__checkErrorsOnEachOrder__, axis=1)
        
        return ordersAndContactsDataframe[self.columns_df]

    def getEmptyOrdersAndContactsData(self) -> pd.DataFrame:
        """
        Returns an empty orders table

        Returns:
            DataFrame: empty orders table
        """
        return pd.DataFrame(columns=self.columns_df)

    # Private methods
    def __load_shipping_order_table__(self, shipDate: dt.datetime, team) -> pd.DataFrame:
        """
        Loads orders table according to date and team

        Args:
            date (dt.datetime): date to process

        Returns:
            DataFrame: orders table
        """
        ordersDataFrame = self.__load_shipping_order_table_with_normalized_columns__(team)

        ordersDataFrame = ordersDataFrame[ordersDataFrame["SHIP_DATE"] == shipDate]
        
        ordersDataFrame = self.__correct_regular_columns_for_shipping_orders_table__(ordersDataFrame)

        ordersDataFrame = team.apply_team_specific_changes_for_orders_tables(ordersDataFrame)
        
        return ordersDataFrame[self.columns_for_orders]

    def __load_shipping_order_table_with_normalized_columns__(self, team) -> pd.DataFrame:
        """
        Loads orders table according to team

        Args:
            team (str): team to process
        """
        path_from_get_data, orders_sheet, _ = team.get_data_path()
            
        columns_names, columns_types = team.get_column_rename_type_config_for_orders_tables()

        ordersDataFrame = team.readOrdersExcel(path_from_get_data, orders_sheet, columns_types)

        ordersDataFrame.rename(columns=columns_names, inplace=True)

        return ordersDataFrame

    def __correct_regular_columns_for_shipping_orders_table__(self, ordersDataFrame: pd.DataFrame) -> pd.Series:
        """
        Corrects regular columns for orders table

        Args:
            ordersDataFrame (DataFrame): orders table
        """
        ordersDataFrame["SHIP_DATE"] = self.__correctDateColumns__(ordersDataFrame, "SHIP_DATE")
        ordersDataFrame["DELIVERY_DATE"] = self.__correctDateColumns__(ordersDataFrame, "DELIVERY_DATE")

        ordersDataFrame["TEMPERATURE"] = ordersDataFrame["TEMPERATURE"].str.strip()
        ordersDataFrame["SITE#"] = ordersDataFrame["SITE#"].astype(object)
        ordersDataFrame["AMOUNT_OF_BOXES_TO_SEND"] = ordersDataFrame["AMOUNT_OF_BOXES_TO_SEND"].replace('', '0').fillna('0').astype(int)
        
        shipSchedules = {"8": "08:00:00", "16.3": "16:30:00", "19": "19:00:00"}
        ordersDataFrame["SHIP_TIME_FROM"] = ordersDataFrame["SHIP_TIME_FROM"].replace(shipSchedules)
        
        ordersDataFrame["SHIP_TIME_FROM"] = pd.to_datetime(ordersDataFrame["SHIP_TIME_FROM"], format='%H:%M:%S', errors='coerce')
        ordersDataFrame["SHIP_TIME_TO"] = ordersDataFrame["SHIP_TIME_FROM"] + dt.timedelta(minutes=30)
        ordersDataFrame["SHIP_TIME_FROM"] = ordersDataFrame["SHIP_TIME_FROM"].dt.strftime('%H:%M')
        ordersDataFrame["SHIP_TIME_TO"] = ordersDataFrame["SHIP_TIME_TO"].dt.strftime('%H:%M')
        
        if "AMOUNT_OF_BOXES_TO_RETURN" in ordersDataFrame.columns:
            ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] = ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"].replace('', '0').fillna('0').astype(int)
        else:
            ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] = 0

        ordersDataFrame["TYPE_OF_RETURN"] = "NA"

        ordersDataFrame.fillna('', inplace=True)
        
        return ordersDataFrame

    def __load_contacts_table__(self, team) -> pd.DataFrame:
        """
        Loads sites info table according to team

        Args:
            team (str): team to process
            path_from_get_data (str): excel file path
            sheet (str): excel sheet name

        Returns:
            DataFrame: sites info table
        """
        contactsDataFrame = self.__load_contacts_table_with_normalized_columns__(team)

        contactsDataFrame = team.apply_team_specific_changes_for_contacts_table(contactsDataFrame)

        contactsDataFrame = self.__correct_regular_columns_for_contacts_table__(contactsDataFrame, team)

        return contactsDataFrame[self.columns_for_sites]
    
    def __load_contacts_table_with_normalized_columns__(self, team) -> pd.DataFrame:
        """
        Loads contacts table according to team

        Args:
            team (str): team to process
        """
        path_from_get_data, _, sites_sheet = team.get_data_path()
        columns_names, columns_types = team.get_column_rename_type_config_for_contacts_table()

        contactsDataFrame = team.readSitesExcel(path_from_get_data, sites_sheet, columns_types)
        contactsDataFrame.rename(columns=columns_names, inplace=True)

        return contactsDataFrame

    def __correct_regular_columns_for_contacts_table__(self, contactsDataFrame: pd.DataFrame, team) -> pd.Series:
        """
        Corrects regular columns for contacts table

        Args:
            contactsDataFrame (DataFrame): contacts table
        """
        contactsDataFrame["DELIVERY_TIME_FROM"] = self.__correctTimeColumns__(contactsDataFrame, "DELIVERY_TIME_FROM")
        contactsDataFrame["DELIVERY_TIME_TO"] = self.__correctTimeColumns__(contactsDataFrame, "DELIVERY_TIME_TO")

        contactsDataFrame["TEAM_EMAILS"] = team.getTeamEmail()

        contactsDataFrame = contactsDataFrame.drop_duplicates(subset=["STUDY", "SITE#"], keep='last')
        
        contactsDataFrame.fillna('', inplace=True)

        return contactsDataFrame

    def __checkErrorsOnEachOrder__(self, row: pd.Series) -> str:
        """
        Checks errors on each order

        Args:
            row (Series): order row
        """
        def assertIfIsNotNull(cell: str) -> bool:
            return cell != "" and cell != "nan"

        def assetIfCustomerIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['CUSTOMER'])
        
        def assertIfStudyIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['STUDY'])
        
        def assertIfSiteIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['SITE#'])
        
        def assertIfShipDateIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['SHIP_DATE'])
        
        def assertIfDeliveryDateIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['DELIVERY_DATE'])
        
        def assertIfAreValidDates(row: pd.Series) -> bool:
            return (row['SHIP_DATE'] <= row['DELIVERY_DATE']) and (row['SHIP_DATE'] >= dt.datetime.now().strftime('%d/%m/%Y'))

        def assertIfShipTimeFromIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['SHIP_TIME_FROM'])
        
        def assertIfShipTimeToIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['SHIP_TIME_TO'])
        
        def assertIfAreValidShipTimes(row: pd.Series) -> bool:
            return row['SHIP_TIME_FROM'] <= row['SHIP_TIME_TO']
        
        def assertIfDeliveryTimeFromIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['DELIVERY_TIME_FROM'])
        
        def assertIfDeliveryTimeToIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['DELIVERY_TIME_TO'])
        
        def assertIfAreValidDeliveryTimes(row: pd.Series) -> bool:
            return row['DELIVERY_TIME_FROM'] <= row['DELIVERY_TIME_TO']
        
        def assertIfTypeOfMaterialIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['TYPE_OF_MATERIAL'])
        
        def assertIfTypeOfMaterialIsValid(row: pd.Series) -> bool:
            return row['TYPE_OF_MATERIAL'] in ["Medicine", "Ancillary", "Equipment"]

        def assertIfAmountOfBoxesAreValid(row: pd.Series) -> bool:
            return row['AMOUNT_OF_BOXES_TO_SEND'] > 0
        
        def assertIfCarrier_IDIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['CARRIER_ID'])
        
        def assertIfTemperatureIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['TEMPERATURE'])

        def assertIfTemperatureIsValid(row: pd.Series) -> bool:
            return row['TEMPERATURE'] in ["Ambient", "Controlled Ambient", "Refrigerated"]
        
        def assertIfNumberOfBoxesToReturnIsValid(row: pd.Series) -> bool:
            return row['AMOUNT_OF_BOXES_TO_RETURN'] <= row['AMOUNT_OF_BOXES_TO_SEND']
        
        errors = ""

        if not assetIfCustomerIsNotEmpty(row):
            errors += "No customer; "

        if not assertIfStudyIsNotEmpty(row):
            errors += "No study; "

        if not assertIfSiteIsNotEmpty(row):
            errors += "No site; "

        if not assertIfShipDateIsNotEmpty(row):
            errors += "No ship date; "

        if not assertIfShipTimeFromIsNotEmpty(row):
            errors += "No ship time from; "

        if not assertIfShipTimeToIsNotEmpty(row):
            errors += "No ship time to; "

        if not assertIfAreValidShipTimes(row):
            errors += "Invalid ship times; "

        if not assertIfDeliveryDateIsNotEmpty(row):
            errors += "No delivery date; "
        
        if not assertIfDeliveryTimeFromIsNotEmpty:
            errors += "No delivery time from; "

        if not assertIfDeliveryTimeToIsNotEmpty:
            errors += "No delivery time to; "
        
        if not assertIfAreValidDeliveryTimes(row):
            errors += "Invalid delivery times; "

        #if not assertIfAreValidDates(row):
        #    errors += "Invalid dates; "

        if not assertIfTypeOfMaterialIsNotEmpty(row):
            errors += "No type of material; "

        if not assertIfTypeOfMaterialIsValid(row):
            errors += "Invalid type of material; "

        if not assertIfAmountOfBoxesAreValid(row):
            errors += "Invalid amount of boxes; "

        if not assertIfCarrier_IDIsNotEmpty(row):
            errors += "No carrier ID; "

        if not assertIfTemperatureIsNotEmpty(row):
            errors += "No temperature; "

        if not assertIfTemperatureIsValid(row):
            errors += "Invalid temperature; "

        if row['HAS_RETURN'] and not assertIfNumberOfBoxesToReturnIsValid(row):
            errors += "Invalid number of boxes to return; "

        return "No error" if errors == "" else errors

    def __correctTimeColumns__(self, dataFrame: pd.DataFrame, column: str) -> str:
        """
        Corrects times columns

        Args:
            dataFrame (DataFrame): Pandas DataFrame
            column (str): column name
        """
        return pd.to_datetime(dataFrame[column], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M')
    
    def __correctDateColumns__(self, dataFrame: pd.DataFrame, column: str) -> str:
        """
        Corrects dates columns

        Args:
            dataFrame (DataFrame): Pandas DataFrame
            column (str): column name
        """
        dataFrame[column] = dataFrame[column].astype("datetime64[ns]")
        dataFrame[column] = pd.to_datetime(dataFrame[column], format='%d/%m/%Y', errors='coerce')
        dataFrame[column] = dataFrame[column].dt.strftime('%d/%m/%Y')
        return dataFrame[column]

class FilesInSystemController:
    def create_folder(self, folder_path_to_download: str):
        """
        Creates a folder

        Args:
            folder_path_to_download (str): folder path to download
        """
        try:
            os.makedirs(folder_path_to_download, exist_ok=True)
        except FileExistsError:
            pass
    
    def getFolderPathToDownload(self, team: str, date: str) -> str:
        """
        Gets the folder path to download

        Args:
            team (str): team to process
            date (str): date to process

        Returns:
            str: folder path to download
        """
        downloads_path = os.path.expanduser("~\\Downloads")

        folder_path_to_download = os.path.join(downloads_path, "Tracking_Number_Generator", team, date)
        
        return folder_path_to_download
    
    def renameAllReturnFiles(self, ordersAndContactsDataframe: pd.DataFrame, folder_path_to_download: str):
        """
        Renames all return files

        Args:
            ordersAndContactsDataframe (DataFrame): return tracking numbers
        """
        dataFrameWithReturnTrackingNumbers = ordersAndContactsDataframe[(ordersAndContactsDataframe["RETURN_TRACKING_NUMBER"] != "") & (ordersAndContactsDataframe["PRINT_RETURN_DOCUMENT"])][["RETURN_TRACKING_NUMBER"]]
        
        if dataFrameWithReturnTrackingNumbers.empty:
            return

        for index, row in dataFrameWithReturnTrackingNumbers.iterrows():
            self.__renameReturnPDFFile__(row["RETURN_TRACKING_NUMBER"], folder_path_to_download)

    def export_to_excel(self, dataFrame: pd.DataFrame, folder_path_to_download: str):
        """
        Exports the orders table to an excel file

        Args:
            dataFrame (DataFrame): orders table
        """
        if not dataFrame.empty:
            dataframe_name = "orders_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            dataFrame.to_excel(folder_path_to_download + "\\" + dataframe_name, index=False)
        else:
            Log().add_log("Empty DataFrame")

    def zip_folder(self, folder_path: str, zip_path: str):
            shutil.make_archive(zip_path, 'zip', folder_path)

    def __renameReturnPDFFile__(self, return_tracking_number: str, folder_path_to_download: str):
        """
        Renames the return file

        Args:
            return_tracking_number (str): return tracking number
        """
        try:
            pdf_path = folder_path_to_download + "\\JOB " + return_tracking_number + ".pdf"
            new_pdf_path = folder_path_to_download + "\\JOB " + return_tracking_number + " RETORNO DE CREDO.pdf"
            os.rename(pdf_path, new_pdf_path)
        except Exception as e:
            Log().add_log(f"Error renaming return file: {e}")

class SingletonMeta(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        """
        Possible changes to the value of the `__init__` argument do not affect
        the returned instance.
        """
        if cls not in cls._instances:
            instance = super().__call__(*args, **kwargs)
            cls._instances[cls] = instance
        return cls._instances[cls]

class Log(metaclass=SingletonMeta):
    def __init__(self):
        self.logs = []

    def add_log(self, text: str):
        date_and_time = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.logs.append(f"{date_and_time} - {text}")

    def add_separator(self):
        self.logs.append("----------------------------------------")

    def print_logs(self) -> str:
        return "\n".join(self.logs)
    
    def print_last_n_logs(self, n: int) -> str:
        n = min(n, len(self.logs))
        return "\n".join(self.logs[-n:])
    
    def save_logs(self, folder_path_to_download: str):
        with open(folder_path_to_download + "\\logs.txt", "w") as file:
            file.write(self.print_logs())

    def clear_logs(self):
        self.logs = []
                
class CarriersWebpage:
    def __init__(self, carrier: str = "", folder_path_to_download: str = ""):
        """
        Class constructor
        """

        self.carrierWebpageNames = ["Transportes Ambientales"]

        match carrier:
            case "Transportes Ambientales":
                self.carrierWebpage = self.TransportesAmbientales(folder_path_to_download, self)
            
            case _:
                self.carrierWebpage = self.NoCarrier(folder_path_to_download, self)

    # Methods of Super class
    def getCarriersWebpagesNames(self):
        return self.carrierWebpageNames
    
    # Methods to be implemented by each sub class
    def build_driver(self):
        """
        Builds the driver
        """
        self.carrierWebpage.build_driver()
    
    def quit_driver(self):
        """
        Quits the browser
        """
        self.carrierWebpage.quit_driver()

    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        """
        Logs in webpage

        Args:
            driver (webdriver): selenium driver
            username (str): username
            password (str): password
        """
        return self.carrierWebpage.check_if_user_and_password_are_correct(username, password)

    def complete_shipping_order_form(self, carrier_id: str, reference: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        """
        Completes carrier form

        Args:
            self.driver (webdriver): selenium self.driver
            carrier_id (int): Site ID on carrier website
            reference (str): order reference
            ship_date (str): ship date
            ship_time_from (str): ship time (from)
            ship_time_to (str): ship time (to)
            delivery_date (str): delivery date
            delivery_time_from (str): delivery time (from)
            delivery_time_to (str): delivery time (to)
            type_of_material (str): type of material
            temperature (str): temperature
            contacts (str): contacts
            amount_of_boxes (int): number of boxes

        Returns:
            str: tracking number
        """
        return self.carrierWebpage.complete_shipping_order_form(carrier_id, reference,
                                    ship_date, ship_time_from, ship_time_to,
                                    delivery_date, delivery_time_from, delivery_time_to,
                                    type_of_material, temperature,
                                    contacts, amount_of_boxes)
    
    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
        """
        Completes carrier return form

        Args:
            self.driver (webdriver): selenium self.driver
            carrier_id (int): Site ID on carrier website
            reference_return (str): return reference
            delivery_date (str): delivery date
            return_time_from (str): return time (from)
            return_time_to (str): return time (to)
            type_of_return (str): type of return
            contacts (str): contacts
            amount_of_boxes_to_return (int): number of boxes to return
            return_to_TA (bool): if True, the return order is sent to carrier depot
            tracking_number (str): tracking number

        Returns:
            str: return tracking number
        """
        return self.carrierWebpage.complete_shipping_order_return_form(carrier_id, reference_return,
                                                delivery_date, return_time_from,
                                                return_time_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)
    
    def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
        """
        Prints waybill documents

        Args:
        """
        self.carrierWebpage.printWayBillDocument(tracking_number, amount_of_copies)

    def printLabelDocument(self, tracking_number: str):
        """
        Prints label documents

        Args:
        """
        self.carrierWebpage.printLabelDocument(tracking_number)

    def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
        """
        Prints return waybill documents

        Args:
        """
        self.carrierWebpage.printReturnWayBillDocument(return_tracking_number, amount_of_copies)

    def print_webpage(self, url: str):
        """
        Prints webpage

        Args:
            self.driver (webdriver): selenium self.driver
            url (str): webpage url
        """
        self.carrierWebpage.print_webpage(url)

    # Private methods
    def __print_webpage__(self, driver, url: str):
        try:
            driver.get(url) # this print since chrome options are set to print automatically
            driver.implicitly_wait(5)

        except Exception as e:
            Log().add_log(f"Error printing documents: {e}")
    
    def __quit_driver__(self, driver):
        try:
            driver.quit()
        except Exception as e:
            Log().add_log(f"Error quitting driver: {e}")

    # Sub classes
    class NoCarrier:
        def __init__(self, folder_path_to_download: str = "", father = None):
            """
            Class constructor for NoCarrier

            Args:
                driver (webdriver): selenium driver
                father (CarriersWebpage): father class
            """
            self.folder_path_to_download = folder_path_to_download
            self.father = father

        def build_driver(self):
            self.browser = Browser(self.folder_path_to_download)
            self.driver = self.browser.driver
            self.wait = WebDriverWait(self.driver, 10)

        def quit_driver(self):
            self.father.__quit_driver__(self.driver)
        
        def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
            return False

        def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                    ship_date: str, ship_time_from: str, ship_time_to: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
            return ""

        def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
            return ""
        
        def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
            return ""
        
        def printLabelDocument(self, tracking_number: str):
            return ""
        
        def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
            return ""
        
        def print_webpage(self, url: str):
            return ""
    
    class TransportesAmbientales:
        def __init__(self, folder_path_to_download: str = "", father = None):
            """
            Class constructor for Transportes Ambientales

            Args:
                driver (webdriver): selenium driver
                father (CarriersWebpage): father class
            """
            self.folder_path_to_download = folder_path_to_download
            self.father = father

        def build_driver(self):
            self.browser = Browser(self.folder_path_to_download)
            self.driver = self.browser.driver
            self.wait = WebDriverWait(self.driver, 10)

        def quit_driver(self):
            self.father.__quit_driver__(self.driver)

        def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
            """
            Checks if user can log in webpage

            Args:
                driver (webdriver): selenium driver
                username (str): username
                password (str): password
            """
            self.driver.get("https://sgi.tanet.com.ar/sgi/index.php")
            
            self.complete_login_form(username, password)

            try:
                time.sleep(1)
                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[2]/h1")))
                return True
            except Exception as e:
                Log().add_log(f"Error logging in webpage: {e}")
            
            return False
        
        def complete_login_form(self, username: str, password: str):
            """
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/input[1]")))
            Completes login form

            Args:
                driver (webdriver): selenium driver
                username (str): username
                password (str): password
            """
            self.driver.find_element(By.XPATH, "/html/body/form/input[1]").send_keys(username)
            self.driver.find_element(By.XPATH, "/html/body/form/input[2]").send_keys(password)
            self.driver.find_element(By.XPATH, "/html/body/form/button").click()

        def complete_shipping_order_form(self, carrier_id: str, reference: str, 
                                    ship_date: str, ship_time_from: str, ship_time_to: str, 
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
            tracking_number = ""
            url = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarEnvio+idubicacion={carrier_id}"

            try:
                self.driver.get(url)

                # Wait for the webpage to load
                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/input")))

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").send_keys(reference)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[1]").send_keys(ship_date)
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[2]").send_keys(ship_time_from)
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[5]/td/input[3]").send_keys(ship_time_to)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[1]").send_keys(delivery_date)
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[2]").send_keys(delivery_time_from)
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[3]").send_keys(delivery_time_to)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/input").send_keys("Fisher Clinical Services FCS")
                self.driver.implicitly_wait(5)

                suggestions_container = self.wait.until(EC.presence_of_element_located((By.ID, "suggest_nomDomOri_list")))
                self.driver.implicitly_wait(5)
                time.sleep(1)

                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[7]/td/div[2]/table/tbody/tr/td[1]/div/div/div[1]/table/tbody")))
                self.driver.implicitly_wait(5)
                time.sleep(1)

                # Finds all suggested buttons inside the suggestions container
                suggested_buttons = suggestions_container.find_elements(By.CLASS_NAME, "suggest")
                # Loops through the suggested buttons and finds the one that contains the text "Fisher Clinical Services FCS"
                for button in suggested_buttons:
                    button_text = button.text.strip()
                    if "Fisher Clinical Services FCS" in button_text:
                        button.click()
                        break

                # Selects the material type
                if type_of_material == "Medicine": it_type_of_material = 3
                elif type_of_material == "Ancillary": it_type_of_material = 5
                elif type_of_material == "Equipment": it_type_of_material = 7
                else: it_type_of_material = 8

                for i in range(0, it_type_of_material):
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[1]").send_keys(Keys.DOWN)

                # Selects the temperature
                if temperature == "Ambient": it_temperature = 0
                elif temperature == "Controlled Ambient": it_temperature = 1
                elif temperature == "Refrigerated": it_temperature = 2
                elif temperature == "Congelado": it_temperature = 3
                else: it_temperature = 4

                for i in range(0, it_temperature):
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[10]/td/select[2]").send_keys(Keys.DOWN)

                observaciones_textbox = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/textarea")
                comments = ""
                for line in observaciones_textbox.text.split("\n"):
                    if "Dias y horarios de entrega:" != line[:27]:
                        comments += line + "\n"

                if temperature != "Ambient":
                    comments += "***ENVÍO CON CAJA CREDO. EL COURIER AGUARDARÁ QUE EL CENTRO ALMACENE LA MEDICACIÓN Y RETORNE EL EMBALAJE***"

                observaciones_textbox.clear()
                observaciones_textbox.send_keys(comments)

                if contacts != "" and contacts != "No contact":
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").clear()
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[11]/td/input[1]").send_keys(contacts)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input").send_keys(amount_of_boxes)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/button").click()

                # Wait for the webpage to load and get the tracking number
                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))

                self.driver.implicitly_wait(5)

                tracking_number = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:14]
                
            except Exception as e:
                Log().add_log(f"Error completing shipping order form: {e}")
                Log().add_log(f"Order: {reference}")

            finally:
                return tracking_number

        def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
            return_tracking_number = ""
            url_return = f"https://sgi.tanet.com.ar/sgi/srv.SrvCliente.editarRetorno+idsp={tracking_number[:7]}&idubicacion={carrier_id}"
            url_return += "&returnToTa=true" if return_to_TA else ""

            try:
                self.driver.get(url_return)
                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select")))
                
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").clear()
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[3]/td/input").send_keys(reference_return)
                
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[1]").send_keys(delivery_date)
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[2]").send_keys(return_time_from)
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[6]/td/input[3]").send_keys(return_time_to)

                # Selects the return type
                if type_of_return == "CREDO": it_tipo_retorno = 1
                elif type_of_return == "DATALOGGER": it_tipo_retorno = 2
                else: it_tipo_retorno = 3
                for i in range(0, it_tipo_retorno):
                    self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[9]/td/select").send_keys(Keys.DOWN)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[12]/td/input[1]").send_keys(contacts)
                
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").clear()
                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[13]/td/input").send_keys(amount_of_boxes_to_return)

                self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[4]/table/tbody/tr[14]/td/button").click()
                
                # Wait for the webpage to load and get the return tracking number
                self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption")))
                
                self.driver.implicitly_wait(5)
                
                return_tracking_number = self.driver.find_element(By.XPATH, "/html/body/form/div[2]/div/div[3]/div[5]/table/caption").text[5:14]
            except Exception as e:
                Log().add_log(f"Error completing shipping order return form: {e}")
                Log().add_log(f"Order: {reference_return}")
                return "Error"
                
            finally:
                return return_tracking_number
        
        def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
            url_guias = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id={tracking_number[:7]}&idservicio={tracking_number[:7]}&copies={amount_of_copies}"
            self.print_webpage(url_guias)

        def printLabelDocument(self, tracking_number: str):
            url_rotulo = f"https://sgi.tanet.com.ar/sgi/srv.RotuloPdf.emitir+id={tracking_number[:7]}&idservicio={tracking_number[:7]}"
            self.print_webpage(url_rotulo)

        def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
            url_guias_return = f"https://sgi.tanet.com.ar/sgi/srv.SrvPdf.emitirOde+id={return_tracking_number[:7]}&idservicio={return_tracking_number[:7]}&copies={amount_of_copies}"
            self.print_webpage(url_guias_return)
        
        def print_webpage(self, url: str):
            self.father.__print_webpage__(self.driver, url)

class Teams:
    def __init__(self, team: str = "", folder_path_to_download: str = ""):
        """
        Class constructor for teams
        """

        self.teamsNames = ["Eli Lilly Argentina", "GPM Argentina", "Test_5_ordenes"]
        
        match team:
            case "Eli Lilly Argentina":
                self.selectedTeam = self.EliLillyArgentinaTeam(folder_path_to_download, self)
            case "GPM Argentina":
                self.selectedTeam = self.GPMArgentinaTeam(folder_path_to_download, self)
            case "Test_5_ordenes":
                self.selectedTeam = self.TestTeam(folder_path_to_download, self)

            # ---------------------
            case _:
                self.selectedTeam = self.NoSelectedTeam("", self)

    # Methods of Super class
    def getTeamsNames(self) -> list:
        """
        Gets teams names
        """
        return self.teamsNames
    
    # Methods to be implemented by each sub class
    def getTeamName(self) -> str:
        """
        Gets team name
        """
        return self.selectedTeam.getTeamName()
    
    def getTeamEmail(self) -> str:
        """
        Gets team email
        """
        return self.selectedTeam.getTeamEmail()

    def get_column_rename_type_config_for_contacts_table(self) -> (dict, dict):
        """
        Loads columns names and types for the sites info table

        Returns:
            dict: columns names
            dict: columns types
        """
        return self.selectedTeam.get_column_rename_type_config_for_contacts_table()
    
    def apply_team_specific_changes_for_contacts_table(self, contactsDataframe: pd.DataFrame) -> pd.DataFrame:
        """
        Applies team specific changes to sites info table

        Args:
            contactsDataframe (DataFrame): sites info table

        Returns:
            DataFrame: sites info table
        """
        return self.selectedTeam.apply_team_specific_changes_for_contacts_table(contactsDataframe)
    
    def get_data_path(self) -> (str, str, str):
        """
        Loads data path

        Returns:
            str: excel file path
            str: excel sheet name
            str: excel sheet name with sites info
        """
        return self.selectedTeam.get_data_path()
    
    def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
        """
        Loads columns names and types for the orders table

        Returns:
            dict: columns names
            dict: columns types
        """
        return self.selectedTeam.get_column_rename_type_config_for_orders_tables()
    
    def apply_team_specific_changes_for_orders_tables(self, ordersDataframe: pd.DataFrame) -> pd.DataFrame:
        """
        Applies team specific changes to orders table

        Args:
            ordersDataframe (DataFrame): orders table

        Returns:
            DataFrame: orders table
        """
        try:
            return self.selectedTeam.apply_team_specific_changes_for_orders_tables(ordersDataframe)
        except Exception as e:
            Log().add_log(f"Error applying team specific changes for orders tables: {e}")
            return ordersDataframe
    
    def sendEmailWithOrdersToTeam(self, folder_path_with_orders_files: str, date: str):
        """
        Sends an email with the orders table to the team

        Args:
            folder_path_with_orders_files (str): folder path with orders files
            date (str): date
        """
        self.selectedTeam.sendEmailWithOrdersToTeam(folder_path_with_orders_files, date)

    def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
        return self.selectedTeam.readOrdersExcel(path_from_get_data, orders_sheet, columns_types)
    
    def readSitesExcel(self, path_from_get_data: str, sites_sheet: str, columns_types: dict) -> pd.DataFrame:
        return self.selectedTeam.readSitesExcel(path_from_get_data, sites_sheet, columns_types)

    def build_driver(self):
        self.selectedTeam.build_driver()

    def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
        return self.selectedTeam.check_if_user_and_password_are_correct(username, password)
    
    def quit_driver(self):
        self.selectedTeam.quit_driver()

    def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                    ship_date: str, ship_time_from: str, ship_time_to: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        return self.selectedTeam.complete_shipping_order_form(carrier_id, reference,
                                    ship_date, ship_time_from, ship_time_to,
                                    delivery_date, delivery_time_from, delivery_time_to,
                                    type_of_material, temperature,
                                    contacts, amount_of_boxes)

    def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                            return_delivery_date: str, return_delivery_hour_from: str,
                                            return_delivery_hour_to: str, type_of_return: str,
                                            contacts: str, amount_of_boxes_to_return: int,
                                            return_to_TA: bool, tracking_number: str) -> str:
        return self.selectedTeam.complete_shipping_order_return_form(carrier_id, reference_return,
                                                return_delivery_date, return_delivery_hour_from,
                                                return_delivery_hour_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)

    def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
        self.selectedTeam.printWayBillDocument(tracking_number, amount_of_copies)

    def printLabelDocument(self, tracking_number: str):
        self.selectedTeam.printLabelDocument(tracking_number)

    def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
        self.selectedTeam.printReturnWayBillDocument(return_tracking_number, amount_of_copies)

    # Private methods
    def __sendEmailWithOrdersToTeam__(self, folder_path_with_orders_files: str, date: str, emails_of_team: str, emails_of_admin: str):
        try:
            # Obtener la carpeta padre
            parent_folder = os.path.dirname(folder_path_with_orders_files)
            
            # Nombre del archivo ZIP sin extensión
            zip_filename = 'orders_' + self.getTeamName() + '_' + date
            
            # Ruta completa del archivo ZIP
            zip_path = os.path.join(parent_folder, zip_filename)

            # Crear el archivo ZIP
            FilesInSystemController().zip_folder(folder_path_with_orders_files, zip_path)

            # Ruta del archivo ZIP con extensión
            zip_file_with_extension = zip_path + '.zip'

            outlook = win32.Dispatch('outlook.application')

            mail = outlook.CreateItem(0)
            mail.To = emails_of_team
            mail.Cc = emails_of_admin
            mail.Subject = f"Ordenes de envio con despacho {date} - {self.getTeamName()}"
            mail.Body = "Adjunto encontrarán las órdenes para el día de la fecha."

                # Adjuntar el archivo ZIP
            mail.Attachments.Add(zip_file_with_extension)

            mail.Send()
        except Exception as e:
            Log().add_log(f"Error sending email with orders to team: {e}")

    def __check_if_user_and_password_are_correct__(self, carrierWebpage: CarriersWebpage, username: str, password: str) -> bool:
        return carrierWebpage.check_if_user_and_password_are_correct(username, password)
    
    def __quit_driver__(self, carrierWebpage: CarriersWebpage):
        carrierWebpage.quit_driver()

    def __complete_shipping_order_form__(self, carrierWebpage: CarriersWebpage, carrier_id: str, reference: str,
                                    ship_date: str, ship_time_from: str, ship_time_to: str,
                                    delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                    type_of_material: str, temperature: str,
                                    contacts: str, amount_of_boxes: int) -> str:
        return carrierWebpage.complete_shipping_order_form(carrier_id, reference,
                                    ship_date, ship_time_from, ship_time_to,
                                    delivery_date, delivery_time_from, delivery_time_to,
                                    type_of_material, temperature,
                                    contacts, amount_of_boxes)

    def __complete_shipping_order_return_form__(self, carrierWebpage: CarriersWebpage, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
        return carrierWebpage.complete_shipping_order_return_form(carrier_id, reference_return,
                                                delivery_date, return_time_from,
                                                return_time_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)
    
    def __printWayBillDocument__(self, carrierWebpage: CarriersWebpage, tracking_number: str, amount_of_copies: int):
        carrierWebpage.printWayBillDocument(tracking_number, amount_of_copies)

    def __printLabelDocument__(self, carrierWebpage: CarriersWebpage, tracking_number: str):
        carrierWebpage.printLabelDocument(tracking_number)

    def __printReturnWayBillDocument__(self, carrierWebpage: CarriersWebpage, return_tracking_number: str, amount_of_copies: int):
        carrierWebpage.printReturnWayBillDocument(return_tracking_number, amount_of_copies)

    def __build_driver__(self, carrierWebpage: CarriersWebpage):
        carrierWebpage.build_driver()

    # Sub classes
    class NoSelectedTeam:
        def __init__(self, folder_path_to_download: str, father = None):
            self.carrierWebpage = CarriersWebpage("", folder_path_to_download)
            self.father = father

        def getTeamName(self) -> str:
            return "No Selected Team"
        
        def getTeamEmail(self) -> str:
            return ""
        
        def get_column_rename_type_config_for_contacts_table(self) -> (dict, dict):
            return {}, {}
        
        def apply_team_specific_changes_for_contacts_table(self, contactsDataframe: pd.DataFrame) -> pd.DataFrame:
            return contactsDataframe
        
        def get_data_path(self) -> (str, str, str):
            return ""
        
        def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
            return {}, {}
        
        def apply_team_specific_changes_for_orders_tables(self, ordersDataframe: pd.DataFrame) -> pd.DataFrame:
            return ordersDataframe

        def sendEmailWithOrdersToTeam(self, df: pd.DataFrame, date: str):
            return ""
        
        def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
            return DataRecolector(Teams()).getEmptyOrdersAndContactsData()
        
        def readSitesExcel(self, path_from_get_data: str, sites_sheet: str, columns_types: dict) -> pd.DataFrame:
            return DataRecolector(Teams()).getEmptyOrdersAndContactsData()

        def build_driver(self):
            self.father.__build_driver__(self.carrierWebpage)

        def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
            return self.father.__check_if_user_and_password_are_correct__(self.carrierWebpage, username, password)
        
        def quit_driver(self):
            self.father.__quit_driver__(self.carrierWebpage)

        def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                        ship_date: str, ship_time_from: str, ship_time_to: str,
                                        delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                        type_of_material: str, temperature: str,
                                        contacts: str, amount_of_boxes: int) -> str:
            return self.father.__complete_shipping_order_form__(self.carrierWebpage, carrier_id, reference,
                                        ship_date, ship_time_from, ship_time_to,
                                        delivery_date, delivery_time_from, delivery_time_to,
                                        type_of_material, temperature,
                                        contacts, amount_of_boxes)
        
        def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
            return self.father.__complete_shipping_order_return_form__(self.carrierWebpage, carrier_id, reference_return,
                                                delivery_date, return_time_from,
                                                return_time_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)
        
        def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
            self.father.__printWayBillDocument__(self.carrierWebpage, tracking_number, amount_of_copies)

        def printLabelDocument(self, tracking_number: str):
            self.father.__printLabelDocument__(self.carrierWebpage, tracking_number)

        def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
            self.father.__printReturnWayBillDocument__(self.carrierWebpage, return_tracking_number, amount_of_copies)
    
    class EliLillyArgentinaTeam:
        def __init__(self, folder_path_to_download: str, father = None):
            self.carrierWebpage = CarriersWebpage("Transportes Ambientales", folder_path_to_download)
            self.father = father

        def getTeamName(self) -> str:
            return "Eli Lilly Argentina"

        def getTeamEmail(self) -> str:
            return "guido.hendl@thermofisher.com; florencia.acosta@thermofisher.com"
            return "AR.Lilly.logistics@fishersci.com"

        def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
            try:
                df = pd.read_excel(path_from_get_data, sheet_name=orders_sheet, dtype=columns_types, header=0)
            except Exception as e:
                Log().add_log(f"Error reading orders excel: {e}")
                df = DataRecolector(Teams()).getEmptyOrdersAndContactsData()

            return df
        
        def readSitesExcel(self, path_from_get_data: str, sites_sheet: str, columns_types: dict) -> pd.DataFrame:
            try:
                df = pd.read_excel(path_from_get_data, sheet_name=sites_sheet, dtype=columns_types, header=0)
            except Exception as e:
                Log().add_log(f"Error reading sites excel: {e}")
                df = DataRecolector(Teams()).getEmptyOrdersAndContactsData()
            return df

        def get_column_rename_type_config_for_contacts_table(self) -> (dict, dict):
            columns_names = {"Protocolo": "STUDY", "Codigo": "CARRIER_ID", "Site": "SITE#",
                            "Horario inicio": "DELIVERY_TIME_FROM", "Horario fin": "DELIVERY_TIME_TO"}
            columns_types = {"Protocolo": str, "Site": str, "Codigo": str, "Horario inicio": str, "Horario fin": str}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_contacts_table(self, contactsDataFrame: pd.DataFrame) -> pd.DataFrame:
            contactsDataFrame["CONTACTS"] = "No contact"
            contactsDataFrame["ROLE_OF_CONTACTS"] = "Can receive medicines"
            contactsDataFrame["MEDICAL_CENTER_EMAILS"] = ""
            contactsDataFrame["CUSTOMER_EMAIL"] = ""
            contactsDataFrame["CRA_EMAILS"] = ""
            return contactsDataFrame
        
        def get_data_path(self) -> (str, str, str):
            path_from_get_data = os.path.expanduser("~\\Thermo Fisher Scientific\Power BI Lilly Argentina - General\Share Point Lilly Argentina.xlsx")
            orders_sheet = "Shipments"
            info_sites_sheet = "Dias y Destinos"
            return path_from_get_data, orders_sheet, info_sites_sheet
        
        def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
            columns_names = {"CT-WIN": "SYSTEM_NUMBER", "IVRS": "IVRS_NUMBER",
                            "Trial Alias": "STUDY", "Site ": "SITE#",
                            "Order received": "ENTER DATE", "Ship date": "SHIP_DATE",
                            "Horario de Despacho": "SHIP_TIME_FROM",  
                            "Dia de entrega": "DELIVERY_DATE", "Destination": "DESTINATION",
                            "CONDICION": "TEMPERATURE", "TT4": "AMOUNT_OF_BOXES_TO_SEND",  
                            "AWB": "TRACKING_NUMBER", "Shipper return AWB": "RETURN_TRACKING_NUMBER"}
            columns_types = {"CT-WIN": str, "IVRS": str, 
                            "Trial Alias": str, "Site ": str, 
                            "Order received": str, "Ship date": dt.datetime,
                            "Horario de Despacho": str,
                            "Dia de entrega": dt.datetime, "Destination": str,
                            "CONDICION": str, "TT4": str, 
                            "AWB": str, "Shipper return AWB": str}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_orders_tables(self, ordersDataFrame: pd.DataFrame) -> pd.DataFrame:
            ordersDataFrame["TYPE_OF_MATERIAL"] = "Medicine"
            ordersDataFrame["CUSTOMER"] = "Eli Lilly and Company"

            temperatures = {"L": "Ambient",
                            "M": "Controlled Ambient", "M + L": "Controlled Ambient",
                            "H": "Controlled Ambient", "H + M": "Controlled Ambient", "H + L": "Controlled Ambient", "H + M + L": "Controlled Ambient",
                            "REF": "Refrigerated", "REF + H": "Refrigerated", "REF + M": "Refrigerated", "REF + L": "Refrigerated",
                            "REF + H + M": "Refrigerated", "REF + H + L": "Refrigerated", "REF + M + L": "Refrigerated",
                            "REF + H + M + L": "Refrigerated"}
            ordersDataFrame["TEMPERATURE"] = ordersDataFrame["TEMPERATURE"].replace(temperatures)
            ordersDataFrame.loc[(ordersDataFrame["TEMPERATURE"] == "Ambient") & (ordersDataFrame["RETURN_TRACKING_NUMBER"] != "N"), "TEMPERATURE"] = "Controlled Ambient"
            
            ordersDataFrame["Cajas (Carton)"] = ordersDataFrame["Cajas (Carton)"].replace("", 0).fillna(0).astype(int)
            ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] = ordersDataFrame["AMOUNT_OF_BOXES_TO_SEND"] - ordersDataFrame["Cajas (Carton)"]
            
            ordersDataFrame["RETURN_TO_CARRIER_DEPOT"] = False
            
            ordersDataFrame["HAS_RETURN"] = (ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & (ordersDataFrame["TEMPERATURE"] != "Ambient")
            ordersDataFrame.loc[ordersDataFrame["HAS_RETURN"], "TYPE_OF_RETURN"] = "CREDO"

            ordersDataFrame["PRINT_RETURN_DOCUMENT"] = ordersDataFrame["HAS_RETURN"]

            return ordersDataFrame

        def sendEmailWithOrdersToTeam(self, folder_path_with_orders_files: str, date: str):
            self.father.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "inaki.costa@thermofisher")

        def build_driver(self):
            self.father.__build_driver__(self.carrierWebpage)

        def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
            return self.father.__check_if_user_and_password_are_correct__(self.carrierWebpage, username, password)
        
        def quit_driver(self):
            self.father.__quit_driver__(self.carrierWebpage)

        def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                        ship_date: str, ship_time_from: str, ship_time_to: str,
                                        delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                        type_of_material: str, temperature: str,
                                        contacts: str, amount_of_boxes: int) -> str:
            return self.father.__complete_shipping_order_form__(self.carrierWebpage, carrier_id, reference,
                                        ship_date, ship_time_from, ship_time_to,
                                        delivery_date, delivery_time_from, delivery_time_to,
                                        type_of_material, temperature,
                                        contacts, amount_of_boxes)
        
        def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
            return self.father.__complete_shipping_order_return_form__(self.carrierWebpage, carrier_id, reference_return,
                                                delivery_date, return_time_from,
                                                return_time_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)
        
        def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
            self.father.__printWayBillDocument__(self.carrierWebpage, tracking_number, amount_of_copies)

        def printLabelDocument(self, tracking_number: str):
            self.father.__printLabelDocument__(self.carrierWebpage, tracking_number)

        def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
            self.father.__printReturnWayBillDocument__(self.carrierWebpage, return_tracking_number, amount_of_copies)
        
    class GPMArgentinaTeam:
        def __init__(self, folder_path_to_download: str, father = None):
            self.carrierWebpage = CarriersWebpage("Transportes Ambientales", folder_path_to_download)
            self.father = father

        def getTeamName(self) -> str:
            return "GPM Argentina"

        def getTeamEmail(self) -> str:
            return "inaki.costa@thermofisher.com"

        def get_column_rename_type_config_for_contacts_table(self) -> (dict, dict):
            columns_names = {"Linea de facturacion" : "STUDY", "Site": "SITE#", "Site ID": "CARRIER_ID",
                             "Persona de contacto" : "CONTACTS", "Emails": "MEDICAL_CENTER_EMAILS", "Emails2": "CUSTOMER_EMAIL", "Emails3": "CRA_EMAILS"}
            columns_types = {"Linea de facturacion": str, "Site": str, "Site ID": str, "Persona de contacto": str}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_contacts_table(self, df: pd.DataFrame) -> pd.DataFrame:
            df["STUDY"] = df["STUDY"].str.strip()
            df["SITE#"] = df["SITE#"].str.strip()
            df["CUSTOMER_EMAIL"] = ""
            df["CRA_EMAILS"] = ""
            return df
        
        def get_data_path(self) -> (str, str, str):
            path_from_get_data = os.path.expanduser("~\\Downloads\orderTracker_GPM.xlsx")
            orders_sheet = "Ordenes"
            info_sites_sheet = "Contactos"
            return path_from_get_data, orders_sheet, info_sites_sheet
        
        def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
            columns_names = {}
            columns_types = {"SYSTEM_NUMBER": str, "IVRS_NUMBER": str,"STUDY": str, "SITE#": str, 
                             "SHIP_DATE": dt.datetime, "SHIP_TIME_FROM": dt.datetime, "SHIP_TIME_TO": dt.datetime,
                             "DELIVERY_DATE": dt.datetime, "DELIVERY_TIME_FROM": dt.datetime, "DELIVERY_TIME_TO": dt.datetime,
                             "TYPE_OF_MATERIAL": str, "TEMPERATURE": str, 
                             "AMOUNT_OF_BOXES_TO_SEND": str, 
                             "HAS_RETURN": bool, 
                             "RETURN_TO_CARRIER_DEPOT": bool, "TYPE_OF_RETURN": str, 
                             "AMOUNT_OF_BOXES_TO_RETURN": str, 
                             "TRACKING_NUMBER": str, 
                             "RETURN_TRACKING_NUMBER": str, "PRINT_RETURN_DOCUMENT": bool, 
                             "CONTACTS": str, "CARRIER_ID": str}
            return columns_names, columns_types

        def apply_team_specific_changes_for_orders_tables(self, ordersDataFrame: pd.DataFrame) -> pd.DataFrame:
            temperatures = {"Ambiente": "Ambient", "Ambiente Controlado": "Controlled Ambient", "Refrigerado": "Refrigerated"}
            ordersDataFrame["TEMPERATURE"] = ordersDataFrame["TEMPERATURE"].replace(temperatures)
            
            ordersDataFrame["PRINT_RETURN_DOCUMENT"] = False
            
            ordersDataFrame["HAS_RETURN"] = (ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & (ordersDataFrame["TEMPERATURE"] != "Ambient")
            ordersDataFrame.loc[ordersDataFrame["HAS_RETURN"], "TYPE_OF_RETURN"] = "CREDO"
            ordersDataFrame["RETURN_TO_CARRIER_DEPOT"] = ordersDataFrame["HAS_RETURN"]

            return ordersDataFrame

        def sendEmailWithOrdersToTeam(self, folder_path_with_orders_files: str, date: str):
            self.father.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "")

        def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
            try:
                df = pd.read_excel(path_from_get_data, sheet_name=orders_sheet, dtype=columns_types, header=0, skiprows=7)
            except Exception as e:
                Log().add_log(f"Error reading orders excel: {e}")
                df = DataRecolector(Teams()).getEmptyOrdersAndContactsData()
            return df
        
        def readSitesExcel(self, path_from_get_data: str, sites_sheet: str, columns_types: dict) -> pd.DataFrame:
            try:
                df = pd.read_excel(path_from_get_data, sheet_name=sites_sheet, dtype=columns_types, header=0)
            except Exception as e:
                Log().add_log(f"Error reading sites excel: {e}")
                df = DataRecolector(Teams()).getEmptyOrdersAndContactsData()
            return df
    
        def build_driver(self):
            self.father.__build_driver__(self.carrierWebpage)

        def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
            return self.father.__check_if_user_and_password_are_correct__(self.carrierWebpage, username, password)
        
        def quit_driver(self):
            self.father.__quit_driver__(self.carrierWebpage)

        def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                        ship_date: str, ship_time_from: str, ship_time_to: str,
                                        delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                        type_of_material: str, temperature: str,
                                        contacts: str, amount_of_boxes: int) -> str:
            return self.father.__complete_shipping_order_form__(self.carrierWebpage, carrier_id, reference,
                                        ship_date, ship_time_from, ship_time_to,
                                        delivery_date, delivery_time_from, delivery_time_to,
                                        type_of_material, temperature,
                                        contacts, amount_of_boxes)
        
        def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
            return self.father.__complete_shipping_order_return_form__(self.carrierWebpage, carrier_id, reference_return,
                                                delivery_date, return_time_from,
                                                return_time_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)
        
        def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
            self.father.__printWayBillDocument__(self.carrierWebpage, tracking_number, amount_of_copies)

        def printLabelDocument(self, tracking_number: str):
            self.father.__printLabelDocument__(self.carrierWebpage, tracking_number)

        def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
            self.father.__printReturnWayBillDocument__(self.carrierWebpage, return_tracking_number, amount_of_copies)
        
    class TestTeam:
        def __init__(self, folder_path_to_download: str, father = None):
            self.carrierWebpage = CarriersWebpage("Transportes Ambientales", folder_path_to_download)
            self.father = father

        def getTeamName(self) -> str:
            return "Test"
        
        def getTeamEmail(self) -> str:
            return "inaki.costa@thermofisher.com"

        def get_column_rename_type_config_for_contacts_table(self) -> (dict, dict):
            columns_names = {}
            columns_types = {"STUDY": str, "SITE#": str, "CARRIER_ID": str, "DELIVERY_TIME_FROM": dt.datetime, "DELIVERY_TIME_TO": dt.datetime}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_contacts_table(self, df: pd.DataFrame) -> pd.DataFrame:
            df["ROLE_OF_CONTACTS"] = "Can receive medicines"
            df["MEDICAL_CENTER_EMAILS"] = "inaki.costa@thermofisher.com"
            df["CUSTOMER_EMAIL"] = ""
            df["CRA_EMAILS"] = ""
            return df
        
        def get_data_path(self) -> (str, str, str):
            path_from_get_data = os.path.expanduser("~\\OneDrive - Thermo Fisher Scientific\Desktop\Tracking_Number_Generator\Automatizacion_Ordenes.xlsx")
            orders_sheet = "Test_5_ordenes" #"Test"
            info_sites_sheet = "SiteInfo"
            return path_from_get_data, orders_sheet, info_sites_sheet
        
        def get_column_rename_type_config_for_orders_tables(self) -> (dict, dict):
            columns_names = {}
            columns_types = {"SITE#": str, "RETURN_TO_CARRIER_DEPOT": bool, "SHIP_TIME_TO": str}
            return columns_names, columns_types
        
        def apply_team_specific_changes_for_orders_tables(self, ordersDataFrame: pd.DataFrame) -> pd.DataFrame:
            ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] = ordersDataFrame["AMOUNT_OF_BOXES_TO_SEND"]

            ordersDataFrame["PRINT_RETURN_DOCUMENT"] = True

            ordersDataFrame["CUSTOMER"]  = "Test Customer"
            ordersDataFrame["HAS_RETURN"] = (ordersDataFrame["AMOUNT_OF_BOXES_TO_RETURN"] > 0) & (ordersDataFrame["TEMPERATURE"] != "Ambient")
            ordersDataFrame["TYPE_OF_RETURN"] = "CREDO"
            ordersDataFrame["RETURN_TO_CARRIER_DEPOT"] = True

            return ordersDataFrame
        
        def sendEmailWithOrdersToTeam(self, folder_path_with_orders_files: str, date: str):
            self.father.__sendEmailWithOrdersToTeam__(folder_path_with_orders_files, date, self.getTeamEmail(), "")

        def readOrdersExcel(self, path_from_get_data: str, orders_sheet: str, columns_types: dict) -> pd.DataFrame:
            try:
                df = pd.read_excel(path_from_get_data, sheet_name=orders_sheet, dtype=columns_types, header=0)
            except Exception as e:
                Log().add_log(f"Error reading orders excel: {e}")
                df = DataRecolector(Teams()).getEmptyOrdersAndContactsData()
            return df
        
        def readSitesExcel(self, path_from_get_data: str, sites_sheet: str, columns_types: dict) -> pd.DataFrame:
            try:
                df = pd.read_excel(path_from_get_data, sheet_name=sites_sheet, dtype=columns_types, header=0)
            except Exception as e:
                Log().add_log(f"Error reading sites excel: {e}")
                df = DataRecolector(Teams()).getEmptyOrdersAndContactsData()
            return df

        def build_driver(self):
            self.father.__build_driver__(self.carrierWebpage)

        def check_if_user_and_password_are_correct(self, username: str, password: str) -> bool:
            return self.father.__check_if_user_and_password_are_correct__(self.carrierWebpage, username, password)
        
        def quit_driver(self):
            self.father.__quit_driver__(self.carrierWebpage)

        def complete_shipping_order_form(self, carrier_id: str, reference: str,
                                        ship_date: str, ship_time_from: str, ship_time_to: str,
                                        delivery_date: str, delivery_time_from: str, delivery_time_to: str,
                                        type_of_material: str, temperature: str,
                                        contacts: str, amount_of_boxes: int) -> str:
            return self.father.__complete_shipping_order_form__(self.carrierWebpage, carrier_id, reference,
                                        ship_date, ship_time_from, ship_time_to,
                                        delivery_date, delivery_time_from, delivery_time_to,
                                        type_of_material, temperature,
                                        contacts, amount_of_boxes)
        
        def complete_shipping_order_return_form(self, carrier_id: str, reference_return: str,
                                                delivery_date: str, return_time_from: str,
                                                return_time_to: str, type_of_return: str,
                                                contacts: str, amount_of_boxes_to_return: int,
                                                return_to_TA: bool, tracking_number: str) -> str:
            return self.father.__complete_shipping_order_return_form__(self.carrierWebpage, carrier_id, reference_return,
                                                delivery_date, return_time_from,
                                                return_time_to, type_of_return,
                                                contacts, amount_of_boxes_to_return,
                                                return_to_TA, tracking_number)
        
        def printWayBillDocument(self, tracking_number: str, amount_of_copies: int):
            self.father.__printWayBillDocument__(self.carrierWebpage, tracking_number, amount_of_copies)

        def printLabelDocument(self, tracking_number: str):
            self.father.__printLabelDocument__(self.carrierWebpage, tracking_number)

        def printReturnWayBillDocument(self, return_tracking_number: str, amount_of_copies: int):
            self.father.__printReturnWayBillDocument__(self.carrierWebpage, return_tracking_number, amount_of_copies)
        
# Controler
class Controller:
    def __init__(self):
        self.mainUserForm = MyUserForm(self)
        self.logInUserForm = None
        self.selected_team = Teams()
        self.ordersAndContactsDataframe = DataRecolector(Teams()).getEmptyOrdersAndContactsData()
        Log().add_log("Application started")

    def showMainUserForm(self):
        self.mainUserForm.show_userform()

    def destroyMainUserForm(self):
        self.mainUserForm.hide_userform()

    def showLogInUserForm(self):
        self.logInUserForm.show_userform()

    def destroyLogInUserForm(self):
        self.logInUserForm.hide_userform()

    def on_log_btn_click(self):
        LogConsole(self).show_userform()

    def on_loadOrders_btn_click(self):
        selected_team = self.mainUserForm.getSelectedTeamName()
        selected_date = self.mainUserForm.getSelectedDate()

        Log().add_separator()
        Log().add_log(f"Start loading orders")
        time0 = time.time()
        self.ordersAndContactsDataframe = self.__loadOrders__(selected_team, selected_date)
        time1 = time.time()
        
        total_time_with_2_decimals = round(time1 - time0, 2)
        Log().add_log(f"End loading orders. Time: {total_time_with_2_decimals}")


        self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)

    def on_clearOrders_btn_click(self):
        self.ordersAndContactsDataframe = DataRecolector(Teams()).getEmptyOrdersAndContactsData()
        self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)
        Log().add_log("Orders table cleaned")

    def on_processOrders_btn_click(self):
        self.logInUserForm = LogInUserForm(self.selected_team, self)
        self.showLogInUserForm()

    def on_login_successful(self):
        self.destroyLogInUserForm()
        
        self.ordersAndContactsDataframe = self.__processOrdersAndCalculateTime__(
                                        self.selected_team, self.selected_date,
                                        self.folder_path_to_download, self.ordersAndContactsDataframe)

        self.mainUserForm.update_ordersAndContactsDataframe_and_widgets(self.ordersAndContactsDataframe)

    def update_tag_color_of_a_treeview_line(self, index: int, tracking_number: str, return_tracking_number: str):
        self.mainUserForm.update_tag_color_of_a_treeview_line(self.ordersAndContactsDataframe, index, tracking_number, return_tracking_number)
        self.mainUserForm.increase_amount_of_orders_processed()

    def __loadOrders__(self, selected_team_name: str, selected_date: str) -> pd.DataFrame:
        self.folder_path_to_download = FilesInSystemController().getFolderPathToDownload(selected_team_name, selected_date.replace("/", "_"))
        self.selected_team = Teams(selected_team_name, self.folder_path_to_download)

        self.selected_date = dt.datetime.strptime(selected_date, '%Y-%m-%d')
        self.ordersAndContactsDataframe = DataRecolector(self.selected_team).recolectOrdersAndContactsData(self.selected_date)

        return self.ordersAndContactsDataframe

    def __processOrders__(self, selected_team, selected_date, folder_path_to_download: str, ordersAndContactsDataframe: pd.DataFrame) -> pd.DataFrame:
        FilesInSystemController().create_folder(folder_path_to_download)

        ordersAndContactsDataframe = OrderProcessor(folder_path_to_download, selected_team, self).processOrdersAndContactsTable(ordersAndContactsDataframe)
        
        selected_date_str = selected_date.strftime("%Y-%m-%d")
        
        selected_team.sendEmailWithOrdersToTeam(folder_path_to_download, selected_date_str)

        return ordersAndContactsDataframe

    def __processOrdersAndCalculateTime__(self, selected_team, selected_date, folder_path_to_download: str, ordersAndContactsDataframe: pd.DataFrame) -> pd.DataFrame:
        amountOfOrdersToProcess = len(ordersAndContactsDataframe[(ordersAndContactsDataframe["TRACKING_NUMBER"] == "") & (ordersAndContactsDataframe["HAS_AN_ERROR"] == "No error")])
        
        Log().add_log(f"Start processing orders")
        time0 = time.time()
        self.ordersAndContactsDataframe = self.__processOrders__(
                                        selected_team, selected_date,
                                        folder_path_to_download, ordersAndContactsDataframe)
        time1 = time.time()
        Log().add_log(f"End processing orders")
        
        total_time = time1 - time0
        total_time_with_2_decimals = round(total_time, 2)
        Log().add_log(f"Total processing time: {total_time_with_2_decimals} s for {amountOfOrdersToProcess} orders")

        if amountOfOrdersToProcess == 0:
            avarage_time_with_2_decimals = 0.00
        else:
            avarage_time_with_2_decimals = round(total_time / amountOfOrdersToProcess, 2)

        Log().add_log(f"Average processing time: {avarage_time_with_2_decimals} s")

        return ordersAndContactsDataframe

# Initializer
if __name__ == "__main__":
    Controller().showMainUserForm()