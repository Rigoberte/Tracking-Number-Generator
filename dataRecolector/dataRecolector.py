import pandas as pd
import datetime as dt
import numpy as np
import queue

from teams.team import Team
from logClass.log import Log

class DataRecolector:
    def __init__(self, aTeam: Team, queue: queue.Queue = queue.Queue(), log: Log = Log()):
        self.selectedTeam = aTeam
        self.queue = queue
        self.log = log
        
        self.memo_of_transit_per_ship_and_delivery_dates = {}
        self.memo_of_return_date_per_delivery_date_and_transit = {}

        self.columns_df = ['SYSTEM_NUMBER', 'IVRS_NUMBER', 'CUSTOMER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO', 
                'TYPE_OF_MATERIAL', 'TEMPERATURE', 'AMOUNT_OF_BOXES_TO_SEND',
                'HAS_RETURN', 'RETURN_TO_CARRIER_DEPOT', 'TYPE_OF_RETURN', 'RETURN_DATE', 'RETURN_DELIVERY_HOUR_FROM', 'RETURN_DELIVERY_HOUR_TO', 'AMOUNT_OF_BOXES_TO_RETURN',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT', 'CONTACTS', 'TYPE_OF_MATERIAL_CAN_RECEIVE', 
                "MEDICAL_CENTER_EMAILS", "CUSTOMER_EMAIL", "CRA_EMAILS", "TEAM_EMAILS",
                'CARRIER_ID', 'HAS_AN_ERROR']
        
        self.columns_for_orders = ['SYSTEM_NUMBER', 'IVRS_NUMBER', 'CUSTOMER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO', 'TYPE_OF_MATERIAL', 
                'TEMPERATURE', 'AMOUNT_OF_BOXES_TO_SEND', 'HAS_RETURN', 
                'RETURN_TO_CARRIER_DEPOT', 'TYPE_OF_RETURN',
                'RETURN_DATE', 'RETURN_DELIVERY_HOUR_FROM', 'RETURN_DELIVERY_HOUR_TO', 'AMOUNT_OF_BOXES_TO_RETURN',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT']
        
        self.columns_for_contacts = ["STUDY", "SITE#", "CARRIER_ID", 
                "DELIVERY_TIME_FROM", "DELIVERY_TIME_TO", 
                "CONTACTS", "TYPE_OF_MATERIAL_CAN_RECEIVE", 
                "MEDICAL_CENTER_EMAILS", "CUSTOMER_EMAIL", "CRA_EMAILS", "TEAM_EMAILS"]
        
    def recolect_orders_and_contacts_dataFrame(self, shipdate: dt.datetime) -> pd.DataFrame:
        """
        Process all orders in the table

        Args:
            team (str): team to process
            shipdate (dt.datetime): date to process

        Returns:
            DataFrame: orders table with standarisized data
        """
        try:
            ordersDataframe = self.__load_shipping_order_table__(shipdate, self.selectedTeam)
            contactsDataframe = self.__load_contacts_table__(self.selectedTeam)
            
            ordersAndContactsDataframe = self.__merge_orders_with_contacts_tables_if_orders_do_not_a_coordinated_delivery_time__(ordersDataframe, contactsDataframe)

            ordersAndContactsDataframe["HAS_AN_ERROR"] = ordersAndContactsDataframe.apply(self.__checkErrorsOnEachOrder__, axis=1)
            ordersAndContactsDataframe.fillna("", inplace=True)

            self.queue.put(ordersAndContactsDataframe[self.columns_df])

            self.memo_of_transit_per_ship_and_delivery_dates = {}
            self.memo_of_return_date_per_delivery_date_and_transit = {}

            return ordersAndContactsDataframe[self.columns_df]
        except Exception as e:
            self.log.add_error_log(f"Error recolecting orders and contacts data: {e}")
            return self.get_empty_ordersAndContactsData()

    def get_empty_ordersAndContactsData(self) -> pd.DataFrame:
        """
        Returns an empty orders table

        Returns:
            DataFrame: empty orders table
        """
        return pd.DataFrame(columns=self.columns_df)

    def get_empty_orders_dataFrame(self) -> pd.DataFrame:
        """
        Returns an empty orders table

        Returns:
            DataFrame: empty orders table
        """
        return pd.DataFrame(columns=self.columns_for_orders)
    
    def get_empty_contacts_dataFrame(self) -> pd.DataFrame:
        """
        Returns an empty contacts table

        Returns:
            DataFrame: empty contacts table
        """
        return pd.DataFrame(columns=self.columns_for_contacts)

    # Private methods
    def __load_shipping_order_table__(self, shipDate: dt.datetime, team: Team) -> pd.DataFrame:
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

        ordersDataFrame = self.__create_undefined_columns__(ordersDataFrame, self.columns_for_orders)

        notWorkingDaysList = self.__load_not_working_days__(team)
        
        ordersDataFrame["TRANSIT"] = ordersDataFrame.apply(lambda x: self.__calculate_transit_days_for_returns__(x["SHIP_DATE"], x["DELIVERY_DATE"], notWorkingDaysList), axis=1)
        ordersDataFrame["RETURN_DATE"] = ordersDataFrame.apply(lambda x: self.__calculate_return_date__(x["HAS_RETURN"], x["DELIVERY_DATE"], x["TRANSIT"], notWorkingDaysList), axis=1)
        
        ordersDataFrame["SHIP_DATE"] = ordersDataFrame["SHIP_DATE"].dt.strftime('%d/%m/%Y')
        ordersDataFrame["DELIVERY_DATE"] = ordersDataFrame["DELIVERY_DATE"].dt.strftime('%d/%m/%Y')
        ordersDataFrame["RETURN_DATE"] = self.__correctDateColumns__(ordersDataFrame, "RETURN_DATE")
        ordersDataFrame["RETURN_DATE"] = ordersDataFrame["RETURN_DATE"].dt.strftime('%d/%m/%Y')

        ordersDataFrame["RETURN_DELIVERY_HOUR_FROM"] = ordersDataFrame.apply(lambda x: "09:00" if x["HAS_RETURN"] else "", axis=1)
        ordersDataFrame["RETURN_DELIVERY_HOUR_TO"] = ordersDataFrame.apply(lambda x: "16:00" if x["HAS_RETURN"] else "", axis=1)

        return ordersDataFrame[self.columns_for_orders]

    def __load_shipping_order_table_with_normalized_columns__(self, team: Team) -> pd.DataFrame:
        """
        Loads orders table according to team

        Args:
            team (str): team to process
        """
        path_from_get_data, orders_sheet = team.get_data_path(["team_excel_path", "team_orders_sheet"])
        
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

        return ordersDataFrame

    def __load_contacts_table__(self, team: Team) -> pd.DataFrame:
        """
        Loads contacts table according to team

        Args:
            team (Team): team to process

        Returns:
            DataFrame: contacts table
        """
        contactsDataFrame = self.__load_contacts_table_with_normalized_columns__(team)
        
        contactsDataFrame = self.__correct_regular_columns_for_contacts_table__(contactsDataFrame, team)

        contactsDataFrame = team.apply_team_specific_changes_for_contacts_table(contactsDataFrame)

        contactsDataFrame = self.__transform_material_receiving_options__(contactsDataFrame)

        contactsDataFrame = contactsDataFrame.drop_duplicates(subset=["STUDY", "SITE#", "TYPE_OF_MATERIAL_CAN_RECEIVE"], keep='last')
        
        contactsDataFrame = self.__create_undefined_columns__(contactsDataFrame, self.columns_for_contacts)

        return contactsDataFrame[self.columns_for_contacts]
    
    def __load_contacts_table_with_normalized_columns__(self, team: Team) -> pd.DataFrame:
        """
        Loads contacts table according to team

        Args:
            team (Team): team to process
        """
        path_from_get_data, contacts_sheet = team.get_data_path(["team_excel_path", "team_contacts_sheet"])
        columns_names, columns_types = team.get_column_rename_type_config_for_contacts_table()

        contactsDataFrame = team.readContactsExcel(path_from_get_data, contacts_sheet, columns_types)
        contactsDataFrame.rename(columns=columns_names, inplace=True)

        return contactsDataFrame

    def __load_not_working_days__(self, team: Team) -> list:
        """
        Loads not working days according to team

        Args:
            team (Team): team to process
        """
        
        path_from_get_data, not_working_days_sheet = team.get_data_path(["team_excel_path", "team_not_working_days_sheet"])
        
        columns_names, columns_types = team.get_column_rename_type_config_for_not_working_days_table()
        notWorkingDaysDataFrame = team.readNotWorkingDaysExcel(path_from_get_data, not_working_days_sheet, columns_types)
        notWorkingDaysDataFrame.rename(columns=columns_names, inplace=True)

        return notWorkingDaysDataFrame["DATE"].tolist()

    def __correct_regular_columns_for_contacts_table__(self, contactsDataFrame: pd.DataFrame, team: Team) -> pd.Series:
        """
        Corrects regular columns for contacts table

        Args:
            contactsDataFrame (DataFrame): contacts table
        """
        contactsDataFrame["DELIVERY_TIME_FROM"] = self.__correctTimeColumns__(contactsDataFrame, "DELIVERY_TIME_FROM")
        contactsDataFrame["DELIVERY_TIME_TO"] = self.__correctTimeColumns__(contactsDataFrame, "DELIVERY_TIME_TO")

        contactsDataFrame["TEAM_EMAILS"] = team.getTeamEmail()

        return contactsDataFrame

    def __checkErrorsOnEachOrder__(self, row: pd.Series) -> str:
        """
        Checks errors on each order

        Args:
            row (Series): order row
        """
        def assertIfIsNotNull(cell: str) -> bool:
            if isinstance(cell, str):
                if cell == "" or cell.lower() == "nan" or pd.isna(cell):
                    return False
            elif pd.isna(cell):
                return False
            
            try:
                return not np.isnan(float(cell))
            except (ValueError, TypeError):
                return True

        def assertIfSystemNumberIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['SYSTEM_NUMBER'])

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
            try:
                today = dt.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

                shipdate = str(row['SHIP_DATE']).split("/")
                shipdate = dt.datetime(year=int(shipdate[2]), month=int(shipdate[1]), day=int(shipdate[0]))

                deliverydate = str(row['DELIVERY_DATE']).split("/")
                deliverydate = dt.datetime(year=int(deliverydate[2]), month=int(deliverydate[1]), day=int(deliverydate[0]))

                return (type(today) == dt.datetime and type(shipdate) == dt.datetime and type(deliverydate) == dt.datetime and 
                        today <= shipdate and shipdate <= deliverydate)
            
            except (ValueError, TypeError):
                return False

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
            return row['TYPE_OF_MATERIAL'] in ["Medicine", "Ancillary Type 1", "Ancillary Type 2", "Equipment"]

        def assertIfAmountOfBoxesAreValid(row: pd.Series) -> bool:
            return type(row['AMOUNT_OF_BOXES_TO_SEND']) == int and row['AMOUNT_OF_BOXES_TO_SEND'] > 0
        
        def assertIfCarrier_IDIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['CARRIER_ID'])
        
        def assertIfTemperatureIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['TEMPERATURE'])

        def assertIfTemperatureIsValid(row: pd.Series) -> bool:
            return row['TEMPERATURE'] in ["Ambient", "Controlled Ambient", "Refrigerated", "Frozen", "Refrigerated with Dry Ice", "Frozen with Liquid Nitrogen"]
        
        def assertIfNumberOfBoxesToReturnIsValid(row: pd.Series) -> bool:
            return type(row['AMOUNT_OF_BOXES_TO_RETURN']) == int and (row['AMOUNT_OF_BOXES_TO_RETURN'] >= 0) and (row['AMOUNT_OF_BOXES_TO_RETURN'] <= row['AMOUNT_OF_BOXES_TO_SEND'])
        
        def assertIfTypeOfReturnIsValid(row: pd.Series) -> bool:
            return row['TYPE_OF_RETURN'] in ["CREDO", "DATALOGGER", "CREDO AND DATALOGGER", "NA"] and not (row['HAS_RETURN'] and row['TYPE_OF_RETURN'] == "NA")
        
        def assertIfHasReturnIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['HAS_RETURN'])
        
        def assertIfReturnToCarrierDepotIsNotEmpty(row: pd.Series) -> bool:
            return assertIfIsNotNull(row['RETURN_TO_CARRIER_DEPOT'])
        
        errors = ""

        if not assertIfSystemNumberIsNotEmpty(row):
            errors += "No system number; "

        if not assetIfCustomerIsNotEmpty(row):
            errors += "No customer; "

        if not assertIfStudyIsNotEmpty(row):
            errors += "No study; "

        if not assertIfSiteIsNotEmpty(row):
            errors += "No site; "

        shipDateIsNotEmpty = assertIfShipDateIsNotEmpty(row)
        if not shipDateIsNotEmpty:
            errors += "No ship date; "

        if not assertIfShipTimeFromIsNotEmpty(row):
            errors += "No ship time from; "

        if not assertIfShipTimeToIsNotEmpty(row):
            errors += "No ship time to; "

        if not assertIfAreValidShipTimes(row):
            errors += "Invalid ship times; "

        deliveryDateIsNotEmpty = assertIfDeliveryDateIsNotEmpty(row)
        if not deliveryDateIsNotEmpty:
            errors += "No delivery date; "
        
        deliveryTimeFromIsNotEmpty = assertIfDeliveryTimeFromIsNotEmpty(row)
        if not deliveryTimeFromIsNotEmpty:
            errors += "No delivery time from; "

        deliveryTimeToIsNotEmpty = assertIfDeliveryTimeToIsNotEmpty(row)
        if not deliveryTimeToIsNotEmpty:
            errors += "No delivery time to; "
        
        if deliveryTimeFromIsNotEmpty and deliveryTimeToIsNotEmpty:
            if not assertIfAreValidDeliveryTimes(row):
                errors += "Invalid delivery times; "

        if shipDateIsNotEmpty and deliveryDateIsNotEmpty:
            if not assertIfAreValidDates(row):
                errors += "Invalid dates; "

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

        if not assertIfTypeOfReturnIsValid(row):
            errors += "Invalid type of return; "

        if not assertIfHasReturnIsNotEmpty(row):
            errors += "No has return; "

        if not assertIfReturnToCarrierDepotIsNotEmpty(row):
            errors += "No return to carrier depot; "

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
        return dataFrame[column]
    
    def __calculate_transit_days_for_returns__(self, shipDate: dt.datetime, deliveryDate: dt.datetime, notWorkingDaysList: list) -> int:
        """
        Calculates the transit days

        Args:
            shipDate (dt.datetime): ship date
            deliveryDate (dt.datetime): delivery date
        """
        if deliveryDate is None or shipDate is None or pd.isna(deliveryDate) or pd.isna(shipDate):
            return 1
        
        if (shipDate, deliveryDate) in self.memo_of_transit_per_ship_and_delivery_dates:
            return self.memo_of_transit_per_ship_and_delivery_dates[(shipDate, deliveryDate)]
        
        amount_of_days_until_next_working_day = self.__amount_of_days_until_next_working_day__(shipDate + dt.timedelta(days=1) , notWorkingDaysList)

        transitWithOutWeekend = (deliveryDate - shipDate).days - amount_of_days_until_next_working_day

        self.memo_of_transit_per_ship_and_delivery_dates[(shipDate, deliveryDate)] = max(transitWithOutWeekend, 1)

        return self.memo_of_transit_per_ship_and_delivery_dates[(shipDate, deliveryDate)]

    def __is_a_working_day__(self, date: dt.datetime, notWorkingDaysList: list) -> bool:
        """
        Checks if a date is a working day

        Args:
            date (dt.datetime): date to process
        """
        return date.weekday() <= 4 and date not in notWorkingDaysList

    def __amount_of_days_until_next_working_day__(self, date: dt.datetime, notWorkingDaysList: list) -> int:
        """
        Returns the amount of days until the next working day

        Args:
            date (dt.datetime): date to process
        """
        days = 0
        while not self.__is_a_working_day__(date + dt.timedelta(days=days), notWorkingDaysList):
            days += 1

        return days

    def __calculate_return_date__(self, hasReturn: bool, deliveryDate: dt.datetime, transitDays: int, notWorkingDaysList: list) -> dt.datetime:
        """
        Calculates the return date

        Args:
            deliveryDate (dt.datetime): delivery date
            transitDays (int): transit days
        """
        if not hasReturn:
            return None

        if pd.isna(deliveryDate) or np.isnan(transitDays):
            return None
        
        if (deliveryDate, transitDays) in self.memo_of_return_date_per_delivery_date_and_transit:
            return self.memo_of_return_date_per_delivery_date_and_transit[(deliveryDate, transitDays)]
        
        nextWorkingDay = self.__nextWorkingDay__(deliveryDate, notWorkingDaysList)

        #next WorkingDay With Transit Days
        returnDate = self.__nextWorkingDay__(nextWorkingDay + dt.timedelta(days=transitDays), notWorkingDaysList)

        self.memo_of_return_date_per_delivery_date_and_transit[(deliveryDate, transitDays)] = returnDate

        return self.memo_of_return_date_per_delivery_date_and_transit[(deliveryDate, transitDays)]

    def __nextWorkingDay__(self, date: dt.datetime, notWorkingDaysList: list) -> dt.datetime:
        """
        Returns the next working day

        Args:
            date (dt.datetime): date to process
        """
        
        date += dt.timedelta(days= self.__amount_of_days_until_next_working_day__(date, notWorkingDaysList) )

        return date

    def __create_undefined_columns__(self, dataFrame: pd.DataFrame, columns: list) -> pd.DataFrame:
        """
        Creates undefined columns

        Args:
            dataFrame (DataFrame): Pandas DataFrame
            columns (list): columns to create
        """
        for column in columns:
            if column not in dataFrame.columns:
                dataFrame[column] = ""
        
        return dataFrame
    
    def __transform_material_receiving_options__(self, contactsDataFrame: pd.DataFrame) -> pd.DataFrame:
        # Columns that should not be transformed
        other_columns = contactsDataFrame.columns.difference(['CAN_RECEIVE_MEDICINES', 'CAN_RECEIVE_ANCILLARIES_TYPE1', 'CAN_RECEIVE_ANCILLARIES_TYPE2', 'CAN_RECEIVE_EQUIPMENTS'])
        
        # Turn the columns 'CAN_RECEIVE_MEDICINES' and 'CAN_RECEIVE_ANCILLARIES' into rows
        df_melted = contactsDataFrame.melt(id_vars=other_columns, 
                            value_vars=['CAN_RECEIVE_MEDICINES', 'CAN_RECEIVE_ANCILLARIES_TYPE1', 'CAN_RECEIVE_ANCILLARIES_TYPE2', 'CAN_RECEIVE_EQUIPMENTS'],
                            var_name='Option',
                            value_name='Chosen')
        
        # Filter only the rows where 'Chosen' is True
        df_filtered = df_melted[df_melted['Chosen']]

        # Drop the column 'Chosen' since we don't need it anymore
        df_filtered = df_filtered.drop(columns='Chosen')

        # Create a new column 'TYPE_OF_MATERIAL_CAN_RECEIVE' with the value of 'Option' without the prefix 'CAN_RECEIVE_'
        df_filtered['TYPE_OF_MATERIAL_CAN_RECEIVE'] = df_filtered['Option'].str.replace('CAN_RECEIVE_', '')

        # Drop the column 'Option' since we don't need it anymore
        df_filtered = df_filtered.drop(columns='Option')

        #Change MEDICINES to Medicine and ANCILLARIES to Ancillary
        types_of_materials = {"MEDICINES": "Medicine", "ANCILLARIES_TYPE1": "Ancillary Type 1", "ANCILLARIES_TYPE2": "Ancillary Type 2", "EQUIPMENTS": "Equipment"}

        df_filtered["TYPE_OF_MATERIAL_CAN_RECEIVE"] = df_filtered["TYPE_OF_MATERIAL_CAN_RECEIVE"].replace(types_of_materials)

        return df_filtered
    
    def __merge_orders_with_contacts_tables_if_orders_do_not_a_coordinated_delivery_time__(self, ordersDataframe: pd.DataFrame, contactsDataframe: pd.DataFrame) -> pd.DataFrame:
        """
        Merges orders with contacts tables if orders do not have a coordinated delivery time

        Args:
            ordersDataframe (DataFrame): orders table
            contactsDataframe (DataFrame): contacts table
        """
        ordersAndContactsDataframe = pd.merge(ordersDataframe, contactsDataframe, 
                                    left_on=["STUDY", "SITE#", "TYPE_OF_MATERIAL"], 
                                    right_on=["STUDY", "SITE#", "TYPE_OF_MATERIAL_CAN_RECEIVE"], 
                                    how="left")

        ordersAndContactsDataframe["DELIVERY_TIME_FROM"] = ordersAndContactsDataframe.apply(
            lambda row: row["DELIVERY_TIME_FROM_y"] if row["DELIVERY_TIME_FROM_x"] in ['', '00:00', None] else row["DELIVERY_TIME_FROM_x"],
            axis=1
        )


        ordersAndContactsDataframe["DELIVERY_TIME_TO"] = ordersAndContactsDataframe.apply(
            lambda row: row["DELIVERY_TIME_TO_y"] if row["DELIVERY_TIME_TO_x"] in ['', '00:00', None] else row["DELIVERY_TIME_TO_x"],
            axis=1
        )

        ordersAndContactsDataframe = ordersAndContactsDataframe.drop(columns=["DELIVERY_TIME_FROM_x", "DELIVERY_TIME_TO_x", "DELIVERY_TIME_FROM_y", "DELIVERY_TIME_TO_y"])

        return ordersAndContactsDataframe