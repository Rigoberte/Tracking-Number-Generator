import pandas as pd
import datetime as dt

from teams.team import Team
from logClass.log import Log

class DataRecolector:
    def __init__(self, team: Team):
        self.selectedTeam = team
        self.columns_df = ['SYSTEM_NUMBER', 'IVRS_NUMBER', 'CUSTOMER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'DELIVERY_TIME_FROM', 'DELIVERY_TIME_TO', 
                'TYPE_OF_MATERIAL', 'TEMPERATURE', 'AMOUNT_OF_BOXES_TO_SEND',
                'HAS_RETURN', 'RETURN_TO_CARRIER_DEPOT', 'TYPE_OF_RETURN', 'AMOUNT_OF_BOXES_TO_RETURN',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT', 'CONTACTS', 'TYPE_OF_MATERIAL_CAN_RECEIVE', 
                "MEDICAL_CENTER_EMAILS", "CUSTOMER_EMAIL", "CRA_EMAILS", "TEAM_EMAILS",
                'CARRIER_ID', 'HAS_AN_ERROR']
        
        self.columns_for_orders = ['SYSTEM_NUMBER', 'IVRS_NUMBER', 'CUSTOMER', 'STUDY', 'SITE#',
                'SHIP_DATE', 'SHIP_TIME_FROM', 'SHIP_TIME_TO', 
                'DELIVERY_DATE', 'TYPE_OF_MATERIAL', 
                'TEMPERATURE', 'AMOUNT_OF_BOXES_TO_SEND', 'HAS_RETURN', 
                'RETURN_TO_CARRIER_DEPOT', 'TYPE_OF_RETURN', 'AMOUNT_OF_BOXES_TO_RETURN',
                'TRACKING_NUMBER', 'RETURN_TRACKING_NUMBER', 'PRINT_RETURN_DOCUMENT']
        
        self.columns_for_contacts = ["STUDY", "SITE#", "CARRIER_ID", 
                "DELIVERY_TIME_FROM", "DELIVERY_TIME_TO", 
                "CONTACTS", "TYPE_OF_MATERIAL_CAN_RECEIVE", 
                "MEDICAL_CENTER_EMAILS", "CUSTOMER_EMAIL", "CRA_EMAILS", "TEAM_EMAILS"]
    
    def recolectOrdersAndContactsData(self, shipdate: dt.datetime) -> pd.DataFrame:
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

            ordersAndContactsDataframe = pd.merge(ordersDataframe, contactsDataframe, 
                                                left_on=["STUDY", "SITE#", "TYPE_OF_MATERIAL"], 
                                                right_on=["STUDY", "SITE#", "TYPE_OF_MATERIAL_CAN_RECEIVE"], 
                                                how="left")

            ordersAndContactsDataframe["HAS_AN_ERROR"] = ordersAndContactsDataframe.apply(self.__checkErrorsOnEachOrder__, axis=1)
            ordersAndContactsDataframe.fillna("", inplace=True)

            return ordersAndContactsDataframe[self.columns_df]
        except Exception as e:
            Log().add_log(f"Error recolecting orders and contacts data: {e}")
            return self.getEmptyOrdersAndContactsData()

    def getEmptyOrdersAndContactsData(self) -> pd.DataFrame:
        """
        Returns an empty orders table

        Returns:
            DataFrame: empty orders table
        """
        return pd.DataFrame(columns=self.columns_df)

    def getEmptyOrdersData(self) -> pd.DataFrame:
        """
        Returns an empty orders table

        Returns:
            DataFrame: empty orders table
        """
        return pd.DataFrame(columns=self.columns_for_orders)
    
    def getEmptyContactsData(self) -> pd.DataFrame:
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
        
        return ordersDataFrame[self.columns_for_orders]

    def __load_shipping_order_table_with_normalized_columns__(self, team: Team) -> pd.DataFrame:
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
        path_from_get_data, _, sites_sheet = team.get_data_path()
        columns_names, columns_types = team.get_column_rename_type_config_for_contacts_table()

        contactsDataFrame = team.readSitesExcel(path_from_get_data, sites_sheet, columns_types)
        contactsDataFrame.rename(columns=columns_names, inplace=True)

        return contactsDataFrame

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
        other_columns = contactsDataFrame.columns.difference(['CAN_RECEIVE_MEDICINES', 'CAN_RECEIVE_ANCILLARIES'])
        
        # Turn the columns 'CAN_RECEIVE_MEDICINES' and 'CAN_RECEIVE_ANCILLARIES' into rows
        df_melted = contactsDataFrame.melt(id_vars=other_columns, 
                            value_vars=['CAN_RECEIVE_MEDICINES', 'CAN_RECEIVE_ANCILLARIES'], 
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
        types_of_materials = {"MEDICINES": "Medicine", "ANCILLARIES": "Ancillary"}

        df_filtered["TYPE_OF_MATERIAL_CAN_RECEIVE"] = df_filtered["TYPE_OF_MATERIAL_CAN_RECEIVE"].replace(types_of_materials)

        return df_filtered