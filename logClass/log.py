import datetime as dt
import pandas as pd

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
        self.logs = pd.DataFrame(columns=["Date and Time", "Text", "Type"])

    def add_error_log(self, text: str):
        self.__add_log__(text, type = "Error")

    def add_warning_log(self, text: str):
        self.__add_log__(text, type = "Warning")

    def add_info_log(self, text: str):
        self.__add_log__(text, type = "Info")

    def add_separator(self):
        self.__add_log__("---------------------------------------------------", type = "Separator")

    def print_logs(self) -> pd.DataFrame:
        return self.print_last_n_logs(len(self.logs))
    
    def print_last_n_logs(self, n: int) -> pd.DataFrame:
        n = min(n, len(self.logs))

        logs_to_print = self.logs.tail(n).copy()
        logs_to_print['Date and Time'] = logs_to_print['Date and Time'].dt.strftime('%Y-%m-%d %H:%M:%S')
        logs_to_print['Date and Time'] = '<' + logs_to_print['Date and Time'] + '>'

        return logs_to_print
    
    def export_to_csv(self, folder_path_to_download: str):
        if not self.logs.empty:
            dataframe_name = "logs_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".csv"
            file_path = folder_path_to_download + "\\" + dataframe_name

            self.logs.to_csv(file_path)
        else:
            self.add_warning_log("Empty DataFrame")

    def clear_logs(self):
        self.logs = pd.DataFrame( self.logs.columns )

    def __add_log__(self, text: str, *, type: str):
        date_and_time = dt.datetime.now()

        log = {"Date and Time": date_and_time, "Text": text, "Type": type}
        log_df = pd.DataFrame(log, index=[0])

        self.logs = pd.concat([self.logs, log_df], ignore_index=True)