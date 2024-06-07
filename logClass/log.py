import datetime as dt

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
           