import json

class DataPathController:
    def __init__(self):
        self.file = "dataPathController/config.json"

    def redefine_a_config_of_a_team(self, team_name: str, new_values: dict) -> None:
        data = self.__get_all_data_from_file__()
        data[team_name] = new_values
        with open(self.file, "w") as f:
            json.dump(data, f)

    def get_config_of_a_team(self, team_name) -> dict:
        if not self.__check_if_a_team_exists__(team_name):
            self.__define_new_team__(team_name)
        
        with open(self.file, "r") as f:
            return json.load(f)[team_name]

    def __define_new_team__(self, team_name) -> None:
        data = self.__get_all_data_from_file__()
        data[team_name] = {
            "team_excel_path": "",
            "team_orders_sheet": "",
            "team_contacts_sheet": "",
            "team_not_working_days_sheet": "",
            "team_send_email_to_medical_centers": False,
            "team_email": ""
        }

        with open(self.file, "w") as f:
            json.dump(data, f)

    def __check_if_a_team_exists__(self, team_name) -> bool:
        data = self.__get_all_data_from_file__()
        return team_name in data
    
    def __get_all_data_from_file__(self) -> dict:
        with open(self.file, "r") as f:
            return json.load(f)