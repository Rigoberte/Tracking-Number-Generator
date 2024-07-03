from logClass.log import Log

class TeamFactory:
    def create_team(self, team_name: str, folder_path_to_download: str = "", log: Log = Log()):
        """
        Class constructor for teams
        """
        match team_name:
            case 'Eli Lilly Argentina':
                from .EliLilly_Argentina import EliLillyArgentinaTeam
                return EliLillyArgentinaTeam(folder_path_to_download, log)
            case 'GPM Argentina':
                from .GPM_Argentina import GPMArgentinaTeam
                return GPMArgentinaTeam(folder_path_to_download, log)
            case 'Team_for_testings':
                from .TestTeam import TestTeam
                return TestTeam(folder_path_to_download, log)

            # ---------------------
            case _:
                from .NoSelectedTeam import NoSelectedTeam
                return NoSelectedTeam("", log)
        
    def get_team_names(self):
        return ['Eli Lilly Argentina', 'GPM Argentina', 'Team_for_testings']