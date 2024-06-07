class TeamFactory:
    def create_team(self, team_name: str, folder_path_to_download: str):
        """
        Class constructor for teams
        """
        match team_name:
            case 'Eli Lilly Argentina':
                from .EliLilly_Argentina import EliLillyArgentinaTeam
                return EliLillyArgentinaTeam(folder_path_to_download)
            case 'GPM Argentina':
                from .GPM_Argentina import GPMArgentinaTeam
                return GPMArgentinaTeam(folder_path_to_download)
            case 'Test_5_ordenes':
                from .TestTeam import TestTeam
                return TestTeam(folder_path_to_download)

            # ---------------------
            case _:
                from .NoSelectedTeam import NoSelectedTeam
                return NoSelectedTeam("")
        
    def get_team_names(self):
        return ['Eli Lilly Argentina', 'GPM Argentina', 'Test_5_ordenes']