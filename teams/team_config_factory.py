"""
Team Configuration Factory.

Creates team configurations based on team name.
"""
from typing import List

from .team_config import TeamConfig
from .eli_lilly_config import EliLillyArgentinaConfig
from .gpm_config import GPMArgentinaConfig
from .test_team_config import TestTeamConfig
from .no_team_config import NoTeamConfig


class TeamConfigFactory:
    """Factory for creating team configurations."""
    
    _TEAMS = {
        "Eli Lilly Argentina": EliLillyArgentinaConfig,
        "GPM Argentina": GPMArgentinaConfig,
        "Team_for_testings": TestTeamConfig,
    }
    
    @classmethod
    def create(cls, team_name: str) -> TeamConfig:
        """
        Create a team configuration by name.
        
        Args:
            team_name: The name of the team.
            
        Returns:
            TeamConfig instance for the specified team.
        """
        config_class = cls._TEAMS.get(team_name, NoTeamConfig)
        return config_class()
    
    @classmethod
    def get_team_names(cls) -> List[str]:
        """Get list of available team names."""
        return list(cls._TEAMS.keys())
