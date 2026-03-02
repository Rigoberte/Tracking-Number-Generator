"""
Teams Package.

Provides team configurations for different clients/teams.
"""
from .team_config import TeamConfig, ColumnMapping
from .team_config_factory import TeamConfigFactory
from .eli_lilly_config import EliLillyArgentinaConfig
from .gpm_config import GPMArgentinaConfig
from .test_team_config import TestTeamConfig
from .no_team_config import NoTeamConfig


__all__ = [
    'TeamConfig',
    'ColumnMapping',
    'TeamConfigFactory',
    'EliLillyArgentinaConfig',
    'GPMArgentinaConfig',
    'TestTeamConfig',
    'NoTeamConfig',
]
