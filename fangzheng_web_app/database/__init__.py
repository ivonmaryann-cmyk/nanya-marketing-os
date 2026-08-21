"""Database infrastructure shared by the automation migration tools."""

from .automation import automation_cursor, close_automation_pool
from .config import (
    AutomationDatabaseConfig,
    ConfigurationDatabaseConfig,
    IdentityDatabaseConfig,
    PlanningDatabaseConfig,
    TranscodeDatabaseConfig,
)
from .configuration import close_configuration_pool, configuration_cursor
from .identity import close_identity_pool, identity_cursor
from .planning import close_planning_pool, planning_cursor
from .transcode import close_transcode_pool, transcode_cursor

__all__ = [
    "AutomationDatabaseConfig",
    "ConfigurationDatabaseConfig",
    "IdentityDatabaseConfig",
    "PlanningDatabaseConfig",
    "TranscodeDatabaseConfig",
    "automation_cursor",
    "close_configuration_pool",
    "configuration_cursor",
    "close_automation_pool",
    "close_identity_pool",
    "identity_cursor",
    "close_planning_pool",
    "planning_cursor",
    "close_transcode_pool",
    "transcode_cursor",
]
