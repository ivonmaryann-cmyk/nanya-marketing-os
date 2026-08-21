"""Database infrastructure shared by the automation migration tools."""

from .automation import automation_cursor, close_automation_pool
from .config import (
    AutomationDatabaseConfig,
    ConfigurationDatabaseConfig,
    IdentityDatabaseConfig,
    TranscodeDatabaseConfig,
)
from .configuration import close_configuration_pool, configuration_cursor
from .identity import close_identity_pool, identity_cursor
from .transcode import close_transcode_pool, transcode_cursor

__all__ = [
    "AutomationDatabaseConfig",
    "ConfigurationDatabaseConfig",
    "IdentityDatabaseConfig",
    "TranscodeDatabaseConfig",
    "automation_cursor",
    "close_configuration_pool",
    "configuration_cursor",
    "close_automation_pool",
    "close_identity_pool",
    "identity_cursor",
    "close_transcode_pool",
    "transcode_cursor",
]
