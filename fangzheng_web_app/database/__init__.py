"""Database infrastructure shared by the automation migration tools."""

from .automation import automation_cursor, close_automation_pool
from .config import AutomationDatabaseConfig, IdentityDatabaseConfig
from .identity import close_identity_pool, identity_cursor

__all__ = [
    "AutomationDatabaseConfig",
    "IdentityDatabaseConfig",
    "automation_cursor",
    "close_automation_pool",
    "close_identity_pool",
    "identity_cursor",
]
