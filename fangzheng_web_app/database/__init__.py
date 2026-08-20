"""Database infrastructure shared by the automation migration tools."""

from .automation import automation_cursor, close_automation_pool
from .config import AutomationDatabaseConfig

__all__ = ["AutomationDatabaseConfig", "automation_cursor", "close_automation_pool"]
