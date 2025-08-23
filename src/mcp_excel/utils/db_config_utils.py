"""
Configuration management for database operations.

This module provides backward compatibility with the old configuration system
while delegating to the new centralized configuration management.
"""

import logging
from typing import cast

from mcp_excel.config import (
    get_postgres_connection_string,
    is_database_configured,
)


logger = logging.getLogger(__name__)


class Settings:
    """
    Legacy settings class for backward compatibility.

    This class maintains the same interface as the original Settings class
    but delegates to the new configuration system.
    """

    @property
    def postgres_connection_string(self) -> str | None:
        """Return the database connection string."""
        return cast(str | None, get_postgres_connection_string())

    @property
    def database_uri(self) -> str | None:
        """Return the database connection string."""
        return self.postgres_connection_string

    @property
    def has_database_config(self) -> bool:
        """Check if database configuration is available."""
        return cast(bool, is_database_configured())


def get_settings() -> Settings:
    """
    Get application settings instance.

    This function maintains backward compatibility with the original interface
    while using the new configuration system internally.
    """
    return Settings()


def validate_postgres_connection_string(connection_string: str) -> bool:
    """
    Validate PostgreSQL connection string format.

    Args:
        connection_string: The connection string to validate

    Returns:
        True if valid, False otherwise
    """
    if not connection_string:
        return False

    return connection_string.startswith(("postgresql://", "postgres://"))


# Create settings instance for backward compatibility
settings = Settings()
