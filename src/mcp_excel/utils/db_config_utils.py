"""Configuration management for the application."""

import os

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings with environment variable support."""

    # Database connection - made optional to allow graceful degradation
    POSTGRES_CONNECTION_STRING: str | None = None

    @property
    def database_uri(self) -> str | None:
        """Return the database connection string."""
        return self.POSTGRES_CONNECTION_STRING

    @property
    def has_database_config(self) -> bool:
        """Check if database configuration is available."""
        return self.POSTGRES_CONNECTION_STRING is not None

    model_config = SettingsConfigDict(
        env_file=".env", env_file_encoding="utf-8", case_sensitive=True, extra="ignore"
    )


def get_settings() -> Settings:
    """
    Get application settings instance.

    This function creates a new Settings instance each time it's called,
    allowing for dynamic configuration updates during runtime.
    """
    return Settings()


def get_postgres_connection_string() -> str | None:
    """
    Get PostgreSQL connection string from environment variables.

    Returns:
        Connection string if available, None otherwise
    """
    return os.environ.get("POSTGRES_CONNECTION_STRING")


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


# Create settings instance - this will not fail if POSTGRES_CONNECTION_STRING is missing
settings = Settings()
