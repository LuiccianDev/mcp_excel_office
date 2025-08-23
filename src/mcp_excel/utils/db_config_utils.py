"""Configuration management for the application."""

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings with environment variable support."""

    # Database connection
    POSTGRES_CONNECTION_STRING: str

    @property
    def database_uri(self) -> str:
        """Return the database connection string."""
        return self.POSTGRES_CONNECTION_STRING

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=True,
        extra="ignore"
    )


# Create settings instance
settings = Settings()
