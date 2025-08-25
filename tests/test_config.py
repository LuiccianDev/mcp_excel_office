"""
Tests for configuration management system.

This module tests the comprehensive configuration management including
environment variable handling, user_config variable substitution,
and validation across different deployment modes.
"""

import os
import tempfile
import pytest
from pathlib import Path
from unittest.mock import patch

from mcp_excel.config import (
    MCPExcelConfig,
    ConfigurationManager,
    ConfigurationError,
    DatabaseConfig,
    FileConfig,
    get_config,
    get_postgres_connection_string,
    get_directory,
    is_database_configured,
    validate_file_path,
)


class TestDatabaseConfig:
    """Test database configuration validation."""

    def test_valid_postgresql_connection_string(self):
        """Test valid PostgreSQL connection string."""
        config = DatabaseConfig(connection_string="postgresql://user:pass@localhost:5432/db")
        assert config.connection_string == "postgresql://user:pass@localhost:5432/db"
        assert config.is_configured is True

    def test_valid_postgres_connection_string(self):
        """Test valid postgres:// connection string."""
        config = DatabaseConfig(connection_string="postgres://user:pass@localhost:5432/db")
        assert config.connection_string == "postgres://user:pass@localhost:5432/db"
        assert config.is_configured is True

    def test_none_connection_string(self):
        """Test None connection string."""
        config = DatabaseConfig(connection_string=None)
        assert config.connection_string is None
        assert config.is_configured is False

    def test_invalid_connection_string_format(self):
        """Test invalid connection string format."""
        with pytest.raises(ValueError, match="must start with 'postgresql://' or 'postgres://'"):
            DatabaseConfig(connection_string="mysql://user:pass@localhost/db")

    def test_malformed_connection_string(self):
        """Test malformed connection string."""
        with pytest.raises(ValueError, match="appears to be malformed"):
            DatabaseConfig(connection_string="postgresql://invalid")


class TestFileConfig:
    """Test file configuration validation."""

    def test_valid_directory(self):
        """Test valid directory configuration."""
        with tempfile.TemporaryDirectory() as temp_dir:
            config = FileConfig(directory=temp_dir)
            assert Path(config.directory).exists()
            assert Path(config.directory).is_dir()

    def test_directory_creation(self):
        """Test directory creation when it doesn't exist."""
        with tempfile.TemporaryDirectory() as temp_dir:
            new_dir = Path(temp_dir) / "new_directory"
            config = FileConfig(directory=str(new_dir))
            assert Path(config.directory).exists()
            assert Path(config.directory).is_dir()

    def test_invalid_directory_empty(self):
        """Test empty directory string."""
        with pytest.raises(ValueError, match="Directory cannot be empty"):
            FileConfig(directory="")

    def test_default_values(self):
        """Test default configuration values."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Change to temp directory to avoid creating ./documents in test environment
            with patch('pathlib.Path.cwd', return_value=Path(temp_dir)):
                config = FileConfig()
                assert config.max_file_size == 100 * 1024 * 1024
                assert config.allowed_extensions == ['.xlsx']


class TestMCPExcelConfig:
    """Test main configuration class."""

    def test_default_configuration(self):
        """Test default configuration values."""
        with tempfile.TemporaryDirectory() as temp_dir:
            config = MCPExcelConfig(directory=temp_dir)
            assert config.postgres_connection_string is None
            assert Path(config.directory).resolve() == Path(temp_dir).resolve()
            assert config.log_level == "INFO"

    def test_environment_variable_configuration(self):
        """Test configuration from environment variables."""
        with tempfile.TemporaryDirectory() as temp_dir:
            config = MCPExcelConfig(
                postgres_connection_string='postgresql://user:pass@localhost:5432/db',
                directory=temp_dir,
                log_level='DEBUG'
            )
            assert config.postgres_connection_string == 'postgresql://user:pass@localhost:5432/db'
            assert Path(config.directory).resolve() == Path(temp_dir).resolve()
            assert config.log_level == 'DEBUG'

    def test_user_config_variable_substitution(self):
        """Test ${user_config.*} variable substitution."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Set environment variables that user_config variables should resolve to
            env_vars = {
                'POSTGRES_CONNECTION_STRING': 'postgresql://user:pass@localhost:5432/db',
                'DIRECTORY': temp_dir
            }
            with patch.dict(os.environ, env_vars, clear=True):
                # Test user_config variable substitution
                config = MCPExcelConfig(
                    postgres_connection_string='${user_config.postgres_connection_string}',
                    directory='${user_config.directory}'
                )
                assert config.postgres_connection_string == 'postgresql://user:pass@localhost:5432/db'
                assert Path(config.directory).resolve() == Path(temp_dir).resolve()

    def test_user_config_variable_not_found(self):
        """Test user_config variable when environment variable not found."""
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.dict(os.environ, {}, clear=True):
                # This should set to None when variable not found
                config = MCPExcelConfig(
                    postgres_connection_string='${user_config.nonexistent_var}',
                    directory=temp_dir
                )
                assert config.postgres_connection_string is None
                assert Path(config.directory).resolve() == Path(temp_dir).resolve()

    def test_configuration_properties(self):
        """Test configuration properties."""
        with tempfile.TemporaryDirectory() as temp_dir:
            config = MCPExcelConfig(
                postgres_connection_string='postgresql://user:pass@localhost:5432/db',
                directory=temp_dir
            )

            # Test database_config property
            db_config = config.database_config
            assert isinstance(db_config, DatabaseConfig)
            assert db_config.is_configured is True

            # Test file_config property
            file_config = config.file_config
            assert isinstance(file_config, FileConfig)
            assert Path(file_config.directory).resolve() == Path(temp_dir).resolve()

            # Test get_effective_config
            effective = config.get_effective_config()
            assert effective['postgres_connection_string'] == 'postgresql://user:pass@localhost:5432/db'
            assert Path(effective['directory']).resolve() == Path(temp_dir).resolve()
            assert effective['database_configured'] is True


class TestConfigurationManager:
    """Test configuration manager singleton."""

    def test_singleton_pattern(self):
        """Test that ConfigurationManager follows singleton pattern."""
        manager1 = ConfigurationManager()
        manager2 = ConfigurationManager()
        assert manager1 is manager2

    def test_configuration_loading(self):
        """Test configuration loading and access."""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = ConfigurationManager()
            manager.reload_configuration(directory=temp_dir)
            config = manager.config
            assert isinstance(config, MCPExcelConfig)
            assert Path(config.directory).resolve() == Path(temp_dir).resolve()

    def test_configuration_reload(self):
        """Test configuration reloading with overrides."""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = ConfigurationManager()

            # Initial configuration
            manager.reload_configuration(directory=temp_dir)
            assert Path(manager.get_directory()).resolve() == Path(temp_dir).resolve()
            assert manager.get_postgres_connection_string() is None

            # Reload with overrides
            manager.reload_configuration(
                directory=temp_dir,
                postgres_connection_string='postgresql://user:pass@localhost:5432/db'
            )

            assert manager.get_postgres_connection_string() == 'postgresql://user:pass@localhost:5432/db'
            assert manager.is_database_configured() is True

    def test_file_path_validation(self):
        """Test file path validation against configured directory."""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = ConfigurationManager()
            manager.reload_configuration(directory=temp_dir)

            # Valid file path within directory
            test_file = Path(temp_dir) / "test.xlsx"
            validated_path = manager.validate_file_path(str(test_file))
            assert Path(validated_path).is_absolute()
            assert str(test_file.resolve()) == validated_path

            # Invalid file path outside directory
            with pytest.raises(ConfigurationError, match="outside allowed directory"):
                manager.validate_file_path("/tmp/outside.xlsx")

    def test_configuration_summary(self):
        """Test configuration summary generation."""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = ConfigurationManager()
            manager.reload_configuration(
                postgres_connection_string='postgresql://user:pass@localhost:5432/db',
                directory=temp_dir
            )
            summary = manager.get_configuration_summary()

            assert "MCP Excel Office Server Configuration:" in summary
            assert "Database: Configured" in summary
            assert "Log Level: INFO" in summary


class TestGlobalFunctions:
    """Test global configuration functions."""

    def test_global_functions(self):
        """Test global configuration access functions."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Force reload of global configuration
            from mcp_excel.config import config_manager
            config_manager.reload_configuration(
                postgres_connection_string='postgresql://user:pass@localhost:5432/db',
                directory=temp_dir
            )

            # Test global functions
            config = get_config()
            assert isinstance(config, MCPExcelConfig)

            assert get_postgres_connection_string() == 'postgresql://user:pass@localhost:5432/db'
            assert Path(get_directory()).resolve() == Path(temp_dir).resolve()
            assert is_database_configured() is True

            # Test file path validation
            test_file = Path(temp_dir) / "test.xlsx"
            validated = validate_file_path(str(test_file))
            assert Path(validated).is_absolute()


class TestUserConfigIntegration:
    """Test user_config variable integration for different deployment modes."""

    def test_dxt_deployment_simulation(self):
        """Test configuration as it would work in DXT deployment."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Simulate DXT environment where user_config variables are passed
            # and environment variables contain the actual values
            env_vars = {
                'POSTGRES_CONNECTION_STRING': 'postgresql://dxt_user:pass@localhost:5432/dxt_db',
                'DIRECTORY': temp_dir
            }
            with patch.dict(os.environ, env_vars, clear=True):
                # Simulate DXT passing user_config variables
                config = MCPExcelConfig(
                    postgres_connection_string='${user_config.postgres_connection_string}',
                    directory='${user_config.directory}'
                )

                # Verify substitution worked
                assert config.postgres_connection_string == 'postgresql://dxt_user:pass@localhost:5432/dxt_db'
                assert Path(config.directory).resolve() == Path(temp_dir).resolve()
                assert config.database_config.is_configured is True

    def test_traditional_mcp_deployment_simulation(self):
        """Test configuration as it would work in traditional MCP deployment."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Simulate traditional MCP where configuration is passed directly
            config = MCPExcelConfig(
                postgres_connection_string='postgresql://mcp_user:pass@localhost:5432/mcp_db',
                directory=temp_dir
            )

            # Verify configuration was used directly
            assert config.postgres_connection_string == 'postgresql://mcp_user:pass@localhost:5432/mcp_db'
            assert Path(config.directory).resolve() == Path(temp_dir).resolve()
            assert config.database_config.is_configured is True

    def test_cli_deployment_simulation(self):
        """Test configuration as it would work in CLI deployment."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Simulate CLI where configuration comes from arguments and environment
            # CLI arguments would override through reload_configuration
            manager = ConfigurationManager()
            manager.reload_configuration(
                postgres_connection_string='postgresql://cli_user:pass@localhost:5432/cli_db',
                directory=temp_dir
            )

            config = manager.config
            assert config.postgres_connection_string == 'postgresql://cli_user:pass@localhost:5432/cli_db'
            assert Path(config.directory).resolve() == Path(temp_dir).resolve()
            assert config.database_config.is_configured is True
