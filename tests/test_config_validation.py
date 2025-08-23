"""
Tests for configuration validation and error handling.

This module tests the comprehensive validation and error handling
capabilities of the configuration system, including database connection
testing, file path validation, and security checks.
"""

import os
import tempfile
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from mcp_excel.config import (
    test_database_connection,
    validate_configuration,
    ConfigurationManager,
    ConfigurationError,
)


class TestDatabaseConnectionValidation:
    """Test database connection validation and error handling."""

    def test_no_database_configured(self):
        """Test database connection test when no database is configured."""
        with patch('mcp_excel.config.get_postgres_connection_string', return_value=None):
            success, error = test_database_connection()
            assert success is False
            assert "No database connection string configured" in error

    def test_psycopg2_not_installed(self):
        """Test database connection test when psycopg2 is not available."""
        with patch('mcp_excel.config.get_postgres_connection_string', return_value='postgresql://test'):
            with patch('builtins.__import__', side_effect=ImportError("No module named 'psycopg2'")):
                success, error = test_database_connection()
                assert success is False
                assert "psycopg2 not installed" in error

    def test_database_connection_failure(self):
        """Test database connection test when connection fails."""
        with patch('mcp_excel.config.get_postgres_connection_string', return_value='postgresql://invalid'):
            # Mock psycopg2 to be available but connection to fail
            mock_psycopg2 = MagicMock()
            mock_psycopg2.connect.side_effect = Exception("Connection failed")
            mock_psycopg2.Error = Exception

            with patch.dict('sys.modules', {'psycopg2': mock_psycopg2}):
                success, error = test_database_connection()
                assert success is False
                assert "Database connection failed" in error

    def test_database_connection_success(self):
        """Test successful database connection test."""
        with patch('mcp_excel.config.get_postgres_connection_string', return_value='postgresql://valid'):
            # Mock psycopg2 to be available and connection to succeed
            mock_conn = MagicMock()
            mock_psycopg2 = MagicMock()
            mock_psycopg2.connect.return_value = mock_conn
            mock_psycopg2.Error = Exception

            with patch.dict('sys.modules', {'psycopg2': mock_psycopg2}):
                success, error = test_database_connection()
                assert success is True
                assert error is None
                mock_conn.close.assert_called_once()


class TestConfigurationValidation:
    """Test comprehensive configuration validation."""

    def test_validate_configuration_all_valid(self, tmp_path):
        """Test configuration validation when everything is valid."""
        manager = ConfigurationManager()
        manager.reload_configuration(
            postgres_connection_string='postgresql://user:pass@localhost:5432/db',
            directory=str(tmp_path),
            log_level='INFO'
        )

        # Mock successful database connection test
        with patch('mcp_excel.config.test_database_connection', return_value=(True, None)):
            issues = validate_configuration()
            assert len(issues) == 0

    def test_validate_configuration_database_issues(self, tmp_path):
        """Test configuration validation with database issues."""
        manager = ConfigurationManager()
        # Use a valid format but non-working connection string
        manager.reload_configuration(
            postgres_connection_string='postgresql://user:pass@nonexistent:5432/db',
            directory=str(tmp_path),
            log_level='INFO'
        )

        # Mock failed database connection test
        with patch('mcp_excel.config.test_database_connection', return_value=(False, "Connection failed")):
            issues = validate_configuration()
            assert len(issues) > 0
            assert any("Database configuration issue" in issue for issue in issues)

    def test_validate_configuration_directory_not_exists(self):
        """Test configuration validation when directory doesn't exist."""
        nonexistent_dir = "/path/that/does/not/exist"
        manager = ConfigurationManager()

        # This should fail during configuration loading due to directory validation
        with pytest.raises(ConfigurationError):
            manager.reload_configuration(directory=nonexistent_dir)

    def test_validate_configuration_invalid_log_level(self, tmp_path):
        """Test configuration validation with invalid log level."""
        manager = ConfigurationManager()

        # Bypass normal validation to test the validate_configuration function
        with patch('mcp_excel.config.get_config') as mock_get_config:
            mock_config = MagicMock()
            mock_config.postgres_connection_string = None
            mock_config.log_level = 'INVALID_LEVEL'
            mock_get_config.return_value = mock_config

            with patch('mcp_excel.config.get_directory', return_value=str(tmp_path)):
                issues = validate_configuration()
                assert len(issues) > 0
                assert any("Invalid log level" in issue for issue in issues)


class TestFilePathValidation:
    """Test file path validation and security."""

    def test_file_path_validation_within_directory(self, tmp_path):
        """Test file path validation for files within allowed directory."""
        manager = ConfigurationManager()
        manager.reload_configuration(directory=str(tmp_path))

        # Test valid file path
        test_file = tmp_path / "test.xlsx"
        validated_path = manager.validate_file_path(str(test_file))
        assert Path(validated_path).is_absolute()
        assert str(test_file.resolve()) == validated_path

    def test_file_path_validation_outside_directory(self, tmp_path):
        """Test file path validation for files outside allowed directory."""
        manager = ConfigurationManager()
        manager.reload_configuration(directory=str(tmp_path))

        # Test invalid file path outside directory
        with pytest.raises(ConfigurationError, match="outside allowed directory"):
            manager.validate_file_path("/tmp/outside.xlsx")

    def test_file_path_validation_directory_traversal(self, tmp_path):
        """Test file path validation prevents directory traversal attacks."""
        manager = ConfigurationManager()
        manager.reload_configuration(directory=str(tmp_path))

        # Test directory traversal attempt
        with pytest.raises(ConfigurationError, match="outside allowed directory"):
            manager.validate_file_path(f"{tmp_path}/../../../etc/passwd")


class TestConfigurationErrorHandling:
    """Test configuration error handling and recovery."""

    def test_configuration_error_on_invalid_postgres_string(self, tmp_path):
        """Test configuration error when PostgreSQL string is invalid."""
        with pytest.raises(ConfigurationError):
            manager = ConfigurationManager()
            manager.reload_configuration(
                postgres_connection_string='invalid://not-postgres',
                directory=str(tmp_path)
            )

    def test_configuration_error_on_invalid_directory(self):
        """Test configuration error when directory is invalid."""
        with pytest.raises(ConfigurationError):
            manager = ConfigurationManager()
            manager.reload_configuration(directory="/path/that/cannot/be/created/due/to/permissions")

    def test_configuration_recovery_after_error(self, tmp_path):
        """Test configuration recovery after an error."""
        manager = ConfigurationManager()

        # First, set valid configuration
        manager.reload_configuration(directory=str(tmp_path))
        assert Path(manager.get_directory()).resolve() == Path(tmp_path).resolve()

        # Try to set invalid configuration (should fail)
        with pytest.raises(ConfigurationError):
            manager.reload_configuration(directory="/invalid/path")

        # Configuration should remain unchanged after error
        assert Path(manager.get_directory()).resolve() == Path(tmp_path).resolve()

    def test_configuration_summary_with_errors(self, tmp_path):
        """Test configuration summary includes error information."""
        manager = ConfigurationManager()
        manager.reload_configuration(directory=str(tmp_path))

        summary = manager.get_configuration_summary()
        assert "MCP Excel Office Server Configuration:" in summary
        assert "Database: Not configured" in summary
        assert "Log Level: INFO" in summary


class TestSecurityValidation:
    """Test security-related validation and error handling."""

    def test_sql_injection_prevention(self, tmp_path):
        """Test that SQL injection attempts are prevented."""
        # This would be tested in the database tools, but we can test
        # that the configuration system properly validates connection strings
        # Test that malicious connection strings are rejected
        with pytest.raises(ConfigurationError):
            manager = ConfigurationManager()
            manager.reload_configuration(
                postgres_connection_string='mysql://user:pass@localhost/db; DROP TABLE users;',
                directory=str(tmp_path)
            )

    def test_path_traversal_prevention(self, tmp_path):
        """Test that path traversal attempts are prevented."""
        manager = ConfigurationManager()
        manager.reload_configuration(directory=str(tmp_path))

        # Test various path traversal attempts
        traversal_attempts = [
            "../../../etc/passwd",
            "..\\..\\..\\windows\\system32\\config\\sam",
            f"{tmp_path}/../outside.xlsx",
            "subdir/../../outside.xlsx"
        ]

        for attempt in traversal_attempts:
            with pytest.raises(ConfigurationError, match="outside allowed directory"):
                manager.validate_file_path(attempt)

    def test_file_extension_validation(self):
        """Test that file extension validation works properly."""
        from mcp_excel.utils.file_utils import ensure_xlsx_extension

        # Test that .xlsx extension is enforced
        assert ensure_xlsx_extension("test.xls") == "test.xls.xlsx"
        assert ensure_xlsx_extension("test") == "test.xlsx"
        assert ensure_xlsx_extension("test.xlsx") == "test.xlsx"
