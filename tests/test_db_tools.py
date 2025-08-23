from pathlib import Path
from unittest.mock import patch

import pytest

from mcp_excel.tools.db_tools import (
    fetch_and_insert_db_to_excel,
    insert_calculated_data_to_db,
)
from mcp_excel.config import ConfigurationManager


# Mock the database connection at the module level
pytestmark = pytest.mark.asyncio

# Using a dummy connection string that matches PostgreSQL format since that's what the code expects
TEST_CONN_STR = "postgresql://user:pass@localhost:5432/testdb"
TEST_QUERY = "SELECT id, name FROM users WHERE active = true"
TEST_SHEET = "Data"
TEST_TABLE = "users"
TEST_COLUMNS = ["id", "name", "email"]
TEST_ROWS = [(1, "John Doe", "john@example.com"), (2, "Jane Smith", "jane@example.com")]
MOCK_DB_RESULT = {
    "columns": ["id", "name", "email"],
    "rows": [
        (1, "John Doe", "john@example.com"),
        (2, "Jane Smith", "jane@example.com"),
    ],
}


@pytest.mark.asyncio  # type: ignore[misc]
async def test_fetch_and_insert_db_to_excel_success(tmp_path) -> None:
    """Test successful database fetch and Excel insert."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_filename = str(tmp_path / "test_output.xlsx")

    with (
        patch(
            "mcp_excel.tools.db_tools.validate_sql_query", return_value=True
        ) as mock_validate,
        patch(
            "mcp_excel.tools.db_tools.fetch_data_from_db", return_value=MOCK_DB_RESULT
        ) as mock_fetch,
        patch(
            "mcp_excel.tools.db_tools.insert_data_to_excel",
            return_value={"status": "success", "message": "Data inserted successfully"},
        ) as mock_insert,
    ):

        result = await fetch_and_insert_db_to_excel(
            TEST_QUERY, test_filename, TEST_SHEET, TEST_CONN_STR
        )

        assert isinstance(result, dict)
        assert "success" in result["message"].lower()

        mock_validate.assert_called_once_with(TEST_QUERY)
        mock_fetch.assert_called_once_with(query=TEST_QUERY, connection_string=TEST_CONN_STR)
        mock_insert.assert_called_once()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_fetch_and_insert_db_to_excel_invalid_query(tmp_path) -> None:
    """Test with invalid SQL query."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_filename = str(tmp_path / "DROP_TABLE_users.xlsx")

    with patch("mcp_excel.tools.db_tools.validate_sql_query", return_value=False):
        result = await fetch_and_insert_db_to_excel(
            "DROP TABLE users", test_filename, TEST_SHEET, TEST_CONN_STR
        )
        assert isinstance(result, dict)
        assert result["status"] == "error"
        assert "invalid or potentially unsafe" in result["message"].lower()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_fetch_and_insert_db_to_excel_db_error(tmp_path) -> None:
    """Test database error handling."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_filename = str(tmp_path / "test_output.xlsx")

    with (
        patch("mcp_excel.tools.db_tools.validate_sql_query", return_value=True),
        patch(
            "mcp_excel.tools.db_tools.fetch_data_from_db",
            return_value={"status": "error", "message": "Connection failed"},
        ),
    ):

        result = await fetch_and_insert_db_to_excel(
            TEST_QUERY, test_filename, TEST_SHEET, TEST_CONN_STR
        )

        assert isinstance(result, dict)
        assert result["status"] == "error"
        assert "connection failed" in result["message"].lower()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_fetch_and_insert_db_to_excel_excel_error(tmp_path) -> None:
    """Test Excel insertion error handling."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_filename = str(tmp_path / "test_output.xlsx")

    with (
        patch("mcp_excel.tools.db_tools.validate_sql_query", return_value=True),
        patch(
            "mcp_excel.tools.db_tools.fetch_data_from_db", return_value=MOCK_DB_RESULT
        ),
        patch(
            "mcp_excel.tools.db_tools.insert_data_to_excel",
            return_value={"status": "error", "message": "Excel error"},
        ),
    ):

        result = await fetch_and_insert_db_to_excel(
            TEST_QUERY, test_filename, TEST_SHEET, TEST_CONN_STR
        )

        assert isinstance(result, dict)
        assert result["status"] == "error"
        assert "excel error" in result["message"].lower()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_insert_calculated_data_to_db_success() -> None:
    """Test successful data insertion to database."""
    mock_result = {
        "status": "success",
        "message": "2 rows inserted",
        "rows_affected": 2,
        "rows_inserted": 2,
        "table": TEST_TABLE,
    }

    with (
        patch(
            "mcp_excel.tools.db_tools.clean_data", return_value=TEST_ROWS
        ) as mock_clean,
        patch(
            "mcp_excel.tools.db_tools.insert_data_to_db", return_value=mock_result
        ) as mock_insert,
    ):

        result = await insert_calculated_data_to_db(
            TEST_TABLE, TEST_COLUMNS, TEST_ROWS, TEST_CONN_STR
        )

        assert isinstance(result, dict)
        assert result.get("status") == "success"
        assert result.get("table") == TEST_TABLE
        assert "2" in result.get("message", "")

        mock_clean.assert_called_once_with(TEST_ROWS, TEST_COLUMNS)
        mock_insert.assert_called_once_with(
            table=TEST_TABLE, columns=TEST_COLUMNS, rows=TEST_ROWS, connection_string=TEST_CONN_STR
        )


@pytest.mark.asyncio  # type: ignore[misc]
async def test_insert_calculated_data_to_db_error() -> None:
    """Test database insertion error handling."""
    error_response = {
        "error": "Duplicate key",
        "status": "error",
        "message": "Database error: Duplicate key",
    }

    with (
        patch("mcp_excel.tools.db_tools.clean_data", return_value=TEST_ROWS),
        patch(
            "mcp_excel.tools.db_tools.insert_data_to_db", return_value=error_response
        ),
    ):

        result = await insert_calculated_data_to_db(
            TEST_CONN_STR, TEST_TABLE, TEST_COLUMNS, TEST_ROWS
        )

        assert result["status"] == "error"
        assert "duplicate" in result["message"].lower()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_insert_calculated_data_to_db_validation_error() -> None:
    """Test validation error during data insertion."""
    with patch(
        "mcp_excel.tools.db_tools.clean_data", side_effect=ValueError("Invalid data")
    ) as mock_clean:

        result = await insert_calculated_data_to_db(
            TEST_CONN_STR, TEST_TABLE, TEST_COLUMNS, [("invalid",)]
        )

        assert result["status"] == "error"
        assert "invalid" in result["message"].lower()
        mock_clean.assert_called_once()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_insert_calculated_data_to_db_connection_error() -> None:
    """Test database connection error during insertion."""
    error_response = {
        "status": "error",
        "message": "Connection failed: Could not connect to database",
    }

    with (
        patch("mcp_excel.tools.db_tools.clean_data", return_value=TEST_ROWS),
        patch(
            "mcp_excel.tools.db_tools.insert_data_to_db", return_value=error_response
        ),
    ):

        result = await insert_calculated_data_to_db(
            "invalid_connection", TEST_TABLE, TEST_COLUMNS, TEST_ROWS
        )

        assert result["status"] == "error"


@pytest.mark.asyncio  # type: ignore[misc]
async def test_insert_calculated_data_to_db_exception() -> None:
    """Test exception handling during database insertion."""
    with (
        patch("mcp_excel.tools.db_tools.clean_data", return_value=TEST_ROWS),
        patch(
            "mcp_excel.tools.db_tools.insert_data_to_db",
            side_effect=Exception("Unexpected error"),
        ) as mock_insert,
    ):

        result = await insert_calculated_data_to_db(
            TEST_CONN_STR, TEST_TABLE, TEST_COLUMNS, TEST_ROWS
        )

        assert result["status"] == "error"
        assert "unexpected" in result["message"].lower()
        mock_insert.assert_called_once()
