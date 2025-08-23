from pathlib import Path
from unittest.mock import patch

import pytest

from mcp_excel.exceptions.exceptions import ValidationError
from mcp_excel.tools import content_tools


# Test data
TEST_DIR = Path(__file__).parent.parent / "documents"
TEST_DIR.mkdir(exist_ok=True)
TEST_FILENAME = str(TEST_DIR / "test_file.xlsx")
TEST_SHEET = "Sheet1"
TEST_DATA = [["Name", "Age"], ["Alice", 30], ["Bob", 25]]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_read_data_from_excel_success() -> None:
    """Test successful read operation from Excel."""

    with patch("mcp_excel.tools.content_tools.read_excel_range") as mock_read:
        # Mock actual function
        mock_read.return_value = ["Row1", "Row2"]

        result = await content_tools.read_data_from_excel(TEST_FILENAME, TEST_SHEET)

    assert result["status"] == "success"
    assert result["data"] == ["Row1", "Row2"]
    mock_read.assert_called_once()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_read_data_from_excel_no_data() -> None:
    """Test read operation when no data is found."""
    with patch("mcp_excel.tools.content_tools.read_excel_range") as mock_read:
        # Mock actual function to return empty data
        mock_read.return_value = []

        result = await content_tools.read_data_from_excel(TEST_FILENAME, TEST_SHEET)

    assert result["status"] == "error"
    assert "No data found" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_read_data_from_excel_error() -> None:
    """Test error handling during read operation."""
    with patch("mcp_excel.core.data.read_excel_range") as mock_read:
        mock_read.side_effect = Exception("Read error")

        result = await content_tools.read_data_from_excel(TEST_FILENAME, TEST_SHEET)

    assert result["status"] == "error"
    assert "Failed to read Excel data" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_write_data_to_excel_success() -> None:
    """Test successful write operation to Excel."""
    with patch("mcp_excel.core.data.write_data") as mock_write:
        mock_write.return_value = {"message": "Data written successfully"}

        result = await content_tools.write_data_to_excel(
            TEST_FILENAME, TEST_SHEET, TEST_DATA
        )

        assert "successfully" in result
        args, _ = mock_write.call_args
        assert str(args[0]).endswith("test_file.xlsx")
        assert args[1] == TEST_SHEET
        assert args[2] == TEST_DATA


@pytest.mark.asyncio  # type: ignore[misc]
async def test_write_data_to_excel_validation_error() -> None:
    """Test validation error during write operation."""
    with patch("mcp_excel.core.data.write_data") as mock_write:
        mock_write.side_effect = ValidationError("Invalid data")

        result = await content_tools.write_data_to_excel(
            TEST_FILENAME, TEST_SHEET, TEST_DATA
        )

        assert "Error: Invalid data" in result


@pytest.mark.asyncio  # type: ignore[misc]
async def test_write_data_to_excel_general_error() -> None:
    """Test general error during write operation."""
    with patch("mcp_excel.core.data.write_data") as mock_write:
        mock_write.side_effect = Exception("Write error")

        result = await content_tools.write_data_to_excel(
            TEST_FILENAME, TEST_SHEET, TEST_DATA
        )

        assert "Error: Failed to write data: Write error" in result
