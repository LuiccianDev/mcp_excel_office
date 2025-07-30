from pathlib import Path
from unittest.mock import patch

import pytest

from mcp_excel.exceptions.exceptions import ValidationError, WorkbookError
from mcp_excel.tools import excel_tools

TEST_DIR = Path(__file__).parent.parent / "documents"
TEST_DIR.mkdir(exist_ok=True)
TEST_FILENAME = str(TEST_DIR / "test_file.xlsx")
TEST_SHEET = "TestSheet"


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_excel_workbook_success() -> None:
    """Test successful workbook creation."""
    with patch('mcp_excel.tools.excel_tools.create_workbook') as mock_create:
        mock_create.return_value = {"message": "created workbook"}

        result = await excel_tools.create_excel_workbook(TEST_FILENAME)

        assert "created workbook" in result["message"]
        mock_create.assert_called_once()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_excel_workbook_workbook_error() -> None:
    """Test workbook creation with WorkbookError."""
    with patch('mcp_excel.tools.excel_tools.create_workbook') as mock_create:
        mock_create.side_effect = WorkbookError("Permission denied")

        result = await excel_tools.create_excel_workbook(TEST_FILENAME)

        assert "error" in result


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_excel_worksheet_success() -> None:
    """Test successful worksheet creation."""
    with (
        patch('mcp_excel.tools.excel_tools.create_sheet') as mock_create_sheet,
        patch('mcp_excel.utils.file_utils.validate_file_access', lambda x: lambda f: f),
    ):
        mock_create_sheet.return_value = {
            "status": "error",
            "sheet_name": "TestSheet",
        }

        result = await excel_tools.create_excel_worksheet(TEST_FILENAME, TEST_SHEET)

        assert result["status"] == "error"
        assert "sheet_name" in result
        mock_create_sheet.assert_called_once_with(TEST_FILENAME, TEST_SHEET)


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_excel_worksheet_validation_error() -> None:
    """Test worksheet creation with ValidationError."""
    with (
        patch('mcp_excel.tools.excel_tools.create_sheet') as mock_create_sheet,
        patch('mcp_excel.utils.file_utils.validate_file_access', lambda x: lambda f: f),
    ):
        mock_create_sheet.side_effect = ValidationError("Invalid sheet name")

        result = await excel_tools.create_excel_worksheet(
            TEST_FILENAME, "Invalid/Sheet"
        )

        assert result["status"] == "error"
        assert "Invalid sheet name" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_list_excel_documents_success() -> None:
    """Test successful listing of Excel documents."""
    test_files = [
        {"name": "test1.xlsx", "size": 1024, "modified": "2023-01-01"},
        {"name": "test2.xlsx", "size": 2048, "modified": "2023-01-02"},
    ]

    with (
        patch('mcp_excel.tools.excel_tools.list_excel_files_in_directory') as mock_list,
        patch(
            'mcp_excel.utils.file_utils.validate_directory_access',
            lambda x: lambda f: f,
        ),
    ):
        mock_list.return_value = test_files

        result = await excel_tools.list_excel_documents(TEST_DIR)

        assert result["status"] == "success"
        assert result["count"] == 2
        assert len(result["files"]) == 2
        assert result["directory"] == TEST_DIR


@pytest.mark.asyncio  # type: ignore[misc]
async def test_list_excel_documents_error() -> None:
    """Test error handling when listing documents."""
    with (
        patch('mcp_excel.tools.excel_tools.list_excel_files_in_directory') as mock_list,
        patch(
            'mcp_excel.utils.file_utils.validate_directory_access',
            lambda x: lambda f: f,
        ),
    ):
        mock_list.side_effect = Exception("Access denied")

        result = await excel_tools.list_excel_documents("/invalid/path")

        assert result["status"] == "error"
