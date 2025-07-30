from pathlib import Path
from unittest.mock import patch

import pytest

from mcp_excel.exceptions.exceptions import ValidationError
from mcp_excel.tools import formulas_excel_tools

# Test data
TEST_DIR = Path(__file__).parent.parent / "documents"
TEST_DIR.mkdir(exist_ok=True)
TEST_FILENAME = str(TEST_DIR / "test_workbook.xlsx")
TEST_SHEET = "Sheet1"
TEST_CELL = "A1"
TEST_FORMULA = "=SUM(A1:A10)"
TEST_INVALID_FORMULA = "=SUM(A1:A10"  # Missing closing parenthesis


# Test for validate_formula_syntax
@pytest.mark.asyncio  # type: ignore[misc]
async def test_validate_formula_syntax_valid() -> None:
    """Test successful formula validation."""
    with patch(
        'mcp_excel.tools.formulas_excel_tools.validate_formula_in_cell_operation'
    ) as mock_validate:
        mock_validate.return_value = {"valid": True, "formula": TEST_FORMULA}

        result = await formulas_excel_tools.validate_formula_syntax(
            filename=TEST_FILENAME,
            sheet_name=TEST_SHEET,
            cell=TEST_CELL,
            formula=TEST_FORMULA,
        )

        assert result["valid"] is True
        assert result["formula"] == TEST_FORMULA
        mock_validate.assert_called_once_with(
            TEST_FILENAME, TEST_SHEET, TEST_CELL, TEST_FORMULA
        )


@pytest.mark.asyncio  # type: ignore[misc]
async def test_validate_formula_syntax_invalid() -> None:
    """Test validation of invalid formula."""
    with patch(
        'mcp_excel.tools.formulas_excel_tools.validate_formula_in_cell_operation'
    ) as mock_validate:
        mock_validate.side_effect = ValidationError("Invalid formula syntax")

        result = await formulas_excel_tools.validate_formula_syntax(
            filename=TEST_FILENAME,
            sheet_name=TEST_SHEET,
            cell=TEST_CELL,
            formula=TEST_INVALID_FORMULA,
        )

        assert result["status"] == "error"
        assert "Invalid formula syntax" in result["message"]


# Test for apply_formula
@patch(
    "mcp_excel.tools.formulas_excel_tools.validate_file_access",
    lambda arg: (lambda f: f),
)
@pytest.mark.asyncio  # type: ignore[misc]
async def test_apply_formula_success() -> None:
    """Test successful formula application."""
    with (
        patch(
            'mcp_excel.tools.formulas_excel_tools.validate_formula_in_cell_operation'
        ) as mock_validate,
        patch('mcp_excel.tools.formulas_excel_tools.apply_formula') as mock_apply,
    ):

        # Simular validación exitosa
        mock_validate.return_value = {
            "status": "success",
            "message": "Formula is valid and matches cell content",
            "valid": True,
            "matches": True,
            "cell": TEST_CELL,
            "formula": TEST_FORMULA,
        }
        print("Mock creado:", mock_validate)
        # Simular aplicación exitosa
        mock_apply.return_value = {
            "status": "success",
            "cell": TEST_CELL,
            "formula": TEST_FORMULA,
            "value": 100,
        }

        result = await formulas_excel_tools.apply_formula_excel(
            filename=TEST_FILENAME,
            sheet_name=TEST_SHEET,
            cell=TEST_CELL,
            formula=TEST_FORMULA,
        )

        # Aserciones
        assert result["status"] == "success"
        assert result["cell"] == TEST_CELL
        assert result["formula"] == TEST_FORMULA
        assert "value" in result

        # Verificar llamadas
        mock_validate.assert_called_once_with(
            TEST_FILENAME, TEST_SHEET, TEST_CELL, TEST_FORMULA
        )
        mock_apply.assert_called_once_with(
            TEST_FILENAME, TEST_SHEET, TEST_CELL, TEST_FORMULA
        )


@patch(
    "mcp_excel.tools.formulas_excel_tools.validate_file_access",
    lambda arg: (lambda f: f),
)
@pytest.mark.asyncio  # type: ignore[misc]
async def test_apply_formula_validation_failure() -> None:
    """Test formula application with invalid formula."""
    with patch(
        'mcp_excel.tools.formulas_excel_tools.validate_formula_in_cell_operation'
    ) as mock_validate:

        # Simular validación exitosa
        mock_validate.return_value = {
            "status": "error",
            "message": "Formula is invalid and doesn't match cell content",
        }

        result = await formulas_excel_tools.apply_formula_excel(
            filename=TEST_FILENAME,
            sheet_name=TEST_SHEET,
            cell=TEST_CELL,
            formula=TEST_INVALID_FORMULA,
        )

        assert result["status"] == "error"
        assert "Formula is invalid and doesn't match cell content" in result["message"]
        mock_validate.assert_called_once()
