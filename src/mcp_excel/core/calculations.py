"""Core functionality for Excel workbook operations.

This module provides functions to apply formulas to cells.
"""

from pathlib import Path
from typing import Any, Final

from openpyxl.worksheet.worksheet import Worksheet

from mcp_excel.core.exceptions import FormulaError, ValidationError
from mcp_excel.core.workbook import get_or_create_workbook
from mcp_excel.utils.cell_utils import validate_cell_reference
from mcp_excel.utils.validation_utils import validate_formula


# Constants
FORMULA_PREFIX: Final[str] = "="


def apply_formula(
    filename: str | Path,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """
    Apply an Excel formula to a specific cell in a worksheet.

    Args:
        filename: Path to the Excel file.
        sheet_name: Name of the worksheet.
        cell: Cell reference (e.g., 'A1').
        formula: Excel formula to apply (with or without '=' prefix).

    Returns:
        Dict containing operation result details.

    Raises:
        ValidationError: If cell reference or sheet is invalid.
        FormulaError: If formula application or file save fails.
    """
    # Input validation
    if not validate_cell_reference(cell):
        raise ValidationError(f"Invalid cell reference: {cell}")

    # Load workbook and validate sheet
    workbook = get_or_create_workbook(str(filename))
    _validate_worksheet_exists(workbook, sheet_name)
    worksheet = workbook[sheet_name]

    # Process and validate formula
    formula = _ensure_formula_format(formula)
    _validate_formula_syntax(formula)

    # Apply formula and save
    _apply_formula_to_cell(worksheet, cell, formula)
    _save_workbook(workbook, str(filename))

    # Return success result
    result = {
        "status": "success",
        "message": f"Applied formula '{formula}' to cell {cell}",
        "cell": cell,
        "formula": formula,
    }
    return result

def _validate_worksheet_exists(workbook: Any, sheet_name: str) -> None:
    """Validate if the specified worksheet exists in the workbook."""
    if sheet_name not in workbook.sheetnames:
        raise ValidationError(f"Sheet '{sheet_name}' not found")


def _ensure_formula_format(formula: str) -> str:
    """Ensure the formula starts with '='."""
    return formula if formula.startswith(FORMULA_PREFIX) else f"={formula}"


def _validate_formula_syntax(formula: str) -> None:
    """Validate the syntax of the formula."""
    is_valid, message = validate_formula(formula)
    if not is_valid:
        raise FormulaError(f"Invalid formula syntax: {message}")


def _apply_formula_to_cell(worksheet: Worksheet, cell_ref: str, formula: str) -> None:
    """Apply formula to the specified cell in the worksheet."""
    try:
        worksheet[cell_ref].value = formula
    except Exception as e:
        raise FormulaError(f"Failed to apply formula to cell: {str(e)}") from e


def _save_workbook(workbook: Any, filename: str) -> None:
    """Save the workbook to the specified file."""
    try:
        workbook.save(filename)
    except Exception as e:
        raise FormulaError(f"Failed to save workbook: {str(e)}") from e
