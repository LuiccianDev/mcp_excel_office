# Import exceptions
from typing import Any

from mcp_excel.exceptions.exceptions import CalculationError, ValidationError
from mcp_excel.utils.file_utils import ensure_xlsx_extension, validate_file_access

# Import core/tools/utils with new structure
from mcp_excel.utils.validation_utils import (
    validate_formula_in_cell_operation as validate_formula_impl,
)


async def validate_formula_syntax(
    filename: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """Validate Excel formula syntax without applying it.
    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        cell: Cell reference (e.g., "A1")
        formula: Formula to validate (e.g., "=SUM(A1:A10)")
    """
    filename = ensure_xlsx_extension(filename)
    try:
        result: dict[str, Any] = validate_formula_impl(
            filename, sheet_name, cell, formula
        )
        return result
    except (ValidationError, CalculationError) as e:
        return {"error": f"Error: {str(e)}"}
    except Exception as e:
        return {"error": f"Failed to validate formula: {str(e)}"}


#! No borrar el type: ignore[misc] que se encuentra en la linea siguiente en caso contraio eliminar disallow_untyped_decorators = true de pyproject.toml
@validate_file_access("filename")  # type: ignore[misc]
async def apply_formula(
    filename: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """
    Apply an Excel formula to a specific cell in a worksheet.
    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        cell: Cell reference (e.g., "A1")
        formula: Formula to apply (e.g., "=SUM(A1:A10)")
    """
    filename = ensure_xlsx_extension(filename)

    try:
        # First validate the formula
        validation: dict[str, Any] = validate_formula_impl(
            filename, sheet_name, cell, formula
        )
        if isinstance(validation, dict) and "error" in validation:
            return validation

        # If valid, apply the formula
        from mcp_excel.core.calculations import apply_formula as apply_formula_impl

        result: dict[str, Any] = apply_formula_impl(filename, sheet_name, cell, formula)
        return result
    except (ValidationError, CalculationError) as e:
        return {"error": f"Error: {str(e)}"}
    except Exception as e:
        return {"error": f"Failed to apply formula: {str(e)}"}
