
from email import message
from mcp_excel_server.utils.file_utils import  ensure_xlsx_extension,check_file_writeable
import os
# Import exceptions
from mcp_excel_server.exceptions.exceptions import (
    ValidationError,
    CalculationError,

)
# Import core/tools/utils with new structure
from mcp_excel_server.utils.validation_utils import (
    validate_formula_in_cell_operation as validate_formula_impl,
    validate_range_in_sheet_operation as validate_range_impl
)
async def validate_formula_syntax(filename: str,sheet_name: str,cell: str,formula: str,) -> str:
    """Validate Excel formula syntax without applying it.
    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        cell: Cell reference (e.g., "A1")
        formula: Formula to validate (e.g., "=SUM(A1:A10)")
    """
    filename = ensure_xlsx_extension(filename)
    try:
        result = validate_formula_impl(filename, sheet_name, cell, formula)
        return result["message"]
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to validate formula: {str(e)}"

async def apply_formula(filename: str,sheet_name: str,cell: str,formula: str,
) -> str:
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
        is_valid, message = check_file_writeable(filename)
        if not is_valid:
            return f"Error: {message}"

        # First validate the formula
        validation = validate_formula_impl(filename, sheet_name, cell, formula)
        if isinstance(validation, dict) and "error" in validation:
            return f"Error: {validation['error']}"

            
        # If valid, apply the formula
        from mcp_excel_server.core.calculations import apply_formula as apply_formula_impl
        result = apply_formula_impl(filename, sheet_name, cell, formula)
        return result["message"]
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to apply formula: {str(e)}"
