from typing import Any
from mcp_excel_server.core.workbook import get_or_create_workbook
from mcp_excel_server.utils.cell_utils import validate_cell_reference
from mcp_excel_server.exceptions.exceptions import ValidationError, CalculationError
from mcp_excel_server.utils.validation_utils import validate_formula


def apply_formula(
    filename: str, sheet_name: str, cell: str, formula: str
) -> dict[str, Any]:
    """Apply any Excel formula to a cell."""
    try:
        if not validate_cell_reference(cell):
            raise ValidationError(f"Invalid cell reference: {cell}")
        wb = get_or_create_workbook(filename)
        if sheet_name not in wb.sheetnames:
            raise ValidationError(f"Sheet '{sheet_name}' not found")
        sheet = wb[sheet_name]
        # Ensure formula starts with =
        if not formula.startswith("="):
            formula = f"={formula}"
        # Validate formula syntax
        is_valid, message = validate_formula(formula)
        if not is_valid:
            raise CalculationError(f"Invalid formula syntax: {message}")
        try:
            # Apply formula to the cell
            cell_obj = sheet[cell]
            cell_obj.value = formula
        except Exception as e:
            raise CalculationError(f"Failed to apply formula to cell: {str(e)}")
        try:
            wb.save(filename)
        except Exception as e:
            raise CalculationError(
                f"Failed to save workbook after applying formula: {str(e)}"
            )
        return {
            "message": f"Applied formula '{formula}' to cell {cell}",
            "cell": cell,
            "formula": formula,
        }
    except (ValidationError, CalculationError) as e:
        raise e
    except Exception as e:
        raise CalculationError(str(e))
