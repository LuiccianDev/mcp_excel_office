import re
from typing import Any

from mcp_excel.utils.cell_utils import parse_cell_range, validate_cell_reference
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


def validate_formula_in_cell_operation(
    filepath: str, sheet_name: str, cell: str, formula: str
) -> dict[str, Any]:
    # Validaciones rápidas
    if not validate_cell_reference(cell):
        return {"error": f"Invalid cell reference: {cell}"}
    try:
        wb = load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            return {"error": f"Sheet '{sheet_name}' not found"}
        # Validar sintaxis de fórmula
        result: tuple[bool, str] = validate_formula(formula)
        is_valid, message = result
        if not is_valid:
            return {"error": f"Invalid formula syntax: {message}"}
        # Validar referencias de celda en la fórmula
        cell_refs = re.findall(r"[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?", formula)
        for ref in cell_refs:
            if ":" in ref:
                start, end = ref.split(":")
                if not (
                    validate_cell_reference(start) and validate_cell_reference(end)
                ):
                    return {"error": f"Invalid cell range reference in formula: {ref}"}
            else:
                if not validate_cell_reference(ref):
                    return {"error": f"Invalid cell reference in formula: {ref}"}
        # Comparar con el contenido actual de la celda
        sheet = wb[sheet_name]
        cell_obj = sheet[cell]
        current_formula = cell_obj.value
        if isinstance(current_formula, str) and current_formula.startswith("="):
            if formula.startswith("="):
                if current_formula != formula:
                    return {
                        "message": "Formula is valid but doesn't match cell content",
                        "valid": True,
                        "matches": False,
                        "cell": cell,
                        "provided_formula": formula,
                        "current_formula": current_formula,
                    }
            else:
                if current_formula != f"={formula}":
                    return {
                        "message": "Formula is valid but doesn't match cell content",
                        "valid": True,
                        "matches": False,
                        "cell": cell,
                        "provided_formula": formula,
                        "current_formula": current_formula,
                    }
                else:
                    return {
                        "message": "Formula is valid and matches cell content",
                        "valid": True,
                        "matches": True,
                        "cell": cell,
                        "formula": formula,
                    }
        else:
            return {
                "message": "Formula is valid but cell contains no formula",
                "valid": True,
                "matches": False,
                "cell": cell,
                "provided_formula": formula,
                "current_content": str(current_formula) if current_formula else "",
            }
    except Exception as e:
        return {"error": str(e)}

    # Fallback de seguridad si ningún return anterior se ejecutó usando mypy or ruff
    return {"error": "Unknown error: no result was returned from the function logic"}


def validate_range_in_sheet_operation(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
) -> dict[str, Any]:
    if not validate_cell_reference(start_cell):
        return {"error": f"Invalid start cell reference: {start_cell}"}
    if end_cell and not validate_cell_reference(end_cell):
        return {"error": f"Invalid end cell reference: {end_cell}"}
    try:
        wb = load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            return {"error": f"Sheet '{sheet_name}' not found"}
        worksheet = wb[sheet_name]
        data_max_row = worksheet.max_row
        data_max_col = worksheet.max_column
        try:
            start_row, start_col, end_row, end_col = parse_cell_range(
                start_cell, end_cell
            )
        except ValueError as e:
            return {"error": f"Invalid range: {str(e)}"}
        if end_row is None:
            end_row = start_row
        if end_col is None:
            end_col = start_col
        is_valid, message = validate_range_bounds(
            worksheet, start_row, start_col, end_row, end_col
        )
        if not is_valid:
            return {"error": message}
        range_str = f"{start_cell}" if end_cell is None else f"{start_cell}:{end_cell}"
        data_range_str = f"A1:{get_column_letter(data_max_col)}{data_max_row}"
        extends_beyond_data = end_row > data_max_row or end_col > data_max_col
        return {
            "message": (
                f"Range '{range_str}' is valid. "
                f"Sheet contains data in range '{data_range_str}'"
            ),
            "valid": True,
            "range": range_str,
            "data_range": data_range_str,
            "extends_beyond_data": extends_beyond_data,
            "data_dimensions": {
                "max_row": data_max_row,
                "max_col": data_max_col,
                "max_col_letter": get_column_letter(data_max_col),
            },
        }
    except Exception as e:
        return {"error": str(e)}


def validate_formula(formula: str) -> tuple[bool, str]:
    if not formula.startswith("="):
        return False, "Formula must start with '='"
    formula = formula[1:]
    parens = 0
    for c in formula:
        if c == "(":
            parens += 1
        elif c == ")":
            parens -= 1
        if parens < 0:
            return False, "Unmatched closing parenthesis"
    if parens > 0:
        return False, "Unclosed parenthesis"
    func_pattern = r"([A-Z]+)\("
    funcs = re.findall(func_pattern, formula)
    unsafe_funcs = {"INDIRECT", "HYPERLINK", "WEBSERVICE", "DGET", "RTD"}
    for func in funcs:
        if func in unsafe_funcs:
            return False, f"Unsafe function: {func}"
    return True, "Formula is valid"


def validate_range_bounds(
    worksheet: Worksheet,
    start_row: int,
    start_col: int,
    end_row: int | None = None,
    end_col: int | None = None,
) -> tuple[bool, str]:
    max_row = worksheet.max_row
    max_col = worksheet.max_column
    try:
        if start_row < 1 or start_row > max_row:
            return False, f"Start row {start_row} out of bounds (1-{max_row})"
        if start_col < 1 or start_col > max_col:
            return False, (
                f"Start column {get_column_letter(start_col)} "
                f"out of bounds (A-{get_column_letter(max_col)})"
            )
        if end_row is not None and end_col is not None:
            if end_row < start_row:
                return False, "End row cannot be before start row"
            if end_col < start_col:
                return False, "End column cannot be before start column"
            if end_row > max_row:
                return False, f"End row {end_row} out of bounds (1-{max_row})"
            if end_col > max_col:
                return False, (
                    f"End column {get_column_letter(end_col)} "
                    f"out of bounds (A-{get_column_letter(max_col)})"
                )
        return True, "Range is valid"
    except Exception as e:
        return False, f"Invalid range: {e!s}"
