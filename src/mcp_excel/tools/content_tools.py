from typing import Any

from mcp_excel.core.data import read_excel_range
from mcp_excel.exceptions.exceptions import DataError, ValidationError
from mcp_excel.utils.file_utils import ensure_xlsx_extension, validate_file_access


#! No borrar el type: ignore[misc] que se encuentra en la linea siguiente en caso contraio eliminar disallow_untyped_decorators = true de pyproject.toml
@validate_file_access("filename")  # type: ignore[misc]
async def read_data_from_excel(
    filename: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str | None = None,
    preview_only: bool = False,
) -> dict[str, Any]:
    """
    Read data from an Excel worksheet and return it in a structured format.

    Args:
        filename: Path to the Excel workbook file (.xlsx)
        sheet_name: Name of the worksheet to read data from
        start_cell: Cell reference where to start reading data (default is "A1")
        end_cell: Optional cell reference where to stop reading data
        preview_only: If True, only returns a preview of the data

    Returns:
        Dict containing:
        - status: "success" or "error"
        - data: 2D list of cell values (if successful)
        - message: Error message (if error occurred)
    """
    try:
        # Ensure filename has .xlsx extension
        filename = ensure_xlsx_extension(filename)

        # Read data from Excel
        data = read_excel_range(
            filename=filename,
            sheet_name=sheet_name,
            start_cell=start_cell,
            end_cell=end_cell,
            preview_only=preview_only,
        )

        # Handle empty results
        if not data:
            return {
                "status": "error",
                "message": "No data found in the specified range",
                "data": [],
            }

        return {
            "status": "success",
            "data": data,
            "message": f"Successfully read {len(data)} rows from {sheet_name}",
        }

    except FileNotFoundError:
        return {
            "status": "error",
            "message": f"File not found: {filename}",
            "data": None,
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to read Excel data: {str(e)}",
            "data": None,
        }


#! No borrar el type: ignore[misc] que se encuentra en la linea siguiente en caso contraio eliminar disallow_untyped_decorators = true de pyproject.toml
@validate_file_access("filename")  # type: ignore[misc]
async def write_data_to_excel(
    filename: str,
    sheet_name: str,
    data: list[list],
    start_cell: str = "A1",
) -> str:
    """
    Write data to Excel worksheet.
    Excel formula will write to cell without any verification.

    Args:
        filename: Path to the workbook file
        sheet_name: Name of the worksheet to write data to
        data: Data to write (list of lists)
        start_cell: Cell reference where to start writing data (default is "A1")

    """
    filename = ensure_xlsx_extension(filename)
    try:
        from mcp_excel.core.data import write_data

        result = write_data(filename, sheet_name, data, start_cell)
        if isinstance(result, dict) and "error" in result:
            return f"Error: {result['error']}"
        return str(result.get("message", "Data written successfully"))
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Error: Failed to write data: {str(e)}"
