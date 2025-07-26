from typing import List, Dict, Any
from mcp_excel_server.utils.file_utils import (
    check_file_writeable,
    ensure_xlsx_extension,
)
from mcp_excel_server.exceptions.exceptions import (
    ValidationError,
    DataError,
)


async def read_data_from_excel(
    filename: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    preview_only: bool = False,
) -> str:
    """
    Read data from Excel worksheet.

    Args:
        filename: Path to the workbook file
        sheet_name: Name of the worksheet to read data from
        start_cell: Cell reference where to start reading data (default is "A1")
        end_cell: Cell reference where to stop reading data (default is None, which means read until the end)
        preview_only: If True, only preview the data without loading the entire range

    returns:
        str: Data read from the specified range in the worksheet
    """
    filename = ensure_xlsx_extension(filename)
    is_readable, error_message = check_file_writeable(filename)
    if not is_readable:
        return f"Error: {error_message}"
    try:
        from mcp_excel_server.core.data import read_excel_range

        result = read_excel_range(
            filename, sheet_name, start_cell, end_cell, preview_only
        )
        if isinstance(result, dict) and "error" in result:
            return f"Error: {result['error']}"
        if not result:
            return "No data found in specified range"
        # Opcional: convertir a JSON si se requiere interoperabilidad
        # import json
        # return json.dumps(result)
        data_str = "\n".join([str(row) for row in result])
        return data_str
    except Exception as e:
        return f"Error: Failed to read data: {str(e)}"


async def write_data_to_excel(
    filename: str,
    sheet_name: str,
    data: List[List],
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
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Error: {error_message}"
    try:
        from mcp_excel_server.core.data import write_data

        result = write_data(filename, sheet_name, data, start_cell)
        if isinstance(result, dict) and "error" in result:
            return f"Error: {result['error']}"
        return result.get("message", "Data written successfully")
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Error: Failed to write data: {str(e)}"
