import os

# Import exceptions
from mcp_excel_server.exceptions.exceptions import (
    ValidationError,
    WorkbookError,
)
from mcp_excel_server.utils.file_utils import (
    check_file_writeable,
    ensure_xlsx_extension,
    is_path_in_allowed_directories,
)


async def create_workbook(filename: str) -> str:
    """Create a new Excel workbook.

    Args:
        filename: Name of the workbook to create (with or without .xlsx extension)
    """
    filename = ensure_xlsx_extension(filename)
    # verificar directorio
    is_valid, error_message = is_path_in_allowed_directories(filename)
    if not is_valid:
        return error_message
    # Check if the file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Error: {error_message}"
    try:
        from mcp_excel_server.core.workbook import (
            create_workbook as create_workbook_impl,
        )

        result = create_workbook_impl(filename)  # Eliminar el parámetro title
        return f"Created workbook at {filename}"
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to create workbook: {str(e)}"


async def create_worksheet(filename: str, sheet_name: str) -> str:
    """Create new worksheet in workbook.
    Args:
        filename: Path to the workbook file
        sheet_name: Name of the new worksheet to create
    """
    filename = ensure_xlsx_extension(filename)
    # Check if the file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Error: {error_message}"
    # Check if the file is in an allowed directory
    is_valid, error_message = is_path_in_allowed_directories(filename)
    if not is_valid:
        return error_message

    try:
        from mcp_excel_server.core.workbook import create_sheet as create_worksheet_impl

        result = create_worksheet_impl(filename, sheet_name)
        return result["message"]
    except (ValidationError, WorkbookError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to create worksheet: {str(e)}"
    except FileNotFoundError as e:
        return f"File not found: {str(e)}"


async def list_available_documents(directory: str = ".") -> str:
    """List all .xlsx files in the specified directory.

    Args:
        directory: Directory to search for Excel documents
    """
    try:
        if not os.path.exists(directory):
            return f"Directory {directory} does not exist"

        excels_files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]

        if not excels_files:
            return f"No Excel documents found in {directory}"

        result = f"Found {len(excels_files)} Excel documents in {directory}:\n"
        for file in excels_files:
            file_path = os.path.join(directory, file)
            size = os.path.getsize(file_path) / 1024  # KB
            result += f"- {file} ({size:.2f} KB)\n"

        return result
    except Exception as e:
        return f"Failed to list documents: {str(e)}"
