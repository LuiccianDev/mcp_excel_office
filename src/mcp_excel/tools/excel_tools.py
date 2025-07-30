# Import exceptions
from typing import Any

from mcp_excel.core.workbook import create_sheet, create_workbook
from mcp_excel.exceptions.exceptions import ValidationError, WorkbookError
from mcp_excel.utils.file_utils import (
    ensure_xlsx_extension,
    list_excel_files_in_directory,
    resolve_safe_path,
    validate_directory_access,
    validate_file_access,
)


# * Create a new Excel workbook
# @validate_file_access()
async def create_excel_workbook(filename: str) -> dict[str, Any]:
    """Create a new Excel workbook.

    Args:
        filename: Name of the workbook to create (with or without .xlsx extension)
    """
    try:
        filename = resolve_safe_path(filename)
        result: dict[str, Any] = create_workbook(filename)
        return result
    except WorkbookError as e:
        return {"error": f"Error: {str(e)}"}
    except Exception as e:
        return {"error": f"Failed to create workbook: {str(e)}"}


# * Create new worksheet in workbook
#! No borrar el type: ignore[misc] que se encuentra en la linea siguiente en caso contraio eliminar disallow_untyped_decorators = true de pyproject.toml
@validate_file_access("filename")  # type: ignore[misc]
async def create_excel_worksheet(filename: str, sheet_name: str) -> dict[str, Any]:
    """Create new worksheet in workbook.
    Args:
        filename: Path to the workbook file
        sheet_name: Name of the new worksheet to create
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = create_sheet(filename, sheet_name)
        return result
    except (ValidationError, WorkbookError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to create worksheet: {str(e)}"}


#! No borrar el type: ignore[misc] que se encuentra en la linea siguiente en caso contraio eliminar disallow_untyped_decorators = true de pyproject.toml
@validate_directory_access("directory")  # type: ignore[misc]
async def list_excel_documents(directory: str = ".") -> dict[str, Any]:
    """
    List all .xlsx files in the specified directory and return info as a dict.

    Args:
        directory: Directory to search for Excel documents
    Returns:
        dict: {"status": "success", "count": int, "directory": str, "files": list[dict]} or error dict
    """
    try:
        files = list_excel_files_in_directory(directory)
        return {
            "status": "success",
            "count": len(files),
            "directory": directory,
            "files": files,
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to list documents: {str(e)}",
            "directory": directory,
        }
