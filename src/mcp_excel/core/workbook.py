"""Core functionality for Excel workbook operations.

This module provides functions to create, modify, and inspect Excel workbooks
following the Model Context Protocol (MCP) standards.
"""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, TypedDict, Union

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# Type aliases
WorkbookResult = Dict[str, Any]
SheetName = str


class WorkbookInfo(TypedDict):
    """Type definition for workbook information dictionary."""
    filename: str
    sheets: List[SheetName]
    size: int
    modified: float
    used_ranges: Optional[Dict[SheetName, str]]


class WorkbookError(Exception):
    """Base exception for workbook-related errors."""
    pass


class WorkbookNotFoundError(WorkbookError):
    """Raised when a workbook file is not found."""
    pass


class SheetExistsError(WorkbookError):
    """Raised when trying to create a sheet that already exists."""
    pass


def create_workbook(
    filename: str,
    sheet_name: str = "Sheet1",
    data_only: bool = False
) -> WorkbookResult:
    """Create a new Excel workbook with an optional custom sheet name.

    Args:
        filename: Path where the workbook will be saved.
        sheet_name: Name for the initial worksheet. Defaults to "Sheet1".
        data_only: Whether to save only values, not formulas. Defaults to False.

    Returns:
        Dictionary with operation status and details:
        - On success: {"message": str, "active_sheet": str}
        - On failure: {"error": str}

    Raises:
        ValueError: If sheet_name is empty or contains invalid characters.
        PermissionError: If the file cannot be written to the specified location.
    """
    if not sheet_name.strip():
        raise ValueError("Sheet name cannot be empty")
    
    path = Path(filename).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    
    if not _is_valid_sheet_name(sheet_name):
        raise ValueError(f"Invalid sheet name: {sheet_name}")
    
    wb = Workbook()
    try:
        # Remove default sheet and create a new one with the specified name
        wb.remove(wb.active)
        wb.create_sheet(sheet_name)
        
        wb.save(str(path), data_only=data_only)
        return {
            "message": f"Created workbook: {path}",
            "active_sheet": sheet_name,
            "filepath": str(path)
        }
    except Exception as e:
        if isinstance(e, PermissionError):
            raise PermissionError(f"Cannot write to {path}: {e}") from e
        raise WorkbookError(f"Failed to create workbook: {e}") from e
    finally:
        wb.close()


def get_or_create_workbook(filename: str, read_only: bool = False) -> Workbook:
    """Get an existing workbook or create a new one if it doesn't exist.
    
    Args:
        filename: Path to the Excel file.
        read_only: Whether to open the workbook in read-only mode.
        
    Returns:
        An openpyxl Workbook instance.
        
    Raises:
        WorkbookError: If there's an error loading or creating the workbook.
    """
    path = Path(filename).resolve()
    
    try:
        if path.exists():
            return load_workbook(str(path), read_only=read_only)
            
        # Create new workbook if it doesn't exist
        wb = Workbook()
        wb.save(str(path))
        return wb
        
    except Exception as e:
        raise WorkbookError(f"Failed to get or create workbook {path}: {e}") from e


def create_sheet(filename: str, sheet_name: str) -> WorkbookResult:
    """Create a new worksheet in the workbook if it doesn't exist.
    
    Args:
        filename: Path to the Excel file.
        sheet_name: Name for the new worksheet.
        
    Returns:
        Dictionary with operation status and details:
        - On success: {"message": str}
        - On failure: {"error": str}
        
    Raises:
        ValueError: If sheet_name is invalid.
        SheetExistsError: If a sheet with the same name already exists.
        WorkbookError: For other workbook-related errors.
    """
    if not _is_valid_sheet_name(sheet_name):
        raise ValueError(f"Invalid sheet name: {sheet_name}")
    
    wb = None
    try:
        wb = load_workbook(filename)
        
        if sheet_name in wb.sheetnames:
            raise SheetExistsError(f"Sheet '{sheet_name}' already exists")
            
        wb.create_sheet(sheet_name)
        wb.save(filename)
        return {"message": f"Sheet '{sheet_name}' created successfully"}
        
    except Exception as e:
        if not isinstance(e, (ValueError, SheetExistsError)):
            raise WorkbookError(f"Failed to create sheet: {e}") from e
        raise
    finally:
        if wb is not None:
            wb.close()


def get_workbook_info(
    filename: str,
    include_ranges: bool = False
) -> Union[WorkbookInfo, Dict[str, str]]:
    """Get metadata about workbook including sheets, ranges, etc.
    
    Args:
        filename: Path to the Excel file.
        include_ranges: Whether to include used ranges for each sheet.
        
    Returns:
        Dictionary containing workbook information with the following keys:
        - filename: Name of the file
        - sheets: List of sheet names
        - size: File size in bytes
        - modified: Last modification timestamp
        - used_ranges: (Optional) Dictionary mapping sheet names to their used ranges
        
    Raises:
        FileNotFoundError: If the specified file doesn't exist.
        WorkbookError: For other workbook-related errors.
    """
    path = Path(filename).resolve()
    
    if not path.exists():
        raise FileNotFoundError(f"File not found: {filename}")
    
    wb = None
    try:
        wb = load_workbook(str(path), read_only=True, data_only=True)
        info: WorkbookInfo = {
            "filename": path.name,
            "sheets": wb.sheetnames,
            "size": path.stat().st_size,
            "modified": path.stat().st_mtime,
            "used_ranges": None
        }
        
        if include_ranges:
            ranges: Dict[str, str] = {}
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if ws.max_row > 0 and ws.max_column > 0:
                    ranges[sheet_name] = (
                        f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
                    )
            info["used_ranges"] = ranges
            
        return info
        
    except Exception as e:
        raise WorkbookError(f"Failed to get workbook info: {e}") from e
    finally:
        if wb is not None:
            wb.close()


def _is_valid_sheet_name(name: str) -> bool:
    """Check if a sheet name is valid according to Excel's rules.
    
    Args:
        name: The sheet name to validate.
        
    Returns:
        bool: True if the name is valid, False otherwise.
    """
    if not name or len(name) > 31:
        return False
    
    # Check for invalid characters
    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
    if any(char in name for char in invalid_chars):
        return False
        
    # Sheet name cannot start or end with an apostrophe
    if name.startswith("'") or name.endswith("'"):
        return False
        
    # Sheet name cannot be empty after trimming
    if not name.strip():
        return False
        
    return True
