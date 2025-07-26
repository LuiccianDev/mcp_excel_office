from typing import Any

from mcp_excel.core.workbook import get_workbook_info

# Import exceptions
from mcp_excel.exceptions.exceptions import (
    FormattingError,
    SheetError,
    ValidationError,
    WorkbookError,
)
from mcp_excel.utils.file_utils import ensure_xlsx_extension
from mcp_excel.utils.sheet_utils import (
    copy_sheet,
    delete_sheet,
    merge_range,
    rename_sheet,
    unmerge_range,
)
from mcp_excel.utils.validation_utils import (
    validate_range_in_sheet_operation as validate_range_impl,
)


async def format_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int = None,
    font_color: str = None,
    bg_color: str = None,
    border_style: str = None,
    border_color: str = None,
    number_format: str = None,
    alignment: str = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: dict[str, Any] = None,
    conditional_format: dict[str, Any] = None,
) -> str:
    """Apply formatting to a range of cells.
    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        start_cell: Starting cell reference (e.g., "A1")
        end_cell: Optional ending cell reference (e.g., "B2")
        bold: Whether to make text bold
        italic: Whether to make text italic
        underline: Whether to underline text
        font_size: Font size in points
        font_color: Font color (hex code)
        bg_color: Background color (hex code)
        border_style: Border style (e.g., "thin", "thick")
        border_color: Border color (hex code)
        number_format: Number format string (e.g., "0.00")
        alignment: Text alignment (e.g., "center", "left", "right")
        wrap_text: Whether to wrap text in the cell
        merge_cells: Whether to merge the specified range of cells
        protection: Cell protection options (e.g., {"locked": True})
        conditional_format: Conditional formatting rules

    """

    filename = ensure_xlsx_extension(filename)
    try:
        from mcp_excel.core.formatting import format_range as format_range_func

        result = format_range_func(
            filename=filename,
            sheet_name=sheet_name,
            start_cell=start_cell,
            end_cell=end_cell,
            bold=bold,
            italic=italic,
            underline=underline,
            font_size=font_size,
            font_color=font_color,
            bg_color=bg_color,
            border_style=border_style,
            border_color=border_color,
            number_format=number_format,
            alignment=alignment,
            wrap_text=wrap_text,
            merge_cells=merge_cells,
            protection=protection,
            conditional_format=conditional_format,
        )
        return "Range formatted successfully"
    except (ValidationError, FormattingError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to format range: {str(e)}"


async def copy_worksheet(filename: str, source_sheet: str, target_sheet: str) -> str:
    """Copy worksheet within workbook.
    Args:
        filename: Path to the Excel file
        source_sheet: Name of the worksheet to copy
        target_sheet: Name of the new worksheet
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result = copy_sheet(filename, source_sheet, target_sheet)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to copy worksheet: {str(e)}"


async def delete_worksheet(filename: str, sheet_name: str) -> str:
    """Delete worksheet from workbook.
    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet to delete
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result = delete_sheet(filename, sheet_name)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to delete worksheet: {str(e)}"


async def rename_worksheet(filename: str, old_name: str, new_name: str) -> str:
    """Rename worksheet in workbook.
    Args:
        filename: Path to the Excel file
        old_name: Current name of the worksheet
        new_name: New name for the worksheet
    """
    filename = ensure_xlsx_extension(filename)

    try:

        result = rename_sheet(filename, old_name, new_name)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to rename worksheet: {str(e)}"


async def get_workbook_metadata(filename: str, include_ranges: bool = False) -> str:
    """Get metadata about workbook including sheets, ranges, etc.
    Args:
        filename: Path to the Excel file
        include_ranges: Whether to include range information
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result = get_workbook_info(filename, include_ranges=include_ranges)
        return str(result)
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to get workbook metadata: {str(e)}"


async def merge_cells(
    filename: str, sheet_name: str, start_cell: str, end_cell: str
) -> str:
    """Merge a range of cells.

    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        start_cell: Starting cell reference (e.g., "A1")
        end_cell: Ending cell reference (e.g., "B2")
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result = merge_range(filename, sheet_name, start_cell, end_cell)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to merge cells: {str(e)}"


async def unmerge_cells(
    filename: str, sheet_name: str, start_cell: str, end_cell: str
) -> str:
    """Unmerge a range of cells.
    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        start_cell: Starting cell reference (e.g., "A1")
        end_cell: Ending cell reference (e.g., "B2")
    """
    filename = ensure_xlsx_extension(filename)

    try:

        result = unmerge_range(filename, sheet_name, start_cell, end_cell)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to unmerge cells: {str(e)}"


async def copy_range(
    filename: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str = None,
) -> str:
    """Copy a range of cells to another location.
    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        source_start: Starting cell reference of the source range (e.g., "A1")
        source_end: Ending cell reference of the source range (e.g., "B2")
        target_start: Starting cell reference of the target range (e.g., "C1")
        target_sheet: Optional name of the target worksheet
    """
    filename = ensure_xlsx_extension(filename)

    try:

        from mcp_excel.utils.sheet_utils import copy_range_operation

        result = copy_range_operation(
            filename, sheet_name, source_start, source_end, target_start, target_sheet
        )
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to copy range: {str(e)}"


async def delete_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up",
) -> str:
    """Delete a range of cells and shift remaining cells.

    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        start_cell: Starting cell reference (e.g., "A1")
        end_cell: Optional ending cell reference (e.g., "B2")
        shift_direction: Direction to shift cells ("up" or "left")
    """
    filename = ensure_xlsx_extension(filename)

    try:

        from mcp_excel.utils.sheet_utils import delete_range_operation

        result = delete_range_operation(
            filename, sheet_name, start_cell, end_cell, shift_direction
        )
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to delete range: {str(e)}"


async def validate_excel_range(
    filename: str, sheet_name: str, start_cell: str, end_cell: str = None
) -> str:
    """Validate if a range exists and is properly formatted.
    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        start_cell: Starting cell reference (e.g., "A1")
        end_cell: Optional ending cell reference (e.g., "B2")
    """
    filename = ensure_xlsx_extension(filename)

    try:

        range_str = start_cell if not end_cell else f"{start_cell}:{end_cell}"
        result = validate_range_impl(filename, sheet_name, range_str)
        return result["message"]
    except ValidationError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to validate range: {str(e)}"
