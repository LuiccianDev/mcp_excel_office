"""Centralized exceptions package for MCP Excel Office.

This package provides all custom exceptions used across the mcp_excel module,
organized into core and tools specific exceptions.
"""

# Re-export all exceptions for convenient imports
from mcp_excel.exceptions.exception_core import (
    CoreError,
    WorkbookError,
    WorkbookNotFoundError,
    SheetError,
    SheetNotFoundError,
    SheetExistsError,
    WorksheetError,
    DataError,
    InvalidDataError,
    CellReferenceError,
    InvalidCellReferenceError,
    RangeError,
    FormulaError,
    PivotError,
    ValidationError,
)

from mcp_excel.exceptions.exception_tools import (
    ExcelMCPError,
    WorkbookError as ToolsWorkbookError,
    SheetError as ToolsSheetError,
    DataError as ToolsDataError,
    ValidationError as ToolsValidationError,
    FormattingError,
    CalculationError,
    PivotError as ToolsPivotError,
    ChartError,
)

__all__ = [
    # Core exceptions
    "CoreError",
    "WorkbookError",
    "WorkbookNotFoundError",
    "SheetError",
    "SheetNotFoundError",
    "SheetExistsError",
    "WorksheetError",
    "DataError",
    "InvalidDataError",
    "CellReferenceError",
    "InvalidCellReferenceError",
    "RangeError",
    "FormulaError",
    "PivotError",
    "ValidationError",
    # Tools exceptions
    "ExcelMCPError",
    "ToolsWorkbookError",
    "ToolsSheetError",
    "ToolsDataError",
    "ToolsValidationError",
    "FormattingError",
    "CalculationError",
    "ToolsPivotError",
    "ChartError",
]
