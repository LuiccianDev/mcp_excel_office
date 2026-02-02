"""Centralized exceptions package for MCP Excel Office.

This package provides all custom exceptions used across the mcp_excel module,
organized into core and tools specific exceptions.
"""

# Re-export all exceptions for convenient imports
from mcp_excel.exceptions.exception_core import (
    CellReferenceError,
    CoreError,
    DataError,
    FormulaError,
    InvalidCellReferenceError,
    InvalidDataError,
    PivotError,
    RangeError,
    SheetError,
    SheetExistsError,
    SheetNotFoundError,
    ValidationError,
    WorkbookError,
    WorkbookNotFoundError,
    WorksheetError,
)
from mcp_excel.exceptions.exception_tools import (
    CalculationError,
    ChartError,
    DataError as ToolsDataError,
    ExcelMCPError,
    FormattingError,
    PivotError as ToolsPivotError,
    SheetError as ToolsSheetError,
    ValidationError as ToolsValidationError,
    WorkbookError as ToolsWorkbookError,
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
