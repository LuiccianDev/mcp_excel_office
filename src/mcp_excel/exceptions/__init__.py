"""
Custom exceptions for MCP Excel Office Server.

This module defines a hierarchy of custom exceptions used throughout the MCP Excel
Office Server to provide clear, specific error handling for different types of
Excel operations and failures.

Exception Hierarchy:
- ExcelMCPError: Base exception for all Excel MCP operations
- WorkbookError: Workbook-related operation failures
- SheetError: Worksheet-related operation failures
- DataError: Data read/write operation failures
- ValidationError: Input validation failures
- FormattingError: Cell formatting operation failures
- CalculationError: Formula calculation failures
- PivotError: Pivot table operation failures
- ChartError: Chart creation operation failures
"""

from mcp_excel.exceptions.exceptions import (
    CalculationError,
    ChartError,
    DataError,
    ExcelMCPError,
    FormattingError,
    PivotError,
    SheetError,
    ValidationError,
    WorkbookError,
)


__all__ = [
    "ExcelMCPError",
    "WorkbookError",
    "SheetError",
    "DataError",
    "ValidationError",
    "FormattingError",
    "CalculationError",
    "PivotError",
    "ChartError",
]
