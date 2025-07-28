"""Custom exceptions for the core module."""


class CoreError(Exception):
    """Base exception for all core module errors."""

    pass


class FormulaError(CoreError):
    """Raised when there's an error with formula operations."""

    pass


class DataError(CoreError):
    """Base exception for data-related errors."""

    pass


class InvalidDataError(DataError):
    """Raised when invalid data is provided."""

    pass


class CellReferenceError(DataError):
    """Base exception for cell reference related errors."""

    pass


class InvalidCellReferenceError(CellReferenceError):
    """Raised when a cell reference is invalid."""

    pass


class RangeError(DataError):
    """Raised when there's an error with cell ranges."""

    pass


class WorkbookError(CoreError):
    """Base exception for all workbook-related errors."""

    pass


class WorkbookNotFoundError(WorkbookError):
    """Raised when a workbook file is not found."""

    pass


class SheetError(WorkbookError):
    """Base exception for sheet-related errors."""

    pass


class SheetNotFoundError(SheetError):
    """Raised when a specified sheet is not found."""

    pass


class SheetExistsError(SheetError):
    """Raised when trying to create a sheet that already exists."""

    pass


class WorksheetError(CoreError):
    """Raised when a worksheet operation fails."""

    pass


class ValidationError(CoreError):
    """Raised when validation fails."""

    pass
