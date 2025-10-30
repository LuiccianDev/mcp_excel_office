class ExcelMCPError(Exception):
    """Base exception for Excel MCP errors.
    Used across src/mcp_excel/tools/*.py and server.py
    """

    pass


class WorkbookError(ExcelMCPError):
    """Raised when workbook operations fail.
    Used in tools/excel_tools.py, tools/content_tools.py
    """

    pass


class SheetError(ExcelMCPError):
    """Raised when sheet operations fail.
    Used in tools/excel_tools.py, tools/content_tools.py
    """

    pass


class DataError(ExcelMCPError):
    """Raised when data operations fail.
    Used in tools/content_tools.py, tools/db_tools.py
    """

    pass


class ValidationError(ExcelMCPError):
    """Raised when validation fails.
    Used in tools/excel_tools.py, tools/format_tools.py, config.py
    """

    pass


class FormattingError(ExcelMCPError):
    """Raised when formatting operations fail.
    Used in tools/format_tools.py
    """

    pass


class CalculationError(ExcelMCPError):
    """Raised when formula calculations fail.
    Used in tools/formulas_excel_tools.py
    """

    pass


class PivotError(ExcelMCPError):
    """Raised when pivot table operations fail.
    Used in tools/graphics_tools.py
    """

    pass


class ChartError(ExcelMCPError):
    """Raised when chart operations fail.
    Used in tools/graphics_tools.py
    """

    pass
