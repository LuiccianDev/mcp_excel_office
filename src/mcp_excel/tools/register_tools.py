"""
Tool registration module for MCP Excel Office Server.

This module handles the registration of all available Excel manipulation tools
with the FastMCP server. It provides a centralized location for tool management
and ensures consistent registration patterns across all tool categories.

Tool Categories:
- Database Tools: PostgreSQL integration for data import/export
- Content Tools: Data read/write operations for Excel files
- Excel Tools: Basic workbook and worksheet operations
- Format Tools: Cell styling, formatting, and layout operations
- Formula Tools: Excel formula application and validation
- Graphics Tools: Charts, pivot tables, and visual elements
"""

import logging

from mcp.server.fastmcp import FastMCP

from mcp_excel.tools import (
    content_tools,
    db_tools,
    excel_tools,
    format_tools,
    formulas_excel_tools,
    graphics_tools,
)


logger = logging.getLogger(__name__)


def register_database_tools(mcp: FastMCP) -> list[str]:
    """
    Register database integration tools with the MCP server.

    Args:
        mcp: FastMCP server instance

    Returns:
        List[str]: Names of successfully registered database tools

    Raises:
        RuntimeError: If tool registration fails
    """
    tools = []
    try:
        # Database tools for PostgreSQL integration
        mcp.tool()(db_tools.fetch_and_insert_db_to_excel)
        tools.append("fetch_and_insert_db_to_excel")

        mcp.tool()(db_tools.insert_calculated_data_to_db)
        tools.append("insert_calculated_data_to_db")

        logger.info(f"Registered {len(tools)} database tools")
        return tools

    except Exception as e:
        logger.error(f"Failed to register database tools: {e}")
        raise RuntimeError(f"Database tool registration failed: {e}") from e


def register_content_tools(mcp: FastMCP) -> list[str]:
    """
    Register content manipulation tools with the MCP server.

    Args:
        mcp: FastMCP server instance

    Returns:
        List[str]: Names of successfully registered content tools

    Raises:
        RuntimeError: If tool registration fails
    """
    tools = []
    try:
        # Content tools for data read/write operations
        mcp.tool()(content_tools.read_data_from_excel)
        tools.append("read_data_from_excel")

        mcp.tool()(content_tools.write_data_to_excel)
        tools.append("write_data_to_excel")

        logger.info(f"Registered {len(tools)} content tools")
        return tools

    except Exception as e:
        logger.error(f"Failed to register content tools: {e}")
        raise RuntimeError(f"Content tool registration failed: {e}") from e


def register_excel_tools(mcp: FastMCP) -> list[str]:
    """
    Register basic Excel workbook and worksheet tools with the MCP server.

    Args:
        mcp: FastMCP server instance

    Returns:
        List[str]: Names of successfully registered Excel tools

    Raises:
        RuntimeError: If tool registration fails
    """
    tools = []
    try:
        # Excel tools for basic workbook operations
        mcp.tool()(excel_tools.create_excel_workbook)
        tools.append("create_excel_workbook")

        mcp.tool()(excel_tools.create_excel_worksheet)
        tools.append("create_excel_worksheet")

        mcp.tool()(excel_tools.list_excel_documents)
        tools.append("list_excel_documents")

        logger.info(f"Registered {len(tools)} Excel tools")
        return tools

    except Exception as e:
        logger.error(f"Failed to register Excel tools: {e}")
        raise RuntimeError(f"Excel tool registration failed: {e}") from e


def register_format_tools(mcp: FastMCP) -> list[str]:
    """
    Register formatting and styling tools with the MCP server.

    Args:
        mcp: FastMCP server instance

    Returns:
        List[str]: Names of successfully registered format tools

    Raises:
        RuntimeError: If tool registration fails
    """
    tools = []
    try:
        # Format tools for cell styling and layout
        mcp.tool()(format_tools.format_range_excel)
        tools.append("format_range_excel")

        mcp.tool()(format_tools.copy_worksheet)
        tools.append("copy_worksheet")

        mcp.tool()(format_tools.delete_worksheet)
        tools.append("delete_worksheet")

        mcp.tool()(format_tools.rename_worksheet)
        tools.append("rename_worksheet")

        mcp.tool()(format_tools.get_workbook_metadata)
        tools.append("get_workbook_metadata")

        mcp.tool()(format_tools.merge_cells)
        tools.append("merge_cells")

        mcp.tool()(format_tools.unmerge_cells)
        tools.append("unmerge_cells")

        mcp.tool()(format_tools.copy_range)
        tools.append("copy_range")

        mcp.tool()(format_tools.delete_range)
        tools.append("delete_range")

        mcp.tool()(format_tools.validate_excel_range)
        tools.append("validate_excel_range")

        logger.info(f"Registered {len(tools)} format tools")
        return tools

    except Exception as e:
        logger.error(f"Failed to register format tools: {e}")
        raise RuntimeError(f"Format tool registration failed: {e}") from e


def register_formula_tools(mcp: FastMCP) -> list[str]:
    """
    Register formula and calculation tools with the MCP server.

    Args:
        mcp: FastMCP server instance

    Returns:
        List[str]: Names of successfully registered formula tools

    Raises:
        RuntimeError: If tool registration fails
    """
    tools = []
    try:
        # Formula tools for Excel calculations
        mcp.tool()(formulas_excel_tools.validate_formula_syntax)
        tools.append("validate_formula_syntax")

        mcp.tool()(formulas_excel_tools.apply_formula_excel)
        tools.append("apply_formula_excel")

        logger.info(f"Registered {len(tools)} formula tools")
        return tools

    except Exception as e:
        logger.error(f"Failed to register formula tools: {e}")
        raise RuntimeError(f"Formula tool registration failed: {e}") from e


def register_graphics_tools(mcp: FastMCP) -> list[str]:
    """
    Register graphics and visualization tools with the MCP server.

    Args:
        mcp: FastMCP server instance

    Returns:
        List[str]: Names of successfully registered graphics tools

    Raises:
        RuntimeError: If tool registration fails
    """
    tools = []
    try:
        # Graphics tools for charts and pivot tables
        mcp.tool()(graphics_tools.create_chart)
        tools.append("create_chart")

        mcp.tool()(graphics_tools.create_pivot_table)
        tools.append("create_pivot_table")

        logger.info(f"Registered {len(tools)} graphics tools")
        return tools

    except Exception as e:
        logger.error(f"Failed to register graphics tools: {e}")
        raise RuntimeError(f"Graphics tool registration failed: {e}") from e


def register_all_tools(mcp: FastMCP) -> None:
    """
    Register all available Excel manipulation tools with the FastMCP server.

    This function coordinates the registration of all tool categories and provides
    comprehensive error handling and logging. It ensures that tool registration
    failures are properly reported and handled.

    Args:
        mcp: FastMCP server instance to register tools with

    Raises:
        RuntimeError: If any tool category registration fails
    """
    logger.info("Starting tool registration process...")

    registered_tools = []
    registration_functions = [
        ("Database", register_database_tools),
        ("Content", register_content_tools),
        ("Excel", register_excel_tools),
        ("Format", register_format_tools),
        ("Formula", register_formula_tools),
        ("Graphics", register_graphics_tools),
    ]

    for category_name, register_func in registration_functions:
        try:
            tools = register_func(mcp)
            registered_tools.extend(tools)
            logger.info(
                f"Successfully registered {category_name} tools: {', '.join(tools)}"
            )

        except Exception as e:
            logger.error(f"Failed to register {category_name} tools: {e}")
            raise RuntimeError(f"{category_name} tool registration failed: {e}") from e

    logger.info(
        f"Tool registration completed successfully. Total tools registered: {len(registered_tools)}"
    )
    logger.info(f"Registered tools: {', '.join(registered_tools)}")

    # Verify all tools are properly registered
    if len(registered_tools) == 0:
        raise RuntimeError("No tools were successfully registered")

    logger.info("All Excel manipulation tools are ready for use")
