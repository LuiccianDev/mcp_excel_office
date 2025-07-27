
from mcp_excel.config import load_config

from mcp.server.fastmcp import FastMCP

from mcp_excel.tools import (
    content_tools,
    db_tools,
    excel_tools,
    format_tools,
    formulas_excel_tools,
    graphics_tools,
)

# This is a FastMCP server that serves as an Excel server.
# Initializes the server with the specified configuration and starts it.
config = load_config()

mcp = FastMCP(**config)


def register_tools():
    """
    Register tools to the FastMCP server.
    """
    # Database tools
    mcp.tool()(db_tools.fetch_and_insert_db_to_excel)
    mcp.tool()(db_tools.insert_calculated_data_to_db)

    # Content tools
    mcp.tool()(content_tools.read_data_from_excel)
    mcp.tool()(content_tools.write_data_to_excel)

    # Excel tools
    mcp.tool()(excel_tools.create_excel_workbook)
    mcp.tool()(excel_tools.create_excel_worksheet)
    mcp.tool()(excel_tools.list_excel_documents)

    # Formar tools

    mcp.tool()(format_tools.format_range)
    mcp.tool()(format_tools.copy_worksheet)
    mcp.tool()(format_tools.delete_worksheet)
    mcp.tool()(format_tools.rename_worksheet)
    mcp.tool()(format_tools.get_workbook_metadata)
    mcp.tool()(format_tools.merge_cells)
    mcp.tool()(format_tools.unmerge_cells)
    mcp.tool()(format_tools.copy_range)
    mcp.tool()(format_tools.delete_range)
    mcp.tool()(format_tools.validate_excel_range)

    # Formula tools

    mcp.tool()(formulas_excel_tools.validate_formula_syntax)
    mcp.tool()(formulas_excel_tools.apply_formula)

    # Graphics tools
    mcp.tool()(graphics_tools.create_chart)
    mcp.tool()(graphics_tools.create_pivot_table)


register_tools()


def run_server():
    """
    Run the FastMCP server.
    """
    mcp.run(transport="stdio")
    return mcp


if __name__ == "__main__":
    # Run the server
    run_server()
