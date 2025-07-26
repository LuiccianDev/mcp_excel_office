from atexit import register
from http import server
import os
from mcp_excel_server.tools import (
    content_tools,
    excel_tools,
    format_tools,
    formulas_excel_tools,
    graphics_tools,
    db_tools
    )

from mcp.server.fastmcp import FastMCP

# This is a FastMCP server that serves as an Excel server.
# Initializes the server with the specified configuration and starts it.

mcp = FastMCP(
    name="ExcelServer",
    description="Excel server for processing Excel files.",
    version="0.1.0",
    author="LuiccianDev",
    )

def register_tools():
    """
    Register tools to the FastMCP server.
    """
    # Database tools
    mcp.tool()(db_tools.fetch_and_insert_db_data)
    mcp.tool()(db_tools.insert_calculated_data)
    
    
    # Content tools
    mcp.tool()(content_tools.read_data_from_excel)
    mcp.tool()(content_tools.write_data_to_excel)

    # Excel tools
    mcp.tool()(excel_tools.create_workbook)
    mcp.tool()(excel_tools.create_worksheet)
    mcp.tool()(excel_tools.list_available_documents)
    
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
    mcp.run(transport='stdio')
    return mcp

if __name__ == "__main__":
    # Run the server
    run_server()