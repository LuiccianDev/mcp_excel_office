from mcp.server.fastmcp import FastMCP

from mcp_excel.config import load_config
from mcp_excel.tools.register_tools import register_all_tools

# This is a FastMCP server that serves as an Excel server.
# Initializes the server with the specified configuration and starts it.


def run_server() -> FastMCP:
    """
    Run the FastMCP server.
    """
    config = load_config()
    mcp = FastMCP(**config)

    register_all_tools(mcp)

    return mcp
