"""Main server implementation for the Git Commit Generator MCP server."""

from mcp.server.fastmcp import FastMCP

from mcp_excel.tools import register_all_tools


# This is a FastMCP server that serves as an Excel server.
# Initializes the server with the specified configuration and starts it.


def run_server() -> FastMCP:
    """
    Run the FastMCP server.
    """

    mcp = FastMCP("Mcp Excel Office")

    register_all_tools(mcp)

    return mcp


if __name__ == "__main__":
    import asyncio
    asyncio.run(run_server().run(transport="stdio"))
