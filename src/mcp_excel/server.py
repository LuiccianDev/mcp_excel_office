"""
Main server implementation for the MCP Excel Office server.

This module provides the core MCP server functionality for Excel manipulation operations.
It initializes the FastMCP server, registers all available tools, and handles the server
lifecycle with proper error handling and configuration management.
"""

import logging
import os
import sys

from mcp.server.fastmcp import FastMCP

from mcp_excel.tools.register_tools import register_all_tools


# Configure logging for server operations
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def validate_server_configuration() -> tuple[str | None, str | None]:
    """
    Validate server configuration from environment variables.

    Returns:
        tuple[Optional[str], Optional[str]]: A tuple containing (postgres_connection, directory)

    Raises:
        ValueError: If configuration is invalid or required settings are missing
    """
    postgres_conn = os.environ.get("POSTGRES_CONNECTION_STRING")
    directory = os.environ.get("DIRECTORY", "./documents")

    # Validate PostgreSQL connection string format if provided
    if postgres_conn:
        if not (
            postgres_conn.startswith("postgresql://")
            or postgres_conn.startswith("postgres://")
        ):
            raise ValueError(
                "Invalid POSTGRES_CONNECTION_STRING format. Must start with 'postgresql://' or 'postgres://'"
            )
        logger.info("PostgreSQL connection configured")
    else:
        logger.info(
            "PostgreSQL connection not configured - database tools will be unavailable"
        )

    # Validate and create directory if needed
    if directory:
        try:
            os.makedirs(directory, exist_ok=True)
            if not os.access(directory, os.R_OK | os.W_OK):
                raise ValueError(f"Directory is not readable/writable: {directory}")
            logger.info(f"File operations directory configured: {directory}")
        except OSError as e:
            raise ValueError(
                f"Cannot create or access directory '{directory}': {e}"
            ) from e

    return postgres_conn, directory


def run_server() -> FastMCP:
    """
    Initialize and configure the FastMCP server for Excel operations.

    This function creates a new FastMCP server instance, validates the configuration,
    registers all available Excel manipulation tools, and prepares the server for
    client connections through the MCP protocol.

    Returns:
        FastMCP: Configured server instance ready to handle MCP requests

    Raises:
        ValueError: If server configuration is invalid
        RuntimeError: If tool registration fails
    """
    try:
        # Validate configuration before starting server
        postgres_conn, directory = validate_server_configuration()

        # Initialize FastMCP server with proper name
        mcp = FastMCP("MCP Excel Office Server")

        # Register all available tools with error handling
        try:
            register_all_tools(mcp)
            logger.info("All Excel tools registered successfully")
        except Exception as e:
            logger.error(f"Failed to register tools: {e}")
            raise RuntimeError(f"Tool registration failed: {e}") from e

        logger.info("MCP Excel Office Server initialized successfully")
        return mcp

    except Exception as e:
        logger.error(f"Server initialization failed: {e}")
        raise


if __name__ == "__main__":
    """
    Entry point for running the server directly.

    This allows the server to be started with: python -m mcp_excel.server
    """
    try:
        import asyncio

        logger.info("Starting MCP Excel Office Server...")
        server = run_server()
        asyncio.run(server.run(transport="stdio"))

    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Server startup failed: {e}")
        sys.exit(1)
