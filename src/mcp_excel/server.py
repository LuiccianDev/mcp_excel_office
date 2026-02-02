"""
Main server implementation for the MCP Excel Office server.

This module provides the core MCP server functionality for Excel manipulation operations.
It initializes the FastMCP server, registers all available tools, and handles the server
lifecycle with proper error handling and configuration management.
"""

import logging
import sys

from mcp.server.fastmcp import FastMCP

from mcp_excel.config import ConfigurationError, config_manager
from mcp_excel.tools.register_tools import register_all_tools


# Configure logging for server operations
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def run_server() -> FastMCP:
    """
    Initialize and configure the FastMCP server for Excel operations.

    This function creates a new FastMCP server instance, validates the configuration,
    registers all available Excel manipulation tools, and prepares the server for
    client connections through the MCP protocol.

    Returns:
        FastMCP: Configured server instance ready to handle MCP requests

    Raises:
        ConfigurationError: If server configuration is invalid
        RuntimeError: If tool registration fails
    """
    try:
        # Load and validate configuration
        config = config_manager.config

        # Log configuration summary
        logger.info("Server configuration:")
        logger.info(f"  File Directory: {config.file_config.directory}")
        logger.info(f"  Log Level: {config.log_level}")

        # Update logging level based on configuration
        logging.getLogger().setLevel(getattr(logging, config.log_level))

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

    except ConfigurationError as e:
        logger.error(f"Server configuration error: {e}")
        raise
    except Exception as e:
        logger.error(f"Server initialization failed: {e}")
        raise RuntimeError(f"Server initialization failed: {e}") from e


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
