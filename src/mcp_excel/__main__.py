"""
Entry point for running mcp_excel as a module.

This module provides the command-line interface for the MCP Excel Office server,
supporting both MCP server mode and standalone CLI operations. It handles
environment variable configuration, command-line argument parsing, and provides
clear error messages for configuration issues.

The module supports multiple execution modes:
- MCP Server: `python -m mcp_excel server` or `uv run mcp_excel_office`
- File listing: `python -m mcp_excel list`
- Default behavior: Start MCP server when no command is specified
"""

import asyncio
import logging
import os
import sys
from typing import Annotated

import typer

from mcp_excel.server import run_server


# Configure logging for CLI operations
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


app = typer.Typer(
    name="mcp_excel",
    help="MCP Excel Office Server - Excel manipulation through Model Context Protocol",
    no_args_is_help=True,
)


def validate_environment_variables() -> tuple[str | None, str | None]:
    """
    Validate and retrieve environment variables for MCP Excel configuration.

    This function checks the POSTGRES_CONNECTION_STRING and DIRECTORY environment
    variables, validating their format and accessibility. It provides clear error
    messages for configuration issues.

    Returns:
        tuple[Optional[str], Optional[str]]: A tuple containing (postgres_connection_string, directory)

    Raises:
        typer.Exit: If environment variables are invalid or inaccessible
    """
    postgres_conn = os.environ.get("POSTGRES_CONNECTION_STRING")
    directory = os.environ.get("DIRECTORY")

    # Validate directory if provided
    if directory:
        try:
            # Create directory if it doesn't exist
            os.makedirs(directory, exist_ok=True)

            if not os.path.exists(directory):
                typer.echo(
                    f"Error: Cannot create directory specified in DIRECTORY environment variable: {directory}",
                    err=True,
                )
                logger.error(f"Directory creation failed: {directory}")
                raise typer.Exit(1)

            if not os.path.isdir(directory):
                typer.echo(
                    f"Error: Path specified in DIRECTORY environment variable is not a directory: {directory}",
                    err=True,
                )
                logger.error(f"Invalid directory path: {directory}")
                raise typer.Exit(1)

            if not os.access(directory, os.R_OK | os.W_OK):
                typer.echo(
                    f"Error: Directory specified in DIRECTORY environment variable is not readable/writable: {directory}",
                    err=True,
                )
                logger.error(f"Directory access denied: {directory}")
                raise typer.Exit(1)

        except OSError as e:
            typer.echo(
                f"Error: Cannot access directory specified in DIRECTORY environment variable: {directory} - {e}",
                err=True,
            )
            logger.error(f"Directory access error: {directory} - {e}")
            raise typer.Exit(1) from e

    # Validate PostgreSQL connection string if provided
    if postgres_conn:
        if not (
            postgres_conn.startswith("postgresql://")
            or postgres_conn.startswith("postgres://")
        ):
            typer.echo(
                "Error: POSTGRES_CONNECTION_STRING must start with 'postgresql://' or 'postgres://'",
                err=True,
            )
            logger.error("Invalid PostgreSQL connection string format")
            raise typer.Exit(1)

        # Basic validation of connection string components
        try:
            # Simple check for required components
            if "@" not in postgres_conn or "/" not in postgres_conn.split("@")[1]:
                typer.echo(
                    "Error: POSTGRES_CONNECTION_STRING appears to be malformed. Expected format: postgresql://user:password@host:port/database",
                    err=True,
                )
                logger.error("Malformed PostgreSQL connection string")
                raise typer.Exit(1)
        except (IndexError, AttributeError) as e:
            typer.echo(
                "Error: POSTGRES_CONNECTION_STRING appears to be malformed. Expected format: postgresql://user:password@host:port/database",
                err=True,
            )
            logger.error("Malformed PostgreSQL connection string")
            raise typer.Exit(1) from e

    return postgres_conn, directory


@app.command("server")  # type: ignore[misc]
def run_mcp_server(
    postgres: Annotated[
        str | None,
        typer.Option(
            "--postgres",
            help="PostgreSQL connection string (overrides POSTGRES_CONNECTION_STRING env var)",
        ),
    ] = None,
    path: Annotated[
        str | None,
        typer.Option(
            "--path",
            help="Base directory for file operations (overrides DIRECTORY env var)",
        ),
    ] = None,
) -> None:
    """
    Start the MCP Excel Office server.

    The server provides Excel manipulation capabilities through the Model Context Protocol.
    Configuration can be provided via command-line arguments or environment variables.
    Command-line arguments take precedence over environment variables.

    Environment Variables:
        POSTGRES_CONNECTION_STRING: PostgreSQL connection string for database operations
        DIRECTORY: Base directory for file operations (defaults to ./documents)

    Examples:
        python -m mcp_excel server
        python -m mcp_excel server --postgres "postgresql://user:pass@localhost/db"
        python -m mcp_excel server --path "/path/to/excel/files"
    """
    try:
        logger.info("Initializing MCP Excel Office Server...")

        # Get and validate environment variables
        env_postgres, env_directory = validate_environment_variables()

        # Command-line arguments override environment variables
        final_postgres = postgres or env_postgres
        final_directory = path or env_directory or "./documents"

        # Validate command-line provided PostgreSQL connection if given
        if postgres and not (
            postgres.startswith("postgresql://") or postgres.startswith("postgres://")
        ):
            typer.echo(
                "Error: --postgres argument must start with 'postgresql://' or 'postgres://'",
                err=True,
            )
            logger.error("Invalid PostgreSQL connection string from command line")
            raise typer.Exit(1)

        # Validate command-line provided directory if given
        if path:
            try:
                os.makedirs(path, exist_ok=True)
                if not os.access(path, os.R_OK | os.W_OK):
                    typer.echo(
                        f"Error: Directory specified in --path is not readable/writable: {path}",
                        err=True,
                    )
                    logger.error(f"Directory access denied: {path}")
                    raise typer.Exit(1)
            except OSError as e:
                typer.echo(
                    f"Error: Cannot create or access directory specified in --path: {path} - {e}",
                    err=True,
                )
                logger.error(f"Directory creation failed: {path} - {e}")
                raise typer.Exit(1) from e

        # Set environment variables for the server to use
        if final_postgres:
            os.environ["POSTGRES_CONNECTION_STRING"] = final_postgres
        if final_directory:
            os.environ["DIRECTORY"] = final_directory

        # Provide user feedback about configuration
        typer.echo("Starting MCP Excel Office Server...")
        typer.echo("Server Name: MCP Excel Office Server")

        if final_postgres:
            typer.echo("Database: Configured (PostgreSQL)")
            logger.info("PostgreSQL database configured")
        else:
            typer.echo("Database: Not configured (database tools will be unavailable)")
            logger.info("No database configuration - database tools disabled")

        typer.echo(f"File operations directory: {final_directory}")
        logger.info(f"File operations directory: {final_directory}")

        typer.echo("Server ready. Listening on stdio transport...")
        logger.info("MCP server starting on stdio transport")

        # Start the MCP server with proper error handling
        try:
            server = run_server()
            asyncio.run(server.run(transport="stdio"))
        except Exception as server_error:
            typer.echo(f"Error initializing MCP server: {server_error}", err=True)
            logger.error(f"MCP server initialization failed: {server_error}")
            raise typer.Exit(1) from server_error

    except KeyboardInterrupt:
        typer.echo("\nServer stopped by user.", err=True)
        logger.info("Server stopped by user interrupt")
        sys.exit(0)
    except typer.Exit:
        # Re-raise typer.Exit to preserve exit codes
        raise
    except Exception as e:
        typer.echo(f"Unexpected error starting server: {e}", err=True)
        logger.error(f"Unexpected server startup error: {e}")
        sys.exit(1)


@app.command("list")  # type: ignore[misc]
def list_excel_files(
    path: Annotated[
        str | None,
        typer.Option(
            "--path",
            help="Directory to scan for Excel files (overrides DIRECTORY env var)",
        ),
    ] = None,
) -> None:
    """
    List all Excel files in the specified directory.

    This command provides a standalone way to discover Excel files without starting the MCP server.
    """
    try:
        # Get directory from argument or environment variable
        _, env_directory = validate_environment_variables()
        target_directory = path or env_directory or "./documents"

        # Validate directory
        if not os.path.exists(target_directory):
            typer.echo(f"Error: Directory does not exist: {target_directory}", err=True)
            raise typer.Exit(1)

        if not os.path.isdir(target_directory):
            typer.echo(f"Error: Path is not a directory: {target_directory}", err=True)
            raise typer.Exit(1)

        # List Excel files
        excel_files = []
        try:
            for entry in os.scandir(target_directory):
                if entry.is_file() and entry.name.lower().endswith((".xlsx", ".xls")):
                    stat_info = entry.stat()
                    excel_files.append(
                        {
                            "name": entry.name,
                            "path": entry.path,
                            "size": stat_info.st_size,
                            "modified": stat_info.st_mtime,
                        }
                    )
        except OSError as e:
            typer.echo(f"Error scanning directory: {e}", err=True)
            raise typer.Exit(1) from e

        # Display results
        if not excel_files:
            typer.echo(f"No Excel files found in: {target_directory}")
        else:
            typer.echo(f"Found {len(excel_files)} Excel file(s) in: {target_directory}")
            for file_info in excel_files:
                size_mb = float(file_info["size"]) / (1024 * 1024)  # type: ignore[arg-type]
                typer.echo(f"  - {file_info['name']} ({size_mb:.2f} MB)")

    except Exception as e:
        typer.echo(f"Error listing Excel files: {e}", err=True)
        sys.exit(1)


def main() -> None:
    """
    Main entry point for the mcp_excel module.

    This function serves as the primary entry point when the module is executed
    with `python -m mcp_excel` or through the `mcp_excel_office` script.

    It supports multiple execution patterns:
    - `python -m mcp_excel` (defaults to server mode)
    - `python -m mcp_excel server` (explicit server mode)
    - `python -m mcp_excel list` (list Excel files)
    - `uv run mcp_excel_office` (entry point script)

    The function provides proper error handling and logging for all execution modes.
    """
    try:
        logger.info("MCP Excel Office CLI starting...")

        # If no arguments provided, default to server mode for backward compatibility
        if len(sys.argv) == 1:
            # Default behavior: start the MCP server
            typer.echo(
                "No command specified. Starting MCP server (use --help for options)..."
            )
            logger.info("Defaulting to server mode - no command specified")
            run_mcp_server()
        else:
            # Use Typer CLI for command parsing
            logger.info(f"Processing command: {' '.join(sys.argv[1:])}")
            app()

    except KeyboardInterrupt:
        typer.echo("\nOperation cancelled by user.", err=True)
        logger.info("Operation cancelled by user interrupt")
        sys.exit(0)
    except typer.Exit as e:
        # Re-raise typer.Exit to preserve exit codes
        logger.info(f"CLI exited with code: {e.exit_code}")
        raise
    except Exception as e:
        typer.echo(f"Unexpected error in MCP Excel CLI: {e}", err=True)
        logger.error(f"Unexpected CLI error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
