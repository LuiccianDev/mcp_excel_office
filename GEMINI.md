# GEMINI.md - MCP Office Excel Server

## Project Overview

This project is a Python-based MCP (Model Context Protocol) server designed for programmatic manipulation of Excel (`.xlsx`) documents. It provides a set of tools to create, read, and modify Excel files, including operations for formatting, creating charts, and generating pivot tables.

The server is built using the `fastmcp` library and utilizes `openpyxl` for core Excel functionalities. The project is well-structured, with a clear separation of concerns between the server implementation, core logic, and utility functions. It also includes a comprehensive test suite built with `pytest`.

**Key Technologies:**

* **Python 3.11+**
* **MCP (Model Context Protocol)**
* **FastMCP** (for server implementation)
* **openpyxl** (for Excel manipulation)
* **uv** (for dependency management)
* **pytest** (for testing)
* **ruff** (for linting and formatting)

## Building and Running

The project uses `uv` for dependency management and running tasks.

### Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/LuiccianDev/mcp_excel_office.git
    cd mcp_excel_office
    ```

2. **Install dependencies using `uv`:**
    * For development (includes testing and linting tools):

        ```bash
        uv sync --dev
        ```

    * For production (only required dependencies):

        ```bash
        uv sync --production
        ```

### Running the Server

The server can be run in two ways:

1. **Directly via Python:**

    ```bash
    python src/mcp_excel/server.py
    ```

2. **Using the `mcp` CLI:**
    Add the following configuration to your `mcp_config.json` file:

    ```json
    {
        "mcpServers": {
            "officeExcel": {
                "command": "uv",
                "args": ["run", "mcp-office-excel"],
                "env": {
                    "DIRECTORY": "user/to/path"
                }
            }
        }
    }
    ```

### Testing

To run the test suite, use the following command:

```bash
uv run pytest
```

The project aims for a test coverage of at least 80%.

## Development Conventions

The project follows a strict set of development conventions to ensure code quality and maintainability.

* **Code Style:** Adheres to the **PEP 8** style guide.
* **Linting and Formatting:** Uses **ruff** for linting and formatting. The configuration can be found in the `ruff.toml` file.
* **Type Hinting:** All function signatures must include type hints.
* **Docstrings:** All public APIs must be documented with docstrings following the **Google style**.
* **Modularity:** The project is divided into modules with well-defined responsibilities.
* **Testing:** All new features should be accompanied by unit tests.

For more detailed information on the project's rules and structure, please refer to the documents in the `.windsurf/rules` directory.
