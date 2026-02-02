<div align="center">
  <h1>MCP Excel Office Server</h1>
  <p>
    <em>Powerful MCP server for programmatic Excel (.xlsx) manipulation and automation</em>
  </p>

  [![Python Version](https://img.shields.io/badge/python-3.11%2B-blue.svg)](https://www.python.org/downloads/)
  [![Code style: ruff](https://img.shields.io/badge/code%20style-ruff-000000.svg)](https://github.com/astral-sh/ruff)
  [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
  [![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)
  [![Type Checked: mypy](https://img.shields.io/badge/type%20checked-mypy-blue.svg)](https://mypy-lang.org/)
</div>

## Table of Contents

- [Key Features](#key-features)
- [Installation](#installation)
- [Deployment Modes](#deployment-modes)
- [Configuration](#configuration)
- [Available Tools](#available-tools)
- [Project Structure](#project-structure)
- [Testing](#testing)
- [Development](#development)
- [Contributing](#contributing)
- [License](#license)

## Key Features

- **Comprehensive Excel Operations**: Create, read, modify workbooks and worksheets with full data manipulation support
- **Advanced Formatting**: Apply styles, fonts, colors, borders, and cell formatting with precision
- **Data Visualization**: Generate charts and pivot tables programmatically
- **Formula Support**: Apply and validate Excel formulas with error handling
- **Security First**: File path validation, access control, and robust error handling
- **Multiple Deployment Modes**: DXT package, traditional MCP server, or standalone CLI
- **AI-Ready**: Optimized for AI assistant integration via Model Context Protocol

## Installation

### Prerequisites

- **Python 3.11+**: Modern Python with type hints support
- **UV Package Manager**: [Install UV](https://docs.astral.sh/uv/getting-started/installation/) (recommended)
- **Git**: For cloning the repository
- **Desktop Extensions (DXT)**: For creating .dxt packages for Claude Desktop [Install DXT](https://github.com/anthropics/dxt)

### Clone the Repository

```bash
git clone https://github.com/LuiccianDev/mcp_excel_office.git
cd mcp_excel_office
```

### Installation with UV (Recommended)

```bash
# Install production dependencies
uv sync

# Install with development dependencies
uv sync --dev

# Install all dependency groups (dev + test)
uv sync --all-groups
```

### Alternative: Installation with pip

```bash
# Install the package
pip install .

# Development installation (editable)
pip install -e ".[dev,test]"
```

### Build and Package

```bash
# Build distributable package
uv build

# Install from built package
uv pip install dist/mcp_excel-*.whl
```

## Deployment Modes

The MCP Excel Office Server supports three deployment modes:

### DXT Package Deployment

Best for: Integrated DXT ecosystem users who want seamless configuration management.

1. **Package the project**:

   ```bash
   dxt pack
   ```

2. **Configuration**: The DXT package automatically handles dependencies and provides user-friendly configuration through the manifest.json:
   - `directory`: Base directory for file operations

3. **Usage**: Once packaged, the tool integrates directly with DXT-compatible clients with automatic user configuration variable substitution.

4. **Server Configuration**: This project includes [manifest.json](./manifest.json) for building .dxt packages.

For more details see [DXT Package Documentation](https://github.com/anthropics/dxt).

### Traditional MCP Server

Best for: Standard MCP server deployments with existing MCP infrastructure.

Add to your MCP configuration file (e.g., Claude Desktop's `mcp_config.json`):

```json
{
  "mcpServers": {
    "mcp_excel": {
      "command": "uv",
      "args": ["run", "mcp_excel"],
      "env": {
        "DIRECTORY": "/path/to/your/files"
      }
    }
  }
}
```

**Alternative with CLI arguments:**

```json
{
  "mcpServers": {
    "mcp_excel": {
      "command": "uv",
      "args": [
        "run", "-m", "mcp_excel",
        "--directory", "/path/to/files"
      ]
    }
  }
}
```

### Standalone CLI

Best for: Direct command-line usage, scripting, and automation without MCP protocol overhead.

```bash
# Run with environment variables
export DIRECTORY="/path/to/your/files"
python -m mcp_excel

# Or run with command-line arguments
python -m mcp_excel --directory "/path/to/files"

# Using UV
uv run mcp_excel --help
```

### Docker Deployment

You can install and run the MCP Excel Office Server using Docker for an isolated and reproducible environment.

See [Docker.md](./Docker.md) for details and advanced configuration options.

## Configuration

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `DIRECTORY` | Base directory for file operations | Yes |
| `LOG_LEVEL` | Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL) | No (default: INFO) |

### Configuration Validation

The server validates all configuration on startup and provides clear error messages for:

- Missing required environment variables
- Invalid directory paths
- File access permissions

### Configuration Sources

Configuration is loaded with the following precedence (highest to lowest):

1. Command-line arguments
2. Environment variables
3. User config variables (`${user_config.*}`)
4. Default values

## Available Tools

The MCP Excel Office Server provides **19 tools** organized into 5 categories:

### Content Tools (2)

- `read_data_from_excel` - Read data from Excel worksheets
- `write_data_to_excel` - Write data to Excel worksheets

### Excel Tools (5)

- `create_excel_workbook` - Create new Excel workbooks
- `create_excel_worksheet` - Add worksheets to workbooks
- `list_excel_documents` - List Excel files in directory
- `copy_worksheet` - Copy worksheets within workbooks
- `delete_worksheet` - Delete worksheets from workbooks
- `rename_worksheet` - Rename worksheets
- `get_workbook_metadata` - Get workbook information

### Format Tools (6)

- `format_range_excel` - Apply comprehensive cell formatting
- `merge_cells` - Merge cell ranges
- `unmerge_cells` - Unmerge cell ranges
- `copy_range` - Copy cell ranges
- `delete_range` - Delete cell ranges
- `validate_excel_range` - Validate range references

### Formula Tools (2)

- `apply_formula_excel` - Apply Excel formulas
- `validate_formula_syntax` - Validate formula syntax

### Graphics Tools (2)

- `create_chart` - Create charts
- `create_pivot_table` - Create pivot tables

For detailed documentation of all tools, see [TOOLS.md](./TOOLS.md).

## Project Structure

```
mcp_excel_office/
├── src/mcp_excel/              # Main package
│   ├── __init__.py            # Package initialization
│   ├── __main__.py            # CLI entry point
│   ├── server.py              # MCP server implementation
│   ├── config.py              # Configuration management
│   ├── tools/                 # MCP tool implementations
│   │   ├── excel_tools.py     # Workbook/worksheet operations (7 tools)
│   │   ├── content_tools.py   # Data read/write operations (2 tools)
│   │   ├── format_tools.py    # Cell formatting (6 tools)
│   │   ├── formulas_excel_tools.py  # Formula operations (2 tools)
│   │   ├── graphics_tools.py  # Charts/pivot tables (2 tools)
│   │   └── register_tools.py  # Tool registration
│   ├── core/                  # Core functionality
│   │   ├── workbook.py        # Workbook operations
│   │   ├── formatting.py      # Cell formatting logic
│   │   ├── data.py            # Data read/write logic
│   │   ├── calculations.py    # Formula calculations
│   │   ├── chart.py           # Chart creation
│   │   └── pivot.py           # Pivot table creation
│   ├── utils/                 # Utility functions
│   │   ├── file_utils.py      # File validation and operations
│   │   ├── sheet_utils.py     # Sheet operations
│   │   ├── cell_utils.py      # Cell utilities
│   │   └── validation_utils.py # Validation utilities
│   └── exceptions/            # Custom exceptions
│       ├── exception_core.py  # Core exceptions
│       └── exception_tools.py # Tool-specific exceptions
├── tests/                     # Test suite
│   ├── conftest.py            # Pytest configuration and fixtures
│   ├── test_excel_tools.py    # Excel tools tests
│   ├── test_content_tools.py  # Content tools tests
│   ├── test_format_tools.py   # Format tools tests
│   ├── test_formulas_excel_tools.py  # Formula tools tests
│   ├── test_graphics_tools.py # Graphics tools tests
│   ├── test_workbook.py       # Workbook core tests
│   ├── test_data.py           # Data core tests
│   ├── test_file_utils.py     # File utilities tests
│   ├── test_file_security.py  # Security tests
│   └── ...                    # Additional test files
├── pyproject.toml             # Project configuration
├── manifest.json              # DXT package configuration
├── ruff.toml                  # Ruff configuration
├── mypy.ini                   # MyPy configuration
├── pytest.ini                 # Pytest configuration
├── .pre-commit-config.yaml    # Pre-commit hooks
├── TOOLS.md                   # Detailed tool documentation
├── README.md                  # This file
├── AGENTS.md                  # Development guidelines for AI agents
└── Docker.md                  # Docker deployment guide
```

## Testing

### Run All Tests

```bash
uv run pytest
```

### Run Specific Test File

```bash
uv run pytest tests/test_excel_tools.py
```

### Run Tests Matching Pattern

```bash
uv run pytest -k "test_name"
```

### List Tests Without Running

```bash
uv run pytest --co
```

### Verbose Output

```bash
uv run pytest -v --tb=short
```

### Test Coverage

The test suite includes coverage reporting. Coverage reports are generated in:

- `htmlcov/` - HTML coverage report
- `coverage.xml` - XML coverage report

## Development

### Development Setup

```bash
# Install development dependencies
uv sync --dev

# Install pre-commit hooks
uv run pre-commit install

# Run quality checks
uv run pre-commit run --all-files
```

### Code Quality Standards

```bash
# Format code with Ruff
uv run ruff format

# Check code style
uv run ruff check

# Auto-fix issues
uv run ruff check --fix

# Type checking with MyPy
uv run mypy src/

# Run all quality checks
uv run pre-commit run --all-files
```

### Development Commands Summary

| Command | Description |
|---------|-------------|
| `uv sync --dev` | Install development dependencies |
| `uv run ruff check` | Code style and quality checks |
| `uv run ruff format` | Format code |
| `uv run mypy src/` | Type checking (strict) |
| `uv run pytest` | Run test suite with coverage |
| `uv build` | Build distributable package |
| `dxt pack` | Create DXT package |

### Code Style Guidelines

- **Python Version**: 3.11+
- **Type Hints**: Strict (mypy with `disallow_untyped_defs`)
- **Line Length**: 88 characters
- **Formatting**: Ruff/Black-compatible
- **Imports**: Organized with Ruff's isort
- **Documentation**: Google-style docstrings

See [AGENTS.md](./AGENTS.md) for detailed development guidelines.

## Contributing

Contributions are welcome! Please follow these guidelines:

1. Fork the repository
2. Create a feature branch
3. Install development dependencies with `uv sync --dev`
4. Run code quality checks: `uv run ruff check && uv run mypy src/`
5. Ensure tests pass: `uv run pytest`
6. Submit a pull request

## Issues and Support

- **Bug Reports**: [Open an issue](https://github.com/LuiccianDev/mcp_excel_office/issues) with detailed reproduction steps
- **Feature Requests**: Describe your use case and proposed solution
- **Questions**: Check existing issues or start a discussion

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

<div align="center">
  <p><strong>MCP Excel Office Server</strong></p>
  <p>Empowering AI assistants with comprehensive Excel manipulation capabilities</p>
  <p>
    <a href="https://github.com/LuiccianDev/mcp_excel_office">GitHub</a> |
    <a href="https://modelcontextprotocol.io">MCP Protocol</a> |
    <a href="https://github.com/LuiccianDev/mcp_excel_office/blob/main/TOOLS.md">Tool Documentation</a>
  </p>
  <p><em>Created by LuiccianDev</em></p>
</div>
