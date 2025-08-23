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

## 📖 Description

A comprehensive MCP (Model Context Protocol) server that provides AI assistants with powerful Excel manipulation capabilities. This server enables programmatic creation, modification, and management of Excel files through standardized MCP tools, supporting data operations, formatting, formulas, charts, and PostgreSQL database integration.

## 📋 Table of Contents

- [✨ Key Features](#-key-features)
- [🚀 Installation](#-installation)
- [⚙️ Deployment Modes](#️-deployment-modes)
  - [DXT Package Deployment](#dxt-package-deployment)
  - [Traditional MCP Server](#traditional-mcp-server)
  - [Standalone CLI](#standalone-cli)
- [🔧 Configuration](#-configuration)
- [🚀 Quick Start](#-quick-start)
- [📚 Available Tools](#-available-tools)
- [🧪 Testing](#-testing)
- [🧩 Project Structure](#-project-structure)
- [🔧 Development](#-development)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)

## ✨ Key Features

- **📊 Comprehensive Excel Operations**: Create, read, modify workbooks and worksheets with full data manipulation support
- **🎨 Advanced Formatting**: Apply styles, fonts, colors, borders, and cell formatting with precision
- **📈 Data Visualization**: Generate charts, pivot tables, and graphics programmatically
- **🗄️ Database Integration**: Direct PostgreSQL integration for seamless data import/export
- **⚡ Formula Support**: Apply and validate Excel formulas with error handling
- **🔒 Security First**: File path validation, access control, and robust error handling
- **🚀 Multiple Deployment Modes**: DXT package, traditional MCP server, or standalone CLI
- **🤖 AI-Ready**: Optimized for AI assistant integration via Model Context Protocol

## 🚀 Installation

### 📋 Prerequisites

- **Python 3.11+**: Modern Python with type hints support
- **UV Package Manager**: [Install UV](https://docs.astral.sh/uv/getting-started/installation/) (recommended) or use pip
- **Git**: For cloning the repository

### 🔄 Clone the Repository

```bash
git clone https://github.com/LuiccianDev/mcp_excel_office.git
cd mcp_excel_office
```

### ⚡ Installation with UV (Recommended)

```bash
# Install production dependencies
uv sync

# Install with development dependencies
uv sync --dev

# Install all dependency groups (dev + test)
uv sync --all-groups
```

### 🐍 Alternative: Installation with pip

```bash
# Install the package
pip install .

# Development installation (editable)
pip install -e ".[dev,test]"
```

### 🏗️ Build and Package

```bash
# Build distributable package
uv build

# Install from built package
uv pip install dist/mcp_excel-*.whl
```

## ⚙️ Deployment Modes

The MCP Excel Office Server supports three deployment modes to fit different workflows and environments:

### DXT Package Deployment

**Best for**: Integrated DXT ecosystem users who want seamless configuration management.

1. **Package the project**:

   ```bash
   dxt pack
   ```

2. **Configuration**: The DXT package automatically handles dependencies and provides user-friendly configuration through the manifest.json:
   - `directory`: Base directory for file operations
   - `postgres_connection_string`: PostgreSQL database connection (marked as sensitive)

3. **Usage**: Once packaged, the tool integrates directly with DXT-compatible clients with automatic user configuration variable substitution.

4. **Server Configuration**: The DXT package includes a default server configuration in the manifest.json:

```json
"server": {
    "type": "python",
    "entry_point": "src/mcp_excel/server.py",
    "mcp_config": {
      "command": "python",
      "args": [
        "${__dirname}/src/mcp_excel/server.py"
      ],
      "env": {
        "DIRECTORY": "${user_config.directory}",
        "PYTHONPATH": "${__dirname}/src",
        "POSTGRES_CONNECTION_STRING": "${user_config.postgres_connection_string}"
      }
    }
}
```

for more details see [DXT Package Documentation](https://github.com/anthropics/dxt).

### Traditional MCP Server

**Best for**: Standard MCP server deployments with existing MCP infrastructure.

Add to your MCP configuration file (e.g., Claude Desktop's `mcp_config.json`):

```json
{
  "mcpServers": {
    "mcp_excel": {
      "command": "uv",
      "args": ["run", "mcp_excel_office"],
      "env": {
        "DIRECTORY": "user/to/path/directory",
        "POSTGRES_CONNECTION_STRING": "postgres_connection_string"
      }
    }
  }
}
```

**Alternative configuration with CLI arguments**:

```json
{
  "mcpServers": {
    "mcp_excel": {
      "command": "uv",
      "args": [
        "run", "-m", "mcp_excel",
        "--path", "user/to/path/directory",
        "--postgres", "postgres_connection_string"
      ]
    }
  }
}
```

### Standalone CLI

**Best for**: Direct command-line usage, scripting, and automation without MCP protocol overhead.

```bash
# Run with environment variables
export DIRECTORY="/path/to/your/files"
export POSTGRES_CONNECTION_STRING="postgresql://user:pass@localhost/db"
python -m mcp_excel

# Or run with command-line arguments
python -m mcp_excel --path "/path/to/files" --postgres "postgresql://user:pass@localhost/db"

# Using UV
uv run mcp_excel_office --help

```

for more details see [DXT Package Documentation](https://github.com/anthropics/dxt).

## 🔧 Configuration

### Environment Variables

- **`DIRECTORY`**: Base directory for file operations (required for security)
- **`POSTGRES_CONNECTION_STRING`**: PostgreSQL connection string for database operations (optional)

### Configuration Validation

The server validates all configuration on startup and provides clear error messages for:

- Missing required environment variables
- Invalid directory paths
- Malformed database connection strings
- File access permissions

## 🚀 Quick Start

### 1. Basic Setup

```bash
# Clone and install
git clone https://github.com/LuiccianDev/mcp_excel_office.git
cd mcp_excel_office
uv sync --dev

# Set up environment
export DIRECTORY="/path/to/your/excel/files"
export POSTGRES_CONNECTION_STRING="postgresql://user:pass@localhost/db"  # Optional
```

### 2. Test the Installation

```bash
# Run tests to verify installation
uv run pytest

# Check code quality
uv run ruff check
uv run mypy src/
```

### 3. Start the MCP Server

```bash
# Start MCP server
uv run mcp_excel_office

# Or use CLI mode
python -m mcp_excel --help
```

## 📚 Available Tools

The server provides comprehensive Excel manipulation through these MCP tool categories:

### 📊 Data Operations

- **`write_data_to_excel`**: Write data to spreadsheet ranges with type validation
- **`read_data_from_excel`**: Read data from spreadsheet ranges with flexible formatting
- **`append_data_to_excel`**: Append data to existing sheets with automatic range detection

### 📋 Workbook Management

- **`create_workbook`**: Create new Excel workbooks with customizable settings
- **`create_worksheet`**: Add worksheets to existing workbooks with naming validation
- **`get_workbook_metadata`**: Retrieve comprehensive workbook information and statistics

### 🎨 Formatting Operations

- **`format_range`**: Apply comprehensive cell formatting (fonts, colors, borders, alignment)
- **`set_column_width`**: Adjust column dimensions with validation
- **`set_row_height`**: Adjust row dimensions with validation

### 🧮 Formula Operations

- **`apply_formula`**: Apply Excel formulas to cells or ranges with validation
- **`validate_formula`**: Validate formula syntax before application
- **`calculate_range`**: Perform calculations on data ranges

### 📈 Graphics and Visualization

- **`create_chart`**: Generate various chart types (bar, line, pie, scatter, etc.)
- **`create_pivot_table`**: Create pivot tables from data with customizable aggregations
- **`add_image`**: Insert images into worksheets with positioning control

### 🗄️ Database Integration

- **`import_from_database`**: Import PostgreSQL data directly into Excel with query support
- **`export_to_database`**: Export Excel data to PostgreSQL with table creation
- **`execute_query_to_excel`**: Execute SQL queries and write results to Excel

For detailed documentation of all tools, parameters, and examples, see [TOOLS.md](TOOLS.md).

## 🧩 Project Structure

```text
mcp_excel_office/
├── src/mcp_excel/              # Main package
│   ├── __init__.py            # Package initialization
│   ├── __main__.py            # CLI entry point
│   ├── server.py              # MCP server implementation
│   ├── register_tools.py      # Tool registration
│   ├── tools/                 # MCP tool implementations
│   │   ├── content_tools.py   # Data read/write operations
│   │   ├── excel_tools.py     # Basic workbook operations
│   │   ├── format_tools.py    # Cell formatting and styling
│   │   ├── formulas_excel_tools.py  # Formula operations
│   │   ├── graphics_tools.py  # Charts and visualizations
│   │   └── db_tools.py        # Database integration
│   ├── core/                  # Core functionality
│   ├── utils/                 # Utility functions
│   └── exceptions/            # Custom exceptions
├── tests/                     # Test suite
│   ├── test_content_tools.py  # Content operations tests
│   ├── test_excel_tools.py    # Basic operations tests
│   ├── test_format_tools.py   # Formatting tests
│   ├── test_formulas_excel_tools.py  # Formula tests
│   ├── test_graphics_tools.py # Graphics tests
│   └── test_db_tools.py       # Database tests
├── documents/                 # Test Excel files
├── .kiro/                     # Kiro configuration
│   ├── specs/                 # Feature specifications
│   └── steering/              # Development guidelines
├── pyproject.toml             # Project configuration
├── manifest.json              # DXT package configuration
├── TOOLS.md                   # Detailed tool documentation
└── README.md                  # This file
```

## 🧪 Testing

### Run Tests

```bash
# Run all tests
uv run pytest

# Run with coverage report
uv run pytest --cov=src/mcp_excel/tools --cov-report=html

# Run specific test categories
uv run pytest -m unit          # Unit tests only
uv run pytest -m integration   # Integration tests only
uv run pytest -m excel         # Excel-specific tests
```

### Test Categories

- **Unit Tests**: Individual function and method testing
- **Integration Tests**: Component interaction testing
- **Excel Tests**: Excel file manipulation testing
- **MCP Tests**: MCP protocol compliance testing

### Coverage Requirements

- Minimum 20% test coverage (configured in pyproject.toml)
- Coverage reports generated in `htmlcov/` directory
- XML coverage reports for CI/CD integration

## 🔧 Development

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

# Type checking with MyPy
uv run mypy src/

# Run all quality checks
uv run pre-commit run --all-files
```

### Development Commands

- **`uv sync --dev`**: Install development dependencies
- **`uv run ruff check`**: Code style and quality checks
- **`uv run mypy src/`**: Type checking with strict configuration
- **`uv run pytest`**: Run test suite with coverage
- **`uv build`**: Build distributable package

## 🤝 Contributing

We welcome contributions! Please follow these guidelines:

### Getting Started

1. **Fork the repository** and clone your fork
2. **Create a feature branch**: `git checkout -b feature/amazing-feature`
3. **Install development dependencies**: `uv sync --dev`
4. **Install pre-commit hooks**: `uv run pre-commit install`

### Development Workflow

1. **Make your changes** following the code standards
2. **Add tests** for new functionality
3. **Run quality checks**: `uv run pre-commit run --all-files`
4. **Ensure tests pass**: `uv run pytest`
5. **Update documentation** as needed

### Code Standards

- **Python 3.11+** with complete type hints
- **88-character line limit** (enforced by Ruff)
- **Double quotes, snake_case naming**
- **MyPy strict mode compliance**
- **Minimum 20% test coverage**

### Submitting Changes

1. **Commit your changes**: `git commit -m 'Add amazing feature'`
2. **Push to your fork**: `git push origin feature/amazing-feature`
3. **Open a Pull Request** with a clear description

## 🐛 Issues and Support

- **Bug Reports**: [Open an issue](https://github.com/LuiccianDev/mcp_excel_office/issues) with detailed reproduction steps
- **Feature Requests**: Describe your use case and proposed solution
- **Questions**: Check existing issues or start a discussion

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

<div align="center">
  <p><strong>MCP Excel Office Server</strong></p>
  <p>Empowering AI assistants with comprehensive Excel manipulation capabilities</p>
  <p>
    <a href="https://github.com/LuiccianDev/mcp_excel_office">🏠 GitHub</a> •
    <a href="https://modelcontextprotocol.io">🔗 MCP Protocol</a> •
    <a href="https://github.com/LuiccianDev/mcp_excel_office/blob/main/TOOLS.md">📚 Tool Documentation</a>
  </p>
  <p><em>Created with by LuiccianDev</em></p>
</div>
