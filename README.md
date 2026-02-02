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

A comprehensive MCP (Model Context Protocol) server that provides AI assistants with powerful Excel manipulation capabilities. This server enables programmatic creation, modification, and management of Excel files through standardized MCP tools, supporting data operations, formatting, formulas, and charts.

## 📋 Table of Contents

- [✨ Key Features](#-key-features)
- [🚀 Installation](#-installation)
- [⚙️ Deployment Modes](#️-deployment-modes)
  - [DXT Package Deployment](#dxt-package-deployment)
  - [Traditional MCP Server](#traditional-mcp-server)
  - [Standalone CLI](#standalone-cli)
- [🔧 Configuration](#-configuration)
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
- **⚡ Formula Support**: Apply and validate Excel formulas with error handling
- **🔒 Security First**: File path validation, access control, and robust error handling
- **🚀 Multiple Deployment Modes**: DXT package, traditional MCP server, or standalone CLI
- **🤖 AI-Ready**: Optimized for AI assistant integration via Model Context Protocol

## 🚀 Installation

### 📋 Prerequisites

- **Python 3.11+**: Modern Python with type hints support
- **UV Package Manager**: [Install UV](https://docs.astral.sh/uv/getting-started/installation/) (recommended) or use pip
- **Git**: For cloning the repository
- **Desktop Extensions (DXT)** : for create packages .dxt for claude desktop [Install DXT](https://github.com/anthropics/dxt)

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

3. **Usage**: Once packaged, the tool integrates directly with DXT-compatible clients with automatic user configuration variable substitution.

4. **Server Configuration**: this proyect include the files [manifest.json](./manifest.json)  for building package .dxt

for more details see [DXT Package Documentation](https://github.com/anthropics/dxt).

### Traditional MCP Server

**Best for**: Standard MCP server deployments with existing MCP infrastructure.

Add to your MCP configuration file (e.g., Claude Desktop's `mcp_config.json`):

```bash
# create packages
uv build
#install packages
pip install dist/archivo*.whl
```

The next steps is configuractiosn en mcp

```json
{
  "mcpServers": {
    "mcp_excel": {
      "command": "uv",
      "args": ["run", "mcp_excel"],
      "env": {
        "DIRECTORY": "user/to/path/directory"
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
        "--path", "user/to/path/directory"
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
python -m mcp_excel

# Or run with command-line arguments
python -m mcp_excel --path "/path/to/files"

# Using UV
uv run mcp_excel_office --help

```

## 🐳 Instalación con Docker

Puedes instalar y ejecutar el servidor MCP Excel Office fácilmente usando Docker. Esto garantiza un entorno aislado y reproducible.

Para más detalles y opciones avanzadas de configuración con Docker, consulta el archivo [`Docker.md`](./Docker.md).

## 🔧 Configuration

### Environment Variables

- **`DIRECTORY`**: Base directory for file operations (required for security)

### Configuration Validation

The server validates all configuration on startup and provides clear error messages for:

- Missing required environment variables
- Invalid directory paths
- File access permissions

## 📚 Available Tools

Todas las herramientas disponibles para manipulación de Excel, operaciones de datos, formato, fórmulas y gráficos están documentadas en detalle en el archivo [`TOOLS.md`](TOOLS.md). Consulta ese archivo para ver la lista completa de herramientas, sus parámetros y ejemplos de uso.

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
│   │   └── graphics_tools.py  # Charts and visualizations
│   ├── core/                  # Core functionality
│   └── utils/                 # Utility functions
├── tests/                     # Test suite
│   ├── test_content_tools.py  # Content operations tests
│   ├── test_excel_tools.py    # Basic operations tests
│   ├── test_format_tools.py   # Formatting tests
│   ├── test_formulas_excel_tools.py  # Formula tests
│   └── test_graphics_tools.py # Graphics tests
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
```

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
