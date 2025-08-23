# WARP.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

This is an MCP (Model Context Protocol) server for programmatic Excel file manipulation. It provides a comprehensive set of tools for creating, reading, modifying, and formatting Excel workbooks through standardized MCP interfaces.

## Development Commands

### Environment Setup
```bash
# Install dependencies with UV (recommended)
uv sync --dev

# Alternative setup with pip
pip install -e ".[dev]"

# Create virtual environment
uv venv
.venv\Scripts\activate  # Windows
```

### Code Quality & Testing
```bash
# Run pre-commit hooks on all files
uv run pre-commit run --all-files

# Run tests with coverage
uv run pytest

# Run specific test file
uv run pytest tests/test_excel_tools.py

# Type checking with mypy
uv run mypy src/

# Linting and formatting with ruff
uv run ruff check src/
uv run ruff format src/
```

### Build & Package
```bash
# Build wheel package
uv build

# Install built package locally
uv pip install dist/mcp_excel_office-*.whl
```

### Pre-commit Workflow (from external context)
```bash
# Stage changes
git add .

# Run hooks explicitly before commit
pre-commit run --all-files

# If corrections were made, re-stage
git add .

# Commit with conventional message
git commit -m "feat: description of changes"
```

## Architecture Overview

### Core Structure
- **`src/mcp_excel/`** - Main package directory
- **`core/`** - Business logic modules (workbook, data, formatting, calculations)
- **`tools/`** - MCP tool implementations that expose functionality
- **`utils/`** - Helper functions for validation, file handling, etc.
- **`exceptions/`** - Custom exception classes

### Key Components

**Server Architecture:**
- `server.py` - FastMCP server setup and initialization
- `__main__.py` - Entry point for CLI execution
- `tools/register_tools.py` - Central tool registration

**Core Modules:**
- `core/workbook.py` - Workbook creation and manipulation
- `core/data.py` - Data read/write operations
- `core/formatting.py` - Cell and range formatting
- `core/calculations.py` - Formula and calculation handling

**Tool Categories:**
- `excel_tools.py` - Workbook and worksheet operations
- `content_tools.py` - Data reading and writing
- `format_tools.py` - Formatting and styling
- `formulas_excel_tools.py` - Formula validation and application
- `graphics_tools.py` - Charts and pivot tables
- `db_tools.py` - Database integration

### Security & Validation
The codebase implements comprehensive security measures:
- File path validation and sanitization
- Directory access controls via `@validate_file_access` decorators
- Secure path resolution to prevent directory traversal
- Input validation for Excel ranges, formulas, and sheet names

### Dependencies
- **openpyxl** - Excel file manipulation
- **mcp[cli]** - Model Context Protocol implementation
- **psycopg2** - PostgreSQL database connectivity
- **typer** - CLI interface

### Tool Organization Pattern
Each tool module follows this pattern:
1. Import validation decorators and core functions
2. Async tool functions with comprehensive docstrings
3. Error handling with custom exception types
4. Return standardized dictionaries with status/error information

## Configuration Files

### pyproject.toml
- Contains comprehensive test configuration with pytest markers
- MyPy type checking settings (strict mode)
- Coverage reporting configuration
- Development and test dependency groups

### ruff.toml
- Code formatting (88 character line length)
- Import sorting with isort integration
- Comprehensive linting rules (pycodestyle, flake8-bugbear, etc.)
- Per-file ignore patterns for tests

### .pre-commit-config.yaml
- Ruff for linting and formatting
- MyPy for type checking
- Standard hooks for file formatting

## Testing Strategy
- Unit tests for each tool category in `tests/`
- Test markers: `@pytest.mark.unit`, `@pytest.mark.integration`, `@pytest.mark.excel`
- Coverage threshold set to 20% (configured in pyproject.toml)
- AsyncIO testing support enabled

## Common Development Patterns

### Adding New Tools
1. Create function in appropriate `tools/` module
2. Add validation decorators if file/directory access needed
3. Include comprehensive docstring with AI context
4. Register tool in `register_tools.py`
5. Add corresponding tests

### Error Handling
- Use custom exceptions from `exceptions/` module
- Return dictionaries with `status` and `message` keys
- Wrap external library exceptions in domain-specific exceptions

### File Operations
- Always use `resolve_safe_path()` for user-provided paths
- Apply `@validate_file_access` or `@validate_directory_access` decorators
- Use `ensure_xlsx_extension()` for Excel files
