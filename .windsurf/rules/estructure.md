---
trigger: manual
---

# MCP Excel Project Rules

## Project Structure
- Keep all source code in the `src/mcp_excel` directory
- Use separate modules for different Excel functionalities (e.g., `workbook.py`, `worksheet.py`, `cell.py`)
- Store tests in the `tests` directory with matching structure to source code

## Code Style
- Follow PEP 8 style guide
- Use type hints for all function signatures
- Keep functions small and focused on single responsibility
- Document all public APIs with docstrings following Google style

## Excel-Specific Rules
1. **Workbook Management**
   - Support both `.xlsx` and `.xls` formats
   - Handle file operations with proper error handling
   - Close files properly after operations

2. **Worksheet Operations**
   - Implement CRUD operations for worksheets
   - Handle large datasets efficiently
   - Support common operations like sorting, filtering, and formatting

3. **Cell Operations**
   - Support different data types (strings, numbers, dates, formulas)
   - Handle cell references (A1, R1C1)
   - Implement cell formatting options

4. **Formulas and Functions**
   - Support common Excel functions
   - Handle formula parsing and evaluation
   - Support array formulas

## Testing
- Write unit tests for all public functions
- Use pytest as the testing framework
- Achieve at least 80% test coverage
- Include integration tests for end-to-end scenarios

## Dependencies
- Use `openpyxl` as the primary Excel library
- Add `pandas` for advanced data manipulation
- Include `pytest` for testing
- Specify exact versions in `requirements.txt`

## Documentation
- Include a comprehensive README.md
- Document all public APIs
- Add examples for common use cases
- Keep a CHANGELOG.md for version history

## Error Handling
- Use custom exceptions for Excel-specific errors
- Provide meaningful error messages
- Log errors appropriately

## Performance
- Optimize for large Excel files
- Use generators for large datasets
- Implement proper memory management

## Security
- Validate all input data
- Handle sensitive information securely
- Protect against Excel injection attacks
