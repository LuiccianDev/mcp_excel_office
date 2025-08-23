# Excel MCP Server Tools

This document provides detailed information about all available tools in the Excel MCP server, organized by functional categories.

## Table of Contents

- [Workbook Operations](#workbook-operations)
- [Worksheet Operations](#worksheet-operations)
- [Data Operations](#data-operations)
- [Formatting Operations](#formatting-operations)
- [Formula Operations](#formula-operations)
- [Database Operations](#database-operations)
- [Chart and Pivot Table Operations](#chart-and-pivot-table-operations)

## Workbook Operations

### create_excel_workbook

Creates a new Excel workbook.

```python
create_excel_workbook(filename: str) -> dict[str, Any]
```

- `filename`: Path where to create the workbook (with or without .xlsx extension)
- Returns: Dictionary with operation status and details

### list_excel_documents

Lists all Excel documents in the configured directory.

```python
list_excel_documents() -> dict[str, Any]
```

- Returns: Dictionary containing list of Excel files with their information (name, path, size, etc.)

### get_workbook_metadata

Get metadata about workbook including sheets and ranges.

```python
get_workbook_metadata(filename: str, include_ranges: bool = False) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `include_ranges`: Whether to include range information (default: False)
- Returns: Dictionary containing workbook metadata

## Worksheet Operations

### create_excel_worksheet

Creates a new worksheet in an existing workbook.

```python
create_excel_worksheet(filename: str, sheet_name: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Name for the new worksheet
- Returns: Dictionary with operation status and details

### copy_worksheet

Creates a copy of an existing worksheet.

```python
copy_worksheet(filename: str, source_sheet: str, target_sheet: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `source_sheet`: Name of the sheet to copy
- `target_sheet`: Name for the new sheet
- Returns: Dictionary with operation status

### delete_worksheet

Deletes a worksheet from a workbook.

```python
delete_worksheet(filename: str, sheet_name: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Name of the sheet to delete
- Returns: Dictionary with operation status

### rename_worksheet

Renames an existing worksheet.

```python
rename_worksheet(filename: str, old_name: str, new_name: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `old_name`: Current name of the sheet
- `new_name`: New name for the sheet
- Returns: Dictionary with operation status

## Data Operations

### write_data_to_excel

Write data to Excel worksheet.

```python
write_data_to_excel(
    filename: str,
    sheet_name: str,
    data: list[list[Any]],
    start_cell: str = "A1",
    headers: list[str] | None = None
) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data`: List of lists containing data to write
- `start_cell`: Starting cell (default: "A1")
- `headers`: Optional list of column headers
- Returns: Dictionary with operation status

### read_data_from_excel

Read data from Excel worksheet.

```python
read_data_from_excel(
    filename: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str | None = None,
    preview_only: bool = False
) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Source worksheet name
- `start_cell`: Starting cell (default: "A1")
- `end_cell`: Optional ending cell
- `preview_only`: Whether to return only a preview
- Returns: Dictionary containing data and status information

## Formatting Operations

### format_range_excel

Apply formatting to a range of cells.

```python
format_range_excel(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    border_style: str | None = None,
    border_color: str | None = None,
    number_format: str | None = None,
    alignment: str | None = None,
    wrap_text: bool = False
) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Optional ending cell of range
- Various formatting options (see parameters)
- Returns: Dictionary with operation status

### merge_cells

Merge a range of cells.

```python
merge_cells(filename: str, sheet_name: str, start_cell: str, end_cell: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Dictionary with operation status

### unmerge_cells

Unmerge a previously merged range of cells.

```python
unmerge_cells(filename: str, sheet_name: str, start_cell: str, end_cell: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Dictionary with operation status

### copy_range

Copy a range of cells to another location.

```python
copy_range(
    filename: str,
    sheet_name: str,
    source_start_cell: str,
    source_end_cell: str,
    target_start_cell: str,
    include_formatting: bool = True
) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Worksheet name
- `source_start_cell`: Starting cell of source range
- `source_end_cell`: Ending cell of source range
- `target_start_cell`: Top-left cell of target range
- `include_formatting`: Whether to copy cell formatting (default: True)
- Returns: Dictionary with operation status

### delete_range

Clear contents and formatting from a range of cells.

```python
delete_range(filename: str, sheet_name: str, start_cell: str, end_cell: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Worksheet name
- `start_cell`: Starting cell of range to clear
- `end_cell`: Ending cell of range to clear
- Returns: Dictionary with operation status

### validate_excel_range

Validate if a range reference is valid for the specified worksheet.

```python
validate_excel_range(filename: str, sheet_name: str, start_cell: str, end_cell: str | None = None) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Worksheet name
- `start_cell`: Starting cell of range to validate
- `end_cell`: Optional ending cell of range to validate
- Returns: Dictionary with validation result and details

## Formula Operations

### apply_formula_excel

Apply Excel formula to cell.

```python
apply_formula_excel(filename: str, sheet_name: str, cell: str, formula: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to apply
- Returns: Dictionary with operation status

### validate_formula_syntax

Validate Excel formula syntax without applying it.

```python
validate_formula_syntax(filename: str, sheet_name: str, cell: str, formula: str) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to validate
- Returns: Dictionary with validation result and details

## Database Operations

### fetch_and_insert_db_to_excel

Fetch data from a database and insert it into an Excel worksheet.

```python
fetch_and_insert_db_to_excel(
    query: str,
    filename: str,
    sheet_name: str,
    connection_string: str | None = None
) -> dict[str, Any]
```

- `query`: SQL SELECT query to fetch data
- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `connection_string`: Database connection string (optional, uses environment variable if not provided)
- Returns: Dictionary with operation status

### insert_calculated_data_to_db

Insert calculated data into a database table.

```python
insert_calculated_data_to_db(
    table: str,
    columns: list[str],
    rows: list[tuple],
    connection_string: str | None = None
) -> dict[str, Any]
```

- `table`: Target database table name
- `columns`: List of column names
- `rows`: List of tuples containing row data
- `connection_string`: Database connection string (optional, uses environment variable if not provided)
- Returns: Dictionary with operation status

## Chart and Pivot Table Operations

### create_chart

Create a chart in the specified worksheet.

```python
create_chart(
    filename: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data_range`: Range containing chart data (e.g., "A1:B10")
- `chart_type`: Type of chart (e.g., "line", "bar", "pie", "scatter", "area")
- `target_cell`: Cell where to place the top-left corner of the chart
- `title`: Chart title (optional)
- `x_axis`: X-axis title (optional)
- `y_axis`: Y-axis title (optional)
- Returns: Dictionary with operation status

### create_pivot_table

Create a pivot table in the specified worksheet.

```python
create_pivot_table(
    filename: str,
    sheet_name: str,
    data_range: str,
    rows: list[str],
    columns: list[str] | None = None,
    values: list[tuple[str, str]] | None = None,
    filters: list[str] | None = None,
    pivot_table_name: str = "PivotTable1"
) -> dict[str, Any]
```

- `filename`: Path to Excel file
- `sheet_name`: Name of the sheet where to create the pivot table
- `data_range`: Range containing source data (e.g., "A1:D100")
- `rows`: List of field names to use as rows
- `columns`: List of field names to use as columns (optional)
- `values`: List of tuples (field_name, aggregate_function) for values
- `filters`: List of field names to use as filters (optional)
- `pivot_table_name`: Name for the pivot table (default: "PivotTable1")
- Returns: Dictionary with operation status
