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

### create_workbook

Creates a new Excel workbook.

```python
create_workbook(filepath: str) -> str
```

- `filepath`: Path where to create the workbook (with or without .xlsx extension)
- Returns: Success message with created file path

### list_excel_documents

Lists all Excel documents in a directory.

```python
list_excel_documents(directory: str) -> List[Dict[str, str]]
```

- `directory`: Path to the directory to search in
- Returns: List of dictionaries containing file information (name, path, size, etc.)

### get_workbook_metadata

Get metadata about workbook including sheets and ranges.

```python
get_workbook_metadata(filepath: str, include_ranges: bool = False) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
- `include_ranges`: Whether to include range information (default: False)
- Returns: Dictionary containing workbook metadata

## Worksheet Operations

### create_worksheet

Creates a new worksheet in an existing workbook.

```python
create_worksheet(filepath: str, sheet_name: str) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
- `sheet_name`: Name for the new worksheet
- Returns: Dictionary with operation status and details

### copy_worksheet

Creates a copy of an existing worksheet.

```python
copy_worksheet(filepath: str, source_sheet: str, target_sheet: str) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
- `source_sheet`: Name of the sheet to copy
- `target_sheet`: Name for the new sheet
- Returns: Dictionary with operation status

### delete_worksheet

Deletes a worksheet from a workbook.

```python
delete_worksheet(filepath: str, sheet_name: str) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the sheet to delete
- Returns: Dictionary with operation status

### rename_worksheet

Renames an existing worksheet.

```python
rename_worksheet(filepath: str, old_name: str, new_name: str) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
- `old_name`: Current name of the sheet
- `new_name`: New name for the sheet
- Returns: Dictionary with operation status

## Data Operations

### write_data_to_excel

Write data to Excel worksheet.

```python
write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[Dict],
    start_cell: str = "A1"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data`: List of dictionaries containing data to write
- `start_cell`: Starting cell (default: "A1")
- Returns: Success message

### read_data_from_excel

Read data from Excel worksheet.

```python
read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    preview_only: bool = False
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Source worksheet name
- `start_cell`: Starting cell (default: "A1")
- `end_cell`: Optional ending cell
- `preview_only`: Whether to return only a preview
- Returns: String representation of data

## Formatting Operations

### format_range

Apply formatting to a range of cells.

```python
format_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int = None,
    font_color: str = None,
    bg_color: str = None,
    border_style: str = None,
    border_color: str = None,
    number_format: str = None,
    alignment: str = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Dict[str, Any] = None,
    conditional_format: Dict[str, Any] = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Optional ending cell of range
- Various formatting options (see parameters)
- Returns: Success message

### merge_cells

Merge a range of cells.

```python
merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Success message

### unmerge_cells

Unmerge a previously merged range of cells.

```python
unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Success message

### copy_range

Copy a range of cells to another location.

```python
copy_range(
    filepath: str,
    sheet_name: str,
    source_range: str,
    target_cell: str,
    include_formatting: bool = True
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Worksheet name
- `source_range`: Range to copy (e.g., "A1:B10")
- `target_cell`: Top-left cell of target range
- `include_formatting`: Whether to copy cell formatting (default: True)
- Returns: Success message

### delete_range

Clear contents and formatting from a range of cells.

```python
delete_range(filepath: str, sheet_name: str, cell_range: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Worksheet name
- `cell_range`: Range to clear (e.g., "A1:B10")
- Returns: Success message

### validate_excel_range

Validate if a range reference is valid for the specified worksheet.

```python
validate_excel_range(filepath: str, sheet_name: str, cell_range: str) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
- `sheet_name`: Worksheet name
- `cell_range`: Range to validate (e.g., "A1:B10")
- Returns: Dictionary with validation result and details

## Formula Operations

### apply_formula

Apply Excel formula to cell.

```python
apply_formula(filepath: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to apply
- Returns: Success message

### validate_formula_syntax

Validate Excel formula syntax without applying it.

```python
validate_formula_syntax(filepath: str, sheet_name: str, cell: str, formula: str) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to validate
- Returns: Dictionary with validation result and details

## Database Operations

### fetch_and_insert_db_to_excel

Fetch data from a database and insert it into an Excel worksheet.

```python
fetch_and_insert_db_to_excel(
    connection_string: str,
    query: str,
    filename: str,
    sheet_name: str
) -> str
```

- `connection_string`: Database connection string
- `query`: SQL SELECT query to fetch data
- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- Returns: Status message

### insert_calculated_data_to_db

Insert calculated data from Excel into a database.

```python
insert_calculated_data_to_db(
    connection_string: str,
    table_name: str,
    filename: str,
    sheet_name: str,
    range_ref: str
) -> str
```

- `connection_string`: Database connection string
- `table_name`: Target database table
- `filename`: Path to Excel file
- `sheet_name`: Source worksheet name
- `range_ref`: Cell range containing data to insert
- Returns: Status message

## Chart and Pivot Table Operations

### create_chart

Create a chart in the specified worksheet.

```python
create_chart(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
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
    filepath: str,
    source_sheet: str,
    target_sheet: str,
    data_range: str,
    rows: List[str],
    columns: List[str] = None,
    values: List[Tuple[str, str]] = None,
    filters: List[str] = None,
    pivot_table_name: str = "PivotTable1"
) -> Dict[str, Any]
```

- `filepath`: Path to Excel file
- `source_sheet`: Name of the sheet containing source data
- `target_sheet`: Name of the sheet where to create the pivot table
- `data_range`: Range containing source data (e.g., "A1:D100")
- `rows`: List of field names to use as rows
- `columns`: List of field names to use as columns (optional)
- `values`: List of tuples (field_name, aggregate_function) for values
- `filters`: List of field names to use as filters (optional)
- `pivot_table_name`: Name for the pivot table (default: "PivotTable1")
- Returns: Dictionary with operation status
