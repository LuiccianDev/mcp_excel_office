import asyncio
from typing import Any

from mcp_excel.config import (
    ConfigurationError,
    is_database_configured,
    validate_file_path,
)
from mcp_excel.core.db_conection import (
    clean_data,
    fetch_data_from_db,
    insert_data_to_db,
    insert_data_to_excel,
    validate_sql_query,
)
from mcp_excel.utils.file_utils import ensure_xlsx_extension


async def fetch_and_insert_db_to_excel(
    query: str,
    filename: str,
    sheet_name: str,
    connection_string: str | None = None,
) -> dict[str, Any]:
    """
    Fetch tabular data from a database using a validated SELECT query, clean the results, and insert them into a specified Excel worksheet.

    Context for AI/LLM:
        Use this tool to automate the extraction of data from a SQL database and its transfer into Excel for reporting, analysis, or archival. The function ensures only safe SELECT queries are executed, cleans the data for Excel compatibility, and writes the results to the specified worksheet.

    Typical use cases:
        1. Migrating data from a database to Excel for business reporting.
        2. Automating ETL (Extract, Transform, Load) workflows.
        3. Creating Excel-based dashboards from live database queries.

    Args:
        connection_string (str): Database connection string (driver, server, credentials, etc.).
        query (str): SQL SELECT query to fetch data. Must be validated as safe.
        filename (str): Path to the target Excel file. The .xlsx extension is enforced.
        sheet_name (str): Name of the worksheet to insert data into. Must exist.

    Returns:
        dict[str, Any]:
            - status (str): "success" or "error".
            - message (str): Details of the operation or error encountered.

    Notes:
        • Only SELECT queries are allowed; unsafe queries are rejected.
        • The worksheet must already exist in the Excel file.
        • Data is cleaned for Excel compatibility before insertion.
        • Blocking I/O is handled in threads for async compatibility.
    """
    try:
        # Check if database is configured
        if not connection_string and not is_database_configured():
            return {
                "status": "error",
                "message": "Database is not configured. Please set POSTGRES_CONNECTION_STRING environment variable or provide connection_string parameter.",
            }

        # Validate and secure file path
        try:
            filename = ensure_xlsx_extension(filename)
            validated_filename = validate_file_path(filename)
        except ConfigurationError as e:
            return {
                "status": "error",
                "message": f"File path validation failed: {e}",
            }

        # Validate SQL query for security
        if not validate_sql_query(query):
            return {
                "status": "error",
                "message": "Invalid or potentially unsafe SQL query. Only SELECT queries are allowed.",
            }

        # Execute the query and get results
        result = await asyncio.to_thread(
            fetch_data_from_db, query=query, connection_string=connection_string
        )

        if result.get("status") == "error":
            return {
                "status": "error",
                "message": f"Database query failed: {result.get('message', 'Unknown error')}",
            }

        columns = result.get("columns", [])
        rows = result.get("rows", [])

        # clean_data is CPU-bound but likely fast enough to not need a thread
        cleaned_rows = clean_data(rows, columns)

        # Run blocking Excel call in a separate thread
        excel_result = await asyncio.to_thread(
            insert_data_to_excel, validated_filename, sheet_name, columns, cleaned_rows
        )

        if excel_result.get("status") == "error":
            return {
                "status": "error",
                "message": f"Excel operation failed: {excel_result.get('message', 'Unknown error')}",
            }

        return {
            "status": "success",
            "message": excel_result.get("message", "Data inserted successfully."),
        }

    except Exception as e:
        return {
            "status": "error",
            "message": f"Unexpected error during database to Excel operation: {e}",
        }


async def insert_calculated_data_to_db(
    table: str,
    columns: list[str],
    rows: list[tuple],
    connection_string: str | None = None,
) -> dict[str, Any]:
    """
    Insert calculated or cleaned tabular data into a database table.

    Context for AI/LLM:
        Use this tool to automate the process of persisting processed or cleaned data into a SQL database. This is useful for updating reporting tables, storing results from data pipelines, or archiving analytics outputs.

    Typical use cases:
        1. Saving processed analytics results to a database for later retrieval.
        2. Batch-inserting cleaned records from Excel or other sources.
        3. Automating the final step of a data processing workflow.

    Args:
        connection_string (str): Database connection string.
        table (str): Target table name for data insertion.
        columns (List): List of column names corresponding to the data.
        rows (List): List of tuples or lists, each representing a row to insert.

    Returns:
        Dict[str, Any]:
            - status (str): "success" or "error".
            - message (str): Details of the operation or error encountered.
            - rows_inserted (int, optional): Number of rows successfully inserted (on success).
            - table (str): Target table name.
            - details (Dict, optional): Additional error details (on failure).

    Notes:
        • Input data is cleaned before insertion to match DB schema.
        • Blocking I/O is handled in threads for async compatibility.
        • On error, returns descriptive message and error details.
    """
    try:
        # Check if database is configured
        if not connection_string and not is_database_configured():
            return {
                "status": "error",
                "message": "Database is not configured. Please set POSTGRES_CONNECTION_STRING environment variable or provide connection_string parameter.",
                "table": table,
            }

        # Validate input parameters
        if not table or not isinstance(table, str):
            return {
                "status": "error",
                "message": "Table name must be a non-empty string.",
                "table": table,
            }

        if not columns or not all(isinstance(col, str) for col in columns):
            return {
                "status": "error",
                "message": "Columns must be a non-empty list of strings.",
                "table": table,
            }

        if not rows:
            return {
                "status": "error",
                "message": "No data rows provided for insertion.",
                "table": table,
            }

        # Clean input rows
        try:
            cleaned_rows = clean_data(rows, columns)
        except ValueError as e:
            return {
                "status": "error",
                "message": f"Data validation failed: {e}",
                "table": table,
            }

        # Insert data into the database
        result = await asyncio.to_thread(
            insert_data_to_db,
            table=table,
            columns=columns,
            rows=cleaned_rows,
            connection_string=connection_string,
        )

        if result.get("status") == "error":
            return {
                "status": "error",
                "message": f"Database insert failed: {result.get('message', 'Unknown error')}",
                "table": table,
                "details": result,
            }

        return {
            "status": "success",
            "message": result.get("message", "Data inserted successfully."),
            "rows_inserted": len(cleaned_rows),
            "table": table,
        }

    except Exception as e:
        return {
            "status": "error",
            "message": f"Unexpected error during database insert: {e}",
            "table": table,
            "error": repr(e),
        }
