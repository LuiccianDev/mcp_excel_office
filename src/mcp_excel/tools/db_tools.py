import asyncio
from typing import Any

from mcp_excel.core.db_conection import (
    clean_data,
    fetch_data_from_db,
    insert_data_to_db,
    insert_data_to_excel,
    validate_sql_query,
)
from mcp_excel.utils.file_utils import ensure_xlsx_extension


async def fetch_and_insert_db_to_excel(
    connection_string: str, query: str, filename: str, sheet_name: str
) -> dict[str, Any]:
    """
    Fetch data from DB (safe SELECT), clean it, and insert into Excel.
    Args:
        connection_string (str): Database connection string.
        query (str): SQL SELECT query to fetch data.
        filename (str): Path to the Excel file.
        sheet_name (str): Name of the sheet to insert data into.
    """
    filename = ensure_xlsx_extension(filename)
    if not validate_sql_query(query):
        return {
            "status": "error",
            "message": "Invalid or potentially unsafe SQL query.",
        }

    # Run blocking DB call in a separate thread
    db_result = await asyncio.to_thread(fetch_data_from_db, connection_string, query)

    if "error" in db_result:
        return {"status": "error", "message": f"Error: {db_result['error']}"}

    columns = db_result.get("columns", [])
    rows = db_result.get("rows", [])

    # clean_data is CPU-bound but likely fast enough to not need a thread
    cleaned_rows = clean_data(rows, columns)

    # Run blocking Excel call in a separate thread
    excel_result = await asyncio.to_thread(
        insert_data_to_excel, filename, sheet_name, columns, cleaned_rows
    )

    if "error" in excel_result:
        return {"status": "error", "message": f"Error: {excel_result['error']}"}

    return {
        "status": "success",
        "message": excel_result.get("message", "Data inserted successfully."),
    }


async def insert_calculated_data_to_db(
    connection_string: str, table: str, columns: list, rows: list
) -> dict[str, Any]:
    """
    Insert calculated/cleaned data into the database.
    Args:
        connection_string (str): Database connection string.
        table (str): Target table name.
        columns (list): List of column names.
        rows (list): List of tuples containing data to insert.
    Returns:
        dict[str, Any]: Status dictionary with message and optional error.
    """
    try:
        # Clean input rows
        cleaned_rows = clean_data(rows, columns)

        # Perform DB insert in a separate thread
        result = await asyncio.to_thread(
            insert_data_to_db, connection_string, table, columns, cleaned_rows
        )

        if result.get("status") == "error":
            return {
                "status": "error",
                "message": f"Database insert failed: {result['error']}",
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
            "message": str(e),
            "error": repr(e),
            "table": table,
        }
