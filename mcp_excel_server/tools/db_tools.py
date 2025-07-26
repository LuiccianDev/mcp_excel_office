from mcp_excel_server.core.db_conection import (
    fetch_data_from_db,
    insert_data_to_excel,
    validate_sql_query,
    clean_data,
    insert_calculated_data_to_db,
)

from mcp_excel_server.utils.file_utils import ensure_xlsx_extension


async def fetch_and_insert_db_data(
    connection_string: str, query: str, filename: str, sheet_name: str
) -> str:
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
        return "Error: Invalid or potentially unsafe SQL query."
    db_result = fetch_data_from_db(connection_string, query)
    if "error" in db_result:
        return f"Error: {db_result['error']}"
    columns = db_result.get("columns", [])
    rows = db_result.get("rows", [])
    cleaned_rows = clean_data(rows, columns)
    excel_result = insert_data_to_excel(filename, sheet_name, columns, cleaned_rows)
    if "error" in excel_result:
        return f"Error: {excel_result['error']}"
    return excel_result.get("message", "Data inserted successfully.")


async def insert_calculated_data(
    connection_string: str, table: str, columns: list, rows: list
) -> str:
    """
    Insert calculated/cleaned data into the database.
    Args:
        connection_string (str): Database connection string.
        table (str): Target table name.
        columns (list): List of column names.
        rows (list): List of tuples containing data to insert.
    """
    cleaned_rows = clean_data(rows, columns)
    result = insert_calculated_data_to_db(
        connection_string, table, columns, cleaned_rows
    )
    if "error" in result:
        return f"Error: {result['error']}"
    return result.get("message", "Data inserted successfully.")
