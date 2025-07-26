from mcp_excel_server.core.chart import create_chart_in_sheet as create_chart_impl
from mcp_excel_server.core.pivot import create_pivot_table as create_pivot_table_impl

# Import exceptions
from mcp_excel_server.exceptions.exceptions import (
    ChartError,
    PivotError,
    ValidationError,
)
from mcp_excel_server.utils.file_utils import (
    check_file_writeable,
    ensure_xlsx_extension,
)


async def create_chart(
    filename: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
) -> str:
    """Create chart in worksheet.
    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet.
        data_range (str): The range of data to be used for the chart.
        chart_type (str): The type of chart to create.
        target_cell (str): The cell where the chart will be placed.
        title (str, optional): The title of the chart. Defaults to "".
        x_axis (str, optional): The label for the x-axis. Defaults to "".
        y_axis (str, optional): The label for the y-axis. Defaults to "".
    """
    filename = ensure_xlsx_extension(filename)
    # Check if the file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Error: {error_message}"

    try:
        result = create_chart_impl(
            filename=filename,
            sheet_name=sheet_name,
            data_range=data_range,
            chart_type=chart_type,
            target_cell=target_cell,
            title=title,
            x_axis=x_axis,
            y_axis=y_axis,
        )
        return result["message"]
    except (ValidationError, ChartError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to create chart: {str(e)}"


async def create_pivot_table(
    filename: str,
    sheet_name: str,
    data_range: str,
    rows: list[str],
    values: list[str],
    columns: list[str] = None,
    agg_func: str = "mean",
) -> str:
    """Create pivot table in worksheet.
    Args:
        filename (str): The name of the Excel file.
        sheet_name (str): The name of the worksheet.
        data_range (str): The range of data to be used for the pivot table.
        rows (List[str]): The rows to be used in the pivot table.
        values (List[str]): The values to be used in the pivot table.
        columns (List[str], optional): The columns to be used in the pivot table. Defaults to None.
        agg_func (str, optional): The aggregation function to be used. Defaults to "mean".
    """
    filename = ensure_xlsx_extension(filename)
    # Check if the file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Error: {error_message}"
    try:

        result = create_pivot_table_impl(
            filename=filename,
            sheet_name=sheet_name,
            data_range=data_range,
            rows=rows,
            values=values,
            columns=columns or [],
            agg_func=agg_func,
        )
        return result["message"]
    except (ValidationError, PivotError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Failed to create pivot table: {str(e)}"
