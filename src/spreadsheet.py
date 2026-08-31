import pywintypes
import win32com.client as win32

from error_handler import handle_errors
from exceptions import ExcelConnectionError, DataCleaningError

FILTER_VALUES_OPERATOR = 7
SORT_ASCENDING = 1
HEADER_PRESENT = 1
SORT_TOP_TO_BOTTOM = 1


def get_column_number_of_heading(ws, column_name):
    for cell in ws.UsedRange.Rows(1).Cells:
        if cell.Value == column_name:
            return cell.Column
    return None


def apply_filter(ws, column_number: int, filter_values: list[str]):

    used = ws.UsedRange
    used.AutoFilter(
        # - used.Column + 1 ensures this would work if ever starting cell changed from A1
        Field = column_number - used.Column + 1,
        Criteria1 = tuple(filter_values),
        Operator = FILTER_VALUES_OPERATOR
    )


def sort_rows(ws, column_header):
    used = ws.UsedRange
    for cell in used.Rows(1).Cells:

        if cell.Value == column_header:
            starting_cell = ws.Cells(cell.Row + 1, cell.Column)
            used.Sort(
                Key1=starting_cell,
                Order1=SORT_ASCENDING,
                Header=HEADER_PRESENT,
                Orientation=SORT_TOP_TO_BOTTOM
            )


@handle_errors
def launch_spreadsheet_app():
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        return excel
    except pywintypes.com_error as e:
        raise ExcelConnectionError(f"Failed to launch spreadsheet: {e}")


@handle_errors
def open_spreadsheet(excel, file_loc, worksheet_name):
    try:
        wb = excel.Workbooks.Open(file_loc)
        ws = wb.Worksheets(worksheet_name)
        return wb, ws
    except pywintypes.com_error as e:
        raise ExcelConnectionError(f"Failed to open spreadsheet: {e}")


@handle_errors
def remove_filters(ws):
    try:
        if ws.FilterMode:
            ws.ShowAllData()
    except pywintypes.com_error as e:
        raise ExcelConnectionError(f"Excel connection error: {e}")


@handle_errors
def update_spreadsheet_from_dataframe(df, ws):
    try:
        values = df.where(df.notna(), None).values.tolist()
        remove_filters(ws)

        target = ws.Range(
            ws.Cells(2, 1),
            ws.Cells(len(values) + 1, len(values[0]))
        )

        target.Value = values

    except IndexError:
        raise DataCleaningError("DataFrame is empty or has invalid structure")
    except ValueError as e:
        raise DataCleaningError(f"Invalid data structure: {e}")
    except pywintypes.com_error as e:
        raise ExcelConnectionError(f"Excel connection error: {e}")
    except Exception as e:
        raise DataCleaningError(f"Unexpected error: {e}")
