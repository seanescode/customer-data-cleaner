import os.path

import pywintypes
import win32com.client as win32

from error_handler import handle_errors
from exceptions import ExcelConnectionError, DataCleaningError

FILTER_VALUES_OPERATOR = 7
SORT_ASCENDING = 1
HEADER_PRESENT = 1
SORT_TOP_TO_BOTTOM = 1


def get_column_number_of_heading(ws: object, column_name: str) -> int | None:
    for cell in ws.UsedRange.Rows(1).Cells:
        if cell.Value == column_name:
            return cell.Column
    return None


def apply_filter(ws: object, column_number: int, filter_values: list[str]) -> None:

    used = ws.UsedRange
    used.AutoFilter(
        # - used.Column + 1 ensures this would work if ever starting cell changed from A1
        Field = column_number - used.Column + 1,
        Criteria1 = tuple(filter_values),
        Operator = FILTER_VALUES_OPERATOR
    )


def sort_rows(ws: object, column_header: str) -> None:
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
def launch_spreadsheet_app() -> object:
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        return excel
    except pywintypes.com_error as e:
        raise ExcelConnectionError(f"Failed to launch spreadsheet: {e}")


@handle_errors
def open_spreadsheet(excel: object, file_path: str, worksheet_name: str) -> tuple[object, object]:
    try:
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Worksheets(worksheet_name)
        return wb, ws
    except pywintypes.com_error as e:
        raise ExcelConnectionError(f"Failed to open spreadsheet: {e}")


@handle_errors
def remove_filters(ws: object) -> None:
    try:
        if ws.FilterMode:
            ws.ShowAllData()
    except pywintypes.com_error as e:
        raise ExcelConnectionError(f"Excel connection error: {e}")

def input_values_to_spreadsheet(df:object, ws: object) -> None:
    headers = df.columns.tolist()
    values = df.where(df.notna(), None).values.tolist()
    ws.Range(
        ws.Cells(1, 1),
        ws.Cells(1, len(headers))
    ).Value = headers

    ws.Range(
        ws.Cells(2, 1),
        ws.Cells(len(values) + 1, len(values[0]))
    ).Value = values

def auto_fit_columns_and_rows(ws: object) -> None:
    used_range = ws.UsedRange
    used_range.EntireColumn.AutoFit()
    used_range.EntireRow.AutoFit()

