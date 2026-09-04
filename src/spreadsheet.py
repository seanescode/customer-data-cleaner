import logging
import pywintypes
import win32com.client as win32

from error_handler import handle_errors
from exceptions import ExcelConnectionError

FILTER_VALUES_OPERATOR = 7
SORT_ASCENDING = 1
HEADER_PRESENT = 1
SORT_TOP_TO_BOTTOM = 1
XL_H_ALIGN_LEFT= 2

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
    log = logging.getLogger("customer_data_cleaner")
    try:
        log.info("Launching Excel application")
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        log.info("Excel application launched successfully")
        return excel
    except pywintypes.com_error as e:
        log.error(f"Failed to launch spreadsheet: {e}")
        raise ExcelConnectionError(f"Failed to launch spreadsheet: {e}")


@handle_errors
def remove_filters(ws: object) -> None:
    log = logging.getLogger("customer_data_cleaner")
    try:
        if ws.FilterMode:
            log.info("Removing existing filters from worksheet")
            ws.ShowAllData()
            log.info("Filters removed successfully")
    except pywintypes.com_error as e:
        log.error(f"Excel connection error: {e}")
        raise ExcelConnectionError(f"Excel connection error: {e}")


@handle_errors
def open_spreadsheet(excel: object, file_path: str, worksheet_name: str) -> tuple[object, object]:
    log = logging.getLogger("customer_data_cleaner")
    try:
        log.info(f"Opening spreadsheet: {file_path}")
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Worksheets(worksheet_name)
        log.info(f"Worksheet '{worksheet_name}' opened successfully")
        return wb, ws
    except pywintypes.com_error as e:
        log.error(f"Failed to open spreadsheet: {e}")
        raise ExcelConnectionError(f"Failed to open spreadsheet: {e}")


def create_new_workbook(excel: object, file_path: str, worksheet_name: str) -> tuple[object, object]:
    log = logging.getLogger("customer_data_cleaner")
    log.info(f"Creating new workbook: {file_path}")
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = worksheet_name
    wb.SaveAs(file_path)
    log.info(f"New workbook created and saved as: {file_path}")
    return wb, ws


def input_values_to_spreadsheet(df: object, ws: object) -> None:
    log = logging.getLogger("customer_data_cleaner")
    log.info(f"Writing {len(df)} records to spreadsheet")
    headers = df.columns.tolist()
    values = df.where(df.notna(), "").values.tolist()
    ws.Range(
        ws.Cells(1, 1),
        ws.Cells(1, len(headers))
    ).Value = headers

    ws.Range(
        ws.Cells(2, 1),
        ws.Cells(len(values) + 1, len(values[0]))
    ).Value = values
    log.info("Data written to spreadsheet successfully")


def auto_fit_columns_and_rows(ws: object) -> None:
    log = logging.getLogger("customer_data_cleaner")
    log.info("Auto-fitting columns and rows")
    used_range = ws.UsedRange
    used_range.EntireColumn.AutoFit()
    used_range.EntireRow.AutoFit()
    log.info("Auto-fit completed")

def left_align_text(ws: object) -> None:
    log = logging.getLogger("customer_data_cleaner")
    log.info("Left-aligning text in spreadsheet")
    used_range = ws.UsedRange
    used_range.Columns.HorizontalAlignment = XL_H_ALIGN_LEFT
    log.info("Text alignment completed")
