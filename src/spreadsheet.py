import pywintypes
import win32com.client as win32

import dialogs

import pandas


def launch_spreadsheet_app():
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        return excel
    except pywintypes.com_error as e:
        dialogs.show_error_dialog("Error", f"Failed to launch spreadsheet: {e}")
        raise


def open_spreadsheet(excel, file_loc, worksheet_name):
    try:
        wb = excel.Workbooks.Open(file_loc)
        ws = wb.Worksheets(worksheet_name)
        return wb, ws
    except pywintypes.com_error as e:
        dialogs.show_error_dialog("Error", f"Failed to open spreadsheet: {e}")
        raise


def remove_filters(ws):
    try:
        if ws.FilterMode:
            ws.ShowAllData()
    except pywintypes.com_error as e:
        dialogs.show_error_dialog("Error", f"Excel connection error: {e}")


def get_dataframe(file_path):
    try:
        return pandas.read_excel(file_path)
    except FileNotFoundError:
        dialogs.show_error_dialog("Error", f"Spreadsheet not found: {file_path}")
        raise
    except (PermissionError, ValueError, IsADirectoryError) as e:
        dialogs.show_error_dialog("Error", f"Error reading Excel file: {e}")
        raise


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
        dialogs.show_error_dialog("Error", "DataFrame is empty or has invalid structure.")
    except ValueError as e:
        dialogs.show_error_dialog("Error", f"Invalid data structure: {e}")
    except pywintypes.com_error as e:
        dialogs.show_error_dialog("Error", f"Excel connection error: {e}")
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Unexpected error: {e}")
