import pandas
import cleanup
import spreadsheet
import dialogs


def get_dataframe(file_loc):
    try:
        return pandas.read_excel(file_loc)
    except FileNotFoundError:
        dialogs.show_error_dialog("Error", f"Spreadsheet not found: {file_loc}")
        raise
    except (PermissionError, ValueError, IsADirectoryError) as e:
        dialogs.show_error_dialog("Error", f"Error reading Excel file: {e}")
        raise


def apply_dataframe_to_spreadsheet(df, ws):
    try:
        spreadsheet.update_spreadsheet_from_dataframe(df, ws)
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Failed to apply data to spreadsheet: {e}")
