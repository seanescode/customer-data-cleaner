import pywintypes
import tomllib
from pathlib import Path

import pandas
import win32com.client as win32

import cleanup
import dialogs


def handle_duplicate_filter(file_loc, ws, root):
    cleanup.filter_duplicate_customers(
        file_loc,
        ws,
        lambda text: dialogs.show_temporary_message(root, text)
    )


def handle_email_filter(file_loc, ws, email_column, root):
    cleanup.filter_to_invalid_emails(
        file_loc,
        ws,
        email_column,
        lambda text: dialogs.show_temporary_message(root, text)
    )


def handle_remove_filters(ws):
    cleanup.remove_filters(ws)


def main():

    # Load configuration
    config_path = Path(__file__).parent.parent / "config.toml"
    try:
        with open(config_path, "rb") as f:
            config = tomllib.load(f)
    except FileNotFoundError:
        dialogs.show_error_dialog("Error", f"Config file not found at {config_path}")
        return
    except tomllib.TOMLDecodeError as e:
        dialogs.show_error_dialog("Error", f"Invalid TOML in config file: {e}")
        return
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Error loading config: {e}")
        return

    try:
        file_loc = config["spreadsheet"]["file_path"]
        worksheet_name = config["spreadsheet"]["worksheet_name"]
    except KeyError as e:
        dialogs.show_error_dialog("Error", f"Missing required config key: {e}")
        return

    try:
        df = pandas.read_excel(file_loc)
    except FileNotFoundError:
        dialogs.show_error_dialog("Error", f"Spreadsheet file not found at {file_loc}")
        return
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Error reading spreadsheet: {e}")
        return

    # clean data
    columns_config = config["columns"]
    cleanup.clean_column(df, columns_config["name"], cleanup.clean_name)
    cleanup.clean_column(df, columns_config["email"], cleanup.clean_email)
    cleanup.clean_column(df, columns_config["phone"], cleanup.clean_phone_number)
    cleanup.clean_column(df, columns_config["address"], cleanup.clean_address)
    cleanup.clean_column(df, columns_config["city"], cleanup.clean_city)
    cleanup.clean_column(df, columns_config["postcode"], cleanup.clean_postcode)

    # Open spreadsheet
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
    except pywintypes.com_error as e:
        dialogs.show_error_dialog("Error", f"Excel COM error: {e}")
        return
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Unexpected error launching Excel: {e}")
        return

    try:
        wb = excel.Workbooks.Open(file_loc)
        ws = wb.Worksheets(worksheet_name)
    except pywintypes.com_error as e:
        dialogs.show_error_dialog("Error", f"Excel COM error opening spreadsheet: {e}")
        return
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Unexpected error opening spreadsheet: {e}")
        return

    try:
        cleanup.update_spreadsheet_from_dataframe(df, ws)
    except pywintypes.com_error as e:
        dialogs.show_error_dialog("Error", f"Excel COM error updating spreadsheet: {e}")
        return
    except ValueError as e:
        dialogs.show_error_dialog("Error", f"Data conversion error: {e}")
        return
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Unexpected error updating spreadsheet: {e}")
        return

    root = dialogs.create_dialog_root()

    dialogs.show_cleanup_panel(
        root,
        excel.Hwnd,
        excel,
        wb,
        lambda: handle_duplicate_filter(file_loc, ws, root),
        lambda: handle_email_filter(file_loc, ws, config["columns"]["email"], root),
        lambda: handle_remove_filters(ws)
    )
    root.mainloop()

if __name__ == '__main__':
    main()
