import tomllib
from pathlib import Path

import pandas
import pywintypes
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


def setup_and_run_dialog(excel, wb, file_loc, ws, config):
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


def load_config():
    config_path = Path(__file__).parent.parent / "config.toml"
    with open(config_path, "rb") as f:
        return tomllib.load(f)


def get_spreadsheet(config):
    return config["spreadsheet"]["file_path"], config["spreadsheet"]["worksheet_name"]


def get_dataframe(file_loc):
    return pandas.read_excel(file_loc)


def clean_data(df, config):
    columns_config = config["columns"]
    cleanup.clean_column(df, columns_config["name"], cleanup.clean_name)
    cleanup.clean_column(df, columns_config["email"], cleanup.clean_email)
    cleanup.clean_column(df, columns_config["phone"], cleanup.clean_phone_number)
    cleanup.clean_column(df, columns_config["address"], cleanup.clean_address)
    cleanup.clean_column(df, columns_config["city"], cleanup.clean_city)
    cleanup.clean_column(df, columns_config["postcode"], cleanup.clean_postcode)


def launch_spreadsheet_app():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    return excel


def open_spreadsheet(excel, file_loc, worksheet_name):
    wb = excel.Workbooks.Open(file_loc)
    ws = wb.Worksheets(worksheet_name)
    return wb, ws


def apply_dataframe_to_spreadsheet(df, ws):
    cleanup.update_spreadsheet_from_dataframe(df, ws)


def main():
    try:
        config = load_config()
        file_loc, worksheet_name = get_spreadsheet(config)
        df = get_dataframe(file_loc)
        clean_data(df, config)
        excel = launch_spreadsheet_app()
        wb, ws = open_spreadsheet(excel, file_loc, worksheet_name)
        apply_dataframe_to_spreadsheet(df, ws)
        setup_and_run_dialog(excel, wb, file_loc, ws, config)

    except FileNotFoundError as e:
        dialogs.show_error_dialog("Error", f"File not found: {e}")
    except tomllib.TOMLDecodeError as e:
        dialogs.show_error_dialog("Error", f"Invalid TOML in config file: {e}")
    except KeyError as e:
        dialogs.show_error_dialog("Error", f"Missing required config key: {e}")
    except pywintypes.com_error as e:
        dialogs.show_error_dialog("Error", f"Excel COM error: {e}")
    except ValueError as e:
        dialogs.show_error_dialog("Error", f"Data conversion error: {e}")
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Unexpected error: {e}")


if __name__ == '__main__':
    main()
