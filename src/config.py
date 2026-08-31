import tomllib
from pathlib import Path

import dialogs


def load_config():
    try:
        config_path = Path(__file__).parent.parent / "config.toml"
        with open(config_path, "rb") as f:
            return tomllib.load(f)
    except FileNotFoundError:
        dialogs.show_error_dialog("Error", "Config file not found: config.toml")
        raise
    except tomllib.TOMLDecodeError as e:
        dialogs.show_error_dialog("Error", f"Invalid TOML in config file: {e}")
        raise


def get_spreadsheet_path_and_worksheet(config):
    try:
        return config["spreadsheet"]["file_path"], config["spreadsheet"]["worksheet_name"]
    except KeyError as e:
        dialogs.show_error_dialog("Error", f"Missing required config key: {e}")
        raise
