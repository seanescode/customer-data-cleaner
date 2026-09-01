import tomllib
from pathlib import Path

from error_handler import handle_errors
from exceptions import ConfigNotFoundError, ConfigParseError, ConfigKeyError


@handle_errors
def load_config() -> dict:
    try:
        config_path = Path(__file__).parent.parent / "config.toml"
        with open(config_path, "rb") as f:
            return tomllib.load(f)
    except FileNotFoundError:
        raise ConfigNotFoundError(str(config_path))
    except tomllib.TOMLDecodeError as e:
        raise ConfigParseError(str(e))


@handle_errors
def get_spreadsheet_paths_and_worksheet(config: dict) -> tuple[str, str, str]:
    try:
        input_file_path = config["spreadsheet"]["input_file_path"]
        output_file_path = config["spreadsheet"]["output_file_path"]
        output_worksheet_name = config["spreadsheet"]["output_worksheet_name"]

        return input_file_path, output_file_path, output_worksheet_name
    except KeyError as e:
        raise ConfigKeyError(str(e))
