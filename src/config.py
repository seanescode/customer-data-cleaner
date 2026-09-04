import logging
import sys
import tomllib
from pathlib import Path

from error_handler import handle_errors
from exceptions import ConfigNotFoundError, ConfigParseError, ConfigKeyError
import logger


def get_config_path() -> Path:
    """Return the path to config.toml."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "config.toml"

    return Path(__file__).parent.parent / "config.toml"


@handle_errors
def load_config() -> dict:
    """Load the application configuration from the TOML file."""
    try:
        config_path = get_config_path()
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


@handle_errors
def initialize_logging(config: dict) -> logging.Logger:
    """Initialize logging based on configuration."""
    try:
        logging_config = config.get("logging", {})
        enabled = logging_config.get("enabled", True)

        if not enabled:
            # Return a null logger if logging is disabled
            return logging.getLogger("null")

        log_file = logging_config.get("log_file", "logs/customer_data_cleaner.log")
        level_str = logging_config.get("level", "INFO")
        max_bytes = logging_config.get("max_bytes", 10 * 1024 * 1024)
        backup_count = logging_config.get("backup_count", 5)

        # Convert string level to logging constant
        level = getattr(logging, level_str.upper(), logging.INFO)

        return logger.setup_logger(
            name="customer_data_cleaner",
            log_file=log_file,
            level=level,
            max_bytes=max_bytes,
            backup_count=backup_count,
        )
    except KeyError as e:
        raise ConfigKeyError(str(e))
