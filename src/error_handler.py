import logging
import dialogs
from functools import wraps
from typing import Any, Callable

from exceptions import (
    ConfigNotFoundError,
    ConfigParseError,
    ConfigKeyError,
    SpreadsheetNotFoundError,
    SpreadsheetAccessError,
    DataCleaningError,
    ExcelConnectionError,
    DuplicateDetectionError,
    EmailValidationError,
    ColumnNotFoundError,
    DialogSetupError,
    CustomerDataCleanerError,
)


def handle_errors(func: Callable[..., Any]) -> Callable[..., Any]:
    """Decorator to handle custom exceptions and show appropriate dialog messages."""
    @wraps(func)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        log = logging.getLogger("customer_data_cleaner")
        try:
            return func(*args, **kwargs)

        except (
            ConfigNotFoundError,
            ConfigParseError,
            ConfigKeyError,
        ) as e:
            log.error(f"Configuration error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Config Error", str(e))

        except (
            SpreadsheetNotFoundError,
            SpreadsheetAccessError,
        ) as e:
            log.error(f"Spreadsheet error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Spreadsheet Error", str(e))

        except DataCleaningError as e:
            log.error(f"Data cleaning error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Data Error", str(e))

        except ExcelConnectionError as e:
            log.error(f"Excel connection error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Excel Error", str(e))

        except DuplicateDetectionError as e:
            log.error(f"Duplicate detection error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Duplicate Detection Error", str(e))

        except EmailValidationError as e:
            log.error(f"Email validation error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Email Validation Error", str(e))

        except ColumnNotFoundError as e:
            log.error(f"Column not found error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Column Error", str(e))

        except DialogSetupError as e:
            log.error(f"Dialog setup error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Dialog Error", str(e))

        except CustomerDataCleanerError as e:
            log.error(f"Customer data cleaner error in {func.__name__}: {e}")
            dialogs.show_error_dialog("Error", str(e))

        except Exception as e:
            log.error(f"Unexpected error in {func.__name__}: {e}", exc_info=True)
            dialogs.show_error_dialog("Error", f"Unexpected error: {e}")

    return wrapper