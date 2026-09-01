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
    CustomerDataCleanerError
)


def handle_errors(func: Callable[..., Any]) -> Callable[..., Any]:
    """Decorator to handle custom exceptions and show appropriate dialog messages."""
    @wraps(func)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return func(*args, **kwargs)
        except ConfigNotFoundError as e:
            dialogs.show_error_dialog("Config Error", str(e))
        except ConfigParseError as e:
            dialogs.show_error_dialog("Config Error", str(e))
        except ConfigKeyError as e:
            dialogs.show_error_dialog("Config Error", str(e))
        except SpreadsheetNotFoundError as e:
            dialogs.show_error_dialog("Spreadsheet Error", str(e))
        except SpreadsheetAccessError as e:
            dialogs.show_error_dialog("Spreadsheet Error", str(e))
        except DataCleaningError as e:
            dialogs.show_error_dialog("Data Error", str(e))
        except ExcelConnectionError as e:
            dialogs.show_error_dialog("Excel Error", str(e))
        except DuplicateDetectionError as e:
            dialogs.show_error_dialog("Duplicate Detection Error", str(e))
        except EmailValidationError as e:
            dialogs.show_error_dialog("Email Validation Error", str(e))
        except ColumnNotFoundError as e:
            dialogs.show_error_dialog("Column Error", str(e))
        except DialogSetupError as e:
            dialogs.show_error_dialog("Dialog Error", str(e))
        except CustomerDataCleanerError as e:
            dialogs.show_error_dialog("Error", str(e))
        except Exception as e:
            dialogs.show_error_dialog("Error", f"Unexpected error: {e}")

    return wrapper
