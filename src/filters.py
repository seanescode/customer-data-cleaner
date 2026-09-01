import pywintypes
from typing import Callable
import spreadsheet
import dataframe

from error_handler import handle_errors
from exceptions import (
    SpreadsheetNotFoundError,
    SpreadsheetAccessError,
    ExcelConnectionError,
    ColumnNotFoundError,
    DuplicateDetectionError,
    EmailValidationError
)


@handle_errors
def filter_to_invalid_emails(file_path: str,
                             ws: object,
                             email_column: str,
                             no_records_found_msg: Callable[[str], None]) -> None:
    try:
        spreadsheet.remove_filters(ws)
        df = dataframe.get_dataframe(file_path)
        invalid_emails = dataframe.get_invalid_emails(df, email_column)

        if not invalid_emails:
            no_records_found_msg("No invalid emails were found.")
            return

        email_column_number = spreadsheet.get_column_number_of_heading(ws, email_column)
        spreadsheet.apply_filter(ws, email_column_number, invalid_emails)

    except FileNotFoundError:
        raise SpreadsheetNotFoundError(file_path)
    except (PermissionError, ValueError, IsADirectoryError) as e:
        raise SpreadsheetAccessError(str(e))
    except KeyError:
        raise ColumnNotFoundError(email_column)
    except pywintypes.com_error:
        raise ExcelConnectionError()
    except Exception as e:
        raise EmailValidationError(str(e))


@handle_errors
def filter_duplicate_customers(file_path: str,
                               ws: object,
                               name_column: str,
                               customer_id_column: str,
                               no_records_found_msg: Callable[[str], None]) -> None:
    try:
        spreadsheet.remove_filters(ws)
        df = dataframe.get_dataframe(file_path)
        duplicate_customers = dataframe.get_duplicate_customers(df)

        duplicate_customer_ids = []
        for cust_id in duplicate_customers["customer_id"]:
            duplicate_customer_ids.append(str(cust_id))
        
        if not duplicate_customer_ids:
            no_records_found_msg("No duplicate customers were found.")
            return

        # sort by names in alphabetical order
        spreadsheet.sort_rows(ws, name_column)
        # filter to show only duplicate customers
        customer_id_column_number = spreadsheet.get_column_number_of_heading(ws, customer_id_column)
        spreadsheet.apply_filter(ws, customer_id_column_number, duplicate_customer_ids)

    except FileNotFoundError:
        raise SpreadsheetNotFoundError(file_path)
    except (PermissionError, ValueError, IsADirectoryError) as e:
        raise SpreadsheetAccessError(str(e))
    except KeyError as e:
        raise ColumnNotFoundError(str(e))
    except pywintypes.com_error:
        raise ExcelConnectionError()
    except Exception as e:
        raise DuplicateDetectionError(str(e))
