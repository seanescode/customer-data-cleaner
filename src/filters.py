import logging
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
    EmailValidationError,
)


@handle_errors
def filter_to_invalid_emails(
    file_path: str,
    ws: object,
    email_column: str,
    no_records_found_msg: Callable[[str], None],
) -> None:
    """Filter the spreadsheet to show customers with invalid emails."""
    log = logging.getLogger("customer_data_cleaner")
    try:
        log.info(f"Filtering to invalid emails in column: {email_column}")
        spreadsheet.remove_filters(ws)
        df = dataframe.get_dataframe(file_path)
        invalid_emails = dataframe.get_invalid_emails(df, email_column)

        if not invalid_emails:
            log.info("No invalid emails found")
            no_records_found_msg("No invalid emails were found.")
            return

        log.info(f"Found {len(invalid_emails)} invalid emails")
        email_column_number = spreadsheet.get_column_number_of_heading(
            ws,
            email_column,
        )
        spreadsheet.apply_filter(ws, email_column_number, invalid_emails)
        log.info("Filter applied successfully")

    except FileNotFoundError:
        log.error(f"Spreadsheet not found: {file_path}")
        raise SpreadsheetNotFoundError(file_path)
    except (PermissionError, ValueError, IsADirectoryError) as e:
        log.error(f"Spreadsheet access error: {e}")
        raise SpreadsheetAccessError(str(e))
    except KeyError:
        log.error(f"Column not found: {email_column}")
        raise ColumnNotFoundError(email_column)
    except pywintypes.com_error:
        log.error("Excel connection error")
        raise ExcelConnectionError()
    except EmailValidationError:
        raise
    except Exception as e:
        log.error(f"Email validation error: {e}")
        raise EmailValidationError(str(e))


@handle_errors
def filter_duplicate_customers(
    file_path: str,
    ws: object,
    name_column: str,
    customer_id_column: str,
    no_records_found_msg: Callable[[str], None],
) -> None:
    """Filter the spreadsheet to show duplicate customers."""
    log = logging.getLogger("customer_data_cleaner")
    try:
        log.info("Filtering to duplicate customers")
        spreadsheet.remove_filters(ws)
        df = dataframe.get_dataframe(file_path)
        duplicate_customers = dataframe.get_duplicate_customers(df)

        duplicate_customer_ids = []
        for cust_id in duplicate_customers["customer_id"]:
            duplicate_customer_ids.append(str(cust_id))

        if not duplicate_customer_ids:
            log.info("No duplicate customers found")
            no_records_found_msg("No duplicate customers were found.")
            return

        log.info(f"Found {len(duplicate_customer_ids)} duplicate customers")

        # Sort by names in alphabetical order
        log.info(f"Sorting by column: {name_column}")
        spreadsheet.sort_rows(ws, name_column)

        # Filter to show only duplicate customers
        customer_id_column_number = spreadsheet.get_column_number_of_heading(
            ws,
            customer_id_column,
        )
        spreadsheet.apply_filter(
            ws,
            customer_id_column_number,
            duplicate_customer_ids,
        )
        log.info("Filter applied successfully")

    except FileNotFoundError:
        log.error(f"Spreadsheet not found: {file_path}")
        raise SpreadsheetNotFoundError(file_path)
    except (PermissionError, ValueError, IsADirectoryError) as e:
        log.error(f"Spreadsheet access error: {e}")
        raise SpreadsheetAccessError(str(e))
    except ColumnNotFoundError:
        raise
    except KeyError:
        log.error(f"Column not found: {customer_id_column}")
        raise ColumnNotFoundError(customer_id_column)
    except pywintypes.com_error:
        log.error("Excel connection error")
        raise ExcelConnectionError()
    except DuplicateDetectionError:
        raise
    except Exception as e:
        log.error(f"Duplicate detection error: {e}")
        raise DuplicateDetectionError(str(e))
