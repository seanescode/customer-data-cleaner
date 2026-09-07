import pandas
from email_validator import EmailNotValidError, validate_email

from error_handler import handle_errors
from exceptions import (
    SpreadsheetNotFoundError,
    SpreadsheetAccessError,
    DuplicateDetectionError,
    ColumnNotFoundError,
)


@handle_errors
def get_dataframe(file_path: str) -> pandas.DataFrame:
    try:
        return pandas.read_excel(file_path)
    except FileNotFoundError:
        raise SpreadsheetNotFoundError(file_path)
    except (PermissionError, ValueError, IsADirectoryError) as e:
        raise SpreadsheetAccessError(str(e))


@handle_errors
def get_duplicate_customers(df: pandas.DataFrame) -> pandas.DataFrame:
    try:
        required_columns = ['name', 'email', 'phone', 'address', 'postcode']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ColumnNotFoundError(f"Missing required columns: {missing_columns}")

        # find duplicates by different table data - to have multiple check points to help catch all duplicates
        dup_email = df.duplicated(subset=['name', 'email'], keep=False)
        dup_phone = df.duplicated(subset=['name', 'phone'], keep=False)
        dup_address = df.duplicated(subset=['name', 'address'], keep=False)
        dup_postcode = df.duplicated(subset=['name', 'postcode'], keep=False)

        # Combine all duplicate checks using OR logic
        all_duplicates = df[dup_email | dup_phone | dup_address | dup_postcode]
        return all_duplicates
    except ColumnNotFoundError:
        raise
    except Exception as e:
        raise DuplicateDetectionError(str(e))


@handle_errors
def get_invalid_emails(df: pandas.DataFrame, email_column: str) -> list[str]:
    try:
        invalid_emails = []
        for email in df[email_column]:
            # Empty Excel cells are read by pandas as NaN (a float). The email validator only accepts strings or bytes.
            if pandas.isna(email):
                continue
            try:
                validate_email(str(email).strip(), check_deliverability=False)
            except EmailNotValidError:
                invalid_emails.append(str(email))
        return invalid_emails
    except KeyError:
        raise ColumnNotFoundError(email_column)
