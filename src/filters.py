import pandas
import pywintypes
from email_validator import EmailNotValidError, validate_email
import cleanup
import dialogs

# Excel AutoFilter constants
CUSTOMER_ID_COLUMN = 1
FILTER_VALUES_OPERATOR = 7


def filter_to_invalid_emails(file_loc, ws, email_column, show_message):
    try:
        df = pandas.read_excel(file_loc)
        cleanup.remove_filters(ws)

        invalid_customer_ids = []
        for index, email in enumerate(df[email_column]):
            # Empty Excel cells are read by pandas as NaN (a float).  The email
            # validator only accepts strings or bytes, so include blank emails in
            # the invalid results instead of passing NaN to the validator.
            if pandas.isna(email):
                invalid_customer_ids.append(
                    str(df.iloc[index]["customer_id"])
                )
                continue

            try:
                validate_email(
                    str(email).strip(),
                    check_deliverability=False
                )
            except EmailNotValidError:
                invalid_customer_ids.append(
                    str(df.iloc[index]["customer_id"])
                )

        if not invalid_customer_ids:
            cleanup.remove_filters(ws)
            show_message("No invalid emails were found.")
            return

        ws.UsedRange.AutoFilter(
            Field=CUSTOMER_ID_COLUMN,
            Criteria1=tuple(invalid_customer_ids),
            Operator=FILTER_VALUES_OPERATOR
        )

    except FileNotFoundError:
        show_message("Spreadsheet not found. Please check the file location.")
    except (PermissionError, ValueError, IsADirectoryError) as e:
        show_message(f"Error reading Excel file: {e}")
    except KeyError:
        show_message(f"Column '{email_column}' not found in spreadsheet.")
    except pywintypes.com_error:
        show_message("Excel connection error. Please ensure Excel is still open.")
    except Exception as e:
        show_message(f"Unexpected error: {e}")


def get_duplicate_customers(df):
    try:
        required_columns = ['name', 'email', 'phone', 'address', 'postcode']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise KeyError(f"Missing required columns: {missing_columns}")

        # find duplicates by different table data - to have multiple check points to help catch all duplicates
        dup_email = df.duplicated(subset=['name', 'email'], keep=False)
        dup_phone = df.duplicated(subset=['name', 'phone'], keep=False)
        dup_address_first_line = df.duplicated(subset=['name', 'address'], keep=False)
        dup_postcode = df.duplicated(subset=['name', 'postcode'], keep=False)

        # Combine all duplicate variables using "OR" logic (|) to show all them
        all_duplicates = df[dup_email | dup_phone | dup_address_first_line | dup_postcode]
        return all_duplicates
    except KeyError as e:
        raise KeyError(f"Missing required columns for duplicate detection: {e}")
    except Exception as e:
        raise Exception(f"Error detecting duplicates: {e}")


def filter_duplicate_customers(file_loc, ws, show_message):
    try:
        duplicate_customers = get_duplicate_customers(pandas.read_excel(file_loc))
        cleanup.remove_filters(ws)

        duplicate_customer_ids = []
        for index, cust_id in enumerate(duplicate_customers["customer_id"]):
            duplicate_customer_ids.append(str(cust_id))

        if not duplicate_customer_ids:
            cleanup.remove_filters(ws)
            show_message("No duplicate customers were found.")
            return

        ws.UsedRange.AutoFilter(
            Field=CUSTOMER_ID_COLUMN,
            Criteria1=tuple(duplicate_customer_ids),
            Operator=FILTER_VALUES_OPERATOR
        )

    except FileNotFoundError:
        show_message("Spreadsheet not found. Please check the file location.")
    except (PermissionError, ValueError, IsADirectoryError) as e:
        show_message(f"Error reading Excel file: {e}")
    except KeyError as e:
        show_message(f"Column not found: {e}")
    except pywintypes.com_error:
        show_message("Excel connection error. Please ensure Excel is still open.")
    except Exception as e:
        show_message(f"Unexpected error: {e}")
