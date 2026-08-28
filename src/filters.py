import pandas
import pywintypes
from email_validator import EmailNotValidError, validate_email

import spreadsheet


# Excel AutoFilter constants
CUSTOMER_ID_COLUMN = 1
FILTER_VALUES_OPERATOR = 7
SORT_ASCENDING = 1
HEADER_PRESENT = 1
SORT_TOP_TO_BOTTOM = 1


def filter_to_invalid_emails(file_loc, ws, email_column, no_records_found_msg):
    try:
        df = pandas.read_excel(file_loc)
        spreadsheet.remove_filters(ws)

        invalid_emails = []
        for index, email in enumerate(df[email_column]):
            # Empty Excel cells are read by pandas as NaN (a float).  The email
            # validator only accepts strings or bytes, so include blank emails in
            # the invalid results instead of passing NaN to the validator.
            if pandas.isna(email):
                invalid_emails.append(
                    str(df.iloc[index][email_column])
                )
                continue

            try:
                validate_email(
                    str(email).strip(),
                    check_deliverability=False
                )
            except EmailNotValidError:
                invalid_emails.append(
                    str(df.iloc[index][email_column])
                )

        if not invalid_emails:
            spreadsheet.remove_filters(ws)
            no_records_found_msg("No invalid emails were found.")
            return

        used = ws.UsedRange
        email_header_column = None
        for cell in used.Rows(1).Cells:
            if cell.Value == email_column:
                email_header_column = cell.Column
                break

        used.AutoFilter(
            Field=email_header_column - used.Column + 1,
            # - used.Column + 1 ensures this would work if ever starting cell changed from A1
            Criteria1=tuple(invalid_emails),
            Operator=FILTER_VALUES_OPERATOR
        )

    except FileNotFoundError:
        no_records_found_msg("Spreadsheet not found. Please check the file location.")
    except (PermissionError, ValueError, IsADirectoryError) as e:
        no_records_found_msg(f"Error reading Excel file: {e}")
    except KeyError:
        no_records_found_msg(f"Column '{email_column}' not found in spreadsheet.")
    except pywintypes.com_error:
        no_records_found_msg("Excel connection error. Please ensure Excel is still open.")
    except Exception as e:
        no_records_found_msg(f"Unexpected error: {e}")


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


def filter_duplicate_customers(file_loc, ws, name_column, no_records_found_msg):
    try:
        duplicate_customers = get_duplicate_customers(pandas.read_excel(file_loc))
        spreadsheet.remove_filters(ws)

        duplicate_customer_ids = []
        for index, cust_id in enumerate(duplicate_customers["customer_id"]):
            duplicate_customer_ids.append(str(cust_id))

        if not duplicate_customer_ids:
            spreadsheet.remove_filters(ws)
            no_records_found_msg("No duplicate customers were found.")
            return

        # sort by names in alphabetical order
        used = ws.UsedRange
        for cell in used.Rows(1).Cells:

            if cell.Value == name_column:
                # Get the first data cell
                key = ws.Cells(cell.Row + 1, cell.Column)
                # Explicitly tell Excel to sort the whole range
                used.Sort(
                    Key1=key,
                    Order1=SORT_ASCENDING,
                    Header=HEADER_PRESENT,  # xlYes
                    Orientation=SORT_TOP_TO_BOTTOM  # xlTopToBottom
                )
                break

        # filter to show only duplicate customers
        ws.UsedRange.AutoFilter(
            Field=CUSTOMER_ID_COLUMN,
            Criteria1=tuple(duplicate_customer_ids),
            Operator=FILTER_VALUES_OPERATOR
        )

    except FileNotFoundError:
        no_records_found_msg("Spreadsheet not found. Please check the file location.")
    except (PermissionError, ValueError, IsADirectoryError) as e:
        no_records_found_msg(f"Error reading Excel file: {e}")
    except KeyError as e:
        no_records_found_msg(f"Column not found: {e}")
    except pywintypes.com_error:
        no_records_found_msg("Excel connection error. Please ensure Excel is still open.")
    except Exception as e:
        no_records_found_msg(f"Unexpected error: {e}")
