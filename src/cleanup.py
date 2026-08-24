from email_validator import validate_email, EmailNotValidError
import phonenumbers
import pandas
import time


def clean_name(customer_name):
    customer_name = " ".join(customer_name.split())
    customer_name = customer_name.title()
    return customer_name


def clean_email(email):
    email = str(email)
    email = email.lower()
    email = email.replace(" ", "")
    return email


def clean_phone_number(phone_number):
    phone_number = str(phone_number)

    phone_number = phone_number.replace(" ", "")
    phone_number = phone_number.replace("-", "")

    # remove invalid short phone numbers
    if len(phone_number) < 5:
        return ""

    # remove Irish prefix code
    if phone_number.startswith("+353"):
        phone_number = phone_number[4:]

    # clean mobile numbers
    if phone_number.startswith(("83", "85", "86", "87", "89")) and len(phone_number) == 9:
        phone_number = "0" + phone_number

    if phone_number.startswith(("083", "085", "086", "087", "089"))and len(phone_number) == 10:
        return phone_number[0:3] + " " + phone_number[3:6] + " " + phone_number[6:10]

    return ""

def clean_address(address):
    address = address.title()
    address = " ".join(address.split())
    return address

def clean_city(city):
    city = city.title()
    city = city.replace(" ", "")
    return city

def clean_postcode(postcode):
    postcode = str(postcode)
    postcode = postcode.replace(" ", "")
    postcode = postcode.upper()
    return postcode[:3] + " " + postcode[3:]


def clean_column(df, column, cleaning_function):
    if column in df.columns:
        df[column] = df[column].fillna("").apply(cleaning_function)


def remove_filters(ws):
    if ws.FilterMode:
        ws.ShowAllData()

def filter_to_invalid_emails(df, ws, email_column):

    remove_filters(ws)

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

    filter_customer_ids = invalid_customer_ids

    if not filter_customer_ids:
        # No invalid emails, so deliberately show no customers
        filter_customer_ids = ["NO_CUSTOMER_MATCH"]

    ws.UsedRange.AutoFilter(
        Field=1,
        Criteria1=tuple(filter_customer_ids),
        Operator=7
    )


def filter_to_duplicates(customer_ids, ws, header_name):

    remove_filters(ws)

    if not customer_ids:
        return
    used_range = ws.UsedRange
    # Apply filter to column A - Field=1 means the first column of used_range

    used_range.AutoFilter(
        Field=1,
        Criteria1=tuple(customer_ids),
        Operator=7
    )

    column_of_name = find_column(ws, header_name)

    if column_of_name is None:
        raise ValueError(f"Column '{header_name}' not found")

    starting_cell_to_sort = ws.Cells(2, column_of_name)

    last_row = find_last_row(ws)
    last_cell_to_sort = ws.Cells(last_row, column_of_name)

    #sort_range = ws.Range("A1:G191")
    ws.Sort.SortFields.Clear()
    ws.Sort.SortFields.Add(
        Key=ws.Range(starting_cell_to_sort, last_cell_to_sort),
        SortOn=0,
        Order=1
    )
    ws.Sort.SetRange(used_range)
    ws.Sort.Header = 1
    ws.Sort.Orientation = 1
    #ws.Sort.Apply()


def find_column(ws, header_name):
    last_column = ws.UsedRange.Columns.Count

    for column in range(1, last_column + 1):
        value = ws.Cells(1, column).Value

        if value is not None and str(value).strip().lower() == header_name.lower():
            return column

    return None

def find_last_row(ws):
    return ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

def update_excel_from_dataframe(df, ws):
    values = df.where(df.notna(), None).values.tolist()
    remove_filters(ws)
    ws.Range(
        ws.Cells(2, 1),
        ws.Cells(len(values) + 1, len(values[0]))
    ).Value = values

    target = ws.Range(
        ws.Cells(2, 1),
        ws.Cells(len(values) + 1, len(values[0]))
    )

    target.Value = values

    print("First five rows read back from Excel:")
    for row in target.Value[:5]:
        print(row)

def find_duplicate_customers(df):
    # find duplicates by different table data - to have multiple check points to help catch all duplicates
    dup_email = df.duplicated(subset=['name', 'email'], keep=False)
    dup_phone = df.duplicated(subset=['name', 'phone'], keep=False)
    dup_address_first_line = df.duplicated(subset=['name', 'address'], keep=False)
    dup_postcode = df.duplicated(subset=['name', 'postcode'], keep=False)

    # Combine all duplicate variables using "OR" logic (|) to show all them
    all_duplicates = df[dup_email | dup_phone | dup_address_first_line | dup_postcode]
    return all_duplicates
