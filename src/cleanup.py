import pandas

from email_validator import EmailNotValidError, validate_email

# Excel AutoFilter constants
CUSTOMER_ID_COLUMN = 1
FILTER_VALUES_OPERATOR = 7


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

    # clean mobile numbers (Excel takes out the zero so have to add back)
    if phone_number.startswith(("83", "85", "86", "87", "89")) and len(phone_number) == 9:
        phone_number = "0" + phone_number

    if phone_number.startswith(("083", "085", "086", "087", "089")) and len(phone_number) == 10:
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


def update_spreadsheet_from_dataframe(df, ws):
    values = df.where(df.notna(), None).values.tolist()
    remove_filters(ws)

    target = ws.Range(
        ws.Cells(2, 1),
        ws.Cells(len(values) + 1, len(values[0]))
    )

    target.Value = values


def remove_filters(ws):
    if ws.FilterMode:
        ws.ShowAllData()


def filter_to_invalid_emails(file_loc, ws, email_column, show_message=None):
    df = pandas.read_excel(file_loc)
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
        remove_filters(ws)
        if show_message is not None:
            show_message("No invalid emails were found.")
        return

    ws.UsedRange.AutoFilter(
        Field=CUSTOMER_ID_COLUMN,
        Criteria1=tuple(filter_customer_ids),
        Operator=FILTER_VALUES_OPERATOR
    )


def get_duplicate_customers(df):
    # find duplicates by different table data - to have multiple check points to help catch all duplicates
    dup_email = df.duplicated(subset=['name', 'email'], keep=False)
    dup_phone = df.duplicated(subset=['name', 'phone'], keep=False)
    dup_address_first_line = df.duplicated(subset=['name', 'address'], keep=False)
    dup_postcode = df.duplicated(subset=['name', 'postcode'], keep=False)

    # Combine all duplicate variables using "OR" logic (|) to show all them
    all_duplicates = df[dup_email | dup_phone | dup_address_first_line | dup_postcode]
    return all_duplicates


def filter_duplicate_customers(file_loc, ws, show_message=None):
    duplicate_customers = get_duplicate_customers(pandas.read_excel(file_loc))
    duplicate_customer_ids = []
    for index, cust_id in enumerate(duplicate_customers["customer_id"]):
        duplicate_customer_ids.append(str(cust_id))

    if not duplicate_customer_ids:
        remove_filters(ws)
        if show_message is not None:
            show_message("No duplicate customers were found.")
        return

    ws.UsedRange.AutoFilter(
        Field=CUSTOMER_ID_COLUMN,
        Criteria1=tuple(duplicate_customer_ids),
        Operator=FILTER_VALUES_OPERATOR
    )
