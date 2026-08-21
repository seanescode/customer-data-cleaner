from email_validator import validate_email, EmailNotValidError

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

    # remove invalid phone numbers
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
    if len(str(postcode)) != 7:
        return ""
    postcode = postcode.upper()
    return postcode[:3] + " " + postcode[3:]

def clean_column(df, column, cleaning_function):
    if column in df.columns:
        df[column] = df[column].fillna("").apply(cleaning_function)


def is_valid_email(email):
    try:
        validate_email(email)
        return True
    except EmailNotValidError:
        return False

def cleanup_customer_data(customer_ids, ws):
    used_range = ws.UsedRange
    # Apply filter to column A - Field=1 means the first column of used_range
    used_range.AutoFilter(
        Field=1,
        Criteria1=customer_ids,
        Operator=7
    )

    sort_range = ws.Range("A1:G201")
    ws.Sort.SortFields.Clear()
    ws.Sort.SortFields.Add(
        Key=ws.Range("B2:B201"),
        SortOn=0,
        Order=1
    )
    ws.Sort.SetRange(sort_range)
    ws.Sort.Header = 1
    ws.Sort.Orientation = 1
    ws.Sort.Apply()


def update_excel_from_dataframe(df, ws):
    values = df.where(df.notna(), None).values.tolist()

    ws.Range(
        ws.Cells(2, 1),
        ws.Cells(len(values) + 1, len(values[0]))
    ).Value = values