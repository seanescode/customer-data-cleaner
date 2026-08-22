from email_validator import validate_email, EmailNotValidError
import phonenumbers


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
    #phonenumbers package takes a number like 16241234 Dublin number and translates to 062 41234
    #am putting the 0 in front then phonenumbers package will then format that
    phone_number = str(phone_number).replace(" ", "")
    if phone_number.startswith("1") and len(phone_number) == 8:
        phone_number = "0" + phone_number

    try:
        number = phonenumbers.parse(phone_number, "IE")
        if not phonenumbers.is_valid_number(number):
            return ""

        formatted_number = phonenumbers.format_number(number, phonenumbers.PhoneNumberFormat.NATIONAL)
        return formatted_number.replace("(", "").replace(")", "")

    except phonenumbers.NumberParseException:
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



def filter_to_invalid_emails(df, email_column, ws):

    customer_id_of_invalid_emails = []

    for index, email in enumerate(df[email_column]):
        try:
            validate_email(email)
        except EmailNotValidError:
            customer_id_of_invalid_emails.append(
                str(df.iloc[index]["customer_id"])
            )
    used_range = ws.UsedRange
    # Apply filter to column A - Field=1 means the first column of used_range
    used_range.AutoFilter(
        Field=1,
        Criteria1=customer_id_of_invalid_emails,
        Operator=7
    )
    print(str(customer_id_of_invalid_emails))


def cleanup_customer_data(customer_ids, ws):

    used_range = ws.UsedRange
    # Apply filter to column A - Field=1 means the first column of used_range
    used_range.AutoFilter(
        Field=1,
        Criteria1=customer_ids,
        Operator=7
    )

    sort_range = ws.Range("A1:G191")
    ws.Sort.SortFields.Clear()
    ws.Sort.SortFields.Add(
        Key=ws.Range("B2:B191"),
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