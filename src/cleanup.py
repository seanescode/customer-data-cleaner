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
    if len(phone_number) < 7:
        return ""

    # remove Irish prefix code
    if phone_number.startswith("+353"):
        phone_number = phone_number[4:]

    # clean mobile numbers
    if phone_number.startswith(("83", "85", "86", "87", "89")) and len(phone_number) == 9:
        phone_number = "0" + phone_number

    if phone_number.startswith(("083", "085", "086", "087", "089"))and len(phone_number) == 10:
        return phone_number[0:3] + " " + phone_number[3:6] + " " + phone_number[6:10]

    # clean landline numbers
    if len(phone_number) == 7:
        phone_number = "01" + phone_number
    elif len(phone_number) == 8 and phone_number.startswith("1"):
        phone_number = "0" + phone_number

    if phone_number.startswith("01") and len(phone_number) == 9:
        return phone_number[0:2] + " " + phone_number[2:9]

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