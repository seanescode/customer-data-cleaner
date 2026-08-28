import dialogs


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


def clean_data(df, config):
    try:
        columns_config = config["columns"]
        clean_column(df, columns_config["name"], clean_name)
        clean_column(df, columns_config["email"], clean_email)
        clean_column(df, columns_config["phone"], clean_phone_number)
        clean_column(df, columns_config["address"], clean_address)
        clean_column(df, columns_config["city"], clean_city)
        clean_column(df, columns_config["postcode"], clean_postcode)
    except KeyError as e:
        dialogs.show_error_dialog("Error", f"Missing required column in config: {e}")
        raise
