def clean_name(customer_name):
    customer_name = str(customer_name) # convert to string first so if user inputs a number for example won't crash
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
    address = str(address)
    address = address.title()
    address = " ".join(address.split())
    return address


def clean_city(city):
    city = str(city)
    city = city.title()
    city = " ".join(city.split())
    return city


def clean_postcode(postcode):
    postcode = str(postcode)
    postcode = postcode.replace(" ", "")
    postcode = postcode.upper()
    if len(postcode) != 7:
        return ""
    return postcode[:3] + " " + postcode[3:]


def clean_column(df, column, cleaning_function):
    if column in df.columns:
        df[column] = df[column].fillna("").apply(cleaning_function)


def clean_data(df, config):
    try:
        clean_column(df, config["columns"]["name"], clean_name)
        clean_column(df, config["columns"]["email"], clean_email)
        clean_column(df, config["columns"]["phone"], clean_phone_number)
        clean_column(df, config["columns"]["address"], clean_address)
        clean_column(df, config["columns"]["city"], clean_city)
        clean_column(df, config["columns"]["postcode"], clean_postcode)
    except KeyError as e:
        raise KeyError(f"Missing required column in config: {e}") from e