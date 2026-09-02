from typing import Any, Callable

from error_handler import handle_errors
from exceptions import DataCleaningError, ConfigKeyError


def clean_name(customer_name: Any) -> str:
    customer_name = str(customer_name) # convert to string first so if user inputs a number for example won't crash
    customer_name = " ".join(customer_name.split())
    customer_name = customer_name.title()
    return customer_name


def clean_email(email: Any) -> str:
    email = str(email)
    email = email.lower()
    email = email.replace(" ", "")
    return email


def clean_phone_number(phone_number: Any) -> str:
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


def clean_address(address: Any) -> str:
    address = str(address)
    address = address.title()
    address = " ".join(address.split())
    return address


def clean_city(city: Any) -> str:
    city = str(city)
    city = city.title()
    city = " ".join(city.split())
    return city


def clean_postcode(postcode: Any) -> str:
    postcode = str(postcode)
    postcode = postcode.replace(" ", "")
    postcode = postcode.upper()
    if len(postcode) != 7:
        return ""
    return postcode[:3] + " " + postcode[3:]


def clean_column(df, column, cleaning_function) -> int:
    if column not in df.columns:
        return 0
    original_values = df[column].fillna("")
    cleaned_values = original_values.apply(cleaning_function)
    changes = (original_values != cleaned_values).sum()
    df[column] = cleaned_values
    return changes

@handle_errors
def clean_data(df: Any, config: dict) -> None:
    try:
        name_updates = clean_column(df, config["columns"]["name"], clean_name)
        email_updates= clean_column(df, config["columns"]["email"], clean_email)
        phone_updates = clean_column(df, config["columns"]["phone"], clean_phone_number)
        address_updates = clean_column(df, config["columns"]["address"], clean_address)
        city_updates = clean_column(df, config["columns"]["city"], clean_city)
        postcode_updates= clean_column(df, config["columns"]["postcode"], clean_postcode)

        total_updates = name_updates + email_updates + phone_updates + address_updates + city_updates + postcode_updates
        print("CLEANING STATISTICS:")
        print("name updates: " + str(name_updates))
        print("email updates: " + str(email_updates))
        print("phone updates: " + str(phone_updates))
        print("address updates: " + str(address_updates))
        print("city updates: " + str(city_updates))
        print("postcode updates: " + str(postcode_updates))
        print("total updates: " + str(total_updates))


    except KeyError as e:
        raise ConfigKeyError(str(e))
    except Exception as e:
        raise DataCleaningError(str(e))