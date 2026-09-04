import logging
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

    # Remove invalid short phone numbers.
    if len(phone_number) < 5:
        return ""

    # Remove Irish prefix code.
    if phone_number.startswith("353"):
        phone_number = phone_number[3:]

    if phone_number.startswith("+353"):
        phone_number = phone_number[4:]

    # Clean mobile numbers (Excel takes out the zero so have to add back).
    if phone_number.startswith(("83", "85", "86", "87", "89")) and len(phone_number) == 9:
        phone_number = "0" + phone_number

    if phone_number.startswith(("083", "085", "086", "087", "089")) and len(phone_number) == 10:
        return phone_number[0:3] + " " + phone_number[3:6] + " " + phone_number[6:10]

    return phone_number


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
    if not postcode:
        return ""

    postcode = str(postcode)
    postcode = postcode.replace(" ", "")
    postcode = postcode.upper()
    return postcode[:3] + " " + postcode[3:]


def clean_column(df: Any, column: str, cleaning_function: Callable[[Any], str]) -> None:
    if column in df.columns:
        df[column] = df[column].fillna("").apply(cleaning_function)
    else:
        df[column] = df[column].fillna("")


@handle_errors
def clean_data(df: Any, config: dict) -> None:
    log = logging.getLogger("customer_data_cleaner")
    try:
        log.info("Cleaning name column")
        clean_column(df, config["columns"]["name"], clean_name)

        log.info("Cleaning email column")
        clean_column(df, config["columns"]["email"], clean_email)

        log.info("Cleaning phone number column")
        clean_column(df, config["columns"]["phone"], clean_phone_number)

        log.info("Cleaning address column")
        clean_column(df, config["columns"]["address"], clean_address)

        log.info("Cleaning city column")
        clean_column(df, config["columns"]["city"], clean_city)

        log.info("Cleaning postcode column")
        clean_column(df, config["columns"]["postcode"], clean_postcode)

        log.info("Data cleaning completed successfully")

    except KeyError as e:
        log.error(f"Configuration key error during cleaning: {e}")
        raise ConfigKeyError(str(e))
    except Exception as e:
        log.error(f"Error during data cleaning: {e}")
        raise DataCleaningError(str(e))