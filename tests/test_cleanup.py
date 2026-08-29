import pandas as pd
from cleanup import (
    clean_name,
    clean_email,
    clean_phone_number,
    clean_address,
    clean_city,
    clean_postcode,
    clean_column,
    clean_data
)

# test clean_name
def test_clean_name_normal_case():
    assert clean_name("john doe") == "John Doe"

def test_clean_name_multiple_spaces():
    assert clean_name("john   doe") == "John Doe"

def test_clean_name_leading_trailing_spaces():
    assert clean_name("  john doe  ") == "John Doe"

def test_clean_name_all_uppercase():
    assert clean_name("JOHN DOE") == "John Doe"

def test_clean_name_mixed_case():
    assert clean_name("jOhN dOe") == "John Doe"

def test_clean_name_three_words():
    assert clean_name("john middle doe") == "John Middle Doe"

def test_clean_name_single_word():
    assert clean_name("john") == "John"

def test_clean_name_empty_string():
    assert clean_name("") == ""

def test_clean_name_only_spaces():
    assert clean_name("   ") == ""

# test clean_email
def test_clean_email_normal_case():
    assert clean_email("John.Doe@Example.COM") == "john.doe@example.com"

def test_clean_email_with_spaces():
    assert clean_email("john doe@example.com") == "johndoe@example.com"

def test_clean_email_leading_trailing_spaces():
    assert clean_email("  john@example.com  ") == "john@example.com"

def test_clean_email_multiple_spaces():
    assert clean_email("john  doe@example.com") == "johndoe@example.com"

def test_clean_email_uppercase():
    assert clean_email("JOHN@EXAMPLE.COM") == "john@example.com"

def test_clean_email_empty_string():
    assert clean_email("") == ""

def test_clean_email_only_spaces():
    assert clean_email("   ") == ""

def test_clean_email_numeric_input():
    assert clean_email(123) == "123"

# test clean_phone_number
def test_clean_phone_number_no_spaces():
    assert clean_phone_number("0861234567") == "086 123 4567"

def test_clean_phone_number_country_code():
    assert clean_phone_number("+353 086 123 4567") == "086 123 4567"

def test_clean_phone_number_no_leading_zero():
    assert clean_phone_number("861234567") == "086 123 4567"

def test_clean_phone_number_dashes():
    assert clean_phone_number("086-123-4567") == "086 123 4567"

def test_clean_phone_number_empty_string():
    assert clean_phone_number("") == ""

def test_clean_phone_number_only_spaces():
    assert clean_phone_number("   ") == ""

def test_clean_phone_number_short_number():
    assert clean_phone_number("1234") == ""

def test_clean_phone_number_non_irish_prefix():
    assert clean_phone_number("0811234567") == ""

def test_clean_phone_number_landline():
    assert clean_phone_number("0211234567") == ""

def test_clean_phone_number_mixed_formatting():
    assert clean_phone_number("+353-86-123-4567") == "086 123 4567"

def test_clean_phone_number_already_formatted():
    assert clean_phone_number("086 123 4567") == "086 123 4567"

def test_clean_phone_number_numeric_input():
    assert clean_phone_number(861234567) == "086 123 4567"

def test_clean_phone_number_too_long():
    assert clean_phone_number("086123456789") == ""

# test clean_address
def test_clean_address_normal_case():
    assert clean_address("123 Main Street") == "123 Main Street"
def test_clean_address_mixed_case():
    assert clean_address("123 mAin stReEt") == "123 Main Street"
def test_clean_address_extra_spaces():
    assert clean_address("  123  Main  Street  ") == "123 Main Street"
def test_clean_address_empty_string():
    assert clean_address("") == ""
def test_clean_address_numeric_input():
    assert clean_address(123) == "123"
def test_clean_address_only_spaces():
    assert clean_address("   ") == ""

# test clean_city
def test_clean_city_normal_case():
    assert clean_city("dublin") == "Dublin"
def test_clean_city_mixed_case():
    assert clean_city("dUBlIn") == "Dublin"
def test_clean_city_is_empty():
    assert clean_city("") == ""
def test_clean_city_is_numeric():
    assert clean_city(123) == "123"
def test_clean_city_only_spaces():
    assert clean_city("   ") == ""
def test_clean_city_leading_and_trailing_spaces():
    assert clean_city("  Dublin  ") == "Dublin"
def test_clean_city_multiple_words():
    assert clean_city("  New   York  ") == "New York"

# test clean_postcode
def test_clean_postcode_lower_case():
    assert clean_postcode("d25 tm70") == "D25 TM70"
def test_clean_postcode_mixed_case():
    assert clean_postcode("D25 tM70") == "D25 TM70"
def test_clean_postcode_is_empty():
    assert clean_postcode("") == ""
def test_clean_postcode_numeric_input():
    assert clean_postcode(123) == ""
def test_clean_postcode_only_spaces():
    assert clean_postcode("   ") == ""
def test_clean_postcode_leading_and_trailing_spaces():
    assert clean_postcode("  D25 TM70  ") == "D25 TM70"
def test_clean_postcode_shorter_than_expected():
    assert clean_postcode("D25") == ""
def test_clean_postcode_longer_than_expected():
    assert clean_postcode("D25 TM70 123") == ""

#test clean_column
def test_clean_column_normal_case():
    df = pd.DataFrame({"name": ["  john doe  ", "jane smith"]})
    clean_column(df,"name", clean_name)
    assert df.loc[0, "name"]=="John Doe"
    assert df.loc[1, "name"]=="Jane Smith"

def test_clean_column_handles_empty_cells():
    df = pd.DataFrame({"name": ["john doe", None]})
    clean_column(df, "name", clean_name)
    assert df.loc[1, "name"] == ""

def test_clean_column_column_not_in_dataframe():
    df = pd.DataFrame({"name": ["john doe"]})
    # Trying to clean an email column that doesn't exist shouldn't crash
    clean_column(df, "email", clean_email)
    assert "email" not in df.columns

# test clean_data
def test_clean_data_routes_columns_to_correct_cleaning_functions():
    """Check that clean_data sends each column to the correct cleaning tool.

    This test makes sure the configuration dictionary pairs everything up right.
    It stops bugs if someone accidentally deletes a cleaning line or mixes up
    the columns (like sending the address column to the email cleaner).
    """
    # 1. Setup a single row of dirty data for every column type
    df = pd.DataFrame({
        "Full Name": ["  mick jones  "],
        "Email Address": [" MICK@mail.com "],
        "Phone Number": ["0861234567"],
        "Street Address": ["123  main  st"],
        "Town/City": ["dublin"],
        "Eircode": ["d25 tm70"]
    })

    # 2. Map the configuration keys to our dataframe columns
    config = {
        "columns": {
            "name": "Full Name",
            "email": "Email Address",
            "phone": "Phone Number",
            "address": "Street Address",
            "city": "Town/City",
            "postcode": "Eircode"
        }
    }

    # 3. Run the orchestration pipeline
    clean_data(df, config)

    # 4. Assert that every column matched up with its proper cleaning partner
    assert df.loc[0, "Full Name"] == "Mick Jones"
    assert df.loc[0, "Email Address"] == "mick@mail.com"
    assert df.loc[0, "Phone Number"] == "086 123 4567"
    assert df.loc[0, "Street Address"] == "123 Main St"
    assert df.loc[0, "Town/City"] == "Dublin"
    assert df.loc[0, "Eircode"] == "D25 TM70"
