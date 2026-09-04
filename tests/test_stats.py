import pandas as pd
from stats import (
    get_processing_time,
    count_records_processed,
    count_invalid_emails)

#test get_processing_time
def test_get_processing_time_normal():
    start = 10.5
    end = 25.5
    assert get_processing_time(start,end) == 15.0

def test_get_processing_time_rounded_to_1_decimal():
    # Test that the result is rounded to 1 decimal place exactly
    end = 7.699
    start = 5.0
    result = get_processing_time(start, end)

    assert result == 2.7

    # Check that there are no floating-point artifacts past the 1st decimal place
    # Converting the float to a string splits it at '.' to count digits after it
    decimal_part = str(result).split(".")[1]

    assert len(decimal_part) == 1

# test count_records_processed
def test_count_records_processed_normal():
    df = pd.DataFrame({
        "Customer_id": ["cust1", "cust2", "cust3"],
        "name": ["Sean", "Joe", "Bob"]
    })
    assert count_records_processed(df) == 3

def test_count_records_processed_empty_dataframe():
    """Verify that an empty list of records safely returns 0."""
    df = pd.DataFrame()
    assert count_records_processed(df) == 0

#count_invalid_emails
def test_count_invalid_emails_valid_emails():
    df = pd.DataFrame({
        "email": ["sean@example.com", "joe@example.com", "bob@example.com"]
    })
    assert count_invalid_emails(df, "email") == 0

def test_count_invalid_emails_invalid_emails():
    df = pd.DataFrame({
        "email": ["sean@example.com", "joe@example.com", "bob@@example.com"]
    })
    assert count_invalid_emails(df, "email") == 1

def test_count_invalid_emails_empty_dataframe():
    df = pd.DataFrame(columns=["email"])
    assert count_invalid_emails(df, "email") == 0






