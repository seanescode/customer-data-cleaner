import pandas as pd
import pytest

from stats import (
    get_processing_time,
    count_records_processed,
    count_invalid_emails,
    count_duplicate_customers,
    get_statistics,
    validate_and_find_changes,
    count_fields_updated,
    count_records_updated,
)


# ============================================================
# Tests for get_processing_time()
# ============================================================

def test_get_processing_time_normal():
    start = 10.5
    end = 25.5
    assert get_processing_time(start, end) == 15.0


def test_get_processing_time_rounded_to_1_decimal():
    # Test that the result is rounded to 1 decimal place exactly
    end = 7.699
    start = 5.0
    result = get_processing_time(start, end)

    assert result == 2.7


# ============================================================
# Tests for count_records_processed()
# ============================================================

def test_count_records_processed_normal():
    df = pd.DataFrame({
        "Customer_id": ["cust1", "cust2", "cust3"],
        "name": ["Sean", "Joe", "Bob"]
    })
    assert count_records_processed(df) == 3


def test_count_records_processed_empty_dataframe():
    """Verify that an empty dataframe of records safely returns 0."""
    df = pd.DataFrame()
    assert count_records_processed(df) == 0


# ============================================================
# Tests for count_invalid_emails()
# ============================================================

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


# ============================================================
# Tests for count_duplicate_customers()
# ============================================================

def test_count_duplicate_customers_no_duplicates():
    df = pd.DataFrame({
        "customer_id": ["CUST-1234", "CUST-1235"],
        "name": ["sean Ryan", "Bill Murs"],
        "email": ["sean@example.com", "bill@example..com"],
        "phone": ["087-456-7890", "083-654-3210"],
        "address": ["123 Dublin Road", "15 Main Street"],
        "city": ["Dublin", "Carlow"],
        "postcode": ["D20 DB123", "C10 CW123"],
    })
    assert count_duplicate_customers(df) == 0


def test_count_duplicate_customers_with_duplicates():
    df = pd.DataFrame({
        "customer_id": ["CUST-1234", "CUST-1235", "CUST-1234", "CUST-1236"],
        "name": ["Sean Ryan", "Bill Murs", "Sean Ryan", "bob Davis"],
        "email": ["sean@example.com", "bill@example..com", "sean@example.com", "bob@@example.com"],
        "phone": ["087-456-7890", "083-654-3210", "087-456-7890", "083-654-3210"],
        "address": ["123 Dublin Road", "15 Main Street", "123 Dublin Road", "123 Dunboyne Road"],
        "city": ["Dublin", "Carlow", "Dublin", "Meath"],
        "postcode": ["D20 DB123", "C10 CW123", "D20 DB123", "M15 MH123"],
    })
    assert count_duplicate_customers(df) == 2


def test_count_duplicate_customers_empty_dataframe():
    df = pd.DataFrame(columns=[
        "name",
        "email",
        "phone",
        "address",
        "postcode",
    ])
    assert count_duplicate_customers(df) == 0


# ============================================================
# Tests for get_statistics()
# ============================================================

def test_get_statistics_end_to_end():
    original_df = pd.DataFrame({
        "customer_id": ["CUST-1234", "CUST-1235", "CUST-1234", "CUST-1236"],
        "name": ["Sean Ryan", "Joe Murs", "Sean Ryan", "bob Davis"],
        "email": ["sean@example.com", "joe@example..com", "sean@example.com", "bob@@example.com"],
        "phone": ["087-456-7890", "083-654-3210", "087-456-7890", "083-654-3210"],
        "address": ["123 Dublin Road", "15 Main Street", "123 Dublin Road", "123 Dunboyne Road"],
        "city": ["Dublin", "Carlow", "Dublin", "Meath"],
        "postcode": ["D20 DB123", "C10 CW123", "D20 DB123", "M15 MH123"]
    })

    cleaned_df = original_df.copy()
    cleaned_df.loc[3, "name"] = "Bob Davis"

    cleaning_statistics = get_statistics(
        program_start_time=15.0,
        finished_cleaning_time=25.5,
        original_df=original_df,
        cleaned_df=cleaned_df,
        email_col="email")

    # make sure the data type is a dictionary
    assert isinstance(cleaning_statistics, dict)

    # make sure the dictionary contains all the keys
    expected_keys = {
        "records_processed",
        "records_cleaned",
        "fields_updated",
        "invalid_email_count",
        "duplicate_customer_count",
        "processing_time"
    }
    assert set(cleaning_statistics.keys()) == expected_keys

    assert cleaning_statistics["records_processed"] == 4
    assert cleaning_statistics["records_cleaned"] == 1
    assert cleaning_statistics["fields_updated"] == 1
    assert cleaning_statistics["invalid_email_count"] == 2
    assert cleaning_statistics["duplicate_customer_count"] == 2
    assert cleaning_statistics["processing_time"] == 10.5


# ============================================================
# Tests for validate_and_find_changes()
# ============================================================

def test_validate_and_find_changes_standard_changes():
    original_df = pd.DataFrame({
        "email": ["sean@example.com", "joe@example.com", "BOB@EXAMPLE.COM"]})
    cleaned_df = pd.DataFrame({
        "email": ["sean@example.com", "joe@example.com", "bob@example.com"]})
    assert validate_and_find_changes(original_df, cleaned_df).sum().sum() == 1


def test_validate_and_find_changes_empty_dataframes():
    raw_df = pd.DataFrame(columns=["email"])
    updated_df = pd.DataFrame(columns=["email"])
    assert validate_and_find_changes(raw_df, updated_df).sum().sum() == 0


def test_validate_and_find_changes_ignores_matching_blanks():
    """Verify that if a cell was blank and stays blank, it is NOT counted as a change."""
    original_df = pd.DataFrame({"email": ["sean@example.com", None, "BOB@EXAMPLE.COM"]})
    cleaned_df = pd.DataFrame({"email": ["sean@example.com", None, "bob@example.com"]})
    # Even though row 1 has None (NaN) in both, it should be ignored.
    # Only row 2 (BOB -> bob) is a true change. Total changes should equal 1.
    result = validate_and_find_changes(original_df, cleaned_df)
    assert result.sum().sum() == 1


def test_validate_and_find_changes_mismatched_rows_raises_error():
    original_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com"]})
    cleaned_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com", "bob@example.com"]})

    with pytest.raises(ValueError, match="The rows do not match between files."):
        validate_and_find_changes(original_df, cleaned_df)


def test_validate_and_find_changes_mismatched_columns_raises_error():
    original_df = pd.DataFrame(
        {"name": ["sean", "joe"],
         "email": ["sean@example.com", "joe@example.com"]
         }
    )
    cleaned_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com"]})
    with pytest.raises(ValueError, match="The column names do not match between files."):
        validate_and_find_changes(original_df, cleaned_df)


# ============================================================
# Tests for count_fields_updated()
# ============================================================

def test_count_fields_updated_without_changes():
    original_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com", "bob@example.com"]})
    cleaned_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com", "bob@example.com"]})
    assert count_fields_updated(original_df, cleaned_df) == 0


def test_count_fields_updated_with_changes():
    original_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com ", "BOB@EXAMPLE.COM"]})
    cleaned_df = pd.DataFrame({"email": ["sean@example.com", "  joe@example.com ", "bob@example.com"]})
    assert count_fields_updated(original_df, cleaned_df) == 2


def test_count_fields_updated_empty_dataframes():
    original_df = pd.DataFrame(columns=["email"])
    cleaned_df = pd.DataFrame(columns=["email"])
    assert count_fields_updated(original_df, cleaned_df) == 0


# ============================================================
# Tests for count_records_updated()
# ============================================================

def test_count_records_updated_without_changes():
    original_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com", "bob@example.com"]})
    cleaned_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com", "bob@example.com"]})
    assert count_records_updated(original_df, cleaned_df) == 0


def test_count_records_updated_with_changes():
    original_df = pd.DataFrame({"email": [" sean@example.com ", "joe@example.com", "BOB@EXAMPLE.COM"]})
    cleaned_df = pd.DataFrame({"email": ["sean@example.com", "joe@example.com", "bob@example.com"]})
    assert count_records_updated(original_df, cleaned_df) == 2


def test_count_records_updated_empty_dataframes():
    original_df = pd.DataFrame(columns=["email"])
    cleaned_df = pd.DataFrame(columns=["email"])
    assert count_records_updated(original_df, cleaned_df) == 0
