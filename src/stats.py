import logging

import dataframe


def get_processing_time(
    program_start_time: float,
    finished_cleaning_time: float,
) -> float:
    """Calculate the processing time in seconds."""
    return round(finished_cleaning_time - program_start_time, 1)


def count_records_processed(df) -> int:
    return len(df)


def count_invalid_emails(df, email_col: str) -> int:
    return len(
        dataframe.get_invalid_emails(
            df,
            email_col,
        )
    )


def count_duplicate_customers(df) -> int:
    return len(dataframe.get_duplicate_customers(df))


def count_fields_updated(original_df, cleaned_df) -> int:
    """Count the total number of fields that were updated."""
    if not original_df.index.equals(cleaned_df.index):
        raise ValueError("DataFrames have different indexes.")

    if not original_df.columns.equals(cleaned_df.columns):
        raise ValueError("DataFrames have different columns.")

    differences = original_df.ne(cleaned_df)
    differences &= ~(original_df.isna() & cleaned_df.isna())

    return differences.sum().sum()


def count_records_cleaned(original_df, cleaned_df) -> int:
    """Count the number of records that were changed."""
    if not original_df.index.equals(cleaned_df.index):
        raise ValueError("DataFrames have different indexes.")

    if not original_df.columns.equals(cleaned_df.columns):
        raise ValueError("DataFrames have different columns.")

    differences = original_df.ne(cleaned_df)
    differences &= ~(original_df.isna() & cleaned_df.isna())

    return differences.any(axis=1).sum()


def get_statistics(
    program_start_time: float,
    finished_cleaning_time: float,
    original_df,
    cleaned_df,
    email_col: str,
) -> dict:
    """Calculate and return all processing statistics."""
    log = logging.getLogger("customer_data_cleaner")
    log.info("Calculating processing statistics")

    stats = {
        "records_processed": count_records_processed(cleaned_df),
        "records_cleaned": count_records_cleaned(
            original_df,
            cleaned_df,
        ),
        "fields_updated": count_fields_updated(
            original_df,
            cleaned_df,
        ),
        "invalid_email_count": count_invalid_emails(
            cleaned_df,
            email_col,
        ),
        "duplicate_customer_count": count_duplicate_customers(
            cleaned_df,
        ),
        "processing_time": get_processing_time(
            program_start_time,
            finished_cleaning_time,
        ),
    }

    log.info(f"Statistics calculated: {stats}")
    return stats