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
        "records_cleaned": count_records_updated(
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


def validate_and_find_changes(original_df, cleaned_df):
    """
    Check if tables match, then find cells that are truly different.
    """
    # Safety checks: Ensure the sheets match in size and layout
    if not original_df.index.equals(cleaned_df.index):
        raise ValueError("The rows do not match between files.")
    if not original_df.columns.equals(cleaned_df.columns):
        raise ValueError("The column names do not match between files.")
    # Find every cell that is NOT EQUAL (.ne)
    raw_changes = original_df.ne(cleaned_df)
    # Fix the blank cell trick: If both are blank, it's NOT a change
    both_are_blank = original_df.isna() & cleaned_df.isna()
    # Keep changes ONLY where they weren't both blank
    true_changes = raw_changes & ~both_are_blank
    return true_changes


def count_fields_updated(original_df, cleaned_df) -> int:
    """Count the total number of individual cells that changed."""
    changes = validate_and_find_changes(original_df, cleaned_df)
    return changes.sum().sum()  # Adds up every single changed cell


def count_records_updated(original_df, cleaned_df) -> int:
    """Count how many rows had at least one change."""
    changes = validate_and_find_changes(original_df, cleaned_df)
    return changes.any(axis=1).sum()  # Counts rows with any True changes