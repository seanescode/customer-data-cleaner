import dataframe


def get_processing_time(
    program_start_time,
    finished_cleaning_time,
):
    return round(
        finished_cleaning_time - program_start_time,
        1,
    )


def get_records_processed(df):
    return len(df)


def get_invalid_email_count(df, email_col):
    return len(
        dataframe.get_invalid_emails(
            df,
            email_col,
        )
    )


def get_duplicate_customer_count(df):
    return len(
        dataframe.get_duplicate_customers(df)
    )


def get_fields_updated(original_df, cleaned_df) -> int:
    return (original_df != cleaned_df).sum().sum()


def get_records_cleaned(original_df, cleaned_df) -> int:
    changed_rows = (original_df != cleaned_df).any(axis=1)
    return changed_rows.sum()


def get_statistics(
    program_start_time,
    finished_cleaning_time,
    original_df,
    cleaned_df,
    email_col,
):
    return {

        "records_processed": get_records_processed(
            cleaned_df
        ),
        "records_cleaned": get_records_cleaned(
            original_df,
            cleaned_df,
        ),
        "fields_updated": get_fields_updated(
            original_df,
            cleaned_df,
        ),
        "invalid_email_count": get_invalid_email_count(
            cleaned_df,
            email_col,
        ),
        "duplicate_customer_count": get_duplicate_customer_count(
            cleaned_df
        ),
        "processing_time": get_processing_time(
            program_start_time,
            finished_cleaning_time,
        ),
    }