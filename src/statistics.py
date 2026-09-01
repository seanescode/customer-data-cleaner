import dataframe

def get_processing_time(program_start_time, finished_cleaning_time):
    return round(finished_cleaning_time - program_start_time, 1)


def get_records_processed(df):
    return len(df)


def get_invalid_email_count(df, email_col):
    return len(dataframe.get_invalid_emails(df, email_col))


def get_duplicate_customer_count(df):
    return len(dataframe.get_duplicate_customers(df))

def get_statistics(program_start_time, finished_cleaning_time, df, email_col):
    return {
        "processing_time": get_processing_time(program_start_time, finished_cleaning_time),
        "records_processed": get_records_processed(df),
        "invalid_email_count": get_invalid_email_count(df, email_col),
        "duplicate_customer_count": get_duplicate_customer_count(df)
    }











