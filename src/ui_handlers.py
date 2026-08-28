import dialogs

import filters
import spreadsheet


def handle_duplicate_customer_filter(file_loc, ws, name_column, root):
    filters.filter_duplicate_customers(
        file_loc,
        ws,
        name_column,
        lambda text: dialogs.show_temporary_message(root, text)
    )


def handle_email_filter(file_loc, ws, email_column, root):
    filters.filter_to_invalid_emails(
        file_loc,
        ws,
        email_column,
        lambda text: dialogs.show_temporary_message(root, text)
    )


def handle_remove_filters(ws):
    spreadsheet.remove_filters(ws)
