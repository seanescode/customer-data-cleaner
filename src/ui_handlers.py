import filters
import dialogs


def handle_duplicate_filter(file_loc, ws, root):
    filters.filter_duplicate_customers(
        file_loc,
        ws,
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
    filters.remove_filters(ws)
