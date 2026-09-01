import dialogs
import spreadsheet
import filters

from error_handler import handle_errors
from exceptions import DialogSetupError


@handle_errors
def setup_and_run_dialog(excel: object,
                         wb: object,
                         file_loc: str,
                         ws: object,
                         config: dict
                         ) -> None:
    try:
        root = dialogs.create_dialog_root()

        review_duplicates = lambda: filters.filter_duplicate_customers(
            file_loc,
            ws,
            config["columns"]["name"],
            config["columns"]["customer_id"],
            lambda text: dialogs.show_temporary_message(root, text)
        )

        review_invalid_emails = lambda: filters.filter_to_invalid_emails(
            file_loc,
            ws,
            config["columns"]["email"],
            lambda text: dialogs.show_temporary_message(root, text)
        )

        remove_filters = lambda: spreadsheet.remove_filters(ws)

        dialogs.show_cleanup_panel(
            root=root,
            excel_hwnd=excel.Hwnd,
            excel_app=excel,
            workbook=wb,
            review_duplicate_customers_button=review_duplicates,
            review_invalid_emails_button=review_invalid_emails,
            remove_filters_button=remove_filters
        )

        root.mainloop()

    except Exception as e:
        raise DialogSetupError(str(e))