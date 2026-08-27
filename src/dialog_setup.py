import dialogs
from ui_handlers import handle_duplicate_filter, handle_email_filter, handle_remove_filters


def setup_and_run_dialog(excel, wb, file_loc, ws, config):
    try:
        root = dialogs.create_dialog_root()
        dialogs.show_cleanup_panel(
            root,
            excel.Hwnd,
            excel,
            wb,
            lambda: handle_duplicate_filter(file_loc, ws, root),
            lambda: handle_email_filter(file_loc, ws, config["columns"]["email"], root),
            lambda: handle_remove_filters(ws)
        )
        root.mainloop()
    except Exception as e:
        dialogs.show_error_dialog("Error", f"Dialog setup error: {e}")
