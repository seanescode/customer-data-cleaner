import tkinter.messagebox
from tkinter import Button, Frame, Label, Tk, Toplevel

import win32con
import win32gui


PANEL_WIDTH = 310
PANEL_HEIGHT = 200
PANEL_RIGHT_PADDING = 35


def create_dialog_root():
    root = Tk()
    root.withdraw()
    return root


def show_error_dialog(title, message):
    error_root = Tk()
    error_root.withdraw()

    tkinter.messagebox.showerror(title, message)

    error_root.destroy()


def show_temporary_message(root, text, duration_ms=2500):
    message = Toplevel(root)

    message.withdraw()
    message.title("")
    message.resizable(False, False)

    _create_temporary_message_content(message, text)
    _centre_window(message)

    message.deiconify()
    message.lift()
    message.attributes("-topmost", True)

    _configure_temporary_message_close(
        message,
        duration_ms
    )


def _create_temporary_message_content(message, text):
    frame = Frame(
        message,
        padx=25,
        pady=20
    )
    frame.pack()

    Label(
        frame,
        text=text,
        font=("Arial", 9),
        justify="center",
        wraplength=300
    ).pack(
        padx=10,
        pady=(0, 12)
    )

    Label(
        frame,
        text="This message will close automatically.",
        font=("Arial", 8),
        foreground="#666666"
    ).pack()


def _centre_window(window):
    window.update_idletasks()

    width = window.winfo_reqwidth()
    height = window.winfo_reqheight()

    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    window.geometry(
        f"{width}x{height}+{x}+{y}"
    )


def _configure_temporary_message_close(message, duration_ms):
    closed = False

    def close_message():
        nonlocal closed

        if closed:
            return

        closed = True

        try:
            if message.winfo_exists():
                message.destroy()
        except Exception:
            pass

    def remove_topmost():
        if not closed:
            try:
                if message.winfo_exists():
                    message.attributes(
                        "-topmost",
                        False
                    )
            except Exception:
                pass

    message.after(200, remove_topmost)
    message.after(duration_ms, close_message)

    message.protocol(
        "WM_DELETE_WINDOW",
        close_message
    )


def show_cleanup_panel(
    root,
    excel_hwnd,
    excel_app,
    workbook,
    review_duplicate_customers_button=None,
    review_invalid_emails_button=None,
    remove_filters_button=None
):
    panel = _create_cleanup_panel(root)

    _enable_minimize_button(panel)

    _configure_panel_close_behavior(panel)

    _create_cleanup_panel_content(
        panel,
        review_duplicate_customers_button,
        review_invalid_emails_button,
        remove_filters_button
    )

    _monitor_and_position_panel(
        root,
        panel,
        excel_hwnd,
        excel_app,
        workbook
    )

    return panel


def _create_cleanup_panel(root):
    panel = Toplevel(root)

    panel.title("Data Cleanup Panel")
    panel.geometry("310x225")
    panel.resizable(False, False)
    panel.update_idletasks()

    return panel


def _enable_minimize_button(panel):
    panel_hwnd = panel.winfo_id()

    panel_style = win32gui.GetWindowLong(
        panel_hwnd,
        win32con.GWL_STYLE
    )

    win32gui.SetWindowLong(
        panel_hwnd,
        win32con.GWL_STYLE,
        panel_style | win32con.WS_MINIMIZEBOX
    )

    window_flags = (
        win32con.SWP_NOMOVE
        | win32con.SWP_NOSIZE
        | win32con.SWP_NOZORDER
        | win32con.SWP_FRAMECHANGED
    )

    win32gui.SetWindowPos(
        panel_hwnd,
        0,
        0,
        0,
        0,
        0,
        window_flags
    )


def _configure_panel_close_behavior(panel):
    def show_close_message():
        panel.update_idletasks()

        _, panel_y, panel_right, _ = (
            win32gui.GetWindowRect(
                panel.winfo_id()
            )
        )

        message = Toplevel(panel)

        message.withdraw()
        message.overrideredirect(True)

        message_label = Label(
            message,
            text=(
                "The cleanup panel stays open while Excel is running. "
                "Close Excel to close the panel."
            ),
            font=("Arial", 9),
            justify="center",
            relief="solid",
            borderwidth=1
        )

        message_label.pack(
            fill="both",
            expand=True
        )

        message.update_idletasks()

        message_width = message_label.winfo_reqwidth() + 2
        message_height = message_label.winfo_reqheight() + 2

        message_x = panel_right - message_width
        message_y = panel_y - message_height - 35

        message.geometry(
            f"{message_width}x{message_height}"
            f"+{message_x}+{message_y}"
        )

        message.deiconify()
        message.lift()

        message.after(
            2500,
            message.destroy
        )

    panel.protocol(
        "WM_DELETE_WINDOW",
        show_close_message
    )


def _create_cleanup_panel_content(
    panel,
    review_duplicate_customers_button,
    review_invalid_emails_button,
    remove_filters_button
):
    Label(
        panel,
        text="Steps to be done in order:",
        font=("Arial", 13, "bold")
    ).pack(
        pady=(10, 2)
    )

    Label(
        panel,
        text=(
            "    1. Review and remove duplicate records\n"
            "    2. Review and remove invalid syntax emails\n"
            "    3. Remove filters to see cleaned data"
        ),
        font=("Arial", 9),
        justify="left",
        anchor="w"
    ).pack(
        anchor="w",
        padx=25,
        pady=(0, 7)
    )

    Button(
        panel,
        text="Review Duplicate Customers",
        width=30,
        command=review_duplicate_customers_button
    ).pack(
        pady=2
    )

    Button(
        panel,
        text="Review Invalid Syntax Emails",
        width=30,
        command=review_invalid_emails_button
    ).pack(
        pady=2
    )

    Button(
        panel,
        text="Remove filters",
        width=30,
        command=remove_filters_button
    ).pack(
        pady=(2, 0)
    )


def _close_panel_and_root(panel, root):
    panel.destroy()
    root.destroy()


def _monitor_and_position_panel(
    root,
    panel,
    excel_hwnd,
    excel_app,
    workbook
):
    excel_was_minimized = False

    def position_panel():
        nonlocal excel_was_minimized

        if not _excel_is_running(excel_app):
            _close_panel_and_root(panel, root)
            return

        if not _workbook_is_open(workbook):
            _close_panel_and_root(panel, root)
            return

        if not _excel_window_is_valid(excel_hwnd):
            _close_panel_and_root(panel, root)
            return

        try:
            excel_is_minimized = win32gui.IsIconic(excel_hwnd)
        except Exception:
            _close_panel_and_root(panel, root)
            return

        if excel_is_minimized:
            if not excel_was_minimized:
                excel_was_minimized = True

                if panel.state() != "iconic":
                    panel.iconify()

        else:
            if excel_was_minimized:
                excel_was_minimized = False

                if panel.state() == "iconic":
                    panel.deiconify()

            if panel.state() != "iconic":
                if not _position_panel_beside_excel(
                    panel,
                    excel_hwnd
                ):
                    _close_panel_and_root(panel, root)
                    return

        panel.after(
            100,
            position_panel
        )

    panel.deiconify()
    panel.lift()

    position_panel()


def _excel_is_running(excel_app):
    try:
        excel_app.Visible
        return True
    except Exception:
        return False


def _workbook_is_open(workbook):
    try:
        workbook.Name
        return True
    except Exception:
        return False


def _excel_window_is_valid(excel_hwnd):
    try:
        return win32gui.IsWindow(excel_hwnd)
    except Exception:
        return False


def _position_panel_beside_excel(panel, excel_hwnd):
    try:
        left, top, right, bottom = (
            win32gui.GetWindowRect(excel_hwnd)
        )

        x = (
            right
            - PANEL_WIDTH
            - PANEL_RIGHT_PADDING
        )

        y = (
            top
            + ((bottom - top) // 2)
            - (PANEL_HEIGHT // 2)
        )

        panel.geometry(
            f"{PANEL_WIDTH}x{PANEL_HEIGHT}"
            f"+{x}+{y}"
        )

        return True

    except Exception:
        return False