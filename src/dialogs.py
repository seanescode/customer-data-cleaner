from tkinter import Button, Label, Tk, Toplevel
import win32con
import win32gui


def create_dialog_root():
    root = Tk()
    root.withdraw()
    return root


def show_cleanup_panel(
    root,
    excel_hwnd,
    action_1=None,
    action_2=None,
    action_3=None
):
    panel = Toplevel(root)

    panel.title("Customer Data Cleanup Assistant")
    panel.geometry("310x225")
    panel.resizable(False, False)
    panel.update_idletasks()

    # Keep the panel fixed-size but explicitly retain its minimize button.
    panel_hwnd = panel.winfo_id()
    panel_style = win32gui.GetWindowLong(panel_hwnd, win32con.GWL_STYLE)
    win32gui.SetWindowLong(
        panel_hwnd,
        win32con.GWL_STYLE,
        panel_style | win32con.WS_MINIMIZEBOX
    )
    win32gui.SetWindowPos(
        panel_hwnd,
        0,
        0,
        0,
        0,
        0,
        win32con.SWP_NOMOVE
        | win32con.SWP_NOSIZE
        | win32con.SWP_NOZORDER
        | win32con.SWP_FRAMECHANGED
    )

    panel_width = 310
    panel_height = 225
    right_padding = 35
    excel_was_minimized = False
    panel_was_minimized_before_excel = False

    def show_close_message():

        panel.update_idletasks()

        # The visible outer edge includes the window border, whereas Tk's
        # geometry refers to the client area.  Use the real edge so the
        # message box is exactly right-aligned with the panel.
        _, panel_y, panel_right, _ = win32gui.GetWindowRect(
            panel.winfo_id()
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

        # Size the box from the label's requested size, so the single-line
        # message is never clipped after its wording is changed.
        message.update_idletasks()
        message_width = message_label.winfo_reqwidth() + 2
        message_height = message_label.winfo_reqheight() + 2

        # Align the message's right edge with the panel and keep it above it.
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

    Label(
        panel,
        text="Customer Data Review",
        font=("Arial", 13, "bold")
    ).pack(pady=(10, 2))

    Label(
        panel,
        text="Complete the following steps in order:",
        font=("Arial", 9)
    ).pack(pady=(0, 7))

    Label(
        panel,
        text=(
            "1.  Review and remove duplicate records\n"
            "2.  Review invalid email addresses\n"
            "3.  Remove filters to see full clean results"
        ),
        font=("Arial", 9),
        justify="left",
        anchor="w"
    ).pack(
        anchor="w",
        padx=35,
        pady=(0, 7)
    )

    Button(
        panel,
        text="Review Duplicate Customers",
        width=30,
        command=action_1
    ).pack(pady=2)

    Button(
        panel,
        text="Review Invalid Emails",
        width=30,
        command=action_2
    ).pack(pady=2)

    Button(
        panel,
        text="Remove filters",
        width=30,
        command=action_3
    ).pack(pady=(2, 0))

    def position_panel():

        nonlocal excel_was_minimized, panel_was_minimized_before_excel

        if not win32gui.IsWindow(excel_hwnd):
            panel.destroy()
            return

        if win32gui.IsIconic(excel_hwnd):
            if not excel_was_minimized:
                panel_was_minimized_before_excel = (
                    panel.state() == "iconic"
                )
                if not panel_was_minimized_before_excel:
                    panel.iconify()
            excel_was_minimized = True

        else:
            # Restore the panel only when Excel returns from being minimized.
            # Do not call deiconify for a panel the user has minimized.
            if excel_was_minimized:
                if not panel_was_minimized_before_excel:
                    panel.deiconify()
                excel_was_minimized = False
                panel_was_minimized_before_excel = False

            if panel.state() != "iconic":
                left, top, right, bottom = (
                    win32gui.GetWindowRect(excel_hwnd)
                )

                x = (
                    right
                    - panel_width
                    - right_padding
                )

                y = (
                    top
                    + ((bottom - top) // 2)
                    - (panel_height // 2)
                )

                panel.geometry(
                    f"{panel_width}x{panel_height}"
                    f"+{x}+{y}"
                )

        panel.after(
            100,
            position_panel
        )

    # Position it immediately.
    position_panel()

    return panel
