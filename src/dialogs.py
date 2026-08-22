from tkinter import Button, Frame, Label, Tk, Toplevel
import win32gui
import win32con


def create_dialog_root():
    root = Tk()
    root.withdraw()
    return root
#
#
# def enable_excel(excel_hwnd):
#     win32gui.EnableWindow(excel_hwnd, True)
#     win32gui.SetForegroundWindow(excel_hwnd)
#
#
# def disable_excel(excel_hwnd):
#     win32gui.EnableWindow(excel_hwnd, False)
#
#
# def remove_minimise_button(window):
#     window.update_idletasks()
#
#     hwnd = window.winfo_id()
#
#     style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
#     style &= ~win32con.WS_MINIMIZEBOX
#
#     win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
#
#     win32gui.SetWindowPos(
#         hwnd,
#         0,
#         0, 0, 0, 0,
#         win32con.SWP_NOMOVE |
#         win32con.SWP_NOSIZE |
#         win32con.SWP_NOZORDER |
#         win32con.SWP_FRAMECHANGED
#     )
#

# def show_cleanup_dialog(root, clean_function, excel_hwnd):
#     dialog = Toplevel(root)
#
#     dialog.attributes("-topmost", True)
#     dialog.title("Data Clean Up")
#     dialog.geometry("425x180")
#     dialog.resizable(False, False)
#
#     remove_minimise_button(dialog)
#
#     disable_excel(excel_hwnd)
#
#     def close_dialog():
#         enable_excel(excel_hwnd)
#         dialog.destroy()
#
#     def yes_clicked():
#         close_dialog()
#         clean_function()
#
#     def no_clicked():
#         close_dialog()
#
#     dialog.protocol("WM_DELETE_WINDOW", close_dialog)
#
#     Label(
#         dialog,
#         text="Potential duplicates found",
#         font=("Arial", 14, "bold")
#     ).pack(pady=(20, 5))
#
#     Label(
#         dialog,
#         text="Would you like to review them before continuing?",
#         font=("Arial", 10)
#     ).pack(pady=(0, 15))
#
#     button_frame = Frame(dialog)
#     button_frame.pack()
#
#     Button(
#         button_frame,
#         text="Yes, show me!",
#         width=18,
#         command=yes_clicked
#     ).pack(side="left", padx=5)
#
#     Button(
#         button_frame,
#         text="No, just show cleaned data.",
#         width=22,
#         command=no_clicked
#     ).pack(side="left", padx=5)
#
#     dialog.grab_set()
#     dialog.focus_force()
#
#     root.wait_window(dialog)


def show_email_cleanup_panel(
    root,
    excel_hwnd,
    action_1=None,
    action_2=None,
    action_3=None
):
    panel = Toplevel(root)

    panel.title("Customer Data Cleanup")
    panel.geometry("320x225")
    panel.resizable(False, False)
    panel.attributes("-topmost", True)

    Label(
        panel,
        text="Customer Data Cleanup",
        font=("Arial", 13, "bold")
    ).pack(pady=(14, 3))

    Label(
        panel,
        text="Complete the following steps in order:",
        font=("Arial", 9)
    ).pack(pady=(0, 10))

    Label(
        panel,
        text=(
            "1. Review and remove duplicate records\n"
            "2. Review invalid email addresses\n"
            "3. Review statistics"
        ),
        font=("Arial", 9),
        justify="left"
    ).pack(anchor="w", padx=35, pady=(0, 10))

    Button(
        panel,
        text="Review Duplicate Customers",
        width=30,
        command=action_2
    ).pack(pady=3)

    Button(
        panel,
        text="Review Invalid Emails",
        width=30,
        command=action_1
    ).pack(pady=3)

    Button(
        panel,
        text="Review Statistics",
        width=30,
        command=action_3
    ).pack(pady=3)

    def position_panel():
        if not win32gui.IsWindow(excel_hwnd):
            panel.destroy()
            return

        left, top, right, bottom = win32gui.GetWindowRect(excel_hwnd)

        x = right - 335
        y = top + 75

        panel.geometry(
            f"320x225+{x}+{y}"
        )

        panel.after(100, position_panel)

    position_panel()

    return panel