import tkinter.messagebox
from tkinter import Button, Frame, Label, Tk, Toplevel

import win32con
import win32gui


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

    # Let Tkinter calculate the complete size of the dialog.
    message.update_idletasks()

    width = message.winfo_reqwidth()
    height = message.winfo_reqheight()

    # Centre on the screen.
    screen_width = message.winfo_screenwidth()
    screen_height = message.winfo_screenheight()

    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    message.geometry(
        f"{width}x{height}+{x}+{y}"
    )

    # Show only after the final position has been calculated.
    message.deiconify()
    message.lift()
    message.attributes("-topmost", True)

    # Each popup has its own independent timer.
    closed = False

    def close_message():
        nonlocal closed

        if closed:
            return

        closed = True

        try:
            if message.winfo_exists():
                message.destroy()
        except:
            pass

    def remove_topmost():
        if not closed:
            try:
                if message.winfo_exists():
                    message.attributes("-topmost", False)
            except:
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
        action_1=None,
        action_2=None,
        action_3=None
):
    panel = Toplevel(root)

    panel.title("Data Cleanup Panel")
    panel.geometry("310x225")
    panel.resizable(False, False)
    panel.update_idletasks()

    # Keep the panel fixed-size but retain the minimise button.
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
    panel_height = 200
    right_padding = 35

    excel_was_minimized = False

    def show_close_message():

        panel.update_idletasks()

        _, panel_y, panel_right, _ = (
            win32gui.GetWindowRect(panel.winfo_id())
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

    Label(
        panel,
        text="Steps to be done in order:",
        font=("Arial", 13, "bold")
    ).pack(pady=(10, 2))

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
        command=action_1
    ).pack(pady=2)

    Button(
        panel,
        text="Review Invalid Syntax Emails",
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

        nonlocal excel_was_minimized

        # Check if Excel is still running using COM object
        try:
            # Try to access Excel application - this will fail if Excel is closed
            excel_app.Visible
        except Exception:
            # Excel has been closed - destroy panel and root to exit program
            panel.destroy()
            root.destroy()
            return

        # Check if the workbook is still open
        try:
            workbook.Name
        except Exception:
            # Workbook has been closed - destroy panel and root to exit program
            panel.destroy()
            root.destroy()
            return

        # Excel has been closed (window handle check as backup)
        try:
            if not win32gui.IsWindow(excel_hwnd):
                panel.destroy()
                root.destroy()
                return
        except Exception:
            # Window handle became invalid
            panel.destroy()
            root.destroy()
            return

        # Excel is minimised.
        try:
            if win32gui.IsIconic(excel_hwnd):
                if not excel_was_minimized:
                    excel_was_minimized = True

                    if panel.state() != "iconic":
                        panel.iconify()
        except Exception:
            # Window handle became invalid
            panel.destroy()
            root.destroy()
            return

        # Excel is visible.
        try:
            if not win32gui.IsIconic(excel_hwnd):
                if excel_was_minimized:
                    excel_was_minimized = False

                    # Excel has just been restored.
                    if panel.state() == "iconic":
                        panel.deiconify()

                # Keep the panel positioned beside Excel.
                if panel.state() != "iconic":

                    try:
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
                    except Exception:
                        # Excel window handle became invalid during positioning
                        panel.destroy()
                        root.destroy()
                        return
        except Exception:
            # Window handle became invalid
            panel.destroy()
            root.destroy()
            return

        panel.after(
            100,
            position_panel
        )

    # Show immediately.
    panel.deiconify()
    panel.lift()

    # Start positioning/monitoring Excel.
    position_panel()

    return panel
