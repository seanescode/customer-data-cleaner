from tkinter import Button, Frame, Label, Tk
import win32gui
import win32con


def show_cleanup_dialog(clean_function, excel_hwnd):
    root = Tk()
    root.attributes("-topmost", True)
    root.title("Data Cleanup")
    root.geometry("475x180")
    root.resizable(False, False)

    root.update_idletasks()

    # Get the actual Windows HWND for the Tkinter window
    hwnd = win32gui.FindWindow(None, "Data Cleanup")

    # Remove minimise button
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    style &= ~win32con.WS_MINIMIZEBOX
    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)

    # Force Windows to redraw the title bar
    win32gui.SetWindowPos(
        hwnd,
        0,
        0, 0, 0, 0,
        win32con.SWP_NOMOVE |
        win32con.SWP_NOSIZE |
        win32con.SWP_NOZORDER |
        win32con.SWP_FRAMECHANGED
    )

    # Prevent interaction with Excel
    win32gui.EnableWindow(excel_hwnd, False)

    def close_dialog():
        win32gui.EnableWindow(excel_hwnd, True)
        win32gui.SetForegroundWindow(excel_hwnd)
        root.destroy()

    def yes_clicked():
        root.destroy()
        win32gui.EnableWindow(excel_hwnd, True)
        win32gui.SetForegroundWindow(excel_hwnd)
        clean_function()

    def no_clicked():
        root.destroy()
        win32gui.EnableWindow(excel_hwnd, True)
        win32gui.SetForegroundWindow(excel_hwnd)

    # Make the X button use close_dialog()
    root.protocol("WM_DELETE_WINDOW", close_dialog)

    Label(
        root,
        text="Potential duplicates found",
        font=("Arial", 14, "bold")
    ).pack(pady=(20, 5))

    Label(
        root,
        text="Would you like to review them before continuing?",
        font=("Arial", 10)
    ).pack(pady=(0, 15))

    button_frame = Frame(root)
    button_frame.pack()

    Button(
        button_frame,
        text="Yes, show me!",
        width=18,
        command=yes_clicked
    ).pack(side="left", padx=5)

    Button(
        button_frame,
        text="No, just show cleaned data.",
        width=22,
        command=no_clicked
    ).pack(side="left", padx=5)

    root.mainloop()