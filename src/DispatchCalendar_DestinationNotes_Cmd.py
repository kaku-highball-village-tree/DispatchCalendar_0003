"""Display the specified Excel file name in a message box."""

from pathlib import Path
import sys
import tkinter as tk
from tkinter import messagebox


ALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
SCRIPT_TITLE = "DispatchCalendar DestinationNotes"


def show_message_box(message_text: str, is_error: bool = False) -> None:
    """Show the specified text in a Windows-style message box."""
    root_window = tk.Tk()
    root_window.withdraw()

    if is_error:
        messagebox.showerror(SCRIPT_TITLE, message_text, parent=root_window)
    else:
        messagebox.showinfo(SCRIPT_TITLE, message_text, parent=root_window)

    root_window.destroy()


def validate_excel_file_name(excel_file_name: str) -> str | None:
    """Return an error message when the specified file name is not supported."""
    if not excel_file_name:
        return "Excelファイル名を指定してください。"

    excel_file_path = Path(excel_file_name)
    if excel_file_path.suffix.lower() not in ALLOWED_EXCEL_EXTENSIONS:
        return ".xlsx または .xlsm ファイルを指定してください。"

    return None


def main() -> int:
    """Read an Excel file name from the command line and display it."""
    if len(sys.argv) != 2:
        show_message_box("Excelファイル名を1つ指定してください。", is_error=True)
        return 1

    excel_file_name = sys.argv[1]
    error_message = validate_excel_file_name(excel_file_name)
    if error_message is not None:
        show_message_box(error_message, is_error=True)
        return 1

    show_message_box(excel_file_name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
