"""Print the specified Excel file name to stdout."""

from pathlib import Path
import sys


ALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}


def validate_excel_file_name(excel_file_name: str) -> str | None:
    """Return an error message when the specified file name is not supported."""
    if not excel_file_name:
        return "Excelファイル名を指定してください。"

    excel_file_path = Path(excel_file_name)
    if excel_file_path.suffix.lower() not in ALLOWED_EXCEL_EXTENSIONS:
        return ".xlsx または .xlsm ファイルを指定してください。"

    return None


def main() -> int:
    """Read an Excel file name from the command line and print it."""
    if len(sys.argv) != 2:
        print("Excelファイル名を1つ指定してください。", file=sys.stderr)
        return 1

    excel_file_name = sys.argv[1]
    error_message = validate_excel_file_name(excel_file_name)
    if error_message is not None:
        print(error_message, file=sys.stderr)
        return 1

    print(excel_file_name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
