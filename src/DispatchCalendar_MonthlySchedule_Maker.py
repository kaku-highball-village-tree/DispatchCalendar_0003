#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Create a random monthly dispatch-calendar Excel workbook for testing."""

from __future__ import annotations

from calendar import monthrange
from datetime import datetime
from pathlib import Path
import random
import re
import tkinter as tk
from tkinter import messagebox
import zipfile
from xml.sax.saxutils import escape


WINDOW_TITLE = "DispatchCalendar MonthlySchedule Maker"
OUTPUT_FILE_PREFIX = "配車カレンダー"
MAX_DELIVERY_COUNT_PER_DAY = 20
DELIVERY_ROW_COUNT = 3

DESTINATIONS: list[str] = [
    "ダイセル神崎",
    "ミカタ",
    "アサヒL&C",
    "龍嘻飯店",
    "積水神戸",
    "積水栗東",
    "ホームロジ(朝)",
    "ホームロジ(昼前)",
    "新明和甲南",
    "三菱三田",
    "阪和エ材(小野)",
    "クボタ武庫川",
    "クボタ尼崎",
    "塩野義製薬(豊中)",
    "神東塗料",
    "阪和エ材(吉川)",
    "サントリー",
    "FIG",
    "サン工業",
    "チタニウム尼崎",
    *[f"配送先{i_index:02d}" for i_index in range(1, 21)],
]

VEHICLE_TYPES: list[str] = ["2t", "4t", "大型"]

NOTE_TEXTS: list[str] = [
    "コボレーン注意！指差呼称",
    "指差呼称・周囲確認",
    "荷台に残りが無いか！必ず確認！",
    "当日使用車両は最低限の清掃する事。",
    "ドライブレコーダー",
    "ウォッシャー液",
    "オイル",
    "カード注意・確認",
    "～内バック禁止",
]


class MonthSelectionDialog:
    """Show a small dialog for selecting a target year and month."""

    def __init__(self) -> None:
        now = datetime.now()
        self.i_selected_year: int = now.year
        self.i_selected_month: int = now.month
        self.obj_result: tuple[int, int] | None = None

        self.obj_window = tk.Tk()
        self.obj_window.title(WINDOW_TITLE)
        self.obj_window.resizable(False, False)
        self.obj_window.protocol("WM_DELETE_WINDOW", self.cancel)

        self.obj_month_text = tk.StringVar()
        self.update_month_text()

        obj_message_label = tk.Label(self.obj_window, text="対象年月を選択してください", padx=20, pady=10)
        obj_message_label.grid(row=0, column=0, columnspan=4)

        obj_previous_button = tk.Button(self.obj_window, text="←", width=6, command=self.move_previous_month)
        obj_previous_button.grid(row=1, column=0, padx=10, pady=10)

        obj_month_label = tk.Label(self.obj_window, textvariable=self.obj_month_text, width=16)
        obj_month_label.grid(row=1, column=1, columnspan=2, padx=10, pady=10)

        obj_next_button = tk.Button(self.obj_window, text="→", width=6, command=self.move_next_month)
        obj_next_button.grid(row=1, column=3, padx=10, pady=10)

        obj_ok_button = tk.Button(self.obj_window, text="OK", width=10, command=self.ok)
        obj_ok_button.grid(row=2, column=1, padx=10, pady=10)

        obj_cancel_button = tk.Button(self.obj_window, text="キャンセル", width=10, command=self.cancel)
        obj_cancel_button.grid(row=2, column=2, padx=10, pady=10)

    def update_month_text(self) -> None:
        """Update the displayed year/month text."""
        self.obj_month_text.set(f"{self.i_selected_year}年{self.i_selected_month:02d}月")

    def move_previous_month(self) -> None:
        """Move the displayed target to the previous month."""
        if self.i_selected_month == 1:
            self.i_selected_year -= 1
            self.i_selected_month = 12
        else:
            self.i_selected_month -= 1
        self.update_month_text()

    def move_next_month(self) -> None:
        """Move the displayed target to the next month."""
        if self.i_selected_month == 12:
            self.i_selected_year += 1
            self.i_selected_month = 1
        else:
            self.i_selected_month += 1
        self.update_month_text()

    def ok(self) -> None:
        """Accept the selected year and month."""
        self.obj_result = (self.i_selected_year, self.i_selected_month)
        self.obj_window.destroy()

    def cancel(self) -> None:
        """Cancel without creating an Excel file."""
        self.obj_result = None
        self.obj_window.destroy()

    def show(self) -> tuple[int, int] | None:
        """Show the modal dialog and return the selected year/month."""
        self.obj_window.mainloop()
        return self.obj_result


def select_target_year_month() -> tuple[int, int] | None:
    """Return a target year/month selected by the user."""
    return MonthSelectionDialog().show()


def sanitize_sheet_name(sheet_name: str) -> str:
    """Return an Excel-compatible worksheet name."""
    sanitized_name = re.sub(r"[\\/*?:\[\]]", "_", sheet_name)
    return sanitized_name[:31]


def get_excel_column_name(i_column_index: int) -> str:
    """Convert a 1-based column index to an Excel column name."""
    list_column_name_parts: list[str] = []
    i_remaining_column_index = i_column_index
    while i_remaining_column_index > 0:
        i_remaining_column_index, i_remainder = divmod(i_remaining_column_index - 1, 26)
        list_column_name_parts.append(chr(ord("A") + i_remainder))
    return "".join(reversed(list_column_name_parts))


def build_unique_output_file_path(i_year: int, i_month: int, output_directory_path: Path) -> Path:
    """Build a monthly output path with the first unused four-digit serial number."""
    for i_serial_number in range(1, 10000):
        file_name = f"{OUTPUT_FILE_PREFIX}_{i_year}年{i_month:02d}月_{i_serial_number:04d}.xlsx"
        output_file_path = output_directory_path / file_name
        if not output_file_path.exists():
            return output_file_path

    raise RuntimeError("使用可能な出力ファイル名がありません。")


def build_note_text(destination_text: str) -> str:
    """Return a random note text, replacing the destination placeholder when needed."""
    note_text = random.choice(NOTE_TEXTS)
    if note_text == "～内バック禁止":
        return f"{destination_text}内バック禁止"
    return note_text


def build_monthly_schedule_values(i_year: int, i_month: int) -> list[list[str]]:
    """Build worksheet values in the DestinationNotes monthly layout."""
    i_last_day = monthrange(i_year, i_month)[1]
    i_row_count = 1 + MAX_DELIVERY_COUNT_PER_DAY * DELIVERY_ROW_COUNT
    i_column_count = 1 + i_last_day

    list_values = [["" for _ in range(i_column_count)] for _ in range(i_row_count)]

    for i_day in range(1, i_last_day + 1):
        list_values[0][i_day] = f"{i_year}/{i_month}/{i_day}"

    for i_delivery_index in range(1, MAX_DELIVERY_COUNT_PER_DAY + 1):
        i_row_index = 1 + (i_delivery_index - 1) * DELIVERY_ROW_COUNT
        list_values[i_row_index][0] = str(i_delivery_index)

    for i_day in range(1, i_last_day + 1):
        i_delivery_count = random.randint(1, MAX_DELIVERY_COUNT_PER_DAY)
        for i_delivery_index in range(1, i_delivery_count + 1):
            i_row_index = 1 + (i_delivery_index - 1) * DELIVERY_ROW_COUNT
            destination_text = random.choice(DESTINATIONS)
            vehicle_type_text = random.choice(VEHICLE_TYPES)
            note_text = build_note_text(destination_text)

            list_values[i_row_index][i_day] = destination_text
            list_values[i_row_index + 1][i_day] = vehicle_type_text
            list_values[i_row_index + 2][i_day] = note_text

    return list_values


def build_cell_xml(i_row_index: int, i_column_index: int, cell_text: str) -> str:
    """Build XML for one worksheet cell."""
    cell_reference = f"{get_excel_column_name(i_column_index)}{i_row_index}"
    escaped_cell_text = escape(cell_text)
    return f'<c r="{cell_reference}" t="inlineStr"><is><t>{escaped_cell_text}</t></is></c>'


def build_sheet_xml(list_values: list[list[str]]) -> str:
    """Build worksheet XML for the generated monthly schedule."""
    list_row_xmls: list[str] = []
    for i_row_index, row_values in enumerate(list_values, start=1):
        list_cell_xmls = [
            build_cell_xml(i_row_index, i_column_index, cell_text)
            for i_column_index, cell_text in enumerate(row_values, start=1)
            if cell_text != ""
        ]
        if len(list_cell_xmls) == 0:
            continue
        list_row_xmls.append(f'<row r="{i_row_index}">{"".join(list_cell_xmls)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetData>'
        f'{"".join(list_row_xmls)}'
        '</sheetData>'
        '</worksheet>'
    )


def build_workbook_xml(sheet_name: str) -> str:
    """Build workbook XML."""
    escaped_sheet_name = escape(sheet_name)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets>'
        f'<sheet name="{escaped_sheet_name}" sheetId="1" r:id="rId1"/>'
        '</sheets>'
        '</workbook>'
    )


def build_workbook_relationships_xml() -> str:
    """Build workbook relationship XML."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '</Relationships>'
    )


def build_root_relationships_xml() -> str:
    """Build root relationship XML."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '</Relationships>'
    )


def build_content_types_xml() -> str:
    """Build content type XML."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>'
    )


def write_xlsx_file(output_file_path: Path, list_values: list[list[str]], sheet_name: str) -> None:
    """Write worksheet values to a minimal .xlsx file."""
    with zipfile.ZipFile(output_file_path, "w", compression=zipfile.ZIP_DEFLATED) as xlsx_archive:
        xlsx_archive.writestr("[Content_Types].xml", build_content_types_xml())
        xlsx_archive.writestr("_rels/.rels", build_root_relationships_xml())
        xlsx_archive.writestr("xl/workbook.xml", build_workbook_xml(sheet_name))
        xlsx_archive.writestr("xl/_rels/workbook.xml.rels", build_workbook_relationships_xml())
        xlsx_archive.writestr("xl/worksheets/sheet1.xml", build_sheet_xml(list_values))


def create_monthly_schedule_excel(i_year: int, i_month: int) -> Path:
    """Create a random monthly dispatch-calendar Excel file."""
    output_directory_path = Path(__file__).resolve().parent
    output_file_path = build_unique_output_file_path(i_year, i_month, output_directory_path)
    list_values = build_monthly_schedule_values(i_year, i_month)
    sheet_name = sanitize_sheet_name(f"{i_year}年{i_month:02d}月")
    write_xlsx_file(output_file_path, list_values, sheet_name)
    return output_file_path


def main() -> int:
    """Run the monthly schedule maker."""
    obj_selected_year_month = select_target_year_month()
    if obj_selected_year_month is None:
        return 0

    i_year, i_month = obj_selected_year_month
    output_file_path = create_monthly_schedule_excel(i_year, i_month)
    messagebox.showinfo(
        WINDOW_TITLE,
        f"テスト用配車カレンダーを作成しました。\n\n{output_file_path}",
    )
    print(output_file_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
