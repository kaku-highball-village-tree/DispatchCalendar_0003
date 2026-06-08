#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DispatchCalendar DestinationNotes drag-and-drop launcher."""

from __future__ import annotations

import ctypes
import subprocess
import sys
from pathlib import Path
from typing import Any

import win32api
import win32con
import win32gui


INSTRUCTION_FONT_HEIGHT: int = -17
INSTRUCTION_FONT_FACE: str = "Meiryo UI"
WINDOW_CLASS_NAME: str = "DispatchCalendarDestinationNotesDnDWindowClass"
WINDOW_TITLE: str = "DispatchCalendar DestinationNotes DnD"
MESSAGE_BOX_TITLE: str = "DispatchCalendar DestinationNotes DnD"
COMMAND_SCRIPT_NAME: str = "DispatchCalendar_DestinationNotes_Cmd.py"
ALLOWED_EXCEL_EXTENSIONS: set[str] = {".xlsx", ".xlsm"}


def create_instruction_font() -> int:
    """Create instruction font handle using Win32 GDI."""
    i_font_weight: int = win32con.FW_NORMAL
    i_font_charset: int = win32con.SHIFTJIS_CHARSET
    i_font_output_precision: int = win32con.OUT_DEFAULT_PRECIS
    i_font_clip_precision: int = win32con.CLIP_DEFAULT_PRECIS
    i_font_quality: int = win32con.CLEARTYPE_QUALITY
    i_font_pitch_and_family: int = win32con.DEFAULT_PITCH | win32con.FF_DONTCARE

    i_font_handle: int = ctypes.windll.gdi32.CreateFontW(
        INSTRUCTION_FONT_HEIGHT,
        0,
        0,
        0,
        i_font_weight,
        0,
        0,
        0,
        i_font_charset,
        i_font_output_precision,
        i_font_clip_precision,
        i_font_quality,
        i_font_pitch_and_family,
        INSTRUCTION_FONT_FACE,
    )
    return i_font_handle


def show_error_message_box(h_window: int, psz_message_text: str) -> None:
    """Show an error message box."""
    win32api.MessageBox(
        h_window,
        psz_message_text,
        MESSAGE_BOX_TITLE,
        win32con.MB_ICONERROR | win32con.MB_OK,
    )


def get_cmd_script_path() -> Path:
    """Resolve CMD script path from this script directory."""
    obj_current_script_path: Path = Path(__file__).resolve()
    obj_cmd_script_path: Path = obj_current_script_path.with_name(COMMAND_SCRIPT_NAME)
    return obj_cmd_script_path


def is_valid_excel_file_path(obj_file_path: Path) -> bool:
    """Check whether dropped path points to an existing .xlsx or .xlsm file."""
    b_has_valid_extension: bool = obj_file_path.suffix.lower() in ALLOWED_EXCEL_EXTENSIONS
    b_exists: bool = obj_file_path.exists()
    b_is_file: bool = obj_file_path.is_file()
    return b_has_valid_extension and b_exists and b_is_file


def run_cmd_script(h_window: int, obj_excel_file_path: Path) -> None:
    """Run DestinationNotes CMD script with the dropped Excel file path."""
    obj_cmd_script_path: Path = get_cmd_script_path()
    if not obj_cmd_script_path.exists():
        show_error_message_box(h_window, f"Cmd script not found:\n{obj_cmd_script_path}")
        return

    list_subprocess_arguments: list[str] = [
        sys.executable,
        str(obj_cmd_script_path),
        str(obj_excel_file_path),
    ]

    try:
        obj_completed_process: subprocess.CompletedProcess[bytes] = subprocess.run(
            list_subprocess_arguments,
            cwd=str(obj_excel_file_path.parent),
            check=False,
        )
    except Exception as obj_exception:  # noqa: BLE001
        show_error_message_box(h_window, f"Failed to start command script.\n{obj_exception}")
        return

    if obj_completed_process.returncode != 0:
        show_error_message_box(h_window, "処理に失敗しました。詳細はcmdを確認してください。")


def on_drop_files(h_window: int, h_drop: int) -> None:
    """Handle WM_DROPFILES."""
    i_file_count: int = win32api.DragQueryFile(h_drop, -1)
    if i_file_count < 1:
        win32api.DragFinish(h_drop)
        return

    list_excel_file_paths: list[Path] = []
    for i_file_index in range(i_file_count):
        psz_dropped_file_path: str = win32api.DragQueryFile(h_drop, i_file_index)
        obj_dropped_file_path: Path = Path(psz_dropped_file_path)
        if is_valid_excel_file_path(obj_dropped_file_path):
            list_excel_file_paths.append(obj_dropped_file_path)

    win32api.DragFinish(h_drop)

    if len(list_excel_file_paths) == 0:
        show_error_message_box(h_window, "Excelファイル(.xlsx/.xlsm)のみ受け付けます。")
        return

    if len(list_excel_file_paths) != 1:
        show_error_message_box(h_window, "Excelファイルは1つだけドロップしてください。")
        return

    run_cmd_script(h_window, list_excel_file_paths[0])


def window_procedure(h_window: int, i_message: int, w_param: int, l_param: int) -> int:
    """Main window procedure."""
    if i_message == win32con.WM_CREATE:
        win32api.DragAcceptFiles(h_window, True)
        return 0

    if i_message == win32con.WM_DROPFILES:
        on_drop_files(h_window, w_param)
        return 0

    if i_message == win32con.WM_PAINT:
        h_device_context: int | None = None
        obj_paint_struct: Any = None
        i_font_handle: int = 0
        i_old_font_handle: int = 0
        try:
            h_device_context, obj_paint_struct = win32gui.BeginPaint(h_window)
            i_font_handle = create_instruction_font()
            if i_font_handle != 0:
                i_old_font_handle = win32gui.SelectObject(h_device_context, i_font_handle)

            win32gui.SetBkMode(h_device_context, win32con.TRANSPARENT)
            obj_client_rect: tuple[int, int, int, int] = win32gui.GetClientRect(h_window)
            i_margin: int = 5
            obj_text_rect: tuple[int, int, int, int] = (
                obj_client_rect[0] + i_margin,
                obj_client_rect[1] + i_margin,
                obj_client_rect[2] - i_margin,
                obj_client_rect[3] - i_margin,
            )
            i_draw_text_flags: int = win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK
            psz_instruction_text: str = (
                "配送先備考Excel(.xlsx/.xlsm)をこのウインドウへドラッグ＆ドロップしてください。\n"
                "ドロップされたExcelファイル名を表示します。"
            )
            win32gui.DrawText(
                h_device_context,
                psz_instruction_text,
                -1,
                obj_text_rect,
                i_draw_text_flags,
            )
        except Exception as obj_exception:  # noqa: BLE001
            print(f"WM_PAINT error: {obj_exception}")
        finally:
            if h_device_context is not None and i_font_handle != 0 and i_old_font_handle != 0:
                win32gui.SelectObject(h_device_context, i_old_font_handle)
            if i_font_handle != 0:
                win32gui.DeleteObject(i_font_handle)
            if h_device_context is not None and obj_paint_struct is not None:
                win32gui.EndPaint(h_window, obj_paint_struct)
        return 0

    if i_message == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0

    return win32gui.DefWindowProc(h_window, i_message, w_param, l_param)


def create_and_show_window() -> int:
    """Create and display main DnD window."""
    h_instance: int = win32api.GetModuleHandle(None)
    obj_window_class: Any = win32gui.WNDCLASS()
    obj_window_class.hInstance = h_instance
    obj_window_class.lpszClassName = WINDOW_CLASS_NAME
    obj_window_class.lpfnWndProc = window_procedure
    obj_window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    obj_window_class.hbrBackground = win32con.COLOR_WINDOW + 1

    i_registered_class_atom: int = win32gui.RegisterClass(obj_window_class)

    i_desktop_width: int = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    i_desktop_height: int = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    i_window_width: int = int(i_desktop_width * 0.5)
    i_window_height: int = int(i_desktop_height * 0.5)
    i_window_left: int = int((i_desktop_width - i_window_width) / 2)
    i_window_top: int = int((i_desktop_height - i_window_height) / 2)
    i_extended_style: int = win32con.WS_EX_ACCEPTFILES
    i_window_style: int = (
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX
    )

    h_window: int = win32gui.CreateWindowEx(
        i_extended_style,
        i_registered_class_atom,
        WINDOW_TITLE,
        i_window_style,
        i_window_left,
        i_window_top,
        i_window_width,
        i_window_height,
        0,
        0,
        h_instance,
        None,
    )
    win32gui.SetWindowPos(
        h_window,
        win32con.HWND_TOPMOST,
        0,
        0,
        0,
        0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
    )
    win32gui.ShowWindow(h_window, win32con.SW_SHOW)
    win32gui.UpdateWindow(h_window)
    return h_window


def main() -> int:
    """Program entry point."""
    create_and_show_window()
    win32gui.PumpMessages()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
