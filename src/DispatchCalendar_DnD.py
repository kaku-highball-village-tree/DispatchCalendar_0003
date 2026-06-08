#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""DispatchCalendar drag-and-drop launcher for CMD converter."""

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
WINDOW_CLASS_NAME: str = "DispatchCalendarDnDWindowClass"
WINDOW_TITLE: str = "DispatchCalendar DnD"
MESSAGE_BOX_TITLE: str = "DispatchCalendar DnD"
MODE_RADIO_CREATE_ID: int = 1001
MODE_RADIO_DELETE_ID: int = 1002

g_b_delete_mode: bool = False
g_h_mode_radio_create: int = 0
g_h_mode_radio_delete: int = 0


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
    win32api.MessageBox(h_window, psz_message_text, MESSAGE_BOX_TITLE, win32con.MB_ICONERROR | win32con.MB_OK)


def show_info_message_box(h_window: int, psz_message_text: str) -> None:
    """Show an information message box."""
    win32api.MessageBox(h_window, psz_message_text, MESSAGE_BOX_TITLE, win32con.MB_ICONINFORMATION | win32con.MB_OK)


def show_delete_confirmation_message_box(h_window: int) -> bool:
    """Show delete confirmation and return True only when user selects OK."""
    psz_message_text: str = "現在「削除」が選択されています。対象の予定をGoogleカレンダーから削除します。続行しますか？"
    i_flags: int = (
        win32con.MB_ICONWARNING
        | win32con.MB_OKCANCEL
        | win32con.MB_DEFBUTTON2
    )
    i_result: int = win32api.MessageBox(h_window, psz_message_text, MESSAGE_BOX_TITLE, i_flags)
    return i_result == win32con.IDOK


def show_auto_close_info_message_box(psz_message_text: str, i_timeout_milliseconds: int = 10000) -> None:
    """Show an information message box that auto-closes after timeout."""
    try:
        ctypes.windll.user32.MessageBoxTimeoutW(
            0,
            psz_message_text,
            MESSAGE_BOX_TITLE,
            win32con.MB_ICONINFORMATION | win32con.MB_OK,
            0,
            i_timeout_milliseconds,
        )
    except Exception:
        show_info_message_box(0, psz_message_text)


def show_auto_close_info_message_box(psz_message_text: str, i_timeout_milliseconds: int = 10000) -> None:
    """Show an information message box that auto-closes after timeout."""
    try:
        ctypes.windll.user32.MessageBoxTimeoutW(
            0,
            psz_message_text,
            MESSAGE_BOX_TITLE,
            win32con.MB_ICONINFORMATION | win32con.MB_OK,
            0,
            i_timeout_milliseconds,
        )
    except Exception:
        show_info_message_box(0, psz_message_text)


def show_auto_close_info_message_box(psz_message_text: str, i_timeout_milliseconds: int = 10000) -> None:
    """Show an information message box that auto-closes after timeout."""
    try:
        ctypes.windll.user32.MessageBoxTimeoutW(
            0,
            psz_message_text,
            MESSAGE_BOX_TITLE,
            win32con.MB_ICONINFORMATION | win32con.MB_OK,
            0,
            i_timeout_milliseconds,
        )
    except Exception:
        show_info_message_box(0, psz_message_text)


def get_cmd_script_path() -> Path:
    """Resolve CMD script path from this script directory."""
    obj_current_script_path: Path = Path(__file__).resolve()
    obj_cmd_script_path: Path = obj_current_script_path.with_name("DispatchCalendar_Cmd.py")
    return obj_cmd_script_path


def is_valid_excel_file_path(obj_file_path: Path) -> bool:
    """Check whether dropped path points to an existing .xlsx file."""
    b_has_valid_extension: bool = obj_file_path.suffix.lower() == ".xlsx"
    b_exists: bool = obj_file_path.exists()
    b_is_file: bool = obj_file_path.is_file()
    return b_has_valid_extension and b_exists and b_is_file


def run_cmd_converter(h_window: int, list_excel_file_paths: list[Path], b_delete_mode: bool) -> None:
    """Run CMD converter via subprocess and show result message."""
    obj_cmd_script_path: Path = get_cmd_script_path()
    if not obj_cmd_script_path.exists():
        show_error_message_box(h_window, f"Cmd script not found:\n{obj_cmd_script_path}")
        return

    if len(list_excel_file_paths) == 0:
        show_error_message_box(h_window, "有効なExcelファイル(.xlsx)がありません。")
        return

    psz_mode: str = "delete" if b_delete_mode else "create"
    list_subprocess_arguments: list[str] = [sys.executable, str(obj_cmd_script_path), "--mode", psz_mode] + [str(obj_path) for obj_path in list_excel_file_paths]
    obj_working_directory_path: Path = list_excel_file_paths[0].parent

    try:
        obj_completed_process: subprocess.CompletedProcess[str] = subprocess.run(
            list_subprocess_arguments,
            cwd=str(obj_working_directory_path),
            capture_output=True,
            text=True,
            check=False,
        )
    except Exception as obj_exception:  # noqa: BLE001
        show_error_message_box(h_window, f"Failed to start converter.\n{obj_exception}")
        return

    if obj_completed_process.returncode == 0:
        if b_delete_mode:
            show_auto_close_info_message_box("カレンダーからの削除を完了しました。")
        else:
            show_auto_close_info_message_box("カレンダーへの登録を完了しました。")
        return

    psz_error_text: str = obj_completed_process.stderr.strip()
    if psz_error_text == "":
        psz_error_text = obj_completed_process.stdout.strip()
    if psz_error_text == "":
        psz_error_text = "変換に失敗しました。詳細は _error.txt を確認してください。"

    show_error_message_box(h_window, psz_error_text)


def on_drop_files(h_window: int, h_drop: int) -> None:
    """Handle WM_DROPFILES."""
    i_file_count: int = win32api.DragQueryFile(h_drop, -1)
    if i_file_count < 1:
        win32api.DragFinish(h_drop)
        return

    list_excel_file_paths: list[Path] = []
    i_invalid_file_count: int = 0

    for iFileIndex in range(i_file_count):
        psz_dropped_file_path: str = win32api.DragQueryFile(h_drop, iFileIndex)
        obj_dropped_file_path: Path = Path(psz_dropped_file_path)
        if is_valid_excel_file_path(obj_dropped_file_path):
            list_excel_file_paths.append(obj_dropped_file_path)
        else:
            i_invalid_file_count += 1

    win32api.DragFinish(h_drop)

    if len(list_excel_file_paths) == 0:
        show_error_message_box(h_window, "Excelファイル(.xlsx)のみ受け付けます。")
        return

    if g_b_delete_mode:
        if not show_delete_confirmation_message_box(h_window):
            return

    run_cmd_converter(h_window, list_excel_file_paths, g_b_delete_mode)


def update_mode_radio_buttons() -> None:
    """Update radio checked states from current mode."""
    if g_h_mode_radio_create == 0 or g_h_mode_radio_delete == 0:
        return
    win32gui.SendMessage(g_h_mode_radio_create, win32con.BM_SETCHECK, win32con.BST_UNCHECKED if g_b_delete_mode else win32con.BST_CHECKED, 0)
    win32gui.SendMessage(g_h_mode_radio_delete, win32con.BM_SETCHECK, win32con.BST_CHECKED if g_b_delete_mode else win32con.BST_UNCHECKED, 0)


def ensure_mode_radio_buttons(h_window: int) -> None:
    """Create mode radio buttons once and keep them visible."""
    global g_h_mode_radio_create, g_h_mode_radio_delete
    if g_h_mode_radio_create != 0 and g_h_mode_radio_delete != 0:
        return

    h_instance: int = win32api.GetModuleHandle(None)
    g_h_mode_radio_create = win32gui.CreateWindowEx(
        0,
        "BUTTON",
        "登録",
        win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.WS_TABSTOP | win32con.BS_AUTORADIOBUTTON | win32con.WS_GROUP,
        0,
        0,
        0,
        0,
        h_window,
        MODE_RADIO_CREATE_ID,
        h_instance,
        None,
    )
    g_h_mode_radio_delete = win32gui.CreateWindowEx(
        0,
        "BUTTON",
        "削除",
        win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.WS_TABSTOP | win32con.BS_AUTORADIOBUTTON,
        0,
        0,
        0,
        0,
        h_window,
        MODE_RADIO_DELETE_ID,
        h_instance,
        None,
    )
    update_mode_radio_buttons()


def layout_mode_radio_buttons(i_client_width: int, i_client_height: int) -> None:
    """Layout mode radio buttons at bottom-right."""
    if g_h_mode_radio_create == 0 or g_h_mode_radio_delete == 0:
        return

    i_radio_width = 80
    i_radio_height = 24
    i_margin = 15
    i_radio_gap = 8
    i_group_width: int = i_radio_width * 2 + i_radio_gap
    i_group_x: int = max(i_margin, i_client_width - i_group_width - i_margin)
    i_group_y: int = max(i_margin, i_client_height - i_radio_height - i_margin)
    win32gui.MoveWindow(g_h_mode_radio_create, i_group_x, i_group_y, i_radio_width, i_radio_height, True)
    win32gui.MoveWindow(g_h_mode_radio_delete, i_group_x + i_radio_width + i_radio_gap, i_group_y, i_radio_width, i_radio_height, True)


def window_procedure(h_window: int, i_message: int, w_param: int, l_param: int) -> int:
    """Main window procedure."""
    global g_b_delete_mode, g_h_mode_radio_create, g_h_mode_radio_delete

    if i_message == win32con.WM_CREATE:
        win32api.DragAcceptFiles(h_window, True)
        ensure_mode_radio_buttons(h_window)
        obj_client_rect: tuple[int, int, int, int] = win32gui.GetClientRect(h_window)
        layout_mode_radio_buttons(obj_client_rect[2], obj_client_rect[3])
        return 0

    if i_message == win32con.WM_SHOWWINDOW:
        ensure_mode_radio_buttons(h_window)
        obj_client_rect = win32gui.GetClientRect(h_window)
        layout_mode_radio_buttons(obj_client_rect[2], obj_client_rect[3])
        return 0

    if i_message == win32con.WM_SIZE:
        layout_mode_radio_buttons(win32api.LOWORD(l_param), win32api.HIWORD(l_param))
        return 0

    if i_message == win32con.WM_COMMAND:
        i_control_id: int = win32api.LOWORD(w_param)
        if i_control_id == MODE_RADIO_CREATE_ID:
            g_b_delete_mode = False
            update_mode_radio_buttons()
            return 0
        if i_control_id == MODE_RADIO_DELETE_ID:
            g_b_delete_mode = True
            update_mode_radio_buttons()
            return 0

    if i_message == win32con.WM_SHOWWINDOW:
        ensure_mode_radio_buttons(h_window)
        obj_client_rect = win32gui.GetClientRect(h_window)
        layout_mode_radio_buttons(obj_client_rect[2], obj_client_rect[3])
        return 0

    if i_message == win32con.WM_SIZE:
        layout_mode_radio_buttons(win32api.LOWORD(l_param), win32api.HIWORD(l_param))
        return 0

    if i_message == win32con.WM_COMMAND:
        i_control_id: int = win32api.LOWORD(w_param)
        if i_control_id == MODE_RADIO_CREATE_ID:
            g_b_delete_mode = False
            update_mode_radio_buttons()
            return 0
        if i_control_id == MODE_RADIO_DELETE_ID:
            g_b_delete_mode = True
            update_mode_radio_buttons()
            return 0

    if i_message == win32con.WM_SIZE:
        layout_mode_radio_buttons(win32api.LOWORD(l_param), win32api.HIWORD(l_param))
        return 0

    if i_message == win32con.WM_COMMAND:
        i_control_id: int = win32api.LOWORD(w_param)
        if i_control_id == MODE_RADIO_CREATE_ID:
            g_b_delete_mode = False
            update_mode_radio_buttons()
            return 0
        if i_control_id == MODE_RADIO_DELETE_ID:
            g_b_delete_mode = True
            update_mode_radio_buttons()
            return 0

    if i_message == win32con.WM_SHOWWINDOW:
        ensure_mode_radio_buttons(h_window)
        obj_client_rect = win32gui.GetClientRect(h_window)
        layout_mode_radio_buttons(obj_client_rect[2], obj_client_rect[3])
        return 0

    if i_message == win32con.WM_SIZE:
        layout_mode_radio_buttons(win32api.LOWORD(l_param), win32api.HIWORD(l_param))
        return 0

    if i_message == win32con.WM_COMMAND:
        i_control_id: int = win32api.LOWORD(w_param)
        if i_control_id == MODE_RADIO_CREATE_ID:
            g_b_delete_mode = False
            update_mode_radio_buttons()
            return 0
        if i_control_id == MODE_RADIO_DELETE_ID:
            g_b_delete_mode = True
            update_mode_radio_buttons()
            return 0

    if i_message == win32con.WM_SHOWWINDOW:
        ensure_mode_radio_buttons(h_window)
        obj_client_rect = win32gui.GetClientRect(h_window)
        layout_mode_radio_buttons(obj_client_rect[2], obj_client_rect[3])
        return 0

    if i_message == win32con.WM_SIZE:
        layout_mode_radio_buttons(win32api.LOWORD(l_param), win32api.HIWORD(l_param))
        return 0

    if i_message == win32con.WM_COMMAND:
        i_control_id: int = win32api.LOWORD(w_param)
        if i_control_id == MODE_RADIO_CREATE_ID:
            g_b_delete_mode = False
            update_mode_radio_buttons()
            return 0
        if i_control_id == MODE_RADIO_DELETE_ID:
            g_b_delete_mode = True
            update_mode_radio_buttons()
            return 0

    if i_message == win32con.WM_COMMAND:
        i_control_id: int = win32api.LOWORD(w_param)
        if i_control_id == MODE_RADIO_CREATE_ID:
            g_b_delete_mode = False
            update_mode_radio_buttons()
            return 0
        if i_control_id == MODE_RADIO_DELETE_ID:
            g_b_delete_mode = True
            update_mode_radio_buttons()
            return 0

    if i_message == win32con.WM_SIZE:
        if g_h_delete_toggle_button != 0:
            i_button_width = 120
            i_button_height = 34
            i_margin = 15
            i_client_width: int = win32api.LOWORD(l_param)
            i_client_height: int = win32api.HIWORD(l_param)
            i_button_x: int = max(i_margin, i_client_width - i_button_width - i_margin)
            i_button_y: int = max(i_margin, i_client_height - i_button_height - i_margin)
            win32gui.MoveWindow(g_h_delete_toggle_button, i_button_x, i_button_y, i_button_width, i_button_height, True)
        return 0

    if i_message == win32con.WM_COMMAND:
        i_control_id: int = win32api.LOWORD(w_param)
        if i_control_id == DELETE_TOGGLE_BUTTON_ID:
            g_b_delete_mode = not g_b_delete_mode
            update_delete_toggle_button_caption(h_window)
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
                "配車カレンダーExcel(.xlsx)をこのウインドウへドラッグ＆ドロップしてください。\n"
                "同じフォルダにTSVファイルを作成します。\n"
                "エラー時は _error.txt を出力します。\n"
                "右下で「登録 / 削除」を選択して実行してください。"
            )

            win32gui.DrawText(h_device_context, psz_instruction_text, -1, obj_text_rect, i_draw_text_flags)
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
