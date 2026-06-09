"""Convert a DestinationNotes Excel worksheet to a same-named TSV file."""

from datetime import datetime, timedelta
from pathlib import Path
import csv
import re
import sys
import zipfile
from xml.etree import ElementTree


ALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
EXCEL_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
VEHICLE_TYPE_SORT_ORDER = {"2t": 0, "4t": 1, "大型": 2}
OTHER_VEHICLE_TYPE_SORT_INDEX = 3
BLANK_VEHICLE_TYPE_SORT_INDEX = 4


BUILT_IN_DATE_FORMAT_IDS = {
    14,
    15,
    16,
    17,
    18,
    19,
    20,
    21,
    22,
    27,
    28,
    29,
    30,
    31,
    32,
    33,
    34,
    35,
    36,
    45,
    46,
    47,
    50,
    51,
    52,
    53,
    54,
    55,
    56,
    57,
    58,
}


def validate_excel_file_path(excel_file_path: Path) -> str | None:
    """Return an error message when the specified Excel file path is not supported."""
    if excel_file_path.suffix.lower() not in ALLOWED_EXCEL_EXTENSIONS:
        return ".xlsx または .xlsm ファイルを指定してください。"

    if not excel_file_path.exists():
        return f"Excelファイルが存在しません: {excel_file_path}"

    if not excel_file_path.is_file():
        return f"Excelファイルではありません: {excel_file_path}"

    return None


def namespace_tag(namespace_uri: str, tag_name: str) -> str:
    """Return an ElementTree namespace-qualified tag."""
    return f"{{{namespace_uri}}}{tag_name}"


def column_letters_to_number(column_letters: str) -> int:
    """Convert Excel column letters to a 1-based column number."""
    i_column_number = 0
    for character in column_letters:
        i_column_number = i_column_number * 26 + ord(character.upper()) - ord("A") + 1
    return i_column_number


def get_cell_coordinates(cell_reference: str, fallback_row: int, fallback_column: int) -> tuple[int, int]:
    """Return 1-based row and column numbers from a cell reference."""
    match = re.fullmatch(r"([A-Za-z]+)([0-9]+)", cell_reference or "")
    if match is None:
        return fallback_row, fallback_column

    i_column_number = column_letters_to_number(match.group(1))
    i_row_number = int(match.group(2))
    return i_row_number, i_column_number


def read_xml_from_zip(excel_archive: zipfile.ZipFile, archive_path: str) -> ElementTree.Element | None:
    """Read an XML file from an Excel archive."""
    try:
        with excel_archive.open(archive_path) as xml_file:
            return ElementTree.parse(xml_file).getroot()
    except KeyError:
        return None


def read_text_without_phonetic(root_node: ElementTree.Element) -> str:
    """Read visible string text without Excel phonetic guide text."""
    list_text_parts: list[str] = []
    text_tag = namespace_tag(EXCEL_NAMESPACE, "t")
    rich_text_run_tag = namespace_tag(EXCEL_NAMESPACE, "r")
    phonetic_text_tag = namespace_tag(EXCEL_NAMESPACE, "rPh")
    phonetic_properties_tag = namespace_tag(EXCEL_NAMESPACE, "phoneticPr")

    for child_node in list(root_node):
        if child_node.tag == text_tag:
            list_text_parts.append(child_node.text or "")
            continue

        if child_node.tag == rich_text_run_tag:
            text_node = child_node.find(text_tag)
            if text_node is not None:
                list_text_parts.append(text_node.text or "")
            continue

        if child_node.tag in {phonetic_text_tag, phonetic_properties_tag}:
            continue

    return "".join(list_text_parts)


def read_shared_strings(excel_archive: zipfile.ZipFile) -> list[str]:
    """Read shared strings from an Excel archive without phonetic guide text."""
    shared_strings_root = read_xml_from_zip(excel_archive, "xl/sharedStrings.xml")
    if shared_strings_root is None:
        return []

    list_shared_strings: list[str] = []
    for string_item in shared_strings_root.findall(namespace_tag(EXCEL_NAMESPACE, "si")):
        list_shared_strings.append(read_text_without_phonetic(string_item))

    return list_shared_strings


def is_date_format_code(format_code: str) -> bool:
    """Return whether an Excel number format code appears to be a date/time format."""
    normalized_format_code = re.sub(r'"[^"]*"', "", format_code.lower())
    normalized_format_code = re.sub(r"\\.", "", normalized_format_code)
    normalized_format_code = re.sub(r"\[[^\]]*\]", "", normalized_format_code)
    return any(character in normalized_format_code for character in ("y", "m", "d"))


def read_date_style_indexes(excel_archive: zipfile.ZipFile) -> set[int]:
    """Read style indexes that should be treated as dates."""
    styles_root = read_xml_from_zip(excel_archive, "xl/styles.xml")
    if styles_root is None:
        return set()

    dict_custom_formats: dict[int, str] = {}
    num_fmts_node = styles_root.find(namespace_tag(EXCEL_NAMESPACE, "numFmts"))
    if num_fmts_node is not None:
        for num_fmt_node in num_fmts_node.findall(namespace_tag(EXCEL_NAMESPACE, "numFmt")):
            num_fmt_id = num_fmt_node.attrib.get("numFmtId")
            format_code = num_fmt_node.attrib.get("formatCode", "")
            if num_fmt_id is not None:
                dict_custom_formats[int(num_fmt_id)] = format_code

    set_date_style_indexes: set[int] = set()
    cell_xfs_node = styles_root.find(namespace_tag(EXCEL_NAMESPACE, "cellXfs"))
    if cell_xfs_node is None:
        return set_date_style_indexes

    for i_style_index, xf_node in enumerate(cell_xfs_node.findall(namespace_tag(EXCEL_NAMESPACE, "xf"))):
        num_fmt_id_text = xf_node.attrib.get("numFmtId", "0")
        i_num_fmt_id = int(num_fmt_id_text)
        if i_num_fmt_id in BUILT_IN_DATE_FORMAT_IDS:
            set_date_style_indexes.add(i_style_index)
            continue

        custom_format_code = dict_custom_formats.get(i_num_fmt_id)
        if custom_format_code is not None and is_date_format_code(custom_format_code):
            set_date_style_indexes.add(i_style_index)

    return set_date_style_indexes


def read_date_1904_flag(workbook_root: ElementTree.Element) -> bool:
    """Return whether the workbook uses the 1904 date system."""
    workbook_pr_node = workbook_root.find(namespace_tag(EXCEL_NAMESPACE, "workbookPr"))
    if workbook_pr_node is None:
        return False

    return workbook_pr_node.attrib.get("date1904") in {"1", "true", "True"}


def get_first_sheet_archive_path(excel_archive: zipfile.ZipFile) -> tuple[str, bool]:
    """Return the first worksheet archive path and workbook date system flag."""
    workbook_root = read_xml_from_zip(excel_archive, "xl/workbook.xml")
    if workbook_root is None:
        raise RuntimeError("workbook.xml が見つかりません。")

    b_date_1904 = read_date_1904_flag(workbook_root)
    first_sheet_node = workbook_root.find(f"{namespace_tag(EXCEL_NAMESPACE, 'sheets')}/{namespace_tag(EXCEL_NAMESPACE, 'sheet')}")
    if first_sheet_node is None:
        raise RuntimeError("シートが見つかりません。")

    relationship_id = first_sheet_node.attrib.get(namespace_tag(OFFICE_RELATIONSHIP_NAMESPACE, "id"))
    if relationship_id is None:
        raise RuntimeError("先頭シートの参照IDが見つかりません。")

    workbook_relationships_root = read_xml_from_zip(excel_archive, "xl/_rels/workbook.xml.rels")
    if workbook_relationships_root is None:
        raise RuntimeError("workbook.xml.rels が見つかりません。")

    for relationship_node in workbook_relationships_root.findall(namespace_tag(RELATIONSHIP_NAMESPACE, "Relationship")):
        if relationship_node.attrib.get("Id") != relationship_id:
            continue

        target_path = relationship_node.attrib.get("Target")
        if target_path is None:
            raise RuntimeError("先頭シートのパスが見つかりません。")

        if target_path.startswith("/"):
            return target_path.lstrip("/"), b_date_1904

        return f"xl/{target_path}", b_date_1904

    raise RuntimeError("先頭シートの参照先が見つかりません。")


def excel_serial_date_to_text(serial_value: float, b_date_1904: bool) -> str:
    """Convert an Excel serial date value to yyyy/m/d text."""
    base_date = datetime(1904, 1, 1) if b_date_1904 else datetime(1899, 12, 30)
    converted_datetime = base_date + timedelta(days=serial_value)
    return f"{converted_datetime.year}/{converted_datetime.month}/{converted_datetime.day}"


def format_numeric_text(numeric_text: str) -> str:
    """Format a numeric Excel text value without unnecessary trailing .0."""
    try:
        numeric_value = float(numeric_text)
    except ValueError:
        return numeric_text

    if numeric_value.is_integer():
        return str(int(numeric_value))

    return str(numeric_value)


def read_inline_string(cell_node: ElementTree.Element) -> str:
    """Read an inline string from a worksheet cell without phonetic guide text."""
    inline_string_node = cell_node.find(namespace_tag(EXCEL_NAMESPACE, "is"))
    if inline_string_node is None:
        return ""

    return read_text_without_phonetic(inline_string_node)


def read_cell_text(
    cell_node: ElementTree.Element,
    list_shared_strings: list[str],
    set_date_style_indexes: set[int],
    b_date_1904: bool,
) -> str:
    """Read a worksheet cell value as TSV text."""
    cell_type = cell_node.attrib.get("t", "n")
    value_node = cell_node.find(namespace_tag(EXCEL_NAMESPACE, "v"))
    raw_value = value_node.text if value_node is not None else None

    if cell_type == "inlineStr":
        return read_inline_string(cell_node)

    if raw_value is None:
        return ""

    if cell_type == "s":
        i_shared_string_index = int(raw_value)
        if i_shared_string_index < len(list_shared_strings):
            return list_shared_strings[i_shared_string_index]
        return ""

    if cell_type == "b":
        return "TRUE" if raw_value == "1" else "FALSE"

    style_index_text = cell_node.attrib.get("s")
    if style_index_text is not None and int(style_index_text) in set_date_style_indexes:
        try:
            return excel_serial_date_to_text(float(raw_value), b_date_1904)
        except ValueError:
            return raw_value

    if cell_type == "n":
        return format_numeric_text(raw_value)

    return raw_value


def read_first_worksheet_values(excel_file_path: Path) -> tuple[dict[tuple[int, int], str], int, int]:
    """Read first worksheet values and value bounds from an Excel file."""
    with zipfile.ZipFile(excel_file_path) as excel_archive:
        list_shared_strings = read_shared_strings(excel_archive)
        set_date_style_indexes = read_date_style_indexes(excel_archive)
        worksheet_archive_path, b_date_1904 = get_first_sheet_archive_path(excel_archive)
        worksheet_root = read_xml_from_zip(excel_archive, worksheet_archive_path)
        if worksheet_root is None:
            raise RuntimeError("先頭シートを読み取れません。")

        dict_cell_values: dict[tuple[int, int], str] = {}
        i_last_value_row = 0
        i_last_value_column = 0
        sheet_data_node = worksheet_root.find(namespace_tag(EXCEL_NAMESPACE, "sheetData"))
        if sheet_data_node is None:
            return dict_cell_values, i_last_value_row, i_last_value_column

        for i_fallback_row, row_node in enumerate(sheet_data_node.findall(namespace_tag(EXCEL_NAMESPACE, "row")), start=1):
            i_row_number = int(row_node.attrib.get("r", i_fallback_row))
            for i_fallback_column, cell_node in enumerate(row_node.findall(namespace_tag(EXCEL_NAMESPACE, "c")), start=1):
                cell_reference = cell_node.attrib.get("r", "")
                i_cell_row, i_cell_column = get_cell_coordinates(cell_reference, i_row_number, i_fallback_column)
                cell_text = read_cell_text(cell_node, list_shared_strings, set_date_style_indexes, b_date_1904)
                dict_cell_values[(i_cell_row, i_cell_column)] = cell_text
                if cell_text != "":
                    i_last_value_row = max(i_last_value_row, i_cell_row)
                    i_last_value_column = max(i_last_value_column, i_cell_column)

    return dict_cell_values, i_last_value_row, i_last_value_column


def write_excel_values_to_tsv(excel_file_path: Path) -> Path:
    """Write first worksheet values to a same-named TSV file."""
    dict_cell_values, i_last_value_row, i_last_value_column = read_first_worksheet_values(excel_file_path)
    if i_last_value_row == 0 or i_last_value_column == 0:
        raise RuntimeError("出力対象シートに値が存在しません。")

    tsv_file_path = excel_file_path.with_suffix(".tsv")
    with tsv_file_path.open("w", encoding="utf-8-sig", newline="") as tsv_file:
        tsv_writer = csv.writer(tsv_file, delimiter="\t", lineterminator="\n")
        for i_row_index in range(1, i_last_value_row + 1):
            row_values = [
                dict_cell_values.get((i_row_index, i_column_index), "")
                for i_column_index in range(1, i_last_value_column + 1)
            ]
            tsv_writer.writerow(row_values)

    return tsv_file_path


def read_tsv_rows(tsv_file_path: Path) -> list[list[str]]:
    """Read a TSV file as rows."""
    with tsv_file_path.open("r", encoding="utf-8-sig", newline="") as tsv_file:
        return list(csv.reader(tsv_file, delimiter="\t"))


def parse_tsv_date_header(date_text: str) -> datetime:
    """Parse a TSV date header into a datetime."""
    normalized_date_text = date_text.strip()
    match = re.fullmatch(r"([0-9]{4})[/-]([0-9]{1,2})[/-]([0-9]{1,2})", normalized_date_text)
    if match is not None:
        return datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))

    match = re.fullmatch(r"([0-9]{4})年([0-9]{1,2})月([0-9]{1,2})日", normalized_date_text)
    if match is not None:
        return datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))

    raise RuntimeError(f"日付ヘッダーを解析できません: {date_text}")


def format_step0001_date_text(date_text: str) -> str:
    """Format a TSV date header for a step0001 output file name."""
    parsed_date = parse_tsv_date_header(date_text)
    return f"{parsed_date.year}年{parsed_date.month:02d}月{parsed_date.day:02d}日"


def get_cell_value(row_values: list[str], i_column_index: int) -> str:
    """Return a cell value from a row or an empty string when the column is missing."""
    if i_column_index >= len(row_values):
        return ""

    return row_values[i_column_index]


def get_date_columns(monthly_tsv_rows: list[list[str]]) -> list[tuple[int, str]]:
    """Return date column indexes and date header text from monthly TSV rows."""
    if len(monthly_tsv_rows) == 0:
        raise RuntimeError("月間TSVに行がありません。")

    header_row = monthly_tsv_rows[0]
    list_date_columns: list[tuple[int, str]] = []
    for i_column_index, date_text in enumerate(header_row[1:], start=1):
        if date_text.strip() == "":
            continue

        parse_tsv_date_header(date_text)
        list_date_columns.append((i_column_index, date_text))

    if len(list_date_columns) == 0:
        raise RuntimeError("日付列が見つかりません。")

    return list_date_columns


def build_daily_blocks(monthly_tsv_rows: list[list[str]], i_date_column_index: int) -> list[list[list[str]]]:
    """Build unsorted three-row daily blocks for one date column."""
    list_daily_blocks: list[list[list[str]]] = []
    data_rows = monthly_tsv_rows[1:]

    for i_block_start_index in range(0, len(data_rows), 3):
        block_source_rows = data_rows[i_block_start_index:i_block_start_index + 3]
        if len(block_source_rows) < 3:
            continue

        destination_text = get_cell_value(block_source_rows[0], i_date_column_index)
        vehicle_type_text = get_cell_value(block_source_rows[1], i_date_column_index)
        note_text = get_cell_value(block_source_rows[2], i_date_column_index)
        if destination_text == "" and vehicle_type_text == "" and note_text == "":
            continue

        no_text = get_cell_value(block_source_rows[0], 0)
        daily_block = [
            [no_text, destination_text],
            ["", vehicle_type_text],
            ["", note_text],
        ]
        list_daily_blocks.append(daily_block)

    return list_daily_blocks


def build_daily_tsv_rows(monthly_tsv_rows: list[list[str]], i_date_column_index: int, date_text: str) -> list[list[str]]:
    """Build TSV rows for a single date column."""
    daily_tsv_rows = [[get_cell_value(monthly_tsv_rows[0], 0), date_text]]
    for daily_block in build_daily_blocks(monthly_tsv_rows, i_date_column_index):
        daily_tsv_rows.extend(daily_block)
    return daily_tsv_rows


def build_step0001_tsv_file_path(monthly_tsv_file_path: Path, date_text: str) -> Path:
    """Build a step0001 daily TSV output path."""
    formatted_date_text = format_step0001_date_text(date_text)
    return monthly_tsv_file_path.with_name(f"{monthly_tsv_file_path.stem}_step0001_{formatted_date_text}.tsv")


def write_tsv_rows(tsv_file_path: Path, rows: list[list[str]]) -> None:
    """Write rows to a TSV file."""
    with tsv_file_path.open("w", encoding="utf-8-sig", newline="") as tsv_file:
        tsv_writer = csv.writer(tsv_file, delimiter="\t", lineterminator="\n")
        tsv_writer.writerows(rows)


def write_step0001_daily_tsv_files(monthly_tsv_file_path: Path) -> list[Path]:
    """Create step0001 daily TSV files from a monthly TSV file."""
    monthly_tsv_rows = read_tsv_rows(monthly_tsv_file_path)
    list_daily_tsv_file_paths: list[Path] = []

    for i_date_column_index, date_text in get_date_columns(monthly_tsv_rows):
        daily_tsv_rows = build_daily_tsv_rows(monthly_tsv_rows, i_date_column_index, date_text)
        daily_tsv_file_path = build_step0001_tsv_file_path(monthly_tsv_file_path, date_text)
        write_tsv_rows(daily_tsv_file_path, daily_tsv_rows)
        list_daily_tsv_file_paths.append(daily_tsv_file_path)

    return list_daily_tsv_file_paths


def build_step0002_tsv_file_path(step0001_tsv_file_path: Path) -> Path:
    """Build a step0002 TSV output path from a step0001 TSV path."""
    if "_step0001_" not in step0001_tsv_file_path.name:
        raise RuntimeError(f"step0001 TSVファイル名ではありません: {step0001_tsv_file_path}")

    step0002_file_name = step0001_tsv_file_path.name.replace("_step0001_", "_step0002_", 1)
    return step0001_tsv_file_path.with_name(step0002_file_name)


def build_step0002_error_file_path(step0002_tsv_file_path: Path) -> Path:
    """Build a step0002 error file path by appending _error.txt to the TSV name."""
    return step0002_tsv_file_path.with_name(f"{step0002_tsv_file_path.name}_error.txt")


def get_vehicle_type_sort_index(vehicle_type_text: str) -> int:
    """Return the requested vehicle-type sort index."""
    normalized_vehicle_type_text = vehicle_type_text.strip()
    if normalized_vehicle_type_text == "":
        return BLANK_VEHICLE_TYPE_SORT_INDEX

    return VEHICLE_TYPE_SORT_ORDER.get(normalized_vehicle_type_text, OTHER_VEHICLE_TYPE_SORT_INDEX)


def read_step0001_daily_blocks(step0001_tsv_rows: list[list[str]]) -> list[list[list[str]]]:
    """Read three-row daily blocks from step0001 TSV rows."""
    if len(step0001_tsv_rows) == 0:
        raise RuntimeError("step0001 TSVに行がありません。")

    data_rows = step0001_tsv_rows[1:]
    if len(data_rows) % 3 != 0:
        raise RuntimeError("step0001 TSVのデータ行数が3の倍数ではありません。")

    list_daily_blocks: list[list[list[str]]] = []
    for i_block_start_index in range(0, len(data_rows), 3):
        block_source_rows = data_rows[i_block_start_index:i_block_start_index + 3]
        block_rows = [list(row_values) for row_values in block_source_rows]
        while len(block_rows[0]) < 2:
            block_rows[0].append("")
        while len(block_rows[1]) < 2:
            block_rows[1].append("")
        while len(block_rows[2]) < 2:
            block_rows[2].append("")
        list_daily_blocks.append(block_rows)

    return list_daily_blocks


def sort_step0001_daily_blocks(list_daily_blocks: list[list[list[str]]]) -> list[list[list[str]]]:
    """Sort step0001 daily blocks by vehicle type without changing equal-order blocks."""
    indexed_daily_blocks = list(enumerate(list_daily_blocks))
    indexed_daily_blocks.sort(
        key=lambda indexed_daily_block: (
            get_vehicle_type_sort_index(get_cell_value(indexed_daily_block[1][1], 1)),
            get_cell_value(indexed_daily_block[1][1], 1).strip(),
            indexed_daily_block[0],
        )
    )
    return [daily_block for _, daily_block in indexed_daily_blocks]


def build_step0002_error_lines(list_sorted_daily_blocks: list[list[list[str]]]) -> list[str]:
    """Build error lines for other or blank vehicle types in a step0002 TSV."""
    list_error_lines: list[str] = []
    for i_block_number, daily_block in enumerate(list_sorted_daily_blocks, start=1):
        vehicle_type_text = get_cell_value(daily_block[1], 1)
        i_sort_index = get_vehicle_type_sort_index(vehicle_type_text)
        if i_sort_index not in {OTHER_VEHICLE_TYPE_SORT_INDEX, BLANK_VEHICLE_TYPE_SORT_INDEX}:
            continue

        if i_sort_index == OTHER_VEHICLE_TYPE_SORT_INDEX:
            error_type_text = "その他車種"
        else:
            error_type_text = "車種空欄"

        list_error_lines.extend([
            f"[{error_type_text}]",
            f"NO: {i_block_number}",
            f"配送先: {get_cell_value(daily_block[0], 1)}",
            f"車種: {vehicle_type_text}",
            f"備考: {get_cell_value(daily_block[2], 1)}",
            "",
        ])

    return list_error_lines


def build_step0002_tsv_rows(step0001_tsv_rows: list[list[str]]) -> tuple[list[list[str]], list[str]]:
    """Build sorted step0002 TSV rows and related error lines from step0001 rows."""
    if len(step0001_tsv_rows) == 0:
        raise RuntimeError("step0001 TSVに行がありません。")

    step0002_tsv_rows = [step0001_tsv_rows[0]]
    list_sorted_daily_blocks = sort_step0001_daily_blocks(read_step0001_daily_blocks(step0001_tsv_rows))
    list_error_lines = build_step0002_error_lines(list_sorted_daily_blocks)

    for i_block_number, daily_block in enumerate(list_sorted_daily_blocks, start=1):
        step0002_tsv_rows.extend([
            [str(i_block_number), get_cell_value(daily_block[0], 1)],
            ["", get_cell_value(daily_block[1], 1)],
            ["", get_cell_value(daily_block[2], 1)],
        ])

    return step0002_tsv_rows, list_error_lines


def write_step0002_error_file(
    step0002_error_file_path: Path,
    step0002_tsv_file_path: Path,
    list_error_lines: list[str],
) -> None:
    """Write a step0002 error file."""
    list_output_lines = [
        f"対象ファイル: {step0002_tsv_file_path.name}",
        "",
        *list_error_lines,
    ]
    step0002_error_file_path.write_text("\n".join(list_output_lines).rstrip() + "\n", encoding="utf-8-sig")


def write_step0002_daily_tsv_file(step0001_tsv_file_path: Path) -> list[Path]:
    """Create a sorted step0002 TSV file and its error file when needed."""
    step0001_tsv_rows = read_tsv_rows(step0001_tsv_file_path)
    step0002_tsv_rows, list_error_lines = build_step0002_tsv_rows(step0001_tsv_rows)
    step0002_tsv_file_path = build_step0002_tsv_file_path(step0001_tsv_file_path)
    step0002_error_file_path = build_step0002_error_file_path(step0002_tsv_file_path)

    write_tsv_rows(step0002_tsv_file_path, step0002_tsv_rows)
    list_created_file_paths = [step0002_tsv_file_path]

    if len(list_error_lines) > 0:
        write_step0002_error_file(step0002_error_file_path, step0002_tsv_file_path, list_error_lines)
        list_created_file_paths.append(step0002_error_file_path)
    elif step0002_error_file_path.exists():
        step0002_error_file_path.unlink()

    return list_created_file_paths


def write_step0002_daily_tsv_files(list_step0001_tsv_file_paths: list[Path]) -> list[Path]:
    """Create sorted step0002 TSV files from step0001 TSV files."""
    list_created_file_paths: list[Path] = []
    for step0001_tsv_file_path in list_step0001_tsv_file_paths:
        list_created_file_paths.extend(write_step0002_daily_tsv_file(step0001_tsv_file_path))

    return list_created_file_paths


def parse_step0001_daily_date(step0001_tsv_file_path: Path) -> datetime:
    """Parse a date from a step0001 daily TSV file name."""
    match = re.fullmatch(r"(.+_step0001_)([0-9]{4})年([0-9]{2})月([0-9]{2})日\.tsv", step0001_tsv_file_path.name)
    if match is None:
        raise RuntimeError(f"step0001日別TSVファイル名の日付を解析できません: {step0001_tsv_file_path}")

    return datetime(int(match.group(2)), int(match.group(3)), int(match.group(4)))


def get_step0001_daily_file_prefix(step0001_tsv_file_path: Path) -> str:
    """Return the file-name prefix before the date in a step0001 daily TSV path."""
    match = re.fullmatch(r"(.+_step0001_)([0-9]{4})年([0-9]{2})月([0-9]{2})日\.tsv", step0001_tsv_file_path.name)
    if match is None:
        raise RuntimeError(f"step0001日別TSVファイル名を解析できません: {step0001_tsv_file_path}")

    return match.group(1)


def build_step0001_daily_tsv_file_path(sample_step0001_tsv_file_path: Path, target_date: datetime) -> Path:
    """Build an expected step0001 daily TSV path for a target date."""
    step0001_file_prefix = get_step0001_daily_file_prefix(sample_step0001_tsv_file_path)
    step0001_file_name = f"{step0001_file_prefix}{target_date.year}年{target_date.month:02d}月{target_date.day:02d}日.tsv"
    return sample_step0001_tsv_file_path.with_name(step0001_file_name)


def build_monthly_step0001_tsv_file_path(sample_step0001_tsv_file_path: Path) -> Path:
    """Build a monthly step0001 TSV path from a daily step0001 TSV path."""
    target_date = parse_step0001_daily_date(sample_step0001_tsv_file_path)
    step0001_file_prefix = get_step0001_daily_file_prefix(sample_step0001_tsv_file_path)
    step0001_file_name = f"{step0001_file_prefix}{target_date.year}年{target_date.month:02d}月.tsv"
    return sample_step0001_tsv_file_path.with_name(step0001_file_name)


def format_monthly_step0001_header_date(target_date: datetime) -> str:
    """Format a date header for the monthly step0001 TSV."""
    return f"{target_date.year}/{target_date.month}/{target_date.day}"


def write_missing_step0001_error_file(step0001_tsv_file_path: Path, target_date: datetime) -> Path:
    """Write an error file for a missing step0001 daily TSV file."""
    step0001_error_file_path = step0001_tsv_file_path.with_name(f"{step0001_tsv_file_path.name}_error.txt")
    list_output_lines = [
        f"対象ファイル: {step0001_tsv_file_path.name}",
        "",
        "[日別step0001ファイルなし]",
        f"日付: {target_date.year}年{target_date.month:02d}月{target_date.day:02d}日",
        "内容: 月間step0001作成時に、対象日のstep0001 TSVファイルが見つかりませんでした。",
    ]
    step0001_error_file_path.write_text("\n".join(list_output_lines).rstrip() + "\n", encoding="utf-8-sig")
    return step0001_error_file_path


def get_monthly_step0001_target_dates(sample_step0001_tsv_file_path: Path) -> list[datetime]:
    """Return all dates in the month of a sample step0001 daily TSV path."""
    target_date = parse_step0001_daily_date(sample_step0001_tsv_file_path)
    i_last_day = get_last_day_of_month(target_date)
    return [datetime(target_date.year, target_date.month, i_day) for i_day in range(1, i_last_day + 1)]


def build_monthly_step0001_tsv_rows(
    sample_step0001_tsv_file_path: Path,
    dict_daily_blocks_by_date: dict[datetime, list[list[list[str]]]],
    first_column_header: str,
) -> list[list[str]]:
    """Build monthly step0001 TSV rows from daily step0001 blocks."""
    list_target_dates = get_monthly_step0001_target_dates(sample_step0001_tsv_file_path)
    i_max_block_count = max((len(dict_daily_blocks_by_date.get(target_date, [])) for target_date in list_target_dates), default=0)
    monthly_header_row = [first_column_header] + [format_monthly_step0001_header_date(target_date) for target_date in list_target_dates]
    monthly_step0001_tsv_rows: list[list[str]] = [monthly_header_row]

    for i_block_index in range(i_max_block_count):
        destination_row = [str(i_block_index + 1)]
        vehicle_type_row = [""]
        note_row = [""]

        for target_date in list_target_dates:
            list_daily_blocks = dict_daily_blocks_by_date.get(target_date, [])
            if i_block_index < len(list_daily_blocks):
                daily_block = list_daily_blocks[i_block_index]
                destination_row.append(get_cell_value(daily_block[0], 1))
                vehicle_type_row.append(get_cell_value(daily_block[1], 1))
                note_row.append(get_cell_value(daily_block[2], 1))
            else:
                destination_row.append("")
                vehicle_type_row.append("")
                note_row.append("")

        monthly_step0001_tsv_rows.extend([destination_row, vehicle_type_row, note_row])

    return monthly_step0001_tsv_rows


def write_monthly_step0001_tsv_file(list_step0001_tsv_file_paths: list[Path]) -> list[Path]:
    """Create a monthly step0001 TSV file from daily step0001 TSV files."""
    list_step0001_daily_tsv_file_paths = [
        step0001_tsv_file_path
        for step0001_tsv_file_path in list_step0001_tsv_file_paths
        if step0001_tsv_file_path.suffix.lower() == ".tsv" and "_step0001_" in step0001_tsv_file_path.name
    ]
    if len(list_step0001_daily_tsv_file_paths) == 0:
        return []

    sample_step0001_tsv_file_path = sorted(list_step0001_daily_tsv_file_paths, key=lambda file_path: file_path.name)[0]
    dict_step0001_daily_tsv_paths_by_date = {
        parse_step0001_daily_date(step0001_tsv_file_path): step0001_tsv_file_path
        for step0001_tsv_file_path in list_step0001_daily_tsv_file_paths
    }
    sample_step0001_tsv_rows = read_tsv_rows(sample_step0001_tsv_file_path)
    first_column_header = get_cell_value(sample_step0001_tsv_rows[0], 0) if len(sample_step0001_tsv_rows) > 0 else ""
    dict_daily_blocks_by_date: dict[datetime, list[list[list[str]]]] = {}
    list_created_file_paths: list[Path] = []

    for target_date in get_monthly_step0001_target_dates(sample_step0001_tsv_file_path):
        step0001_tsv_file_path = dict_step0001_daily_tsv_paths_by_date.get(target_date)
        if step0001_tsv_file_path is None:
            missing_step0001_tsv_file_path = build_step0001_daily_tsv_file_path(sample_step0001_tsv_file_path, target_date)
            list_created_file_paths.append(write_missing_step0001_error_file(missing_step0001_tsv_file_path, target_date))
            dict_daily_blocks_by_date[target_date] = []
            continue

        step0001_tsv_rows = read_tsv_rows(step0001_tsv_file_path)
        dict_daily_blocks_by_date[target_date] = read_step0001_daily_blocks(step0001_tsv_rows)

    monthly_step0001_tsv_file_path = build_monthly_step0001_tsv_file_path(sample_step0001_tsv_file_path)
    monthly_step0001_tsv_rows = build_monthly_step0001_tsv_rows(
        sample_step0001_tsv_file_path,
        dict_daily_blocks_by_date,
        first_column_header,
    )
    write_tsv_rows(monthly_step0001_tsv_file_path, monthly_step0001_tsv_rows)
    list_created_file_paths.insert(0, monthly_step0001_tsv_file_path)
    return list_created_file_paths


def parse_step0002_daily_date(step0002_tsv_file_path: Path) -> datetime:
    """Parse a date from a step0002 daily TSV file name."""
    match = re.fullmatch(r"(.+_step0002_)([0-9]{4})年([0-9]{2})月([0-9]{2})日\.tsv", step0002_tsv_file_path.name)
    if match is None:
        raise RuntimeError(f"step0002日別TSVファイル名の日付を解析できません: {step0002_tsv_file_path}")

    return datetime(int(match.group(2)), int(match.group(3)), int(match.group(4)))


def get_step0002_daily_file_prefix(step0002_tsv_file_path: Path) -> str:
    """Return the file-name prefix before the date in a step0002 daily TSV path."""
    match = re.fullmatch(r"(.+_step0002_)([0-9]{4})年([0-9]{2})月([0-9]{2})日\.tsv", step0002_tsv_file_path.name)
    if match is None:
        raise RuntimeError(f"step0002日別TSVファイル名を解析できません: {step0002_tsv_file_path}")

    return match.group(1)


def get_last_day_of_month(target_date: datetime) -> int:
    """Return the last day number of the target month."""
    if target_date.month == 12:
        next_month_date = datetime(target_date.year + 1, 1, 1)
    else:
        next_month_date = datetime(target_date.year, target_date.month + 1, 1)

    return (next_month_date - timedelta(days=1)).day


def build_step0002_daily_tsv_file_path(sample_step0002_tsv_file_path: Path, target_date: datetime) -> Path:
    """Build an expected step0002 daily TSV path for a target date."""
    step0002_file_prefix = get_step0002_daily_file_prefix(sample_step0002_tsv_file_path)
    step0002_file_name = f"{step0002_file_prefix}{target_date.year}年{target_date.month:02d}月{target_date.day:02d}日.tsv"
    return sample_step0002_tsv_file_path.with_name(step0002_file_name)


def build_monthly_step0002_tsv_file_path(sample_step0002_tsv_file_path: Path) -> Path:
    """Build a monthly step0002 TSV path from a daily step0002 TSV path."""
    target_date = parse_step0002_daily_date(sample_step0002_tsv_file_path)
    step0002_file_prefix = get_step0002_daily_file_prefix(sample_step0002_tsv_file_path)
    step0002_file_name = f"{step0002_file_prefix}{target_date.year}年{target_date.month:02d}月.tsv"
    return sample_step0002_tsv_file_path.with_name(step0002_file_name)


def format_monthly_step0002_header_date(target_date: datetime) -> str:
    """Format a date header for the monthly step0002 TSV."""
    return f"{target_date.year}/{target_date.month}/{target_date.day}"


def read_step0002_daily_blocks(step0002_tsv_rows: list[list[str]]) -> list[list[list[str]]]:
    """Read three-row daily blocks from step0002 TSV rows."""
    return read_step0001_daily_blocks(step0002_tsv_rows)


def build_step0003_tsv_file_path(step0002_tsv_file_path: Path) -> Path:
    """Build a step0003 TSV output path from a step0002 TSV path."""
    if "_step0002_" not in step0002_tsv_file_path.name:
        raise RuntimeError(f"step0002 TSVファイル名ではありません: {step0002_tsv_file_path}")

    step0003_file_name = step0002_tsv_file_path.name.replace("_step0002_", "_step0003_", 1)
    return step0002_tsv_file_path.with_name(step0003_file_name)


def build_step0003_tsv_rows(step0002_tsv_rows: list[list[str]]) -> list[list[str]]:
    """Build step0003 TSV rows by joining each step0002 three-row block with comma-space text."""
    if len(step0002_tsv_rows) == 0:
        raise RuntimeError("step0002 TSVに行がありません。")

    step0003_tsv_rows = [step0002_tsv_rows[0]]
    for daily_block in read_step0002_daily_blocks(step0002_tsv_rows):
        no_text = get_cell_value(daily_block[0], 0)
        joined_text = ", ".join([
            get_cell_value(daily_block[0], 1),
            get_cell_value(daily_block[1], 1),
            get_cell_value(daily_block[2], 1),
        ])
        step0003_tsv_rows.append([no_text, joined_text])

    return step0003_tsv_rows


def write_step0003_daily_tsv_file(step0002_tsv_file_path: Path) -> Path:
    """Create a step0003 TSV file from a step0002 TSV file."""
    step0002_tsv_rows = read_tsv_rows(step0002_tsv_file_path)
    step0003_tsv_rows = build_step0003_tsv_rows(step0002_tsv_rows)
    step0003_tsv_file_path = build_step0003_tsv_file_path(step0002_tsv_file_path)
    write_tsv_rows(step0003_tsv_file_path, step0003_tsv_rows)
    return step0003_tsv_file_path


def write_step0003_daily_tsv_files(list_step0002_tsv_file_paths: list[Path]) -> list[Path]:
    """Create step0003 TSV files from provided step0002 TSV files."""
    list_created_file_paths: list[Path] = []
    for step0002_tsv_file_path in list_step0002_tsv_file_paths:
        if step0002_tsv_file_path.suffix.lower() != ".tsv" or "_step0002_" not in step0002_tsv_file_path.name:
            continue

        list_created_file_paths.append(write_step0003_daily_tsv_file(step0002_tsv_file_path))

    return list_created_file_paths


def build_step0004_tsv_file_path(step0003_tsv_file_path: Path) -> Path:
    """Build a step0004 TSV output path from a step0003 TSV path."""
    if "_step0003_" not in step0003_tsv_file_path.name:
        raise RuntimeError(f"step0003 TSVファイル名ではありません: {step0003_tsv_file_path}")

    step0004_file_name = step0003_tsv_file_path.name.replace("_step0003_", "_step0004_", 1)
    return step0003_tsv_file_path.with_name(step0004_file_name)


def build_step0004_tsv_rows(step0003_tsv_rows: list[list[str]]) -> list[list[str]]:
    """Build step0004 TSV rows by removing the first column from each step0003 row."""
    return [step0003_tsv_row[1:] for step0003_tsv_row in step0003_tsv_rows]


def write_step0004_daily_tsv_file(step0003_tsv_file_path: Path) -> Path:
    """Create a step0004 TSV file from a step0003 TSV file."""
    step0003_tsv_rows = read_tsv_rows(step0003_tsv_file_path)
    step0004_tsv_rows = build_step0004_tsv_rows(step0003_tsv_rows)
    step0004_tsv_file_path = build_step0004_tsv_file_path(step0003_tsv_file_path)
    write_tsv_rows(step0004_tsv_file_path, step0004_tsv_rows)
    return step0004_tsv_file_path


def write_step0004_daily_tsv_files(list_step0003_tsv_file_paths: list[Path]) -> list[Path]:
    """Create step0004 TSV files from provided step0003 TSV files."""
    list_created_file_paths: list[Path] = []
    for step0003_tsv_file_path in list_step0003_tsv_file_paths:
        if step0003_tsv_file_path.suffix.lower() != ".tsv" or "_step0003_" not in step0003_tsv_file_path.name:
            continue

        list_created_file_paths.append(write_step0004_daily_tsv_file(step0003_tsv_file_path))

    return list_created_file_paths


def write_missing_step0002_error_file(step0002_tsv_file_path: Path, target_date: datetime) -> Path:
    """Write an error file for a missing step0002 daily TSV file."""
    step0002_error_file_path = build_step0002_error_file_path(step0002_tsv_file_path)
    list_output_lines = [
        f"対象ファイル: {step0002_tsv_file_path.name}",
        "",
        "[日別step0002ファイルなし]",
        f"日付: {target_date.year}年{target_date.month:02d}月{target_date.day:02d}日",
        "内容: 月間step0002作成時に、対象日のstep0002 TSVファイルが見つかりませんでした。",
    ]
    step0002_error_file_path.write_text("\n".join(list_output_lines).rstrip() + "\n", encoding="utf-8-sig")
    return step0002_error_file_path


def get_monthly_step0002_target_dates(sample_step0002_tsv_file_path: Path) -> list[datetime]:
    """Return all dates in the month of a sample step0002 daily TSV path."""
    target_date = parse_step0002_daily_date(sample_step0002_tsv_file_path)
    i_last_day = get_last_day_of_month(target_date)
    return [datetime(target_date.year, target_date.month, i_day) for i_day in range(1, i_last_day + 1)]


def build_monthly_step0002_tsv_rows(
    sample_step0002_tsv_file_path: Path,
    dict_daily_blocks_by_date: dict[datetime, list[list[list[str]]]],
    first_column_header: str,
) -> list[list[str]]:
    """Build monthly step0002 TSV rows from daily step0002 blocks."""
    list_target_dates = get_monthly_step0002_target_dates(sample_step0002_tsv_file_path)
    i_max_block_count = max((len(dict_daily_blocks_by_date.get(target_date, [])) for target_date in list_target_dates), default=0)
    monthly_header_row = [first_column_header] + [format_monthly_step0002_header_date(target_date) for target_date in list_target_dates]
    monthly_step0002_tsv_rows: list[list[str]] = [monthly_header_row]

    for i_block_index in range(i_max_block_count):
        destination_row = [str(i_block_index + 1)]
        vehicle_type_row = [""]
        note_row = [""]

        for target_date in list_target_dates:
            list_daily_blocks = dict_daily_blocks_by_date.get(target_date, [])
            if i_block_index < len(list_daily_blocks):
                daily_block = list_daily_blocks[i_block_index]
                destination_row.append(get_cell_value(daily_block[0], 1))
                vehicle_type_row.append(get_cell_value(daily_block[1], 1))
                note_row.append(get_cell_value(daily_block[2], 1))
            else:
                destination_row.append("")
                vehicle_type_row.append("")
                note_row.append("")

        monthly_step0002_tsv_rows.extend([destination_row, vehicle_type_row, note_row])

    return monthly_step0002_tsv_rows


def write_monthly_step0002_tsv_file(list_step0002_tsv_file_paths: list[Path]) -> list[Path]:
    """Create a monthly step0002 TSV file from daily step0002 TSV files."""
    list_step0002_daily_tsv_file_paths = [
        step0002_tsv_file_path
        for step0002_tsv_file_path in list_step0002_tsv_file_paths
        if step0002_tsv_file_path.suffix.lower() == ".tsv" and "_step0002_" in step0002_tsv_file_path.name
    ]
    if len(list_step0002_daily_tsv_file_paths) == 0:
        return []

    sample_step0002_tsv_file_path = sorted(list_step0002_daily_tsv_file_paths, key=lambda file_path: file_path.name)[0]
    dict_step0002_daily_tsv_paths_by_date = {
        parse_step0002_daily_date(step0002_tsv_file_path): step0002_tsv_file_path
        for step0002_tsv_file_path in list_step0002_daily_tsv_file_paths
    }
    sample_step0002_tsv_rows = read_tsv_rows(sample_step0002_tsv_file_path)
    first_column_header = get_cell_value(sample_step0002_tsv_rows[0], 0) if len(sample_step0002_tsv_rows) > 0 else ""
    dict_daily_blocks_by_date: dict[datetime, list[list[list[str]]]] = {}
    list_created_file_paths: list[Path] = []

    for target_date in get_monthly_step0002_target_dates(sample_step0002_tsv_file_path):
        step0002_tsv_file_path = dict_step0002_daily_tsv_paths_by_date.get(target_date)
        if step0002_tsv_file_path is None:
            missing_step0002_tsv_file_path = build_step0002_daily_tsv_file_path(sample_step0002_tsv_file_path, target_date)
            list_created_file_paths.append(write_missing_step0002_error_file(missing_step0002_tsv_file_path, target_date))
            dict_daily_blocks_by_date[target_date] = []
            continue

        step0002_tsv_rows = read_tsv_rows(step0002_tsv_file_path)
        dict_daily_blocks_by_date[target_date] = read_step0002_daily_blocks(step0002_tsv_rows)

    monthly_step0002_tsv_file_path = build_monthly_step0002_tsv_file_path(sample_step0002_tsv_file_path)
    monthly_step0002_tsv_rows = build_monthly_step0002_tsv_rows(
        sample_step0002_tsv_file_path,
        dict_daily_blocks_by_date,
        first_column_header,
    )
    write_tsv_rows(monthly_step0002_tsv_file_path, monthly_step0002_tsv_rows)
    list_created_file_paths.insert(0, monthly_step0002_tsv_file_path)
    return list_created_file_paths


def main() -> int:
    """Read an Excel file path from the command line and create a same-named TSV."""
    if len(sys.argv) != 2:
        print("Excelファイル名を1つ指定してください。", file=sys.stderr)
        return 1

    excel_file_path = Path(sys.argv[1]).resolve()
    error_message = validate_excel_file_path(excel_file_path)
    if error_message is not None:
        print(error_message, file=sys.stderr)
        return 1

    try:
        tsv_file_path = write_excel_values_to_tsv(excel_file_path)
        list_daily_tsv_file_paths = write_step0001_daily_tsv_files(tsv_file_path)
        list_monthly_step0001_file_paths = write_monthly_step0001_tsv_file(list_daily_tsv_file_paths)
        list_step0002_tsv_file_paths = write_step0002_daily_tsv_files(list_daily_tsv_file_paths)
        list_step0003_tsv_file_paths = write_step0003_daily_tsv_files(list_step0002_tsv_file_paths)
        list_step0004_tsv_file_paths = write_step0004_daily_tsv_files(list_step0003_tsv_file_paths)
        list_monthly_step0002_file_paths = write_monthly_step0002_tsv_file(list_step0002_tsv_file_paths)
    except Exception as exception:
        print(f"TSV作成に失敗しました: {exception}", file=sys.stderr)
        return 1

    print(tsv_file_path)
    for daily_tsv_file_path in list_daily_tsv_file_paths:
        print(daily_tsv_file_path)
    for monthly_step0001_file_path in list_monthly_step0001_file_paths:
        print(monthly_step0001_file_path)
    for step0002_tsv_file_path in list_step0002_tsv_file_paths:
        print(step0002_tsv_file_path)
    for step0003_tsv_file_path in list_step0003_tsv_file_paths:
        print(step0003_tsv_file_path)
    for step0004_tsv_file_path in list_step0004_tsv_file_paths:
        print(step0004_tsv_file_path)
    for monthly_step0002_file_path in list_monthly_step0002_file_paths:
        print(monthly_step0002_file_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
