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


def read_shared_strings(excel_archive: zipfile.ZipFile) -> list[str]:
    """Read shared strings from an Excel archive."""
    shared_strings_root = read_xml_from_zip(excel_archive, "xl/sharedStrings.xml")
    if shared_strings_root is None:
        return []

    list_shared_strings: list[str] = []
    for string_item in shared_strings_root.findall(namespace_tag(EXCEL_NAMESPACE, "si")):
        list_text_parts: list[str] = []
        for text_node in string_item.iter(namespace_tag(EXCEL_NAMESPACE, "t")):
            list_text_parts.append(text_node.text or "")
        list_shared_strings.append("".join(list_text_parts))

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
    """Read an inline string from a worksheet cell."""
    inline_string_node = cell_node.find(namespace_tag(EXCEL_NAMESPACE, "is"))
    if inline_string_node is None:
        return ""

    list_text_parts: list[str] = []
    for text_node in inline_string_node.iter(namespace_tag(EXCEL_NAMESPACE, "t")):
        list_text_parts.append(text_node.text or "")
    return "".join(list_text_parts)


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
    except Exception as exception:
        print(f"TSV作成に失敗しました: {exception}", file=sys.stderr)
        return 1

    print(tsv_file_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
