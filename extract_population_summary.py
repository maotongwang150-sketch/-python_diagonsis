"""Extract provincial population indicators from the source .xls workbooks.

The repository only contains legacy BIFF8 .xls files and the execution
image may not include Excel libraries.  This script therefore implements a
small OLE/BIFF reader for the record types used by the source files and writes
a dependency-free .xlsx workbook.
"""

from __future__ import annotations

import argparse
import html
import struct
import zipfile
from pathlib import Path

END_OF_CHAIN = 0xFFFFFFFE
FREE_SECTOR = 0xFFFFFFFF

SUMMARY_FILE = Path("第一卷、概要1.xls")
AGE_FILE = Path("第三卷、年龄.xls")
DEFAULT_OUTPUT_DIR = r"C:\Users\29979\Desktop\结果"
OUTPUT_FILENAME = "人口指标汇总.xlsx"
HEADERS = [
    "省份",
    "总人口",
    "男性人口",
    "女性人口",
    "性别比",
    "城镇人口",
    "乡村人口",
    "城镇化率(%)",
    "老龄化率(%)",
]


def _ole_stream(path: Path, stream_name: str = "Workbook") -> bytes:
    data = path.read_bytes()
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError(f"{path} is not an OLE compound document")

    def u16(offset: int) -> int:
        return struct.unpack_from("<H", data, offset)[0]

    def u32(offset: int) -> int:
        return struct.unpack_from("<I", data, offset)[0]

    sector_size = 1 << u16(30)
    first_dir_sector = u32(48)
    fat_sector_count = u32(44)
    difat = list(struct.unpack_from("<109I", data, 76))
    fat: list[int] = []
    for sector_id in difat:
        if sector_id in (FREE_SECTOR, END_OF_CHAIN):
            continue
        offset = (sector_id + 1) * sector_size
        fat.extend(struct.unpack_from(f"<{sector_size // 4}I", data, offset))
        if len(fat) >= fat_sector_count * (sector_size // 4):
            break

    def read_chain(start_sector: int) -> bytes:
        output = bytearray()
        sector_id = start_sector
        visited: set[int] = set()
        while (
            sector_id not in (FREE_SECTOR, END_OF_CHAIN)
            and sector_id < len(fat)
            and sector_id not in visited
        ):
            visited.add(sector_id)
            offset = (sector_id + 1) * sector_size
            output.extend(data[offset : offset + sector_size])
            sector_id = fat[sector_id]
        return bytes(output)

    directory = read_chain(first_dir_sector)
    entries: dict[str, tuple[int, int]] = {}
    for offset in range(0, len(directory), 128):
        entry = directory[offset : offset + 128]
        if len(entry) < 128:
            break
        name_size = struct.unpack_from("<H", entry, 64)[0]
        if name_size < 2:
            continue
        name = entry[: name_size - 2].decode("utf-16le", "ignore")
        start_sector = struct.unpack_from("<I", entry, 116)[0]
        size = struct.unpack_from("<Q", entry, 120)[0]
        entries[name] = (start_sector, size)

    start_sector, size = entries.get(stream_name) or entries["Book"]
    return read_chain(start_sector)[:size]


def _parse_unicode_string(buffer: bytes | bytearray, offset: int, char_count: int | None = None) -> tuple[str, int]:
    if char_count is None:
        char_count = struct.unpack_from("<H", buffer, offset)[0]
        offset += 2
    flags = buffer[offset]
    offset += 1
    has_rich_text = flags & 0x08
    has_ext_data = flags & 0x04
    is_utf16 = flags & 0x01
    rich_text_runs = 0
    ext_data_size = 0
    if has_rich_text:
        rich_text_runs = struct.unpack_from("<H", buffer, offset)[0]
        offset += 2
    if has_ext_data:
        ext_data_size = struct.unpack_from("<I", buffer, offset)[0]
        offset += 4
    if is_utf16:
        size = char_count * 2
        value = buffer[offset : offset + size].decode("utf-16le", "ignore")
    else:
        size = char_count
        value = buffer[offset : offset + size].decode("latin1", "ignore")
    offset += size + rich_text_runs * 4 + ext_data_size
    return value, offset


def _decode_rk(raw_value: int) -> float:
    is_divided_by_100 = raw_value & 1
    is_integer = raw_value & 2
    value_bits = raw_value >> 2
    if is_integer:
        if value_bits & (1 << 29):
            value_bits -= 1 << 30
        value = float(value_bits)
    else:
        value = struct.unpack("<d", struct.pack("<Q", (raw_value & 0xFFFFFFFC) << 32))[0]
    return value / 100 if is_divided_by_100 else value


def _parse_first_sheet(path: Path) -> dict[tuple[int, int], str | float]:
    workbook = _ole_stream(path)
    records: list[tuple[int, bytes]] = []
    offset = 0
    while offset + 4 <= len(workbook):
        record_type, size = struct.unpack_from("<HH", workbook, offset)
        offset += 4
        records.append((record_type, workbook[offset : offset + size]))
        offset += size

    sheet_offsets: list[int] = []
    shared_strings: list[str] = []
    index = 0
    while index < len(records):
        record_type, payload = records[index]
        if record_type == 0x0085:  # BOUNDSHEET
            sheet_offsets.append(struct.unpack_from("<I", payload, 0)[0])
        elif record_type == 0x00FC:  # SST, followed by optional CONTINUE records
            sst_payload = bytearray(payload)
            next_index = index + 1
            while next_index < len(records) and records[next_index][0] == 0x003C:
                sst_payload.extend(records[next_index][1])
                next_index += 1
            _, unique_count = struct.unpack_from("<II", sst_payload, 0)
            string_offset = 8
            for _ in range(unique_count):
                value, string_offset = _parse_unicode_string(sst_payload, string_offset)
                shared_strings.append(value)
        index += 1

    start = min(sheet_offsets)
    following_offsets = [sheet_offset for sheet_offset in sheet_offsets if sheet_offset > start]
    end = min(following_offsets) if following_offsets else len(workbook)
    cells: dict[tuple[int, int], str | float] = {}
    offset = start
    while offset + 4 <= end:
        record_type, size = struct.unpack_from("<HH", workbook, offset)
        offset += 4
        payload = workbook[offset : offset + size]
        offset += size
        if record_type == 0x00FD:  # LABELSST
            row, col, _, sst_index = struct.unpack_from("<HHHI", payload, 0)
            cells[(row, col)] = shared_strings[sst_index]
        elif record_type == 0x0203:  # NUMBER
            row, col, _ = struct.unpack_from("<HHH", payload, 0)
            cells[(row, col)] = struct.unpack_from("<d", payload, 6)[0]
        elif record_type == 0x027E:  # RK
            row, col, _, raw_value = struct.unpack_from("<HHHI", payload, 0)
            cells[(row, col)] = _decode_rk(raw_value)
        elif record_type == 0x00BD:  # MULRK
            row, first_col = struct.unpack_from("<HH", payload, 0)
            last_col = struct.unpack_from("<H", payload, len(payload) - 2)[0]
            payload_offset = 4
            for col in range(first_col, last_col + 1):
                _, raw_value = struct.unpack_from("<HI", payload, payload_offset)
                payload_offset += 6
                cells[(row, col)] = _decode_rk(raw_value)
    return cells


def _rows_for_table(cells: dict[tuple[int, int], str | float], title: str) -> dict[str, int]:
    title_row = next(row for (row, col), value in cells.items() if col == 0 and value == title)
    next_title_rows = [
        row
        for (row, col), value in cells.items()
        if col == 0 and row > title_row and isinstance(value, str) and value.startswith("表")
    ]
    stop_row = min(next_title_rows) if next_title_rows else max(row for row, _ in cells) + 1
    return {
        str(value).strip(): row
        for (row, col), value in cells.items()
        if col == 0 and title_row < row < stop_row and str(value).strip() and not str(value).startswith("地")
    }


def _build_summary_rows() -> list[list[str | int | float]]:
    summary_cells = _parse_first_sheet(SUMMARY_FILE)
    age_cells = _parse_first_sheet(AGE_FILE)

    total_rows = _rows_for_table(summary_cells, "表1-1    省、自治区、直辖市的户数、人口数和性别比")
    city_rows = _rows_for_table(summary_cells, "表1-1a    省、自治区、直辖市的户数、人口数和性别比(城市)")
    town_rows = _rows_for_table(summary_cells, "表1-1b    省、自治区、直辖市的户数、人口数和性别比(镇)")
    rural_rows = _rows_for_table(summary_cells, "表1-1c    省、自治区、直辖市的户数、人口数和性别比(乡村)")
    age_rows = _rows_for_table(age_cells, "表3—2    省、自治区、直辖市年龄构成指数")

    rows: list[list[str | int | float]] = []
    for province, total_row in total_rows.items():
        if province == "合计":
            continue
        city_population = int(summary_cells[(city_rows[province], 4)])
        town_population = int(summary_cells[(town_rows[province], 4)])
        urban_population = city_population + town_population
        rural_population = int(summary_cells[(rural_rows[province], 4)])
        total_population = int(summary_cells[(total_row, 4)])
        rows.append(
            [
                province,
                total_population,
                int(summary_cells[(total_row, 5)]),
                int(summary_cells[(total_row, 6)]),
                summary_cells[(total_row, 7)],
                urban_population,
                rural_population,
                round(urban_population / total_population * 100, 2),
                age_cells[(age_rows[province], 8)],
            ]
        )
    return rows


def _column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _cell_xml(row: int, col: int, value: str | int | float) -> str:
    ref = f"{_column_name(col)}{row}"
    if isinstance(value, str):
        return f'<c r="{ref}" t="inlineStr"><is><t>{html.escape(value)}</t></is></c>'
    if isinstance(value, int) or float(value).is_integer():
        display_value = str(int(value))
    else:
        display_value = f"{float(value):.2f}".rstrip("0").rstrip(".")
    return f'<c r="{ref}"><v>{display_value}</v></c>'


def _write_xlsx(path: Path, rows: list[list[str | int | float]]) -> None:
    sheet_rows = [HEADERS, *rows]
    worksheet_rows = []
    for row_number, values in enumerate(sheet_rows, start=1):
        cells = "".join(_cell_xml(row_number, col_number, value) for col_number, value in enumerate(values, start=1))
        worksheet_rows.append(f'<row r="{row_number}">{cells}</row>')
    worksheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:I{last_row}"/>
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="7" width="14" customWidth="1"/><col min="8" max="9" width="14" customWidth="1"/></cols>
<sheetData>{sheet_data}</sheetData>
</worksheet>
""".format(last_row=len(sheet_rows), sheet_data="".join(worksheet_rows))
    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="人口指标汇总" sheetId="1" r:id="rId1"/></sheets></workbook>
"""
    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
"""
    workbook_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>
"""
    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", rels_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成省级人口指标汇总 xlsx 文件。")
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"输出目录，默认写入 {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--output-file",
        help="完整输出文件路径；设置后会覆盖 --output-dir。",
    )
    return parser.parse_args()


def _resolve_output_file(args: argparse.Namespace) -> Path:
    if args.output_file:
        output_file = Path(args.output_file)
    else:
        output_file = Path(args.output_dir) / OUTPUT_FILENAME
    output_file.parent.mkdir(parents=True, exist_ok=True)
    return output_file


def main() -> None:
    args = _parse_args()
    output_file = _resolve_output_file(args)
    rows = _build_summary_rows()
    _write_xlsx(output_file, rows)
    print(f"Wrote {output_file} with {len(rows)} provincial rows.")


if __name__ == "__main__":
    main()
