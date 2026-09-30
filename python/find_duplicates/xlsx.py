"""Excel (.xlsx) report writer. Needs neither Excel nor third-party packages.

Mirrors Export-DuplicateReport in src/DuplicateFinder.psm1 and writes the same
workbook: one row per duplicated file with the columns File Name, Last Modified,
Size (bytes), MD5, Copies, then "Location 1..N" holding the full folder path of
every copy. The header row is frozen and filtered.
"""

from __future__ import annotations

import os
import re
import uuid
import zipfile
from datetime import datetime
from typing import List, Sequence, Union
from xml.sax.saxutils import escape

from .matcher import DuplicateSet

EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
FIXED_COLUMNS = (
    ("File Name", 40),
    ("Last Modified", 20),
    ("Size (bytes)", 14),
    ("MD5", 34),
    ("Copies", 8),
)
LOCATION_COLUMN_WIDTH = 60
STYLE_BOLD = 1
STYLE_DATE = 2

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
EXCEL_EPOCH = datetime(1899, 12, 30)  # day 0 of Excel/OLE Automation dates

# File names may contain control characters (Linux), unpaired surrogates (NTFS,
# or undecodable bytes on Linux) or non-characters that XML 1.0 cannot represent.
_INVALID_XML_CHARS = re.compile(
    "[{}-{}{}{}{}-{}{}-{}{}{}]".format(
        chr(0x00), chr(0x08), chr(0x0B), chr(0x0C), chr(0x0E), chr(0x1F),
        chr(0xD800), chr(0xDFFF), chr(0xFFFE), chr(0xFFFF),
    )
)
REPLACEMENT_CHAR = chr(0xFFFD)

CONTENT_TYPES = (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    "</Types>"
)
ROOT_RELS = (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    "</Relationships>"
)
WORKBOOK_RELS = (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    "</Relationships>"
)
# Style 0 = default, 1 = bold header, 2 = date/time.
STYLES = (
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>'
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="3">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    "</cellXfs>"
    "</styleSheet>"
)

CellValue = Union[str, int, datetime]


def column_name(index: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA ..."""
    if not 1 <= index <= EXCEL_MAX_COLUMNS:
        raise ValueError(f"Column index {index} is outside 1..{EXCEL_MAX_COLUMNS}.")
    name = ""
    while index > 0:
        index, digit = divmod(index - 1, 26)
        name = chr(65 + digit) + name
    return name


def xml_safe_text(text: str) -> str:
    """Escape text for XML, replacing characters XML 1.0 cannot hold with U+FFFD."""
    return escape(_INVALID_XML_CHARS.sub(REPLACEMENT_CHAR, text))


def excel_serial(value: datetime) -> float:
    """A datetime as an Excel date serial number (days since 1899-12-30)."""
    return (value - EXCEL_EPOCH).total_seconds() / 86400


def _cell(reference: str, value: CellValue, style: int = 0) -> str:
    style_attr = f' s="{style}"' if style else ""
    if isinstance(value, datetime):
        return f'<c r="{reference}"{style_attr}><v>{excel_serial(value)!r}</v></c>'
    if isinstance(value, int):
        return f'<c r="{reference}"{style_attr}><v>{value}</v></c>'
    return (
        f'<c r="{reference}"{style_attr} t="inlineStr"><is>'
        f'<t xml:space="preserve">{xml_safe_text(value)}</t></is></c>'
    )


def _row(number: int, values: Sequence[CellValue]) -> str:
    cells = []
    for index, value in enumerate(values, start=1):
        if number == 1:
            style = STYLE_BOLD
        elif isinstance(value, datetime):
            style = STYLE_DATE
        else:
            style = 0
        cells.append(_cell(f"{column_name(index)}{number}", value, style))
    return f'<row r="{number}">{"".join(cells)}</row>'


def _worksheet(duplicates: Sequence[DuplicateSet], location_columns: int, last_column: str) -> str:
    last_row = len(duplicates) + 1
    widths = [width for _, width in FIXED_COLUMNS] + [LOCATION_COLUMN_WIDTH] * location_columns
    cols = "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, start=1)
    )
    headers: List[CellValue] = [header for header, _ in FIXED_COLUMNS]
    headers += [f"Location {n}" for n in range(1, location_columns + 1)]

    # One row per duplicated file; one column per folder holding a copy.
    rows = [_row(1, headers)]
    for number, dup in enumerate(duplicates, start=2):
        values: List[CellValue] = [dup.file_name, dup.last_write_time, dup.size_bytes, dup.md5, dup.count]
        rows.append(_row(number, values + list(dup.folders)))

    return (
        f'{XML_DECLARATION}<worksheet xmlns="{MAIN_NS}">'
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        "</sheetView></sheetViews>"
        f"<cols>{cols}</cols>"
        f'<sheetData>{"".join(rows)}</sheetData>'
        f'<autoFilter ref="A1:{last_column}{last_row}"/>'
        "</worksheet>"
    )


def _workbook(last_column: str, last_row: int) -> str:
    return (
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Duplicates" sheetId="1" r:id="rId1"/></sheets>'
        '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'
        f"Duplicates!$A$1:${last_column}${last_row}</definedName></definedNames>"
        "</workbook>"
    )


def export_duplicate_report(duplicates: Sequence[DuplicateSet], path: str) -> None:
    """Save duplicate sets to an .xlsx workbook at ``path``."""
    path = os.path.abspath(path)
    max_copies = max([1] + [len(dup.folders) for dup in duplicates])

    column_count = len(FIXED_COLUMNS) + max_copies
    if column_count > EXCEL_MAX_COLUMNS:
        raise ValueError(
            f"A file has {max_copies} copies; Excel supports at most "
            f"{EXCEL_MAX_COLUMNS - len(FIXED_COLUMNS)} location columns."
        )
    if len(duplicates) + 1 > EXCEL_MAX_ROWS:
        raise ValueError(
            f"Found {len(duplicates)} duplicated files; Excel supports at most {EXCEL_MAX_ROWS - 1} rows."
        )
    last_column = column_name(column_count)
    last_row = len(duplicates) + 1

    # Build next to the target, then swap in, so a failure never leaves a half-written report.
    temp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    try:
        with zipfile.ZipFile(temp_path, "x", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", XML_DECLARATION + CONTENT_TYPES)
            archive.writestr("_rels/.rels", XML_DECLARATION + ROOT_RELS)
            archive.writestr("xl/workbook.xml", XML_DECLARATION + _workbook(last_column, last_row))
            archive.writestr("xl/_rels/workbook.xml.rels", XML_DECLARATION + WORKBOOK_RELS)
            archive.writestr("xl/styles.xml", XML_DECLARATION + STYLES)
            archive.writestr("xl/worksheets/sheet1.xml", _worksheet(duplicates, max_copies, last_column))
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
