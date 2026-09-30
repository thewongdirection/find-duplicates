"""Excel (.xlsx) report writer and reader. Needs neither Excel nor third-party packages.

Mirrors Export-DuplicateReport, Import-DuplicateReport and Import-DuplicateFolderReport
in src/DuplicateFinder.psm1 and writes the same workbook:

* sheet "Duplicates": one row per duplicated file with the columns File Name, Last
  Modified, Size (bytes), MD5, Copies, then "Location 1..N" holding the full folder
  path of every copy;
* sheet "Duplicate Folders" (only when folder sets are given): one row per
  duplicated folder with the columns Folder Name, Files, Sub Folders, Size (bytes),
  Copies, then "Location 1..N" holding the full path of every copy.

Header rows are frozen and filtered.
"""

from __future__ import annotations

import os
import re
import uuid
import zipfile
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple, Union
from xml.etree import ElementTree
from xml.sax.saxutils import escape

from .folders import DuplicateFolderSet
from .matcher import DuplicateSet

EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
EXCEL_MAX_CELL_TEXT = 32_767  # in UTF-16 code units, as Excel and .NET count them
FIXED_COLUMNS = (
    ("File Name", 40),
    ("Last Modified", 20),
    ("Size (bytes)", 14),
    ("MD5", 34),
    ("Copies", 8),
)
FOLDER_COLUMNS = (
    ("Folder Name", 40),
    ("Files", 10),
    ("Sub Folders", 12),
    ("Size (bytes)", 16),
    ("Copies", 8),
)
FILE_SHEET_NAME = "Duplicates"
FOLDER_SHEET_NAME = "Duplicate Folders"
RULES_SHEET_NAME = "Rules"
RULES_COLUMN_WIDTH = 150

# The matching rules, written on the Rules sheet. Shared word for word with PowerShell.
RULES_INTRO = (
    "Matching rules",
    "These rules decide what counts as a match on the other sheets of this workbook.",
)
FILE_RULES_TITLE = f"Sheet '{FILE_SHEET_NAME}': duplicate files"
FILE_RULES = (
    "A file is listed when another file has ALL of: the same name (ignoring upper/lower case), the same saved date "
    "(last modified, to the whole second) and the same contents (MD5 hash).",
    "Each row is one duplicated file. Each Location column is the full path of a folder that holds a copy.",
    "Files of 0 bytes are included unless the scan used -IgnoreEmptyFiles (Python: --ignore-empty-files).",
)
FOLDER_RULES_TITLE = f"Sheet '{FOLDER_SHEET_NAME}': duplicate folders"
FOLDER_RULES = (
    "A folder is listed when another folder has ALL of: the same name (ignoring upper/lower case), the same tree of "
    "files and sub folders (empty sub folders included), and every file matching the file at the same place in the "
    "other folder (same name, saved date and MD5).",
    "Only the top-most duplicates are listed: a sub folder is listed on its own only when one of its copies is "
    "outside a duplicate folder. Folders that contain no files are not listed.",
    "Each row is one duplicated folder. Each Location column is the full path of one copy.",
)
LOCATION_COLUMN_WIDTH = 60
STYLE_BOLD = 1
STYLE_DATE = 2

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MS_PER_DAY = 86_400_000
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

ROOT_RELS = (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
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
Columns = Sequence[Tuple[str, int]]


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
    """A datetime as an Excel date serial number (days since 1899-12-30).

    Truncated to the millisecond first, exactly like .NET DateTime.ToOADate.
    """
    return ((value - EXCEL_EPOCH) // timedelta(milliseconds=1)) / MS_PER_DAY


def from_excel_serial(serial: float) -> datetime:
    """An Excel date serial number as a datetime, rounded to the millisecond like .NET DateTime.FromOADate."""
    return EXCEL_EPOCH + timedelta(milliseconds=int(serial * MS_PER_DAY + 0.5))


def _cell(reference: str, value: CellValue, style: int = 0) -> str:
    style_attr = f' s="{style}"' if style else ""
    if isinstance(value, datetime):
        return f'<c r="{reference}"{style_attr}><v>{excel_serial(value)!r}</v></c>'
    if isinstance(value, int):
        return f'<c r="{reference}"{style_attr}><v>{value}</v></c>'
    # Only long text needs its UTF-16 length measured (it is at most twice len()).
    if len(value) > EXCEL_MAX_CELL_TEXT // 2:
        length = len(value.encode("utf-16-le", "surrogatepass")) // 2
        if length > EXCEL_MAX_CELL_TEXT:
            raise ValueError(
                f"Cell {reference} would hold {length} characters; Excel allows at most {EXCEL_MAX_CELL_TEXT}."
            )
    return (
        f'<c r="{reference}"{style_attr} t="inlineStr"><is>'
        f'<t xml:space="preserve">{xml_safe_text(value)}</t></is></c>'
    )


def _row(number: int, values: Sequence[CellValue], bold: bool = False) -> str:
    cells = []
    for index, value in enumerate(values, start=1):
        if bold:
            style = STYLE_BOLD
        elif isinstance(value, datetime):
            style = STYLE_DATE
        else:
            style = 0
        cells.append(_cell(f"{column_name(index)}{number}", value, style))
    return f'<row r="{number}">{"".join(cells)}</row>'


@dataclass
class _Sheet:
    name: str
    header_row: int
    headers: List[str]
    widths: List[int]
    rows: List[List[CellValue]]
    last_column: str
    last_row: int


@dataclass
class _RulesSheet:
    name: str
    lines: List[Optional[Tuple[str, bool]]]  # (text, bold), or None for a blank row


def _rules_sheet(include_folders: bool) -> _RulesSheet:
    """The Rules sheet: a title, a section per data sheet, and blank rows between."""
    lines: List[Optional[Tuple[str, bool]]] = [(RULES_INTRO[0], True)] + [(text, False) for text in RULES_INTRO[1:]]
    sections = [(FILE_RULES_TITLE, FILE_RULES)]
    if include_folders:
        sections.append((FOLDER_RULES_TITLE, FOLDER_RULES))
    for title, rules in sections:
        lines += [None, (title, True)] + [(text, False) for text in rules]
    return _RulesSheet(RULES_SHEET_NAME, lines)


def _rules_worksheet(sheet: _RulesSheet) -> str:
    rows = [
        _row(number, [line[0]], bold=line[1])
        for number, line in enumerate(sheet.lines, start=1)
        if line is not None
    ]
    return (
        f'{XML_DECLARATION}<worksheet xmlns="{MAIN_NS}">'
        f'<cols><col min="1" max="1" width="{RULES_COLUMN_WIDTH}" customWidth="1"/></cols>'
        f'<sheetData>{"".join(rows)}</sheetData>'
        "</worksheet>"
    )


def _sheet(name: str, columns: Columns, rows: List[List[CellValue]], noun: str) -> _Sheet:
    """Lay out one table sheet: fixed columns, then as many "Location N" columns as the
    longest row needs."""
    header_row = 1  # the rules are on their own sheet
    max_copies = max([1] + [len(row) - len(columns) for row in rows])
    column_count = len(columns) + max_copies
    if column_count > EXCEL_MAX_COLUMNS:
        raise ValueError(
            f"A {noun} has {max_copies} copies; Excel supports at most "
            f"{EXCEL_MAX_COLUMNS - len(columns)} location columns."
        )
    if len(rows) + header_row > EXCEL_MAX_ROWS:
        raise ValueError(
            f"Found {len(rows)} duplicated {noun}s; Excel supports at most {EXCEL_MAX_ROWS - header_row} rows."
        )
    return _Sheet(
        name=name,
        header_row=header_row,
        headers=[header for header, _ in columns] + [f"Location {n}" for n in range(1, max_copies + 1)],
        widths=[width for _, width in columns] + [LOCATION_COLUMN_WIDTH] * max_copies,
        rows=rows,
        last_column=column_name(column_count),
        last_row=len(rows) + header_row,
    )


def _worksheet(sheet: _Sheet) -> str:
    cols = "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(sheet.widths, start=1)
    )
    # The header row, then one row per duplicated item with one column per copy.
    rows = [_row(sheet.header_row, sheet.headers, bold=True)]
    rows += [_row(number, values) for number, values in enumerate(sheet.rows, start=sheet.header_row + 1)]
    return (
        f'{XML_DECLARATION}<worksheet xmlns="{MAIN_NS}">'
        '<sheetViews><sheetView workbookViewId="0">'
        f'<pane ySplit="{sheet.header_row}" topLeftCell="A{sheet.header_row + 1}" activePane="bottomLeft" state="frozen"/>'
        "</sheetView></sheetViews>"
        f"<cols>{cols}</cols>"
        f'<sheetData>{"".join(rows)}</sheetData>'
        f'<autoFilter ref="A{sheet.header_row}:{sheet.last_column}{sheet.last_row}"/>'
        "</worksheet>"
    )


def _package_parts(sheets: Sequence[Union[_Sheet, _RulesSheet]]) -> Dict[str, str]:
    """Every part of the package except the worksheets, as XML text."""
    numbered = list(enumerate(sheets, start=1))
    content_types = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i, _ in numbered
        )
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )
    workbook = (
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        + "".join(f'<sheet name="{s.name}" sheetId="{i}" r:id="rId{i}"/>' for i, s in numbered)
        + "</sheets><definedNames>"
        + "".join(
            f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i - 1}" hidden="1">'
            f"'{s.name}'!$A${s.header_row}:${s.last_column}${s.last_row}</definedName>"
            for i, s in numbered
            if isinstance(s, _Sheet)
        )
        + "</definedNames></workbook>"
    )
    workbook_rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i}.xml"/>'
            for i, _ in numbered
        )
        + f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    )
    return {
        "[Content_Types].xml": content_types,
        "_rels/.rels": ROOT_RELS,
        "xl/workbook.xml": workbook,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/styles.xml": STYLES,
    }


def export_duplicate_report(
    duplicates: Sequence[DuplicateSet], path: str, folders: Optional[Sequence[DuplicateFolderSet]] = None
) -> None:
    """Save duplicate sets to an .xlsx workbook at ``path``.

    With ``folders`` (even an empty list) the "Duplicate Folders" sheet is added. A
    "Rules" sheet always follows, stating the matching rules behind the other sheets.
    """
    path = os.path.abspath(path)
    sheets: List[Union[_Sheet, _RulesSheet]] = [
        _sheet(
            FILE_SHEET_NAME,
            FIXED_COLUMNS,
            [[d.file_name, d.last_write_time, d.size_bytes, d.md5, d.count] + list(d.folders) for d in duplicates],
            "file",
        )
    ]
    if folders is not None:
        sheets.append(
            _sheet(
                FOLDER_SHEET_NAME,
                FOLDER_COLUMNS,
                [[f.folder_name, f.file_count, f.folder_count, f.size_bytes, f.count] + list(f.folders) for f in folders],
                "folder",
            )
        )
    sheets.append(_rules_sheet(include_folders=folders is not None))

    # Build next to the target, then swap in, so a failure never leaves a half-written report.
    temp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    try:
        with zipfile.ZipFile(temp_path, "x", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, xml in _package_parts(sheets).items():
                archive.writestr(name, XML_DECLARATION + xml)
            for number, sheet in enumerate(sheets, start=1):
                xml = _rules_worksheet(sheet) if isinstance(sheet, _RulesSheet) else _worksheet(sheet)
                archive.writestr(f"xl/worksheets/sheet{number}.xml", xml)
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def column_index(name: str) -> int:
    """A -> 1, Z -> 26, AA -> 27 ..."""
    index = 0
    for letter in name.upper():
        index = index * 26 + ord(letter) - 64
    return index


def _cell_text(node: ElementTree.Element) -> str:
    """Text of an inline or shared string, including rich-text runs (phonetic runs excluded)."""
    parts = []
    for child in node:
        if child.tag == f"{{{MAIN_NS}}}t":
            parts.append(child.text or "")
        elif child.tag == f"{{{MAIN_NS}}}r":
            run_text = child.find(f"{{{MAIN_NS}}}t")
            parts.append((run_text.text or "") if run_text is not None else "")
    return "".join(parts)


def _read_xml(archive: zipfile.ZipFile, name: str) -> Optional[ElementTree.Element]:
    try:
        data = archive.read(name)
    except KeyError:
        return None
    # Reports never contain DTDs, so refuse them (no entity expansion).
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        raise ValueError(f"Part '{name}' contains a DTD, which reports never do.")
    return ElementTree.fromstring(data)


def _workbook_sheets(archive: zipfile.ZipFile) -> List[Tuple[str, str]]:
    """The workbook's sheets in order, as (name, part path) pairs."""
    workbook = _read_xml(archive, "xl/workbook.xml")
    rels = _read_xml(archive, "xl/_rels/workbook.xml.rels")
    if workbook is None or rels is None:
        raise ValueError("The file is not an Excel workbook.")
    targets = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{PACKAGE_RELS_NS}}}Relationship")}
    sheets = []
    for sheet in workbook.findall(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet"):
        target = targets[sheet.get(f"{{{OFFICE_RELS_NS}}}id")]
        sheets.append((sheet.get("name"), target.lstrip("/") if target.startswith("/") else f"xl/{target}"))
    return sheets


def _worksheet_rows(archive: zipfile.ZipFile, sheet_path: str, shared: List[str]) -> List[List[Optional[str]]]:
    """One worksheet as rows of cell text, one entry per column.

    Handles workbooks written by this tool and the same workbook after Excel saved it.
    """
    rows = []
    sheet = _read_xml(archive, sheet_path)
    cell_tag, value_tag, inline_tag = f"{{{MAIN_NS}}}c", f"{{{MAIN_NS}}}v", f"{{{MAIN_NS}}}is"
    # Column numbers by letters, worked out once per sheet: a sheet has few distinct
    # columns but may have millions of cells.
    column_of: Dict[str, int] = {}
    for row in sheet.findall(f"{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row"):
        cells: Dict[int, str] = {}
        column = 0
        for cell in row.findall(cell_tag):
            reference = cell.get("r")
            # Excel may leave out empty cells, so place each by its reference when present.
            if reference:
                letters = reference.rstrip("0123456789")
                if letters not in column_of:
                    column_of[letters] = column_index(letters)
                column = column_of[letters]
            else:
                column += 1
            value = cell.find(value_tag)
            kind = cell.get("t")
            if kind == "s":
                cells[column] = shared[int(value.text)]
            elif kind == "inlineStr":
                cells[column] = _cell_text(cell.find(inline_tag))
            else:
                cells[column] = (value.text or "") if value is not None else ""
        width = max(cells, default=0)
        rows.append([cells.get(c) for c in range(1, width + 1)])
    return rows


def _report_rows(rows: List[List[Optional[str]]], columns: Columns, error: str) -> List[Tuple[List[str], List[str]]]:
    """Find a sheet's table header row (row 1; lower in reports that had the rules above
    the table) and return
    the data rows under it as (fixed values, folders) pairs."""
    expected = [header for header, _ in columns]

    def is_header(row: List[Optional[str]]) -> bool:
        return len(row) >= len(expected) and all(
            (cell or "").casefold() == name.casefold() for cell, name in zip(row, expected)
        )

    header_at = next((index for index, row in enumerate(rows) if is_header(row)), None)
    if header_at is None:
        raise ValueError(f"{error} '{', '.join(expected)}'.")
    result = []
    for row in rows[header_at + 1:]:
        if len(row) < len(expected) or not row[0]:
            continue
        result.append((row, [folder for folder in row[len(expected):] if folder]))
    return result


def _number(text: Optional[str]) -> int:
    # Rounds half to even, like PowerShell's [long][double].
    return int(round(float(text)))


@dataclass
class DuplicateWorkbook:
    """Both sheets of a report; ``folders`` is None when the report has no folder sheet."""

    files: List[DuplicateSet]
    folders: Optional[List[DuplicateFolderSet]]


def read_duplicate_workbook(path: str) -> DuplicateWorkbook:
    """Read both sheets of a report written by export_duplicate_report."""
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Report '{path}' was not found.")
    try:
        with zipfile.ZipFile(path) as archive:
            sheets = _workbook_sheets(archive)
            shared_xml = _read_xml(archive, "xl/sharedStrings.xml")
            shared = [] if shared_xml is None else [_cell_text(si) for si in shared_xml.findall(f"{{{MAIN_NS}}}si")]
            file_rows = _worksheet_rows(archive, sheets[0][1], shared)
            folder_sheet = next((p for name, p in sheets if name.casefold() == FOLDER_SHEET_NAME.casefold()), None)
            folder_rows = _worksheet_rows(archive, folder_sheet, shared) if folder_sheet else None

        files = [
            DuplicateSet(
                file_name=row[0],
                last_write_time=from_excel_serial(float(row[1])),
                size_bytes=_number(row[2]),
                md5=row[3],
                count=len(folders),
                folders=folders,
            )
            for row, folders in _report_rows(
                file_rows, FIXED_COLUMNS, f"'{path}' is not a duplicates report: it has no header row"
            )
        ]
        duplicate_folders = None
        if folder_rows is not None:
            duplicate_folders = [
                DuplicateFolderSet(
                    folder_name=row[0],
                    file_count=_number(row[1]),
                    folder_count=_number(row[2]),
                    size_bytes=_number(row[3]),
                    count=len(folders),
                    folders=folders,
                )
                for row, folders in _report_rows(
                    folder_rows,
                    FOLDER_COLUMNS,
                    f"'{path}' is not a duplicates report: sheet '{FOLDER_SHEET_NAME}' has no header row",
                )
            ]
    except zipfile.BadZipFile:
        raise ValueError(f"'{path}' is not an Excel workbook.") from None
    except (ElementTree.ParseError, KeyError, IndexError, AttributeError, TypeError) as exc:
        # A damaged or foreign workbook: report it plainly instead of with a traceback.
        raise ValueError(f"'{path}' could not be read as a duplicates report: {exc}") from None
    return DuplicateWorkbook(files, duplicate_folders)


def read_duplicate_report(path: str) -> List[DuplicateSet]:
    """Read the duplicate files back from a report written by export_duplicate_report."""
    return read_duplicate_workbook(path).files


def read_duplicate_folder_report(path: str) -> List[DuplicateFolderSet]:
    """Read the duplicate folders back from a report; empty when it has no folder sheet."""
    return read_duplicate_workbook(path).folders or []
