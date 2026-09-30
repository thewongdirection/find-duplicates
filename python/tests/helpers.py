"""Shared test helpers."""

from __future__ import annotations

import contextlib
import os
import sys
import time
import zipfile
from datetime import datetime, timezone
from typing import Iterator, List
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from find_duplicates import scanner  # noqa: E402
from xml.etree import ElementTree

SAVED = datetime(2024, 5, 17, 10, 30, 0, tzinfo=timezone.utc)
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def add_file(root: str, relative: str, content: str = "same content", saved: datetime = SAVED) -> str:
    """Create a file with the given content and saved date; returns its path."""
    path = os.path.join(root, *relative.split("/"))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as stream:
        stream.write(content)
    ns = int(saved.timestamp()) * 1_000_000_000 + saved.microsecond * 1000
    os.utime(path, ns=(ns, ns))
    return path


def as_if_on_a_network_drive():
    """Treat every path as on a network drive, where Python works on several threads (-j):
    locally it does not, being faster without (see scanner.on_network_drive)."""
    return mock.patch.object(scanner, "on_network_drive", return_value=True)


@contextlib.contextmanager
def time_zone(name: str) -> Iterator[None]:
    """Run with the TZ variable set (Linux and macOS only)."""
    previous = os.environ.get("TZ")
    os.environ["TZ"] = name
    time.tzset()
    try:
        yield
    finally:
        if previous is None:
            del os.environ["TZ"]
        else:
            os.environ["TZ"] = previous
        time.tzset()


TABLE_HEADERS = ("File Name", "Folder Name")
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def sheet_names(path: str) -> List[str]:
    """The workbook's sheet names, in order."""
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    return [sheet.get("name") for sheet in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")]


def read_worksheet(path: str, part: str = "xl/worksheets/sheet1.xml", all_rows: bool = False) -> List[List[str]]:
    """A worksheet as a list of rows, each row a list of cell texts: from the table's
    header row down, or with ``all_rows`` every row (for the Rules sheet, or reports
    that had text above the table)."""
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read(part))
    rows = []
    for row in root.findall("s:sheetData/s:row", NS):
        cells = []
        for cell in row.findall("s:c", NS):
            text = cell.find("s:is/s:t", NS)
            cells.append((text.text or "") if text is not None else cell.find("s:v", NS).text)
        rows.append(cells)
    if all_rows:
        return rows
    start = next(i for i, row in enumerate(rows) if row and row[0] in TABLE_HEADERS)
    return rows[start:]
