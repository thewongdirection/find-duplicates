"""Shared test helpers."""

from __future__ import annotations

import os
import zipfile
from datetime import datetime, timezone
from typing import List
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


def read_worksheet(path: str, part: str = "xl/worksheets/sheet1.xml") -> List[List[str]]:
    """A worksheet written by this tool as a list of rows, each row a list of cell texts."""
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read(part))
    rows = []
    for row in root.findall("s:sheetData/s:row", NS):
        cells = []
        for cell in row.findall("s:c", NS):
            text = cell.find("s:is/s:t", NS)
            cells.append(text.text or "" if text is not None else cell.find("s:v", NS).text)
        rows.append(cells)
    return rows
