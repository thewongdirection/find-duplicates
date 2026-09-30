"""Re-check an existing report without rescanning.

Mirrors Update-DuplicateReport and Test-DuplicateCopy in src/DuplicateFinder.psm1.
"""

from __future__ import annotations

import logging
import os
import stat
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, List, Optional

from .matcher import DuplicateSet, ordinal_ignore_case
from .xlsx import export_duplicate_report, read_duplicate_report

log = logging.getLogger("find_duplicates")

PRESENT = "Present"          # still there with the same size and saved date
MISSING = "Missing"          # no longer there
CHANGED = "Changed"          # still there but its size or saved date changed
UNAVAILABLE = "Unavailable"  # its drive or network share cannot be reached (kept as is)

# Called with (file name, rows checked so far, total rows).
RowCallback = Callable[[str, int, int], None]


@dataclass
class ValidationResult:
    """What validate_report found and did."""

    path: str
    rows_checked: int = 0
    copies_checked: int = 0
    copies_removed: int = 0
    copies_unavailable: int = 0
    rows_removed: int = 0
    rows_remaining: int = 0
    saved: bool = False
    duplicates: List[DuplicateSet] = field(default_factory=list)


def _find_ignoring_case(folder: str, file_name: str) -> Optional[str]:
    # Only needed on case-sensitive file systems (Linux, some macOS volumes).
    wanted = ordinal_ignore_case(file_name)
    try:
        with os.scandir(folder) as entries:
            for entry in entries:
                if entry.is_file() and ordinal_ignore_case(entry.name) == wanted:
                    return entry.path
    except (FileNotFoundError, NotADirectoryError):
        pass
    return None


def _path_root(folder: str) -> str:
    """The drive or network share a folder lives on ('' on POSIX, where '/' always exists)."""
    drive = os.path.splitdrive(folder)[0]
    return drive + os.sep if drive and not drive.startswith(("\\\\", "//")) else drive


def check_copy(folder: str, file_name: str, size_bytes: int, last_write_time: datetime) -> str:
    """Check one recorded copy without reading its contents; returns PRESENT, MISSING, CHANGED or UNAVAILABLE."""
    try:
        path: Optional[str] = os.path.join(folder, file_name)
        if not os.path.isfile(path):
            path = _find_ignoring_case(folder, file_name)
        if path is None:
            root = _path_root(folder)
            if root and not os.path.isdir(root):
                return UNAVAILABLE
            return MISSING
        info = os.stat(path)
    except FileNotFoundError:
        return MISSING  # deleted while being checked
    except OSError:
        return UNAVAILABLE

    if not stat.S_ISREG(info.st_mode):
        return MISSING
    # Whole seconds, as when scanning.
    same_second = info.st_mtime_ns // 1_000_000_000 == int(last_write_time.timestamp() // 1)
    return PRESENT if info.st_size == size_bytes and same_second else CHANGED


def validate_report(path: str, dry_run: bool = False, on_row: Optional[RowCallback] = None) -> ValidationResult:
    """Re-check every copy listed in an existing report and remove the ones that no
    longer exist, without rescanning the folders.

    Each copy is checked with a single file lookup (no contents are read or
    downloaded). Copies that are missing, or whose size or saved date changed,
    are removed; rows left with fewer than two copies are removed. Copies on a
    drive or network share that cannot be reached are kept. The report is
    rewritten in place only when something changed, and never with ``dry_run``.
    """
    path = os.path.abspath(path)
    duplicates = read_duplicate_report(path)
    result = ValidationResult(path=path, rows_checked=len(duplicates))

    for number, dup in enumerate(duplicates, start=1):
        if on_row is not None:
            on_row(dup.file_name, number, len(duplicates))
        present = []
        for folder in dup.folders:
            result.copies_checked += 1
            state = check_copy(folder, dup.file_name, dup.size_bytes, dup.last_write_time)
            if state == PRESENT:
                present.append(folder)
            elif state == UNAVAILABLE:
                result.copies_unavailable += 1
                present.append(folder)
                log.warning("Cannot reach '%s'; keeping its copy of '%s'.", folder, dup.file_name)
            else:
                result.copies_removed += 1
                log.info("%s: '%s'", state, os.path.join(folder, dup.file_name))

        if len(present) >= 2:
            result.duplicates.append(
                DuplicateSet(dup.file_name, dup.last_write_time, dup.size_bytes, dup.md5, len(present), present)
            )
        else:
            result.rows_removed += 1

    result.rows_remaining = len(result.duplicates)
    if (result.copies_removed or result.rows_removed) and not dry_run:
        export_duplicate_report(result.duplicates, path)
        result.saved = True
    return result
