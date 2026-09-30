"""Re-check an existing report without rescanning.

Mirrors Update-DuplicateReport, Invoke-CopyCheck, Test-DuplicateCopy,
Test-DuplicateFolderCopy and Test-PathRootReachable in src/DuplicateFinder.psm1.
"""

from __future__ import annotations

import logging
import os
import stat
from dataclasses import dataclass, field, replace
from datetime import datetime
from typing import Callable, Dict, List, Optional, Sequence, TypeVar

from .folders import DuplicateFolderSet
from .matcher import DuplicateSet
from .names import name_key
from .scanner import FolderRecord, iter_files, local_time
from .xlsx import export_duplicate_report, read_duplicate_workbook

log = logging.getLogger("find_duplicates")

PRESENT = "Present"          # still there, unchanged
MISSING = "Missing"          # no longer there
CHANGED = "Changed"          # still there but changed
UNAVAILABLE = "Unavailable"  # its drive or network share cannot be reached (kept as is)

Row = TypeVar("Row", DuplicateSet, DuplicateFolderSet)
RootCache = Dict[str, bool]

# Called with (file or folder name, rows checked so far, total rows).
RowCallback = Callable[[str, int, int], None]


@dataclass
class ValidationResult:
    """What validate_report found and did. The folder_* fields stay 0 and
    ``duplicate_folders`` None when the report has no folder sheet."""

    path: str
    rows_checked: int = 0
    copies_checked: int = 0
    copies_removed: int = 0
    copies_unavailable: int = 0
    rows_removed: int = 0
    rows_remaining: int = 0
    folder_rows_checked: int = 0
    folder_copies_checked: int = 0
    folder_copies_removed: int = 0
    folder_copies_unavailable: int = 0
    folder_rows_removed: int = 0
    folder_rows_remaining: int = 0
    saved: bool = False
    duplicates: List[DuplicateSet] = field(default_factory=list)
    duplicate_folders: Optional[List[DuplicateFolderSet]] = None


def _path_root(folder: str) -> str:
    """The drive or network share a folder lives on ('' on POSIX, where '/' always exists)."""
    drive = os.path.splitdrive(folder)[0]
    return drive + os.sep if drive and not drive.startswith(("\\\\", "//")) else drive


def root_reachable(folder: str, cache: Optional[RootCache] = None) -> bool:
    """True when the drive or network share holding ``folder`` can be reached.

    Each root is checked once per ``cache``: an offline share can take many seconds
    to time out, and a report may list thousands of copies on it.
    """
    root = _path_root(folder)
    if not root:
        return True
    key = name_key(root)
    if cache is not None and key in cache:
        return cache[key]
    reachable = os.path.isdir(root)
    if cache is not None:
        cache[key] = reachable
    return reachable


def _find_by_name_key(folder: str, file_name: str) -> Optional[str]:
    # The file whose name matches ignoring case and Unicode form. Used when the plain
    # lookup fails: case-sensitive file systems (Linux), and names stored in another
    # Unicode form (Windows and Linux keep both forms apart).
    wanted = name_key(file_name)
    try:
        with os.scandir(folder) as entries:
            for entry in entries:
                if entry.is_file() and name_key(entry.name) == wanted:
                    return entry.path
    except (FileNotFoundError, NotADirectoryError):
        pass
    return None


def check_copy(
    folder: str, file_name: str, size_bytes: int, last_write_time: datetime, root_cache: Optional[RootCache] = None
) -> str:
    """Check one recorded file copy without reading its contents.

    PRESENT when it is still there with the same size and saved date (to the second);
    UNAVAILABLE (kept) when its drive or share cannot be reached, or it cannot be checked,
    as for names Windows reserves for devices.
    """
    try:
        if not root_reachable(folder, root_cache):
            return UNAVAILABLE
        path: Optional[str] = os.path.join(folder, file_name)
        if not os.path.isfile(path):
            path = _find_by_name_key(folder, file_name)
        if path is None:
            return MISSING
        info = os.stat(path)
        if not stat.S_ISREG(info.st_mode):
            # Not a plain file: a name Windows reserves for a device (NUL, CON ...). It cannot
            # be checked, so it is kept, never removed.
            return UNAVAILABLE
        # Compare local wall-clock seconds, as PowerShell does; this is also correct in the
        # repeated hour when daylight saving time ends (naive comparisons ignore "fold").
        saved = local_time(info.st_mtime_ns // 1_000_000_000)
    except FileNotFoundError:
        return MISSING  # deleted while being checked
    except (OSError, OverflowError, ValueError):
        return UNAVAILABLE

    same = info.st_size == size_bytes and saved == last_write_time.replace(microsecond=0)
    return PRESENT if same else CHANGED


def check_folder_copy(
    path: str, file_count: int, folder_count: int, size_bytes: int, root_cache: Optional[RootCache] = None
) -> str:
    """Check one recorded folder copy by listing its tree again (no file contents are read).

    PRESENT when it still has the same number of files and sub folders and the same
    total size; UNAVAILABLE when its drive or share, or part of the tree, cannot be read.
    """
    try:
        if not root_reachable(path, root_cache):
            return UNAVAILABLE
        if not os.path.isdir(path):
            return MISSING
        folders: List[FolderRecord] = []
        previous = log.disabled
        log.disabled = True  # the listing's own warnings are summed up as UNAVAILABLE below
        try:
            files = list(iter_files(path, folders=folders))
        finally:
            log.disabled = previous
    except OSError:
        return UNAVAILABLE

    if not all(f.readable for f in folders):
        return UNAVAILABLE
    same = (
        len(files) == file_count
        and len(folders) - 1 == folder_count
        and sum(f.size for f in files) == size_bytes
    )
    return PRESENT if same else CHANGED


@dataclass
class _CheckOutcome:
    kept: list
    checked: int = 0
    removed: int = 0
    unavailable: int = 0
    rows_removed: int = 0


def _check_rows(
    rows: Sequence[Row],
    test_copy: Callable[[Row, str], str],
    name_of: Callable[[Row], str],
    noun: str,
    location_is_item: bool = False,
    on_row: Optional[RowCallback] = None,
) -> _CheckOutcome:
    """Keep copies that are PRESENT or UNAVAILABLE, drop MISSING and CHANGED ones, and
    drop rows left with fewer than two copies."""
    outcome = _CheckOutcome(kept=[])
    for number, row in enumerate(rows, start=1):
        name = name_of(row)
        if on_row is not None:
            on_row(name, number, len(rows))
        present = []
        for location in row.folders:
            outcome.checked += 1
            state = test_copy(row, location)
            if state == PRESENT:
                present.append(location)
            elif state == UNAVAILABLE:
                outcome.unavailable += 1
                present.append(location)
                log.warning("Cannot reach '%s'; keeping its copy of %s'%s'.", location, noun, name)
            else:
                outcome.removed += 1
                log.info("%s: '%s'", state, location if location_is_item else os.path.join(location, name))
        if len(present) >= 2:
            outcome.kept.append(replace(row, count=len(present), folders=present))
        else:
            outcome.rows_removed += 1
    return outcome


def validate_report(path: str, dry_run: bool = False, on_row: Optional[RowCallback] = None) -> ValidationResult:
    """Re-check every copy listed in an existing report and remove the ones that no
    longer exist, without rescanning.

    Each file copy is checked with a single file lookup, and each folder copy by
    listing its tree again; no contents are read or downloaded. Copies that are
    missing or changed are removed; rows left with fewer than two copies are removed.
    Copies on a drive or network share that cannot be reached are kept. The report is
    rewritten in place only when something changed, and never with ``dry_run``.
    """
    path = os.path.abspath(path)
    workbook = read_duplicate_workbook(path)
    root_cache: RootCache = {}

    files = _check_rows(
        workbook.files,
        lambda row, location: check_copy(location, row.file_name, row.size_bytes, row.last_write_time, root_cache),
        lambda row: row.file_name,
        noun="",
        on_row=on_row,
    )
    result = ValidationResult(
        path=path,
        rows_checked=len(workbook.files),
        copies_checked=files.checked,
        copies_removed=files.removed,
        copies_unavailable=files.unavailable,
        rows_removed=files.rows_removed,
        rows_remaining=len(files.kept),
        duplicates=files.kept,
    )
    changed = files.removed or files.rows_removed

    if workbook.folders is not None:
        folders = _check_rows(
            workbook.folders,
            lambda row, location: check_folder_copy(
                location, row.file_count, row.folder_count, row.size_bytes, root_cache
            ),
            lambda row: row.folder_name,
            noun="folder ",
            location_is_item=True,
            on_row=on_row,
        )
        result.folder_rows_checked = len(workbook.folders)
        result.folder_copies_checked = folders.checked
        result.folder_copies_removed = folders.removed
        result.folder_copies_unavailable = folders.unavailable
        result.folder_rows_removed = folders.rows_removed
        result.folder_rows_remaining = len(folders.kept)
        result.duplicate_folders = folders.kept
        changed = changed or folders.removed or folders.rows_removed

    if changed and not dry_run:
        export_duplicate_report(result.duplicates, path, result.duplicate_folders)
        result.saved = True
    return result
