"""Re-check an existing report without rescanning.

Mirrors Update-DuplicateReport, Invoke-CopyCheck, Test-DuplicateCopy,
Test-DuplicateFolderCopy and Test-PathRootReachable in DuplicateFinder.psm1.
"""

from __future__ import annotations

import contextlib
import logging
import os
import re
import stat
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field, replace
from datetime import datetime, timedelta
from typing import Callable, Dict, Iterator, List, Optional, Sequence, Tuple, TypeVar

from .folders import DuplicateFolderSet
from . import scanner
from .matcher import DuplicateSet
from .names import name_key
from .scanner import FileRecord, FolderRecord, iter_files, local_time
from .xlsx import export_duplicate_report, read_duplicate_workbook

log = logging.getLogger("find_duplicates")

PRESENT = "Present"          # still there, unchanged
MISSING = "Missing"          # no longer there
CHANGED = "Changed"          # still there but changed
UNAVAILABLE = "Unavailable"  # its drive or network share cannot be reached (kept as is)

Row = TypeVar("Row", DuplicateSet, DuplicateFolderSet)
RootCache = Dict[str, bool]
_EPOCH = datetime(1970, 1, 1)

# Called with (folder or folder copy being checked, checked so far, total to check).
ProgressCallback = Callable[[str, int, int], None]


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
    folder: str,
    file_name: str,
    size_bytes: int,
    last_write_time: datetime,
    root_cache: Optional[RootCache] = None,
    utc_offset: Optional[timedelta] = None,
) -> str:
    """Check one recorded file copy without reading its contents.

    ``utc_offset`` is the report's offset for ``last_write_time``; without one (older
    reports) the saved date is compared as local time on this computer.

    PRESENT when it is still there with the same size and saved date (to the second);
    UNAVAILABLE (kept) when its drive or share cannot be reached, or it cannot be checked,
    as for names Windows reserves for devices.
    """
    try:
        if not root_reachable(folder, root_cache):
            return UNAVAILABLE
        path = _find_copy(folder, file_name)
    except OSError:
        return UNAVAILABLE  # something that cannot be checked is kept, never removed
    if path is None:
        return MISSING
    return _copy_state(lambda: os.stat(path), size_bytes, last_write_time, utc_offset)


def _find_copy(folder: str, file_name: str) -> Optional[str]:
    # A recorded copy's file: by its exact name, else ignoring case and Unicode form.
    path = os.path.join(folder, file_name)
    return path if os.path.isfile(path) else _find_by_name_key(folder, file_name)


def _copy_state(
    info_of: Callable[[], os.stat_result], size_bytes: int, last_write_time: datetime, utc_offset: Optional[timedelta]
) -> str:
    """PRESENT or CHANGED for a copy's file, by its size and saved date; MISSING when it was
    deleted while being checked; UNAVAILABLE when its details cannot be read."""
    try:
        info = info_of()
        if not stat.S_ISREG(info.st_mode):
            # Not a plain file: a name Windows reserves for a device (NUL, CON ...). It cannot
            # be checked, so it is kept, never removed.
            return UNAVAILABLE
        same_date = same_saved_date(info.st_mtime_ns, last_write_time, utc_offset)
    except FileNotFoundError:
        return MISSING
    except (OSError, OverflowError, ValueError):
        return UNAVAILABLE
    return PRESENT if info.st_size == size_bytes and same_date else CHANGED


CopyKey = Tuple[int, str]  # (row number, location)


@dataclass(frozen=True)
class _CopyCheck:
    key: CopyKey
    file_name: str
    size_bytes: int
    last_write_time: datetime
    utc_offset: Optional[timedelta]


def _check_copies_in_folder(folder: str, checks: Sequence[_CopyCheck]) -> List[Tuple[CopyKey, str]]:
    """Check the recorded copies in one folder (see check_copy), whose drive or share is
    reachable. Several copies are checked from a single listing of the folder rather than a
    lookup each: on a network share every lookup is a round trip."""
    if len(checks) == 1:
        check = checks[0]
        try:
            path = _find_copy(folder, check.file_name)
        except OSError:
            return [(check.key, UNAVAILABLE)]
        if path is None:
            return [(check.key, MISSING)]
        return [(check.key, _copy_state(lambda: os.stat(path), check.size_bytes, check.last_write_time, check.utc_offset))]

    exact: Dict[str, os.DirEntry] = {}
    failure: Optional[str] = None
    try:
        with os.scandir(folder) as entries:
            exact = {entry.name: entry for entry in entries if entry.is_file()}
    except (FileNotFoundError, NotADirectoryError):
        failure = MISSING
    except OSError:
        failure = UNAVAILABLE

    by_key: Optional[Dict[str, os.DirEntry]] = None  # built only when a name is not found as it is
    results = []
    for check in checks:
        state = failure
        if state is None:
            entry = exact.get(check.file_name)
            if entry is None:
                if by_key is None:
                    by_key = {}
                    for candidate in exact.values():
                        by_key.setdefault(name_key(candidate.name), candidate)
                entry = by_key.get(name_key(check.file_name))
            state = (
                MISSING if entry is None
                else _copy_state(entry.stat, check.size_bytes, check.last_write_time, check.utc_offset)
            )
        results.append((check.key, state))
    return results


Job = TypeVar("Job")
Result = TypeVar("Result")


def _run_all(
    jobs: Sequence[Job],
    work: Callable[[Job], Result],
    throttle_limit: int,
    on_progress: Optional[ProgressCallback],
    name_of: Callable[[Job], str],
) -> Iterator[Result]:
    """Run ``work`` on each job, reporting progress; yields each result (in the order they
    finish). Jobs on a network drive run ``throttle_limit`` at a time on threads; the others
    run here meanwhile, one at a time, which is faster for local folders (see
    scanner.on_network_drive)."""
    threaded = [job for job in jobs if scanner.on_network_drive(name_of(job))] if throttle_limit > 1 else []
    in_thread = {id(job) for job in threaded}
    done = 0

    def progress(job: Job) -> None:
        nonlocal done
        done += 1
        if on_progress is not None:
            on_progress(name_of(job), done, len(jobs))

    if not threaded:
        for job in jobs:
            progress(job)
            yield work(job)
        return
    with ThreadPoolExecutor(max_workers=min(throttle_limit, len(threaded))) as pool:
        futures = {pool.submit(work, job): job for job in threaded}
        for job in jobs:
            if id(job) not in in_thread:
                progress(job)
                yield work(job)
        for future in as_completed(futures):
            progress(futures[future])
            yield future.result()


def _file_copy_states(
    rows: Sequence[DuplicateSet], throttle_limit: int, root_cache: RootCache, on_progress: Optional[ProgressCallback]
) -> Dict[CopyKey, str]:
    """Check every copy of every file row; returns (row number, folder) -> state. Copies on a
    drive or share that cannot be reached are UNAVAILABLE without further checks (each drive
    or share is tried once). The others are grouped by folder and each folder checked with
    _check_copies_in_folder, ``throttle_limit`` folders at a time."""
    states: Dict[CopyKey, str] = {}
    by_folder: Dict[str, List[_CopyCheck]] = {}
    for number, row in enumerate(rows):
        for folder in row.folders:
            key = (number, folder)
            if not root_reachable(folder, root_cache):
                states[key] = UNAVAILABLE
                continue
            by_folder.setdefault(folder, []).append(
                _CopyCheck(key, row.file_name, row.size_bytes, row.last_write_time, row.utc_offset)
            )
    jobs = list(by_folder.items())
    for results in _run_all(jobs, lambda job: _check_copies_in_folder(*job), throttle_limit, on_progress, lambda job: job[0]):
        states.update(results)
    return states


def _folder_copy_states(
    rows: Sequence[DuplicateFolderSet],
    throttle_limit: int,
    root_cache: RootCache,
    on_progress: Optional[ProgressCallback],
    exclude_names: Sequence[str] = (),
) -> Dict[CopyKey, str]:
    """Check every copy of every folder row with check_folder_copy, ``throttle_limit`` at a
    time; returns (row number, folder) -> state. Copies on a drive or share that cannot be
    reached are UNAVAILABLE without further checks."""
    states: Dict[CopyKey, str] = {}
    jobs = []
    for number, row in enumerate(rows):
        for folder in row.folders:
            if root_reachable(folder, root_cache):
                jobs.append(((number, folder), row))
            else:
                states[(number, folder)] = UNAVAILABLE

    def work(job: Tuple[CopyKey, DuplicateFolderSet]) -> Tuple[CopyKey, str]:
        key, row = job
        return key, check_folder_copy(key[1], row.file_count, row.folder_count, row.size_bytes, exclude_names=exclude_names)

    states.update(_run_all(jobs, work, throttle_limit, on_progress, lambda job: job[0][1]))
    return states


def same_saved_date(mtime_ns: int, last_write_time: datetime, utc_offset: Optional[timedelta]) -> bool:
    """True when a file's saved time is a report's, to the whole second: compared as instants
    when the report has the UTC offset (correct whatever this computer's time zone), as local
    times otherwise (reports made before the UTC Offset column), as PowerShell does."""
    seconds = mtime_ns // 1_000_000_000
    if utc_offset is None:
        # Also correct in the repeated hour when daylight saving time ends (naive comparisons
        # ignore "fold").
        return local_time(seconds) == last_write_time.replace(microsecond=0)
    return seconds == (last_write_time - utc_offset - _EPOCH) // timedelta(seconds=1)


_MD5 = re.compile("[0-9A-Fa-f]{32}")


def previous_md5(report: str, files: Sequence[FileRecord]) -> Dict[str, str]:
    """MD5 hashes to take from an earlier report instead of reading the files again.

    Returns full path -> MD5 for each scanned file that the report lists (same folder and
    name, ignoring case) whose size and saved date are still the report's. Only those
    files' details are looked up. MD5s that are not 32 hexadecimal digits (an edited
    report) are ignored.
    """
    rows = read_duplicate_workbook(report).files
    wanted = {name_key(row.file_name) for row in rows}
    scanned: Dict[Tuple[str, str], FileRecord] = {}
    for record in files:
        key = name_key(record.name)
        if key in wanted:
            scanned.setdefault((record.folder, key), record)

    previous: Dict[str, str] = {}
    for row in rows:
        if not _MD5.fullmatch(row.md5 or ""):
            continue
        name = name_key(row.file_name)
        for folder in row.folders:
            record = scanned.get((folder, name))
            if record is None:
                continue
            try:
                same = record.size == row.size_bytes and same_saved_date(record.mtime_ns, row.last_write_time, row.utc_offset)
            except OSError:
                continue  # hashed (and reported on) as usual
            if same:
                previous[record.path] = row.md5.upper()
    return previous


_silence = threading.local()


@contextlib.contextmanager
def _silenced() -> Iterator[None]:
    """Drop this thread's log messages (other threads, checking in parallel, keep theirs)."""
    _silence.active = True
    try:
        yield
    finally:
        _silence.active = False


log.addFilter(lambda record: not getattr(_silence, "active", False))


def check_folder_copy(
    path: str,
    file_count: int,
    folder_count: int,
    size_bytes: int,
    root_cache: Optional[RootCache] = None,
    exclude_names: Sequence[str] = (),
) -> str:
    """Check one recorded folder copy by listing its tree again (no file contents are read).

    PRESENT when it still has the same number of files and sub folders and the same
    total size; UNAVAILABLE when its drive or share, or part of the tree, cannot be read.
    Files and folders whose names match ``exclude_names`` (the scan's --exclude) are left out.
    """
    try:
        if not root_reachable(path, root_cache):
            return UNAVAILABLE
        if not os.path.isdir(path):
            return MISSING
        folders: List[FolderRecord] = []
        with _silenced():  # the listing's own warnings are summed up as UNAVAILABLE below
            files = list(iter_files(path, folders=folders, exclude_names=exclude_names))
        if not all(f.readable for f in folders):
            return UNAVAILABLE
        total_size = sum(f.size for f in files)
    except OSError:
        return UNAVAILABLE

    same = len(files) == file_count and len(folders) - 1 == folder_count and total_size == size_bytes
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
    states: Dict[CopyKey, str],
    name_of: Callable[[Row], str],
    noun: str,
    location_is_item: bool = False,
) -> _CheckOutcome:
    """Given every copy's state, keep copies that are PRESENT or UNAVAILABLE, drop MISSING
    and CHANGED ones, and drop rows left with fewer than two copies."""
    outcome = _CheckOutcome(kept=[])
    for number, row in enumerate(rows):
        name = name_of(row)
        present = []
        for location in row.folders:
            outcome.checked += 1
            state = states[(number, location)]
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


def validate_report(
    path: str,
    dry_run: bool = False,
    on_progress: Optional[ProgressCallback] = None,
    throttle_limit: int = 1,
) -> ValidationResult:
    """Re-check every copy listed in an existing report and remove the ones that no
    longer exist, without rescanning.

    Each file copy is checked with a file lookup (copies in the same folder from one
    listing of it), and each folder copy by listing its tree again; no contents are read
    or downloaded. ``throttle_limit`` checks that many folders at the same time. Copies
    that are missing or changed are removed; rows left with fewer than two copies are
    removed. Copies on a drive or network share that cannot be reached are kept. The
    report is rewritten in place only when something changed, and never with ``dry_run``.
    """
    path = os.path.abspath(path)
    workbook = read_duplicate_workbook(path)
    root_cache: RootCache = {}

    file_states = _file_copy_states(workbook.files, throttle_limit, root_cache, on_progress)
    files = _check_rows(workbook.files, file_states, lambda row: row.file_name, noun="")
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
        folder_states = _folder_copy_states(
            workbook.folders, throttle_limit, root_cache, on_progress, workbook.settings.exclude_names
        )
        folders = _check_rows(
            workbook.folders, folder_states, lambda row: row.folder_name, noun="folder ", location_is_item=True
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
        export_duplicate_report(result.duplicates, path, result.duplicate_folders, workbook.settings)
        result.saved = True
    return result
