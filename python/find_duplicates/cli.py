"""Command line entry point. Mirrors Find-Duplicates.ps1."""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import sys
import time
from typing import Dict, List, Optional, TextIO

from .folders import find_duplicate_folders
from .matcher import MAX_THROTTLE_LIMIT, MIN_THROTTLE_LIMIT, find_duplicate_files
from .scanner import FolderRecord, full_path, iter_files
from .validate import previous_md5, validate_report
from .xlsx import export_duplicate_report

log = logging.getLogger("find_duplicates")

DEFAULT_OUTPUT = "duplicates.xlsx"
PROGRESS_INTERVAL_SECONDS = 0.25

DESCRIPTION = """\
Find duplicate files in a folder and all of its sub folders and save them to
Excel, or re-check an existing report without rescanning.

SCAN (default): two files are duplicates only when ALL three of these match:
file name (case-insensitive), saved date (last modified time, to the whole
second) and MD5 hash of the contents. MD5 is only calculated for files whose
name and saved date already match another file (and whose size matches too),
so most files are never read. When the report already exists, unchanged files
keep the MD5 recorded there. Files of 0 bytes are included unless
--ignore-empty-files is used. A "Rules" sheet in the report states the
matching rules in plain words.

DUPLICATE FOLDERS (--folders): also finds folders with the same name and
exactly the same contents: the same tree of file and sub folder names, where
every file is a duplicate (name, saved date, MD5) of the file at the same place
in the other folder. Only the top-most duplicate folders are reported, on a
second sheet, "Duplicate Folders".

VALIDATE (--validate): re-check every copy listed in an existing report with a
quick file lookup (no contents are read) and remove copies that no longer exist
or whose size or saved date changed. Rows left with fewer than two copies are
removed. The report is updated in place. Copies on a drive or network share
that cannot be reached are kept. Duplicate folders, when the report has them,
are re-listed and must still have the same number of files and sub folders and
total size.

Local folders, network shares (\\\\server\\share or mapped drives) and synced
cloud folders (OneDrive, Google Drive, Dropbox ...) are all supported.
Microsoft Excel does NOT need to be installed.
"""

EPILOG = """\
examples:
  python -m find_duplicates D:\\Photos
  python -m find_duplicates \\\\server\\share C:\\Reports\\share-dupes.xlsx -j 8
  python -m find_duplicates "%OneDrive%" --skip-cloud-only
  python -m find_duplicates D:\\Backups --folders
  python -m find_duplicates --validate C:\\Reports\\share-dupes.xlsx --dry-run
"""


class ProgressLine:
    """A single, throttled, self-overwriting status line (only on a terminal)."""

    def __init__(self, stream: TextIO = sys.stderr) -> None:
        self._stream = stream
        self._enabled = stream.isatty()
        self._last_shown = 0.0
        self._shown = False

    def show(self, text: str, force: bool = False) -> None:
        now = time.monotonic()
        if not self._enabled or (not force and now - self._last_shown < PROGRESS_INTERVAL_SECONDS):
            return
        self._last_shown = now
        width = max(shutil.get_terminal_size().columns - 1, 20)
        if len(text) > width:
            text = "..." + text[-(width - 3):]
        self._stream.write("\r" + text.ljust(width))
        self._stream.flush()
        self._shown = True

    def clear(self) -> None:
        if self._shown:
            self._stream.write("\r" + " " * (shutil.get_terminal_size().columns - 1) + "\r")
            self._stream.flush()
            self._shown = False


def _throttle_limit(text: str) -> int:
    try:
        value = int(text)
    except ValueError:
        value = 0
    if not MIN_THROTTLE_LIMIT <= value <= MAX_THROTTLE_LIMIT:
        raise argparse.ArgumentTypeError(f"must be a whole number from {MIN_THROTTLE_LIMIT} to {MAX_THROTTLE_LIMIT}")
    return value


def _parse_args(argv: Optional[List[str]]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="find-duplicates",
        description=DESCRIPTION,
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "path",
        nargs="?",
        help="Folder to scan (default: the current folder). With --validate: the report to check.",
    )
    parser.add_argument(
        "output",
        nargs="?",
        help=f'Report to write. Defaults to "{DEFAULT_OUTPUT}" in the current folder. '
        '".xlsx" is appended when no extension is given.',
    )
    parser.add_argument("-o", "--output-file", dest="output_option", metavar="FILE", help="Same as the OUTPUT argument.")
    parser.add_argument(
        "--skip-cloud-only",
        action="store_true",
        help="Never download online-only cloud files to hash them. "
        "Duplicates among such files are then not reported.",
    )
    parser.add_argument(
        "-j",
        "--throttle-limit",
        type=_throttle_limit,
        default=1,
        metavar="N",
        help="How many files to hash, and folders to list or check, at the same time (1-64, default 1). Try "
        "4-8 for SSDs, network shares and cloud folders; keep 1 for a single spinning hard disk. Also speeds up "
        "--validate.",
    )
    parser.add_argument(
        "--ignore-empty-files",
        action="store_true",
        help="Leave files of 0 bytes out of the duplicate files (they all have the same contents). "
        "Duplicate folders still compare every file.",
    )
    parser.add_argument(
        "--folders", action="store_true", help='Also find duplicate folders and save them on the "Duplicate Folders" sheet.'
    )
    parser.add_argument(
        "--rehash",
        action="store_true",
        help="Read every candidate file again. Without it, when the report already exists (from an earlier "
        "scan), files it lists whose size and saved date have not changed keep the MD5 hash recorded there.",
    )
    parser.add_argument("--validate", action="store_true", help="Re-check an existing report instead of scanning.")
    parser.add_argument(
        "--dry-run", action="store_true", help="Show what would be saved or removed without changing the report."
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Print every folder (or removed copy) as it goes.")
    args = parser.parse_args(argv)

    if args.validate:
        if args.skip_cloud_only or args.folders or args.ignore_empty_files or args.rehash:
            parser.error(
                "--skip-cloud-only, --folders, --ignore-empty-files and --rehash only apply to a scan, not to --validate"
            )
        if len([value for value in (args.path, args.output, args.output_option) if value]) > 1:
            parser.error("--validate takes a single report")
    return args


def report_path(output: Optional[str]) -> str:
    """The absolute report path, adding .xlsx when no extension was given."""
    output = output or DEFAULT_OUTPUT
    if not os.path.splitext(output)[1]:
        output += ".xlsx"
    return os.path.abspath(output)


def _validate(report: str, dry_run: bool, throttle_limit: int = 1) -> int:
    progress = ProgressLine()
    print(f"Validating '{report}' ...")
    try:
        result = validate_report(
            report, dry_run=dry_run, throttle_limit=throttle_limit,
            on_progress=lambda name, done, total: progress.show(f"Checked {done} of {total}  {name}"),
        )
    except (OSError, ValueError) as exc:
        progress.clear()
        print(f"error: {exc}", file=sys.stderr)
        return 1
    progress.clear()

    print(f"Checked {result.copies_checked} copies in {result.rows_checked} rows: "
          f"{result.copies_removed} missing or changed, {result.copies_unavailable} unreachable (kept).")
    print(f"Removed {result.rows_removed} rows that are no longer duplicates; {result.rows_remaining} remain.")
    if result.duplicate_folders is not None:
        print(f"Checked {result.folder_copies_checked} folder copies in {result.folder_rows_checked} rows: "
              f"{result.folder_copies_removed} missing or changed, {result.folder_copies_unavailable} unreachable (kept).")
        print(f"Removed {result.folder_rows_removed} folder rows that are no longer duplicates; "
              f"{result.folder_rows_remaining} remain.")
    if result.saved:
        print(f"Report updated: '{report}'.")
    elif result.copies_removed or result.rows_removed or result.folder_copies_removed or result.folder_rows_removed:
        print("Report not changed (--dry-run).")
    else:
        print("Report is up to date; nothing to change.")
    return 0


def _scan(
    folder: str,
    report: str,
    skip_cloud_only: bool,
    throttle_limit: int,
    include_folders: bool,
    ignore_empty_files: bool,
    dry_run: bool,
    rehash: bool = False,
) -> int:
    scan_root = full_path(folder)
    if not os.path.isdir(scan_root):
        print(f"error: '{folder}' is not a folder.", file=sys.stderr)
        return 1

    progress = ProgressLine()
    print(f"Scanning '{scan_root}' ...")
    folder_records: Optional[List[FolderRecord]] = [] if include_folders else None
    files = list(
        iter_files(
            scan_root,
            exclude=[report],
            on_folder=lambda path, folders, found: progress.show(f"Folders: {folders}  Files: {found}  {path}"),
            folders=folder_records,
            throttle_limit=throttle_limit,
        )
    )
    progress.clear()
    print(f"Found {len(files)} files. Checking for duplicates ...")

    # Hashes are shared so that folder matching never reads a file twice.
    md5_cache: Dict[str, str] = {}

    # An earlier report's hashes are reused for files that have not changed since.
    if not rehash and os.path.isfile(report):
        try:
            previous = previous_md5(report, files)
        except (OSError, ValueError) as exc:
            log.warning("Not reusing MD5 hashes from '%s': %s", report, exc)
        else:
            md5_cache.update(previous)
            if previous:
                print(f"Reusing {len(previous)} MD5 hashes from the previous report (--rehash to read every file again).")
    options = dict(
        skip_cloud_only=skip_cloud_only,
        throttle_limit=throttle_limit,
        md5_cache=md5_cache,
        on_hash=lambda path, done, total: progress.show(f"Comparing MD5 {done}/{total}  {path}"),
    )
    duplicates = find_duplicate_files(files, ignore_empty_files=ignore_empty_files, **options)
    progress.clear()
    copies = sum(dup.count for dup in duplicates)
    print(f"Found {len(duplicates)} duplicated files ({copies} copies in total).")

    duplicate_folders = None
    if folder_records is not None:
        print("Checking for duplicate folders ...")
        duplicate_folders = find_duplicate_folders(files, folder_records, **options)
        progress.clear()
        folder_copies = sum(dup.count for dup in duplicate_folders)
        print(f"Found {len(duplicate_folders)} duplicated folders ({folder_copies} copies in total).")

    if dry_run:
        print(f"Report not saved (--dry-run): '{report}'.")
    else:
        export_duplicate_report(duplicates, report, duplicate_folders)
        print(f"Report saved to '{report}'.")
    return 0


def _safe_console() -> None:
    # A console or redirected output that cannot show every character (e.g. a Windows
    # code page) would otherwise crash on the first Unicode file name; escape instead.
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(errors="backslashreplace")


def main(argv: Optional[List[str]] = None) -> int:
    _safe_console()
    args = _parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s: %(message)s",
        stream=sys.stderr,
        force=True,  # each call gets its own level and stream, even when run repeatedly in one process
    )

    if args.validate:
        return _validate(report_path(args.output_option or args.path or args.output), args.dry_run, args.throttle_limit)
    return _scan(
        args.path or ".",
        report_path(args.output_option or args.output),
        args.skip_cloud_only,
        args.throttle_limit,
        args.folders,
        args.ignore_empty_files,
        args.dry_run,
        args.rehash,
    )
