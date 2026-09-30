"""Command line entry point. Mirrors Find-Duplicates.ps1."""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import sys
import time
from typing import List, Optional, TextIO

from .matcher import find_duplicate_files
from .scanner import iter_files
from .xlsx import export_duplicate_report

DEFAULT_OUTPUT = "duplicates.xlsx"
PROGRESS_INTERVAL_SECONDS = 0.25

DESCRIPTION = """\
Find duplicate files in a folder and all of its sub folders and save them to Excel.

Two files are duplicates only when ALL three of these match: file name
(case-insensitive), saved date (last modified time, to the whole second) and
MD5 hash of the contents. MD5 is only calculated for files whose name and saved
date already match another file (and whose size matches too), so most files
are never read.

Local folders, network shares (\\\\server\\share or mapped drives) and synced
cloud folders (OneDrive, Google Drive, Dropbox ...) are all supported. Cloud
files that are only stored online are downloaded when they have to be hashed,
unless --skip-cloud-only is used. Microsoft Excel does NOT need to be installed.
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


def _parse_args(argv: Optional[List[str]]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="find-duplicates",
        description=DESCRIPTION,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("path", nargs="?", default=".", help="Folder to scan. Defaults to the current folder.")
    parser.add_argument(
        "output",
        nargs="?",
        help=f'Workbook to write. Defaults to "{DEFAULT_OUTPUT}" in the current folder. '
        '".xlsx" is appended when no extension is given.',
    )
    parser.add_argument("-o", "--output-file", dest="output_option", help="Same as the OUTPUT argument.")
    parser.add_argument(
        "--skip-cloud-only",
        action="store_true",
        help="Never download online-only cloud files to hash them. "
        "Duplicates among such files are then not reported.",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Print every folder as it is scanned.")
    return parser.parse_args(argv)


def report_path(output: Optional[str]) -> str:
    """The absolute report path, adding .xlsx when no extension was given."""
    output = output or DEFAULT_OUTPUT
    if not os.path.splitext(output)[1]:
        output += ".xlsx"
    return os.path.abspath(output)


def main(argv: Optional[List[str]] = None) -> int:
    args = _parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s: %(message)s",
        stream=sys.stderr,
    )

    output = report_path(args.output_option or args.output)
    scan_root = os.path.abspath(args.path)
    if not os.path.isdir(scan_root):
        print(f"error: '{args.path}' is not a folder.", file=sys.stderr)
        return 1

    progress = ProgressLine()
    print(f"Scanning '{scan_root}' ...")
    files = list(
        iter_files(
            scan_root,
            exclude=[output],
            on_folder=lambda folder, folders, found: progress.show(
                f"Folders: {folders}  Files: {found}  {folder}"
            ),
        )
    )
    progress.clear()
    print(f"Found {len(files)} files. Checking for duplicates ...")

    duplicates = find_duplicate_files(
        files,
        skip_cloud_only=args.skip_cloud_only,
        on_hash=lambda path, done, total: progress.show(f"Comparing MD5 {done}/{total}  {path}"),
    )
    progress.clear()
    export_duplicate_report(duplicates, output)

    copies = sum(dup.count for dup in duplicates)
    print(f"Found {len(duplicates)} duplicated files ({copies} copies in total).")
    print(f"Report saved to '{output}'.")
    return 0
