"""Recursive folder scanning.

Mirrors Get-FileInventory, Test-FolderLink and Test-CloudOnlyFile in
src/DuplicateFinder.psm1.
"""

from __future__ import annotations

import logging
import os
import stat
import sys
from dataclasses import dataclass
from typing import Callable, Iterable, Iterator, List, Optional

from .names import sort_key

log = logging.getLogger("find_duplicates")

# Windows attributes of cloud placeholders (OneDrive "Files On-Demand" and other
# Cloud Files providers) whose contents are not stored locally. Reading them downloads them.
FILE_ATTRIBUTE_OFFLINE = 0x1000
FILE_ATTRIBUTE_RECALL_ON_OPEN = 0x40000
FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x400000
CLOUD_ONLY_ATTRIBUTES = (
    FILE_ATTRIBUTE_OFFLINE | FILE_ATTRIBUTE_RECALL_ON_OPEN | FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS
)

IO_REPARSE_TAG_MOUNT_POINT = 0xA0000003  # a junction

# Called with (folder path, folders scanned so far, files found so far).
FolderCallback = Callable[[str, int, int], None]


@dataclass(frozen=True)
class FileRecord:
    """The details of one file that duplicate matching needs."""

    path: str
    name: str
    folder: str
    size: int
    mtime_ns: int
    attributes: int = 0  # Windows file attributes; 0 elsewhere


@dataclass(frozen=True)
class FolderRecord:
    """A folder the scan listed, and whether its contents could be read."""

    path: str
    readable: bool


def is_folder_link(entry: os.DirEntry) -> bool:
    """True for symbolic links and junctions, which could loop back on the tree.

    Other reparse points are ordinary folders to the user: OneDrive and other
    cloud-synced folders, deduplicated volumes, DFS links. Those are scanned.
    """
    if entry.is_symlink():
        return True
    is_junction = getattr(entry, "is_junction", None)  # Python 3.12+
    if is_junction is not None:
        return bool(is_junction())
    if sys.platform == "win32":
        tag = getattr(entry.stat(follow_symlinks=False), "st_reparse_tag", 0)
        return tag == IO_REPARSE_TAG_MOUNT_POINT
    return False


def is_cloud_only(record: FileRecord) -> bool:
    """True for cloud placeholders whose contents would have to be downloaded to hash them."""
    return bool(record.attributes & CLOUD_ONLY_ATTRIBUTES)


def full_path(path: str) -> str:
    """Absolute path, with Windows short (8.3) names such as RUNNER~1 expanded.

    Matches the paths PowerShell reports. Links are deliberately not resolved.
    """
    path = os.path.abspath(path)
    if sys.platform == "win32":
        import ctypes

        buffer = ctypes.create_unicode_buffer(32_768)
        length = ctypes.windll.kernel32.GetLongPathNameW(path, buffer, len(buffer))
        if 0 < length < len(buffer):
            return buffer.value
    return path


def _same_path_key(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def iter_files(
    root: str,
    exclude: Iterable[str] = (),
    on_folder: Optional[FolderCallback] = None,
    folders: Optional[List[FolderRecord]] = None,
) -> Iterator[FileRecord]:
    """Yield every file below ``root``, recursing into sub folders.

    Walks the tree iteratively so deep trees cannot overflow the stack. Folders
    that cannot be read (permissions, dropped network connection) are logged as
    warnings and skipped. Symbolic links and junctions are not followed, which
    prevents infinite loops; cloud-synced folders are. Listing folders never
    downloads cloud files. When ``folders`` is given, it receives a record for
    every folder listed (needed to compare folder trees, including empty and
    unreadable folders). A folder that holds an excluded file, and each folder
    link that is not followed, is recorded as not readable: its contents are not
    fully known, so it can never be proven identical to another folder.
    """
    if not os.path.isdir(root):
        raise NotADirectoryError(f"'{root}' is not a folder.")

    # Normalised like the scanned paths (long names), or a short-form path would never match.
    excluded = {_same_path_key(full_path(p)) for p in exclude}
    pending = [full_path(root)]
    folder_count = 0
    file_count = 0

    while pending:
        folder = pending.pop()
        folder_count += 1
        log.info("Scanning %s", folder)
        if on_folder is not None:
            on_folder(folder, folder_count, file_count)

        try:
            with os.scandir(folder) as it:
                # File systems list entries in different orders (alphabetical on NTFS,
                # arbitrary on ext4); ordinal order matches the PowerShell tool.
                entries = sorted(it, key=lambda e: sort_key(e.name))
        except OSError as exc:
            log.warning("Skipping '%s': %s", folder, exc.strerror or exc)
            if folders is not None:
                folders.append(FolderRecord(folder, readable=False))
            continue
        if folders is not None:
            folders.append(FolderRecord(folder, readable=True))

        sub_folders = []
        for entry in entries:
            try:
                if entry.is_dir():
                    if is_folder_link(entry):
                        log.info("Not following link '%s'", entry.path)
                        if folders is not None:
                            folders.append(FolderRecord(entry.path, readable=False))
                    else:
                        sub_folders.append(entry.path)
                    continue
                if not entry.is_file():
                    continue
                if _same_path_key(entry.path) in excluded:
                    if folders is not None:
                        folders.append(FolderRecord(folder, readable=False))
                    continue
                st = entry.stat()
            except OSError as exc:
                log.warning("Skipping '%s': %s", entry.path, exc.strerror or exc)
                continue
            if not stat.S_ISREG(st.st_mode):
                continue
            file_count += 1
            yield FileRecord(
                path=entry.path,
                name=entry.name,
                folder=folder,
                size=st.st_size,
                mtime_ns=st.st_mtime_ns,
                attributes=getattr(st, "st_file_attributes", 0),
            )

        # Push in reverse so folders are visited in alphabetical order.
        pending.extend(reversed(sub_folders))
