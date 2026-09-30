"""Recursive folder scanning.

Mirrors Get-FileInventory, Test-FolderLink and Test-CloudOnlyFile in
src/DuplicateFinder.psm1.
"""

from __future__ import annotations

import functools
import logging
import os
import queue
import re
import sys
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Callable, Dict, Iterable, Iterator, List, Optional, Tuple

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


class FileRecord:
    """The details of one file that duplicate matching needs.

    Size, saved time and attributes (Windows file attributes; 0 elsewhere) come from the
    scan's directory entry and are read only when first used, like .NET's FileInfo: on
    Linux, macOS and network drives each read is a request per file, and most files never
    need them because no other file has their name. Reading them raises OSError when the
    file has gone or cannot be read since the scan.
    """

    __slots__ = ("path", "name", "folder", "_entry", "_size", "_mtime_ns", "_attributes")

    def __init__(
        self,
        path: str,
        name: str,
        folder: str,
        size: Optional[int] = None,
        mtime_ns: Optional[int] = None,
        attributes: int = 0,
        entry: Optional[os.DirEntry] = None,
    ) -> None:
        self.path, self.name, self.folder = path, name, folder
        self._entry = entry
        self._size, self._mtime_ns, self._attributes = size, mtime_ns, attributes

    def _load(self) -> None:
        info = self._entry.stat() if self._entry is not None else os.stat(self.path)
        self._size, self._mtime_ns = info.st_size, info.st_mtime_ns
        self._attributes = getattr(info, "st_file_attributes", 0)
        self._entry = None  # no longer needed

    @property
    def size(self) -> int:
        if self._size is None:
            self._load()
        return self._size

    @property
    def mtime_ns(self) -> int:
        if self._mtime_ns is None:
            self._load()
        return self._mtime_ns

    @property
    def attributes(self) -> int:
        if self._size is None:
            self._load()
        return self._attributes

    def __repr__(self) -> str:
        return f"FileRecord({self.path!r})"


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


# File systems reached over a network, as Linux names them in /proc/self/mounts.
NETWORK_FILE_SYSTEMS = frozenset({
    "9p", "afs", "ceph", "cifs", "davfs", "fuse.davfs2", "fuse.gcsfuse", "fuse.glusterfs", "fuse.rclone",
    "fuse.s3fs", "fuse.sshfs", "glusterfs", "gpfs", "lustre", "ncpfs", "nfs", "nfs4", "smb3", "smbfs",
})
DRIVE_REMOTE = 4  # GetDriveTypeW: a network drive


def on_network_drive(path: str) -> bool:
    """True when ``path`` is on a network share: a UNC path or network drive on Windows, a
    network file system (NFS, SMB ...) on Linux. False when unknown (macOS).

    Python uses several threads (-j) only there, and for hashing large files: for the many
    small operations of listing and checking local folders, threads contending for Python's
    global lock make it many times slower. (PowerShell has no such lock.)
    """
    path = os.path.abspath(path)
    if sys.platform == "win32":
        drive = os.path.splitdrive(path)[0]
        return drive.startswith(("\\\\", "//")) or _windows_drive_type(drive.upper()) == DRIVE_REMOTE
    if sys.platform.startswith("linux"):
        return mount_type(path, _linux_mounts()) in NETWORK_FILE_SYSTEMS
    return False


def mount_type(path: str, mounts: Iterable[Tuple[str, str]]) -> str:
    """The file system type of the mount holding ``path``, given (mount point, type) pairs."""
    best, kind = "", ""
    for point, point_kind in mounts:
        inside = path == point or path.startswith(point.rstrip("/") + "/")
        if inside and len(point) > len(best):
            best, kind = point, point_kind
    return kind


_OCTAL_ESCAPE = re.compile(r"\\([0-7]{3})")


@functools.lru_cache(maxsize=None)
def _linux_mounts() -> Tuple[Tuple[str, str], ...]:
    # /proc/self/mounts writes spaces and some other characters in mount points as octal (\040).
    try:
        with open("/proc/self/mounts", encoding="utf-8", errors="replace") as mounts:
            lines = [line.split() for line in mounts]
    except OSError:
        return ()
    return tuple(
        (_OCTAL_ESCAPE.sub(lambda m: chr(int(m.group(1), 8)), fields[1]), fields[2]) for fields in lines if len(fields) >= 3
    )


@functools.lru_cache(maxsize=None)
def _windows_drive_type(drive: str) -> int:
    import ctypes

    return ctypes.windll.kernel32.GetDriveTypeW(drive + "\\")


_EPOCH = datetime(1970, 1, 1, tzinfo=timezone.utc)
_EPOCH_NAIVE = datetime(1970, 1, 1)


def local_time(seconds: float) -> datetime:
    """Seconds since 1970-01-01 UTC as a naive local date and time, like .NET's LastWriteTime.

    On Windows, datetime.fromtimestamp rejects times before 1970. Those are converted
    with the UTC offset in force on the same date in 1972 (a leap year, so 29 February
    exists): the same seasonal rule .NET applies to dates older than its time zone data.
    """
    try:
        return datetime.fromtimestamp(seconds)
    except (OSError, OverflowError, ValueError):
        utc = _EPOCH + timedelta(seconds=seconds)
        probe = utc.replace(year=1972)
        offset = datetime.fromtimestamp(probe.timestamp()) - probe.replace(tzinfo=None)
        return (utc + offset).replace(tzinfo=None)


def utc_offset(seconds: int) -> timedelta:
    """Local time minus UTC at a moment given in whole seconds since 1970-01-01 UTC, like
    .NET's LastWriteTime - LastWriteTimeUtc."""
    return local_time(seconds) - (_EPOCH_NAIVE + timedelta(seconds=seconds))


def local_utc_offset(local: datetime) -> timedelta:
    """This computer's UTC offset at a naive local date and time, like .NET's
    TimeZoneInfo.Local.GetUtcOffset: the standard offset when the time happens twice or
    not at all (daylight saving changes). Dates Python cannot convert here (before 1970 on
    Windows) use the same date in 1972, as local_time does."""

    def offset(value: datetime) -> timedelta:
        return min(value.replace(fold=0).astimezone().utcoffset(), value.replace(fold=1).astimezone().utcoffset())

    try:
        return offset(local)
    except (OSError, OverflowError, ValueError):
        return offset(local.replace(year=1972))


def _same_path_key(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


@dataclass
class _Listing:
    """One folder's files, sub folders and folder links, each in ordinal name order, or
    the reason it could not be read."""

    files: List[os.DirEntry] = field(default_factory=list)
    folders: List[os.DirEntry] = field(default_factory=list)
    links: List[os.DirEntry] = field(default_factory=list)
    error: Optional[str] = None


def _list_folder(path: str) -> _Listing:
    try:
        with os.scandir(path) as it:
            # File systems list entries in different orders (alphabetical on NTFS,
            # arbitrary on ext4); ordinal order matches the PowerShell tool.
            entries = sorted(it, key=lambda e: sort_key(e.name))
    except OSError as exc:
        return _Listing(error=exc.strerror or str(exc))
    listing = _Listing()
    for entry in entries:
        try:
            if entry.is_dir():
                (listing.links if is_folder_link(entry) else listing.folders).append(entry)
            elif entry.is_file():  # not devices, pipes or sockets
                listing.files.append(entry)
        except OSError as exc:
            log.warning("Skipping '%s': %s", entry.path, exc.strerror or exc)
    return listing


def _tree_listing(root: str, throttle_limit: int, on_folder: Optional[FolderCallback]) -> Dict[str, _Listing]:
    """List every folder below ``root`` (folder links are not followed), ``throttle_limit``
    folders at a time; returns folder path -> listing."""
    listings: Dict[str, _Listing] = {}
    file_count = 0
    finished: "queue.SimpleQueue[Tuple[str, Future]]" = queue.SimpleQueue()

    def submit(path: str) -> None:
        # Finished listings are queued: waiting on the set of pending ones would cost a
        # pass over all of them each time, and thousands can be pending.
        pool.submit(_list_folder, path).add_done_callback(lambda future: finished.put((path, future)))

    with ThreadPoolExecutor(max_workers=throttle_limit) as pool:
        submit(root)
        pending = 1
        while pending:
            path, future = finished.get()
            pending -= 1
            listing = listings[path] = future.result()
            file_count += len(listing.files)
            if on_folder is not None:
                on_folder(path, len(listings), file_count)
            for entry in listing.folders:
                submit(entry.path)
                pending += 1
    return listings


def iter_files(
    root: str,
    exclude: Iterable[str] = (),
    on_folder: Optional[FolderCallback] = None,
    folders: Optional[List[FolderRecord]] = None,
    throttle_limit: int = 1,
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

    With ``throttle_limit`` above 1, that many folders on a network drive are listed at
    the same time (see on_network_drive); the files come out in the same order.
    """
    if not os.path.isdir(root):
        raise NotADirectoryError(f"'{root}' is not a folder.")

    # Normalised like the scanned paths (long names), or a short-form path would never match.
    excluded = {_same_path_key(full_path(p)) for p in exclude}
    root = full_path(root)
    # Listing several folders at a time lists the whole tree first; it is then walked below
    # exactly as when listing one folder at a time, so the output is the same.
    listings = _tree_listing(root, throttle_limit, on_folder) if throttle_limit > 1 and on_network_drive(root) else None
    pending = [root]
    folder_count = 0
    file_count = 0

    while pending:
        folder = pending.pop()
        folder_count += 1
        log.info("Scanning %s", folder)
        if listings is not None:
            listing = listings[folder]
        else:
            if on_folder is not None:
                on_folder(folder, folder_count, file_count)
            listing = _list_folder(folder)

        if listing.error is not None:
            log.warning("Skipping '%s': %s", folder, listing.error)
            if folders is not None:
                folders.append(FolderRecord(folder, readable=False))
            continue
        if folders is not None:
            folders.append(FolderRecord(folder, readable=True))

        for entry in listing.files:
            if _same_path_key(entry.path) in excluded:
                if folders is not None:
                    folders.append(FolderRecord(folder, readable=False))
                continue
            file_count += 1
            yield FileRecord(entry.path, entry.name, folder, entry=entry)

        for entry in listing.links:
            log.info("Not following link '%s'", entry.path)
            if folders is not None:
                folders.append(FolderRecord(entry.path, readable=False))
        # Push in reverse so folders are visited in alphabetical order.
        pending.extend(entry.path for entry in reversed(listing.folders))
