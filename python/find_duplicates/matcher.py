"""Duplicate matching.

Mirrors Find-DuplicateFile in src/DuplicateFinder.psm1. A file is a duplicate
of another when ALL of the following match:

1. File name (case-insensitive)
2. Saved date (last modified time, compared to the whole second)
3. MD5 hash of the contents

MD5 is only computed for files whose name and saved date already match another
file (and whose size matches, since files of different sizes can never share an
MD5), so most files are never read.
"""

from __future__ import annotations

import hashlib
import logging
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Dict, Hashable, Iterable, List, Optional, Sequence, TypeVar

from .scanner import FileRecord, is_cloud_only

log = logging.getLogger("find_duplicates")

NS_PER_SECOND = 1_000_000_000
HASH_CHUNK_BYTES = 1024 * 1024

T = TypeVar("T")

# Called with (file path, files hashed so far, total files to hash).
HashCallback = Callable[[str, int, int], None]


@dataclass(frozen=True)
class DuplicateSet:
    """One duplicated file and every folder holding a copy of it."""

    file_name: str
    last_write_time: datetime  # local time, like PowerShell's LastWriteTime
    size_bytes: int
    md5: str
    count: int
    folders: List[str]


def md5_file(path: str) -> str:
    """MD5 of a file's contents as upper-case hex (the Get-FileHash format)."""
    digest = hashlib.md5(usedforsecurity=False)  # noqa: S324 - part of the duplicate definition
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(HASH_CHUNK_BYTES), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _saved_date_key(record: FileRecord) -> int:
    # Whole seconds: copies made to network shares or other file systems
    # frequently lose sub-second precision.
    return record.mtime_ns // NS_PER_SECOND


def _groups_of_many(items: Iterable[T], key: Callable[[T], Hashable]) -> List[List[T]]:
    """Group items by key, keeping only the groups that hold more than one item."""
    groups: Dict[Hashable, List[T]] = defaultdict(list)
    for item in items:
        groups[key(item)].append(item)
    return [group for group in groups.values() if len(group) > 1]


def find_duplicate_files(
    files: Sequence[FileRecord],
    skip_cloud_only: bool = False,
    on_hash: Optional[HashCallback] = None,
) -> List[DuplicateSet]:
    """Find sets of files whose name, saved date and MD5 hash all match.

    With ``skip_cloud_only`` online-only cloud files are never hashed (hashing
    would download them); duplicates among such files are then not reported.
    """
    # Stage 1: name + saved date.
    name_date_groups = _groups_of_many(files, lambda f: (_ordinal_ignore_case(f.name), _saved_date_key(f)))

    # Stage 2: size. A cheap check that avoids hashing files that cannot match.
    candidate_groups = [
        group for nd in name_date_groups for group in _groups_of_many(nd, lambda f: f.size)
    ]

    if skip_cloud_only:
        skipped = 0
        kept_groups = []
        for group in candidate_groups:
            local = []
            for record in group:
                if is_cloud_only(record):
                    skipped += 1
                    log.info("Not downloading online-only file '%s'", record.path)
                else:
                    local.append(record)
            if len(local) > 1:
                kept_groups.append(local)
        candidate_groups = kept_groups
        if skipped:
            log.warning(
                "%d online-only cloud file(s) were not checked; duplicates among them are not reported.",
                skipped,
            )

    # Stage 3: MD5, only for files that already match on name, date and size.
    to_hash = sum(len(group) for group in candidate_groups)
    hashed = 0
    results: List[DuplicateSet] = []
    for group in candidate_groups:
        hashed_files = []
        for record in group:
            hashed += 1
            if on_hash is not None:
                on_hash(record.path, hashed, to_hash)
            try:
                hashed_files.append((record, md5_file(record.path)))
            except OSError as exc:
                log.warning("Could not hash '%s': %s", record.path, exc.strerror or exc)

        for same in _groups_of_many(hashed_files, lambda pair: pair[1]):
            first, md5 = same[0]
            results.append(
                DuplicateSet(
                    file_name=first.name,
                    last_write_time=datetime.fromtimestamp(first.mtime_ns / NS_PER_SECOND),
                    size_bytes=first.size,
                    md5=md5,
                    count=len(same),
                    folders=sorted((record.folder for record, _ in same), key=_ordinal_ignore_case),
                )
            )

    results.sort(key=lambda s: (_ordinal_ignore_case(s.file_name), s.last_write_time, s.md5))
    return results


def _ordinal_ignore_case(text: str) -> str:
    """Key matching .NET StringComparer.OrdinalIgnoreCase, used by the PowerShell tool.

    Upper-cases one character at a time and leaves characters whose upper case
    is longer (such as German sharp s) alone, so "straße" and "STRASSE" differ,
    exactly as they do in .NET.
    """
    return "".join(upper if len(upper := char.upper()) == 1 else char for char in text)
