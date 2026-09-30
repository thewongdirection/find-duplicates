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
from collections import defaultdict, deque
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Dict, Hashable, Iterable, List, Optional, Sequence, TypeVar

from .names import name_key, path_sort_key, sort_key
from .scanner import FileRecord, is_cloud_only, local_time

log = logging.getLogger("find_duplicates")

NS_PER_SECOND = 1_000_000_000
HASH_CHUNK_BYTES = 1024 * 1024
MIN_THROTTLE_LIMIT = 1
MAX_THROTTLE_LIMIT = 64

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
    with open(path, "rb", buffering=0) as stream:
        for chunk in iter(lambda: stream.read(HASH_CHUNK_BYTES), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _hash_or_none(path: str) -> Optional[str]:
    # Looks md5_file up at call time so tests can replace it.
    try:
        return md5_file(path)
    except OSError as exc:
        log.warning("Could not hash '%s': %s", path, exc.strerror or exc)
        return None


def md5_map(
    paths: Sequence[str],
    throttle_limit: int = 1,
    on_hash: Optional[HashCallback] = None,
    cache: Optional[Dict[str, str]] = None,
) -> Dict[str, str]:
    """Hash files, up to ``throttle_limit`` at a time, returning full path -> MD5.

    Files that cannot be read are logged as warnings and left out of the map.
    With ``cache``, files already in it are not read again and new hashes are
    added to it.
    """
    if not MIN_THROTTLE_LIMIT <= throttle_limit <= MAX_THROTTLE_LIMIT:
        raise ValueError(f"throttle_limit must be {MIN_THROTTLE_LIMIT}-{MAX_THROTTLE_LIMIT}, not {throttle_limit}.")

    if cache is not None:
        hashed = _md5_map(paths=[p for p in paths if p not in cache], throttle_limit=throttle_limit, on_hash=on_hash)
        cache.update(hashed)
        return {p: cache[p] for p in paths if p in cache}
    return _md5_map(paths, throttle_limit, on_hash)


def _md5_map(paths: Sequence[str], throttle_limit: int, on_hash: Optional[HashCallback]) -> Dict[str, str]:
    result: Dict[str, str] = {}

    def record(done: int, path: str, md5: Optional[str]) -> None:
        if on_hash is not None:
            on_hash(path, done, len(paths))
        if md5 is not None:
            result[path] = md5

    if throttle_limit == 1:
        for done, path in enumerate(paths, start=1):
            record(done, path, _hash_or_none(path))
        return result

    # hashlib releases the GIL while hashing, so threads hash in parallel. A bounded
    # window of pending jobs keeps memory flat however many files there are.
    window = throttle_limit * 4
    pending: deque = deque()
    done = 0
    with ThreadPoolExecutor(max_workers=throttle_limit) as pool:
        for path in paths:
            pending.append((path, pool.submit(_hash_or_none, path)))
            if len(pending) >= window:
                done += 1
                oldest, future = pending.popleft()
                record(done, oldest, future.result())
        while pending:
            done += 1
            oldest, future = pending.popleft()
            record(done, oldest, future.result())
    return result


def _saved_date_key(record: FileRecord) -> int:
    # Whole seconds: copies made to network shares or other file systems
    # frequently lose sub-second precision.
    return record.mtime_ns // NS_PER_SECOND


def groups_of_many(items: Iterable[T], key: Callable[[T], Hashable]) -> List[List[T]]:
    """Group items by key, keeping only the groups that hold more than one item."""
    groups: Dict[Hashable, List[T]] = defaultdict(list)
    for item in items:
        groups[key(item)].append(item)
    return [group for group in groups.values() if len(group) > 1]


def find_duplicate_files(
    files: Sequence[FileRecord],
    skip_cloud_only: bool = False,
    on_hash: Optional[HashCallback] = None,
    throttle_limit: int = 1,
    md5_cache: Optional[Dict[str, str]] = None,
    ignore_empty_files: bool = False,
) -> List[DuplicateSet]:
    """Find sets of files whose name, saved date and MD5 hash all match.

    With ``skip_cloud_only`` online-only cloud files are never hashed (hashing
    would download them); duplicates among such files are then not reported.
    ``throttle_limit`` is how many files to hash at the same time (1-64).
    ``md5_cache`` is shared with find_duplicate_folders so no file is read twice.
    ``ignore_empty_files`` leaves files of 0 bytes out (they all share one MD5).
    """
    if ignore_empty_files:
        files = [f for f in files if f.size > 0]

    # Stage 1: name + saved date.
    name_date_groups = groups_of_many(files, lambda f: (name_key(f.name), _saved_date_key(f)))

    # Stage 2: size. A cheap check that avoids hashing files that cannot match.
    candidate_groups = [
        group for nd in name_date_groups for group in groups_of_many(nd, lambda f: f.size)
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
    md5_by_path = md5_map([r.path for group in candidate_groups for r in group], throttle_limit, on_hash, md5_cache)

    results: List[DuplicateSet] = []
    for group in candidate_groups:
        hashed_files = [(r, md5_by_path[r.path]) for r in group if r.path in md5_by_path]
        for same in groups_of_many(hashed_files, lambda pair: pair[1]):
            first, md5 = same[0]
            results.append(
                DuplicateSet(
                    file_name=first.name,
                    last_write_time=local_time(first.mtime_ns / NS_PER_SECOND),
                    size_bytes=first.size,
                    md5=md5,
                    count=len(same),
                    folders=sorted((record.folder for record, _ in same), key=path_sort_key),
                )
            )

    results.sort(key=lambda s: (sort_key(s.file_name, True), s.last_write_time, sort_key(s.md5)))
    return results

