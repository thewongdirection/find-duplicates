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
import os
import sys
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Callable, Dict, Hashable, Iterable, List, Optional, Sequence, TypeVar

from .names import name_key, path_sort_key, sort_key
from . import scanner
from .scanner import FileRecord, is_cloud_only, local_time, utc_offset

log = logging.getLogger("find_duplicates")

NS_PER_SECOND = 1_000_000_000
HASH_CHUNK_BYTES = 1024 * 1024
MIN_THROTTLE_LIMIT = 1
MAX_THROTTLE_LIMIT = 64
# Local files at least this large are worth hashing on a thread (hashlib releases Python's
# global lock while hashing them); smaller local files are faster hashed one at a time.
PARALLEL_HASH_MIN_BYTES = 1024 * 1024
# Large candidates are first compared by the MD5 of their start (see _split_by_start_hash):
# files that share a name, saved date and size but not their contents are then rarely read
# in full. Only for files this large, so that true duplicates cost at most 1/16 more reading.
FIRST_BYTES_TO_HASH = 1024 * 1024
FIRST_BYTES_MIN_SIZE = 16 * 1024 * 1024

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
    # Local time minus UTC at last_write_time; None for rows read from a report made before
    # the UTC Offset column (those are checked by local time).
    utc_offset: Optional[timedelta] = None


def md5_file(path: str, limit: int = 0) -> str:
    """MD5 of a file's contents as upper-case hex (the Get-FileHash format); with
    ``limit`` above 0, of its first ``limit`` bytes only."""
    digest = hashlib.md5(usedforsecurity=False)  # noqa: S324 - part of the duplicate definition
    remaining = limit if limit > 0 else sys.maxsize
    with open(path, "rb", buffering=0) as stream:
        # Chunks of up to 1 MB into a buffer no larger than the file: small files, the most
        # common, cost no large allocation.
        buffer = memoryview(bytearray(max(1, min(HASH_CHUNK_BYTES, os.fstat(stream.fileno()).st_size, remaining))))
        while remaining > 0:
            read = stream.readinto(buffer[:min(len(buffer), remaining)])
            if not read:
                break
            digest.update(buffer[:read])
            remaining -= read
    return digest.hexdigest().upper()


def _hash_or_none(path: str, limit: int = 0) -> Optional[str]:
    # Looks md5_file up at call time so tests can replace it.
    try:
        return md5_file(path, limit) if limit else md5_file(path)
    except OSError as exc:
        log.warning("Could not hash '%s': %s", path, exc.strerror or exc)
        return None


def md5_map(
    paths: Sequence[str],
    throttle_limit: int = 1,
    on_hash: Optional[HashCallback] = None,
    cache: Optional[Dict[str, str]] = None,
    sizes: Optional[Dict[str, int]] = None,
    first_bytes: bool = False,
) -> Dict[str, str]:
    """Hash files, up to ``throttle_limit`` at a time, returning full path -> MD5.

    Files that cannot be read are logged as warnings and left out of the map.
    With ``cache``, files already in it are not read again and new hashes are
    added to it. Only files on a network drive, or of at least
    PARALLEL_HASH_MIN_BYTES by ``sizes`` (all files when ``sizes`` is not given), are
    hashed on threads; the rest are hashed one at a time meanwhile, which is faster.
    With ``first_bytes``, only the first FIRST_BYTES_TO_HASH bytes of each file are
    hashed (and ``cache``, which holds whole-file hashes, must not be given).
    """
    if not MIN_THROTTLE_LIMIT <= throttle_limit <= MAX_THROTTLE_LIMIT:
        raise ValueError(f"throttle_limit must be {MIN_THROTTLE_LIMIT}-{MAX_THROTTLE_LIMIT}, not {throttle_limit}.")
    limit = FIRST_BYTES_TO_HASH if first_bytes else 0

    if cache is not None:
        hashed = _md5_map([p for p in paths if p not in cache], throttle_limit, on_hash, sizes, limit)
        cache.update(hashed)
        return {p: cache[p] for p in paths if p in cache}
    return _md5_map(paths, throttle_limit, on_hash, sizes, limit)


def _worth_a_thread(path: str, sizes: Optional[Dict[str, int]]) -> bool:
    size = None if sizes is None else sizes.get(path)
    return size is None or size >= PARALLEL_HASH_MIN_BYTES or scanner.on_network_drive(path)


def _md5_map(
    paths: Sequence[str],
    throttle_limit: int,
    on_hash: Optional[HashCallback],
    sizes: Optional[Dict[str, int]],
    limit: int = 0,
) -> Dict[str, str]:
    result: Dict[str, str] = {}

    def record(done: int, path: str, md5: Optional[str]) -> None:
        if on_hash is not None:
            on_hash(path, done, len(paths))
        if md5 is not None:
            result[path] = md5

    threaded = [p for p in paths if _worth_a_thread(p, sizes)] if throttle_limit > 1 else []
    if not threaded:
        for done, path in enumerate(paths, start=1):
            record(done, path, _hash_or_none(path, limit))
        return result

    # hashlib releases the GIL while hashing, so threads hash large or remote files in
    # parallel; the others are hashed here in the meantime.
    in_thread = set(threaded)
    done = 0
    with ThreadPoolExecutor(max_workers=throttle_limit) as pool:
        futures = [(path, pool.submit(_hash_or_none, path, limit)) for path in threaded]
        for path in paths:
            if path not in in_thread:
                done += 1
                record(done, path, _hash_or_none(path, limit))
        for path, future in futures:
            done += 1
            record(done, path, future.result())
    return result


def _split_by_start_hash(
    groups: List[List[FileRecord]],
    throttle_limit: int,
    on_hash: Optional[HashCallback],
    md5_cache: Optional[Dict[str, str]],
) -> List[List[FileRecord]]:
    """Split each group of files of FIRST_BYTES_MIN_SIZE or more by the MD5 of their first
    FIRST_BYTES_TO_HASH bytes, keeping the groups that still hold more than one file. Other
    groups, and groups a file of which already has its whole-file MD5 in ``md5_cache``
    (from a previous report), are kept as they are (as Split-ByStartHash in PowerShell)."""
    kept, large = [], []
    for group in groups:
        is_large = group[0].size >= FIRST_BYTES_MIN_SIZE  # the files share their size (stage 2)
        if is_large and md5_cache is not None and any(r.path in md5_cache for r in group):
            is_large = False
        (large if is_large else kept).append(group)
    if not large:
        return kept

    paths = [r.path for group in large for r in group]
    starts = md5_map(paths, throttle_limit, on_hash, sizes={p: FIRST_BYTES_MIN_SIZE for p in paths}, first_bytes=True)
    for group in large:
        read = [(r, starts[r.path]) for r in group if r.path in starts]
        kept.extend([r for r, _ in same] for same in groups_of_many(read, lambda pair: pair[1]))
    return kept


def _saved_date_key(record: FileRecord) -> int:
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
    minimum_size: int = 0,
) -> List[DuplicateSet]:
    """Find sets of files whose name, saved date and MD5 hash all match.

    With ``skip_cloud_only`` online-only cloud files are never hashed (hashing
    would download them); duplicates among such files are then not reported.
    ``throttle_limit`` is how many files to hash at the same time (1-64).
    ``md5_cache`` is shared with find_duplicate_folders so no file is read twice.
    ``ignore_empty_files`` leaves files of 0 bytes out (they all share one MD5), and
    ``minimum_size`` files smaller than that many bytes.
    """
    if minimum_size < 0:
        raise ValueError(f"minimum_size cannot be negative, not {minimum_size}.")
    # Stage 1: name. Grouped first, so the size and saved date of a file whose name no other
    # file has are never read: on Linux, macOS and network drives each is a request per file.
    name_groups = groups_of_many(files, lambda f: name_key(f.name))

    # Stage 2: saved date (whole second: copies made to network shares or other file systems
    # often lose sub-second precision) and size, a cheap check that avoids hashing files that
    # cannot match.
    if ignore_empty_files:
        minimum_size = max(minimum_size, 1)
    candidate_groups: List[List[FileRecord]] = []
    for group in name_groups:
        keyed = []
        for record in group:
            try:
                key = (_saved_date_key(record), record.size)
            except OSError as exc:
                log.warning("Skipping '%s': %s", record.path, exc.strerror or exc)  # gone or unreadable since the scan
                continue
            if key[1] >= minimum_size:
                keyed.append((record, key))
        candidate_groups.extend([record for record, _ in same] for same in groups_of_many(keyed, lambda pair: pair[1]))

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

    # Stage 3: for large files, the MD5 of their start.
    candidate_groups = _split_by_start_hash(candidate_groups, throttle_limit, on_hash, md5_cache)

    # Stage 4: MD5, only for files that already match on name, date and size.
    candidates = [r for group in candidate_groups for r in group]
    md5_by_path = md5_map(
        [r.path for r in candidates], throttle_limit, on_hash, md5_cache, sizes={r.path: r.size for r in candidates}
    )

    results: List[DuplicateSet] = []
    for group in candidate_groups:
        hashed_files = [(r, md5_by_path[r.path]) for r in group if r.path in md5_by_path]
        for same in groups_of_many(hashed_files, lambda pair: pair[1]):
            first, md5 = same[0]
            results.append(
                DuplicateSet(
                    file_name=first.name,
                    last_write_time=local_time(first.mtime_ns / NS_PER_SECOND),
                    utc_offset=utc_offset(first.mtime_ns // NS_PER_SECOND),
                    size_bytes=first.size,
                    md5=md5,
                    count=len(same),
                    folders=sorted((record.folder for record, _ in same), key=path_sort_key),
                )
            )

    results.sort(key=lambda s: (sort_key(s.file_name, True), s.last_write_time, sort_key(s.md5)))
    return results

