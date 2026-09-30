"""Duplicate folder matching.

Mirrors Find-DuplicateFolder, Get-FolderTree and Get-FolderSignature in
src/DuplicateFinder.psm1. Two folders are duplicates when their names match
(ignoring case), they contain the same tree of file and sub folder names, and
every file at the same relative path is a duplicate by the file rule (name,
saved date and MD5). Their total sizes therefore match too.
"""

from __future__ import annotations

import hashlib
import logging
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Set

from .matcher import HashCallback, NS_PER_SECOND, groups_of_many, md5_map, ordinal_ignore_case
from .scanner import FileRecord, FolderRecord, is_cloud_only

log = logging.getLogger("find_duplicates")


@dataclass(frozen=True)
class DuplicateFolderSet:
    """One duplicated folder and the full path of every copy of it."""

    folder_name: str
    file_count: int
    folder_count: int
    size_bytes: int
    count: int
    folders: List[str]


@dataclass
class _FolderTree:
    children: Dict[str, List[str]] = field(default_factory=dict)
    files: Dict[str, List[FileRecord]] = field(default_factory=dict)
    readable: Set[str] = field(default_factory=set)
    deepest_first: List[str] = field(default_factory=list)


@dataclass
class _FolderInfo:
    signature: Optional[str] = None
    file_count: int = 0
    folder_count: int = 0
    size_bytes: int = 0


def _folder_tree(files: Sequence[FileRecord], folders: Sequence[FolderRecord]) -> _FolderTree:
    """Index a scan: the sub folders and files of every folder, and which were readable."""
    tree = _FolderTree()
    for record in folders:
        tree.children[record.path] = []
        tree.files[record.path] = []
        if record.readable:
            tree.readable.add(record.path)
    for record in folders:
        parent = os.path.dirname(record.path)
        if parent != record.path and parent in tree.children:
            tree.children[parent].append(record.path)
    for record in files:
        if record.folder in tree.files:
            tree.files[record.folder].append(record)
    # Deepest first: a sub folder's path is always longer than its parent's.
    tree.deepest_first = sorted(tree.children, key=len, reverse=True)
    return tree


def _signatures(
    tree: _FolderTree, md5: Optional[Dict[str, str]] = None, scope: Optional[Set[str]] = None
) -> Dict[str, _FolderInfo]:
    """Fingerprint folder trees bottom-up.

    A folder's fingerprint covers the name (ignoring case), size and saved second
    of every file below it, the name of every sub folder (including empty ones),
    and with ``md5`` each file's MD5. A folder whose tree could not be fully read,
    or has a file missing from ``md5``, gets no fingerprint.
    """
    result: Dict[str, _FolderInfo] = {}
    for folder in tree.deepest_first:
        if scope is not None and folder not in scope:
            continue
        info = result[folder] = _FolderInfo()
        if folder not in tree.readable:
            continue

        lines = []
        complete = True
        for record in tree.files[folder]:
            line = f"F|{ordinal_ignore_case(record.name)}|{record.mtime_ns // NS_PER_SECOND}|{record.size}"
            if md5 is not None:
                if record.path not in md5:
                    complete = False
                    break
                line += f"|{md5[record.path]}"
            lines.append(line)
            info.file_count += 1
            info.size_bytes += record.size
        for sub in tree.children[folder] if complete else []:
            child = result[sub]
            if child.signature is None:
                complete = False
                break
            lines.append(f"D|{ordinal_ignore_case(os.path.basename(sub))}|{child.signature}")
            info.file_count += child.file_count
            info.folder_count += child.folder_count + 1
            info.size_bytes += child.size_bytes
        if complete:
            text = "\n".join(sorted(lines)).encode("utf-8", "surrogatepass")
            info.signature = hashlib.sha256(text).hexdigest().upper()
    return result


def _name_key(path: str, signature: str) -> str:
    return f"{ordinal_ignore_case(os.path.basename(path))}|{signature}"


def find_duplicate_folders(
    files: Sequence[FileRecord],
    folders: Sequence[FolderRecord],
    skip_cloud_only: bool = False,
    throttle_limit: int = 1,
    md5_cache: Optional[Dict[str, str]] = None,
    on_hash: Optional[HashCallback] = None,
) -> List[DuplicateFolderSet]:
    """Find sets of folders that have the same name and exactly the same contents.

    Files are only hashed inside folders whose names, sizes and saved dates already
    match. Only the top-most duplicates are reported: a set is left out when every
    one of its folders sits inside a folder that is itself a reported duplicate.
    Folders with no files anywhere below them are not reported.
    """
    tree = _folder_tree(files, folders)

    # Pass 1: names, sizes and saved dates only; nothing is read.
    cheap = _signatures(tree)
    candidates = [p for p in tree.deepest_first if cheap[p].signature and cheap[p].file_count > 0]
    candidate_groups = groups_of_many(candidates, lambda p: _name_key(p, cheap[p].signature))
    if not candidate_groups:
        return []

    # Pass 2: hash every file below the candidates (hashes from the file scan are reused).
    scope: Set[str] = set()
    pending = [p for group in candidate_groups for p in group]
    while pending:
        path = pending.pop()
        if path not in scope:
            scope.add(path)
            pending.extend(tree.children[path])

    to_hash = []
    skipped = 0
    for path in scope:
        for record in tree.files[path]:
            if skip_cloud_only and is_cloud_only(record):
                skipped += 1
                log.info("Not downloading online-only file '%s'", record.path)
            else:
                to_hash.append(record.path)
    if skipped:
        log.warning(
            "%d online-only cloud file(s) were not checked; folders containing them are not reported.", skipped
        )
    full = _signatures(tree, md5_map(to_hash, throttle_limit, on_hash, md5_cache), scope)

    # Group the candidates again, now by contents.
    confirmed = [p for group in candidate_groups for p in group if full[p].signature]
    sets = groups_of_many(confirmed, lambda p: _name_key(p, full[p].signature))

    # Report only the top-most duplicates.
    duplicated = {p for s in sets for p in s}
    results = []
    for members in sets:
        if all(os.path.dirname(p) in duplicated for p in members):
            continue
        ordered = sorted(members, key=ordinal_ignore_case)
        info = full[ordered[0]]
        results.append(
            DuplicateFolderSet(
                folder_name=os.path.basename(ordered[0]),
                file_count=info.file_count,
                folder_count=info.folder_count,
                size_bytes=info.size_bytes,
                count=len(ordered),
                folders=ordered,
            )
        )
    results.sort(key=lambda s: (ordinal_ignore_case(s.folder_name), s.size_bytes, ordinal_ignore_case(s.folders[0])))
    return results
