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

from .matcher import HashCallback, NS_PER_SECOND, groups_of_many, md5_map
from .names import name_key, path_sort_key, sort_key
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
    """Index a scan: the sub folders and files of every folder, and which were readable.

    File details are not read here (see _confirm_tree_files): only the folders that could
    be duplicates need them.
    """
    tree = _FolderTree()
    for record in folders:
        tree.children.setdefault(record.path, [])
        tree.files.setdefault(record.path, [])
    # A folder may be recorded more than once; recorded as not readable anywhere (an
    # excluded file, a skipped link) means not readable.
    unreadable = {record.path for record in folders if not record.readable}
    for record in files:
        if record.folder in tree.files:
            tree.files[record.folder].append(record)
    tree.readable = {path for path in tree.children if path not in unreadable}
    for path in list(tree.children):
        parent = os.path.dirname(path)
        if parent != path and parent in tree.children:
            tree.children[parent].append(path)
    # Deepest first: a sub folder's path is always longer than its parent's.
    tree.deepest_first = sorted(tree.children, key=len, reverse=True)
    return tree


def _confirm_tree_files(tree: _FolderTree, files: Sequence[FileRecord], scope: Set[str]) -> None:
    """Read the size and saved date of every file in the ``scope`` folders (in scan order).
    A file gone or unreadable since the scan is dropped with a warning, and leaves its
    folder's contents unknown: the folder is no longer readable."""
    for record in files:
        if record.folder not in scope or record.folder not in tree.files:
            continue
        try:
            record.size, record.mtime_ns
        except OSError as exc:
            log.warning("Skipping '%s': %s", record.path, exc.strerror or exc)
            tree.files[record.folder].remove(record)
            tree.readable.discard(record.folder)


def _folder_scope(tree: _FolderTree, folders: Sequence[str]) -> Dict[str, None]:
    """The given folders and every folder below them. A dict keeps insertion order, so
    files are hashed in the same order as in PowerShell."""
    scope: Dict[str, None] = {}
    pending = list(folders)
    while pending:
        path = pending.pop()
        if path not in scope:
            scope[path] = None
            pending.extend(tree.children[path])
    return scope


def _repeated_name_scope(tree: _FolderTree) -> Set[str]:
    """The folders whose name (ignoring case) another folder has, and every folder below them."""
    repeated = [p for group in groups_of_many(tree.deepest_first, lambda p: name_key(os.path.basename(p))) for p in group]
    return set(_folder_scope(tree, repeated))


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
            line = f"F|{name_key(record.name)}|{record.mtime_ns // NS_PER_SECOND}|{record.size}"
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
            lines.append(f"D|{name_key(os.path.basename(sub))}|{child.signature}")
            info.file_count += child.file_count
            info.folder_count += child.folder_count + 1
            info.size_bytes += child.size_bytes
        if complete:
            text = "\n".join(sorted(lines)).encode("utf-8", "surrogatepass")
            info.signature = hashlib.sha256(text).hexdigest().upper()
    return result


def _name_key(path: str, signature: str) -> str:
    return f"{name_key(os.path.basename(path))}|{signature}"


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
    one of its folders sits inside a folder that is itself a duplicate.
    Folders with no files anywhere below them are not reported.
    """
    tree = _folder_tree(files, folders)

    # Pass 1: names, sizes and saved dates only; nothing is read. Only folders whose name
    # another folder has can be duplicates, so only they (and the folders below them, part
    # of their fingerprint) are fingerprinted.
    named = _repeated_name_scope(tree)
    if not named:
        return []
    _confirm_tree_files(tree, files, named)
    cheap = _signatures(tree, scope=named)
    candidates = [p for p in tree.deepest_first if p in named and cheap[p].signature and cheap[p].file_count > 0]
    candidate_groups = groups_of_many(candidates, lambda p: _name_key(p, cheap[p].signature))
    if not candidate_groups:
        return []

    # Pass 2: hash every file below the candidates (hashes from the file scan are reused).
    scope = _folder_scope(tree, [p for group in candidate_groups for p in group])

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
    sizes = {record.path: record.size for path in scope for record in tree.files[path]}
    full = _signatures(tree, md5_map(to_hash, throttle_limit, on_hash, md5_cache, sizes), set(scope))

    # Group the candidates again, now by contents.
    confirmed = [p for group in candidate_groups for p in group if full[p].signature]
    sets = groups_of_many(confirmed, lambda p: _name_key(p, full[p].signature))

    # Report only the top-most duplicates.
    duplicated = {p for s in sets for p in s}
    results = []
    for members in sets:
        if all(os.path.dirname(p) in duplicated for p in members):
            continue
        ordered = sorted(members, key=path_sort_key)
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
    results.sort(key=lambda s: (sort_key(s.folder_name, True), s.size_bytes, path_sort_key(s.folders[0])))
    return results
