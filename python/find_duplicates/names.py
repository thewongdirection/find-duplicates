"""How names are compared and ordered, identically to the PowerShell tool.

Mirrors ConvertTo-NameKey and the .NET ordinal comparers used in
src/DuplicateFinder.psm1.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Iterable, Optional, Pattern, Tuple


def _upper(text: str) -> str:
    # One character at a time, leaving characters whose upper case is longer (such as
    # German sharp s) alone, so "strasse" and the sharp-s spelling differ, as in .NET.
    return "".join(upper if len(upper := char.upper()) == 1 else char for char in text)


def name_key(text: str) -> str:
    """The form in which names are compared (as ConvertTo-NameKey in PowerShell).

    Unicode-normalised (NFC, so "e + accent" as macOS often stores it equals the
    single character Windows stores) and upper case, so comparisons ignore case.
    """
    return _upper(unicodedata.normalize("NFC", text))


def sort_key(text: str, ignore_case: bool = False) -> bytes:
    """Orders strings as the PowerShell tool does: by UTF-16 code units (.NET's
    String.CompareOrdinal), upper-cased first with ``ignore_case`` (Compare-IgnoringCase).
    UTF-16 order differs from Python's code-point order for characters beyond U+FFFF
    such as emoji."""
    return (_upper(text) if ignore_case else text).encode("utf-16-be", "surrogatepass")


def path_sort_key(path: str) -> tuple:
    """Case-insensitive order for paths, with a fixed order for paths differing only in
    case (as ByPathIgnoringCase in PowerShell)."""
    return sort_key(path, ignore_case=True), sort_key(path)


def path_depth(path: str) -> int:
    """How many folders deep a path is: its parts between separators (\\ or /), so C:\\a\\b
    and /home/a are 3 and 2 deep (as Get-PathDepth in PowerShell)."""
    return len([part for part in re.split(r"[\\/]", path) if part])


def depth_sort_key(path: str) -> tuple:
    """The least nested path first, paths equally deep in path_sort_key order (as
    Get-SortedFolder -ByDepth in PowerShell)."""
    return (path_depth(path),) + path_sort_key(path)


@dataclass(frozen=True)
class _Wildcard:
    """One wildcard pattern: the parts between its stars, each a regular expression."""

    parts: Tuple[Pattern[str], ...]
    last_length: int  # characters the last part matches

    def matches(self, key: str) -> bool:
        # The first part must start the name and the last end it; each part between is taken
        # at its first place after the one before. That is always right for wildcards and
        # never backtracks, so a pattern such as *a*a*a*a*b cannot take exponential time (as
        # the atomic groups of ConvertTo-NameFilter in PowerShell).
        first, *rest = self.parts
        if not rest:
            return first.fullmatch(key) is not None
        found = first.match(key)
        if found is None:
            return False
        position = found.end()
        for part in rest[:-1]:
            found = part.search(key, position)
            if found is None:
                return False
            position = found.end()
        start = len(key) - self.last_length
        return start >= position and rest[-1].fullmatch(key, start) is not None


NameFilter = Tuple[_Wildcard, ...]


def check_name_pattern(pattern: str) -> None:
    """Reject a pattern that could never match a name: empty, or holding a path separator
    (ValueError, worded as Get-NamePatternProblem in PowerShell)."""
    if not pattern:
        raise ValueError("Exclusion patterns cannot be empty.")
    if "/" in pattern or "\\" in pattern:
        raise ValueError(f"Exclusion pattern '{pattern}' contains / or \\: patterns match file and folder names, not paths.")


def _wildcard(pattern: str) -> _Wildcard:
    check_name_pattern(pattern)
    parts = name_key(pattern).split("*")
    return _Wildcard(
        tuple(re.compile("".join("." if c == "?" else re.escape(c) for c in part), re.DOTALL) for part in parts),
        len(parts[-1]),
    )


def name_filter(patterns: Iterable[str]) -> Optional[NameFilter]:
    """A filter matching the name key (see name_key) of any name that matches one of the
    wildcard patterns, so names match whatever their case or Unicode form (as
    ConvertTo-NameFilter in PowerShell). * stands for any run of characters and ? for any
    one character; everything else is literal. None when there are no patterns."""
    wildcards = tuple(_wildcard(pattern) for pattern in patterns)
    return wildcards or None


def matches_name_filter(names: Optional[NameFilter], name: str) -> bool:
    """Whether ``name`` matches a name_filter(); never when there is none."""
    if names is None:
        return False
    key = name_key(name)
    return any(wildcard.matches(key) for wildcard in names)
