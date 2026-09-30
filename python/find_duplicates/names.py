"""How names are compared and ordered, identically to the PowerShell tool.

Mirrors ConvertTo-NameKey and the .NET ordinal comparers used in
src/DuplicateFinder.psm1.
"""

from __future__ import annotations

import unicodedata


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
    """Orders strings exactly as .NET's ordinal comparers do (StringComparer.Ordinal,
    or OrdinalIgnoreCase with ``ignore_case``): by UTF-16 code units, which differs
    from Python's code-point order for characters beyond U+FFFF such as emoji."""
    return (_upper(text) if ignore_case else text).encode("utf-16-be", "surrogatepass")
