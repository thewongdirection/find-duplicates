# find-duplicates

Finds duplicate files (same **name**, **saved date** and **MD5**) in a folder
tree and saves them to an Excel workbook. The tool exists in two languages
that must behave identically:

| Part                | PowerShell (primary)                 | Python (port)                          |
|---------------------|--------------------------------------|----------------------------------------|
| Command line        | `Find-Duplicates.ps1`                | `python/find_duplicates/cli.py`        |
| Folder scanning     | `Get-FileInventory`, `Test-FolderLink`, `Test-CloudOnlyFile`, `Get-SortedByName` | `python/find_duplicates/scanner.py` |
| Duplicate matching  | `Find-DuplicateFile`, `Get-FileMd5`, `Get-FileMd5Map`, `Get-SortedFolder` | `python/find_duplicates/matcher.py` |
| Name comparison     | `ConvertTo-NameKey`, .NET ordinal comparers | `python/find_duplicates/names.py` |
| Excel output/input  | `Export-DuplicateReport`, `Import-DuplicateReport` and helpers | `python/find_duplicates/xlsx.py` |
| Duplicate folders   | `Find-DuplicateFolder`, `Get-FolderTree`, `Get-FolderSignature` | `python/find_duplicates/folders.py` |
| Validation, hash reuse | `Update-DuplicateReport`, `Invoke-CopyCheck`, `Test-DuplicateCopy`, `Test-DuplicateFolderCopy`, `Test-PathRootReachable`, `Test-SameSavedDate`, `Get-PreviousMd5` | `python/find_duplicates/validate.py` |
| Tests               | `tests/DuplicateFinder.Tests.ps1`    | `python/tests/test_find_duplicates.py`, `test_folders.py`, `test_unicode.py`, `test_edge_cases.py` |
| Cross-language test | —                                    | `python/tests/test_parity.py`          |

All PowerShell functions live in `src/DuplicateFinder.psm1`.

## Rule: PowerShell first, then Python parity

Every new feature, behaviour change or bug fix is done in this order:

1. Implement it in the **PowerShell** code and add Pester tests.
2. Then, in the same change, bring the **Python** port to feature parity:
   same behaviour, same defaults, same messages, same report layout, and
   mirrored unit tests (one Python test per Pester test, same names in
   snake_case).
3. If the change affects the report or which files match, extend the fixture
   in `python/tests/test_parity.py` so the cross-language test covers it.
4. Update both `README.md` and `python/README.md`.

Command-line options map one to one: `-Path` ↔ `path`,
`-OutputFile` ↔ `output` / `-o`, `-ThrottleLimit` ↔ `-j` / `--throttle-limit`,
`-IncludeFolders` ↔ `--folders`, `-IgnoreEmptyFiles` ↔ `--ignore-empty-files`,
`-Rehash` ↔ `--rehash`,
`-SkipCloudOnly` ↔ `--skip-cloud-only`, `-Validate` ↔ `--validate`,
`-WhatIf` ↔ `--dry-run`, `-Verbose` ↔ `--verbose`. `-PassThru` corresponds to
calling `find_duplicate_files()` / `validate_report()` from Python. Console
messages are worded identically apart from the option names they mention.

Keep orderings identical: folders and files are visited in ordinal name order,
locations are sorted ordinal-ignore-case, rows by file name
(ordinal-ignore-case), saved date, then MD5. Excel dates are truncated to the
millisecond when written and rounded to it when read, as .NET does.

Names are matched in NFC form, ignoring case (`ConvertTo-NameKey` / `name_key`),
and ordered by UTF-16 code units: `String.CompareOrdinal` on upper-cased text
(`Compare-IgnoringCase`, `$script:ByPathIgnoringCase` / `sort_key`,
`path_sort_key`). Never order with `StringComparer.OrdinalIgnoreCase`: .NET
Framework (Windows PowerShell 5.1) and .NET (PowerShell 7) order characters
beyond U+FFFF differently with it. The text of the Rules sheet
(`$script:RulesIntro` / `RULES_INTRO`, `$script:FileRules` / `FILE_RULES`,
`$script:FolderRules` / `FOLDER_RULES` and their titles) must be identical in both
languages; the parity test compares every sheet.

PowerShell source files must stay ASCII: Windows PowerShell 5.1 reads BOM-less
files in the ANSI code page. Build non-ASCII test data from code points.

PowerShell pitfalls this code base has hit: never assign a collection with
`$x = if (...) { ... }` (an empty or one-item array is unrolled); return
enumerable objects (XmlDocument, namespace managers, lists) with `, $x`;
pass `-WhatIf`/`-Verbose` explicitly into module functions; keep
`Array.Sort` calls on their non-generic overloads, by casting the arguments to
`[System.Array]` and `[System.Collections.IComparer]` (otherwise PowerShell
may pick the generic overload and sort a copy of the items); remember that
`,` binds tighter than `+` (`@('/' + $a, $b)` is `'/' + ($a, $b)`); PowerShell turns a
.NET property whose getter throws into `$null` (so read file details that may fail
with getter methods, e.g. `$file.get_Length()`, inside `try`).

Parallel work in PowerShell goes through `Start-WorkerPool` / `Receive-WorkerResult` /
`Stop-WorkerPool` (runspaces that load this module); in Python through
`ThreadPoolExecutor`. `-ThrottleLimit` / `-j` sets both folder listing and hashing.

Code that runs per file, per cell or per row (scanning, hashing, the Excel
writer and reader) avoids pipelines, script block comparers and advanced
function calls in the inner loop, and throttles `Write-Progress`: PowerShell's
per-call overhead dominates on large trees and reports.

Situations that need real equipment (network shares, cloud drives, Excel
itself) are listed in `tests/MANUAL-TESTS.md`; extend it when adding such a
feature.

## Constraints

- No runtime dependencies: PowerShell uses only .NET; Python uses only the
  standard library (Python 3.9+). No network access at run time.
- PowerShell must run on Windows PowerShell 5.1 and PowerShell 7+.

## Checks before every commit

```sh
# PowerShell (needs Pester 5+)
pwsh -NoProfile -Command "Invoke-Pester ./tests"
pwsh -NoProfile -Command "Invoke-ScriptAnalyzer -Path . -Recurse -Severity Error,Warning -ExcludeRule PSAvoidUsingWriteHost"

# Python (the parity test runs when pwsh is on the PATH)
cd python && python -m unittest discover -s tests -t .
```

Also review the diff, and scan for secrets, keys and tokens before pushing.
CI (`.github/workflows/tests.yml`) runs all of the above on Windows, Linux
and macOS.
