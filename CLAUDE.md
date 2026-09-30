# find-duplicates

Finds duplicate files (same **name**, **saved date** and **MD5**) in a folder
tree and saves them to an Excel workbook. The tool exists in two languages
that must behave identically:

| Part                | PowerShell (primary)                 | Python (port)                          |
|---------------------|--------------------------------------|----------------------------------------|
| Command line        | `Find-Duplicates.ps1`                | `python/find_duplicates/cli.py`        |
| Folder scanning     | `Get-FileInventory`, `Test-FolderLink`, `Test-CloudOnlyFile`, `Get-SortedByName` | `python/find_duplicates/scanner.py` |
| Duplicate matching  | `Find-DuplicateFile`, `Get-SortedFolder` | `python/find_duplicates/matcher.py` |
| Excel output        | `Export-DuplicateReport` and helpers | `python/find_duplicates/xlsx.py`       |
| Tests               | `tests/DuplicateFinder.Tests.ps1`    | `python/tests/test_find_duplicates.py` |
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
`-OutputFile` ↔ `output` / `-o`, `-SkipCloudOnly` ↔ `--skip-cloud-only`,
`-Verbose` ↔ `--verbose`. `-PassThru` corresponds to calling
`find_duplicate_files()` from Python.

Keep orderings identical: folders and files are visited in ordinal name order,
locations are sorted ordinal-ignore-case, rows by file name
(ordinal-ignore-case), saved date, then MD5.

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
