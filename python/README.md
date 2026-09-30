# find-duplicates (Python)

The Python version of the PowerShell tool in the repository root. It has the
same features, defaults and report layout, and an automated test checks that
both produce the same spreadsheet. It uses only the standard library
(Python 3.9+), so there is nothing to install.

See the [main README](../README.md) for the full list of capabilities, what
counts as a duplicate, the report layout, network and cloud drive behaviour,
and performance advice. Everything there applies here; only the command
syntax differs.

## Running it

From this `python` folder:

```sh
python -m find_duplicates [options]
```

Or install it once to get a `find-duplicates` command anywhere:

```sh
pip install ./python
find-duplicates [options]
```

Built-in help: `python -m find_duplicates --help`.

## Command reference

```text
python -m find_duplicates [path] [output] [-o FILE] [-j N] [--skip-cloud-only] [--dry-run] [-v]
python -m find_duplicates --validate [report] [--dry-run] [-v]
```

### Scan (the default)

| Option | Default | What it does |
|---|---|---|
| `path` | current folder | Folder to scan, including all sub folders. |
| `output`, `-o FILE`, `--output-file FILE` | `duplicates.xlsx` in the current folder | Report to write. `.xlsx` is added when there is no extension. An existing report is replaced. |
| `-j N`, `--throttle-limit N` | `1` | How many files to hash at the same time (1-64). Try 4-8 for SSDs, network shares and cloud folders; keep 1 for a single spinning hard disk. |
| `--skip-cloud-only` | off | Never download online-only cloud files to hash them. Duplicates among such files are then not reported. |
| `--dry-run` | off | Scan and report the totals, but do not save the report. |
| `-v`, `--verbose` | off | Print every folder as it is scanned, and every link or online-only file skipped. |

```sh
# Scan the current folder, save ./duplicates.xlsx
python -m find_duplicates

# Scan a folder, save ./duplicates.xlsx
python -m find_duplicates D:\Photos

# Choose where to save the report (".xlsx" is added when missing)
python -m find_duplicates D:\Photos C:\Reports\photo-dupes
python -m find_duplicates D:\Photos -o C:\Reports\photo-dupes.xlsx

# Hash 8 files at a time (SSD, network share or cloud folder)
python -m find_duplicates \\nas\photos -j 8

# OneDrive without downloading online-only files
python -m find_duplicates "%OneDrive%" --skip-cloud-only

# See the totals without writing a report
python -m find_duplicates D:\Photos --dry-run

# Log every folder scanned
python -m find_duplicates D:\Photos --verbose
```

### Validate an existing report (`--validate`)

Re-checks every copy listed in a report without rescanning, and updates the
report in place: copies that are missing, or whose size or saved date changed,
are removed; rows left with fewer than two copies are removed; copies on a
drive or share that cannot be reached are kept. See the
[main README](../README.md#validate-an-existing-report--validate) for details.

| Option | Default | What it does |
|---|---|---|
| `--validate` | — | Switches to validation. |
| `report`, `-o FILE` | `duplicates.xlsx` in the current folder | Report to check and update. |
| `--dry-run` | off | Report what would be removed, but do not change the report. |
| `-v`, `--verbose` | off | Print every copy that is removed and why. |

`--skip-cloud-only` and `--throttle-limit` only apply to a scan.

```sh
# Check ./duplicates.xlsx
python -m find_duplicates --validate

# Check a named report
python -m find_duplicates --validate C:\Reports\photo-dupes.xlsx

# Preview what would be removed, listing each copy
python -m find_duplicates --validate C:\Reports\photo-dupes.xlsx --dry-run --verbose
```

## Using it from Python

```python
from find_duplicates import (
    iter_files, find_duplicate_files, export_duplicate_report,
    read_duplicate_report, validate_report,
)

files = list(iter_files(r"D:\Photos"))                           # every file, recursively
duplicates = find_duplicate_files(files, throttle_limit=4)       # like -PassThru
for dup in duplicates:
    print(dup.file_name, dup.count, dup.folders)
export_duplicate_report(duplicates, "duplicates.xlsx")

rows = read_duplicate_report("duplicates.xlsx")                  # read a report back
result = validate_report("duplicates.xlsx", dry_run=True)        # summary: removed, kept...
```

## PowerShell ↔ Python

| PowerShell | Python |
|---|---|
| `-Path` | `path` (first argument) |
| `-OutputFile` | `output` (second argument) or `-o` |
| `-ThrottleLimit N` | `-j N` / `--throttle-limit N` |
| `-SkipCloudOnly` | `--skip-cloud-only` |
| `-Validate` | `--validate` |
| `-WhatIf` | `--dry-run` |
| `-Verbose` | `-v` / `--verbose` |
| `-PassThru` | `find_duplicate_files()` / `validate_report()` |
| `Get-FileInventory` | `iter_files()` |
| `Find-DuplicateFile` | `find_duplicate_files()` |
| `Export-DuplicateReport` | `export_duplicate_report()` |
| `Import-DuplicateReport` | `read_duplicate_report()` |
| `Update-DuplicateReport` | `validate_report()` |

## Running the tests

```sh
cd python
python -m unittest discover -s tests -t .
```

`tests/test_parity.py` runs both tools on the same folder tree (plain scan,
parallel hashing, and validation) and requires identical reports. It needs
PowerShell 7 (`pwsh`) and is skipped without it; CI always runs it.
