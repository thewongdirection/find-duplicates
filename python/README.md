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

Or install it once, from the repository root, to get a `find-duplicates` command
anywhere:

```sh
pip install ./python
find-duplicates [options]
```

## Deployment

Python 3.9 or later is all it needs: no packages, no internet access. Copy the
`find_duplicates` folder to the computer (or install it with `pip` as above), and
make sure the account that runs it can read the folders to scan and write the
report. Scheduled runs, network shares and cloud drives are covered in the
[main README](../README.md#deployment).

Built-in help: `python -m find_duplicates --help`.

## Command reference

```text
python -m find_duplicates [path] [output] [-o FILE] [-j N] [--folders] [--ignore-empty-files] [--exclude PATTERN ...]
                          [--minimum-size SIZE] [--skip-cloud-only] [--rehash] [--dry-run] [-v]
python -m find_duplicates --validate [report] [--dry-run] [-v]
```

### Scan (the default)

| Option | Default | What it does |
|---|---|---|
| `path` | current folder | Folder to scan, including all sub folders. |
| `output`, `-o FILE`, `--output-file FILE` | `duplicates.xlsx` in the current folder | Report to write. `.xlsx` is added when there is no extension. An existing report is replaced. |
| `-j N`, `--throttle-limit N` | `4` on a network share or drive, `1` otherwise | How many files to hash, and folders to list, at the same time (1-64); see [Threads](#threads). Try 4-8 for SSDs, network shares and cloud folders; keep 1 for a single spinning hard disk. |
| `--folders` | off | Also find [duplicate folders](../README.md#duplicate-folders) and save them on the *Duplicate Folders* sheet. |
| `--ignore-empty-files` | off | Leave files of 0 bytes out of the duplicate files. |
| `--exclude PATTERN` | none | Leave files and folders with this name out; repeat for more names. `*` stands for any characters and `?` for any one character; upper/lower case is ignored; patterns match names, not paths. Left-out folders are not scanned, and duplicate folders are compared as if left-out names were not there. Recorded on the *Rules* sheet, so `--validate` leaves the same names out. |
| `--minimum-size SIZE` | `0` | Leave files smaller than this out of the duplicate files: a number of bytes, optionally followed by `KB`, `MB`, `GB`, `TB` or `PB` (1024-based, so `1.5MB` is 1572864 bytes, as in PowerShell). Recorded on the *Rules* sheet. |
| `--rehash` | off | Read every candidate file again. Without it, when the report already exists (from an earlier scan), files it lists whose size and saved date have not changed keep the MD5 recorded there instead of being read again. |
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

# Also find duplicate folders (second sheet)
python -m find_duplicates D:\Backups --folders

# Leave out empty (0-byte) files
python -m find_duplicates D:\Photos --ignore-empty-files

# Leave out thumbnail caches, Git folders, temporary files and files under 100 KB
python -m find_duplicates D:\Photos --exclude Thumbs.db --exclude .git --exclude "*.tmp" --minimum-size 100KB

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
| `-j N`, `--throttle-limit N` | `1` | How many folders to check at the same time. Try 4-8 for network shares. |
| `--dry-run` | off | Report what would be removed, but do not change the report. |
| `-v`, `--verbose` | off | Print every copy that is removed and why. |

`--skip-cloud-only`, `--folders`, `--ignore-empty-files`, `--rehash`, `--exclude` and `--minimum-size`
only apply to a scan; validation reads the names to leave out from the report.
Duplicate folders, when the report has them, are re-checked too.

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
    iter_files, find_duplicate_files, find_duplicate_folders, export_duplicate_report,
    read_duplicate_report, read_duplicate_folder_report, validate_report,
)

records = []                                                     # folder records, for folders
files = list(iter_files(r"D:\Photos", folders=records))          # every file, recursively
cache = {}                                                       # share hashes between the two
duplicates = find_duplicate_files(files, throttle_limit=4, md5_cache=cache)   # like -PassThru
folders = find_duplicate_folders(files, records, md5_cache=cache)
for dup in duplicates:
    print(dup.file_name, dup.count, dup.folders)
export_duplicate_report(duplicates, "duplicates.xlsx", folders)  # omit folders: no folder sheet

rows = read_duplicate_report("duplicates.xlsx")                  # read the file rows back
folder_rows = read_duplicate_folder_report("duplicates.xlsx")    # read the folder rows back
result = validate_report("duplicates.xlsx", dry_run=True)        # summary: removed, kept...
```

## PowerShell ↔ Python

| PowerShell | Python |
|---|---|
| `-Path` | `path` (first argument) |
| `-OutputFile` | `output` (second argument) or `-o` |
| `-ThrottleLimit N` | `-j N` / `--throttle-limit N` |
| `-IncludeFolders` | `--folders` |
| `-IgnoreEmptyFiles` | `--ignore-empty-files` |
| `-Rehash` | `--rehash` |
| `-Exclude a, b` | `--exclude a --exclude b` |
| `-MinimumSize N` | `--minimum-size N` |
| `-SkipCloudOnly` | `--skip-cloud-only` |
| `-Validate` | `--validate` |
| `-WhatIf` | `--dry-run` |
| `-Verbose` | `-v` / `--verbose` |
| `-PassThru` | `find_duplicate_files()` / `validate_report()` |
| `Get-FileInventory` | `iter_files()` |
| `Find-DuplicateFile` | `find_duplicate_files()` |
| `Find-DuplicateFolder` | `find_duplicate_folders()` |
| `Export-DuplicateReport` | `export_duplicate_report()` |
| `Import-DuplicateReport` | `read_duplicate_report()` |
| `Import-DuplicateFolderReport` | `read_duplicate_folder_report()` |
| `Update-DuplicateReport` | `validate_report()` |
| `Get-PreviousMd5` | `previous_md5()` |

## Unicode

File and folder names in any language work the same as in PowerShell: they are
compared ignoring case and Unicode form (NFC), and sorted in the same (UTF-16)
order as .NET, so both tools produce identical reports. If the console cannot
show a character (for example output redirected to a file in a Windows code
page), it is printed as an escape such as `\u65e5` instead of stopping the run.

## Threads

`-j` works like PowerShell's `-ThrottleLimit`, with one difference: Python uses
its threads only where they are faster, for folders on a network drive (UNC
paths and network drives on Windows; NFS, SMB and other network file systems on
Linux) and for hashing files of 1 MB or more. Everything else is done one at a
time: for the many small operations of listing and checking local folders,
threads contending for Python's global lock make it many times slower. Reports
are identical either way. Without `-j`, a scan of a folder on a network drive
uses 4 threads, as PowerShell does.

The other speed-ups in the [main README](../README.md#performance) apply too:
large files are first compared by their first 1 MB, and duplicate folders are
only looked for among folders whose name another folder has.

## Limits and edge cases

The same as for PowerShell: see
[Limits and edge cases](../README.md#limits-and-edge-cases). Specific to Python:

- On Windows, paths longer than 260 characters need the system's *long paths*
  setting (`LongPathsEnabled`); without it such folders are skipped with a
  warning.
- On Windows, saved dates before 1970 are converted to local time with the rule
  for the same date in 1972 (Python's own conversion cannot handle them there),
  which is also what PowerShell does for dates older than Windows' time zone data.

## Running the tests

```sh
cd python
python -m unittest discover -s tests -t .
```

`tests/test_parity.py` runs both tools on the same folder tree (plain scan,
parallel hashing, and validation, all with duplicate folders) and requires
identical reports, sheet by sheet. It needs
PowerShell 7 (`pwsh`) and is skipped without it; CI always runs it.

`tests/test_edge_cases.py` mirrors the PowerShell *Edge cases* tests. The
LibreOffice round trip needs LibreOffice Calc (`soffice`) and is skipped
without it, unless `FIND_DUPLICATES_REQUIRE_LIBREOFFICE` is set (as in CI).
