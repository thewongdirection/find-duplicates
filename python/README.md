# find-duplicates (Python)

The Python version of the PowerShell tool in the repository root. It has the
same features, defaults and report layout, and a test checks that both
produce the same spreadsheet. It uses only the standard library (Python 3.9+),
so there is nothing to install.

## Usage

Run from this `python` folder:

```sh
# Scan the current folder, save to ./duplicates.xlsx
python -m find_duplicates

# Scan a folder, save to ./duplicates.xlsx
python -m find_duplicates D:\Photos

# Choose the report name (".xlsx" is added when missing)
python -m find_duplicates D:\Photos C:\Reports\photo-dupes
python -m find_duplicates D:\Photos -o C:\Reports\photo-dupes

# Network share / cloud folder; never download online-only cloud files
python -m find_duplicates \\nas\photos
python -m find_duplicates "%OneDrive%" --skip-cloud-only

# Print every folder as it is scanned
python -m find_duplicates D:\Photos --verbose
```

Or install it to get a `find-duplicates` command anywhere:

```sh
pip install ./python
find-duplicates D:\Photos
```

While it runs, a status line shows the folder currently being scanned and,
afterwards, the file currently being hashed.

## What counts as a duplicate

Exactly as in the PowerShell tool: file name (case-insensitive), saved date
(to the whole second) and MD5 must all match. MD5 is only calculated for files
that already match on name, saved date and size. See the
[main README](../README.md) for the report layout and network and cloud drive
behaviour.

## Using it from Python

```python
from find_duplicates import iter_files, find_duplicate_files, export_duplicate_report

files = list(iter_files(r"D:\Photos"))
duplicates = find_duplicate_files(files)          # like -PassThru
for dup in duplicates:
    print(dup.file_name, dup.count, dup.folders)
export_duplicate_report(duplicates, "duplicates.xlsx")
```

## PowerShell ↔ Python

| PowerShell         | Python                         |
|--------------------|--------------------------------|
| `-Path`            | `path` (first argument)        |
| `-OutputFile`      | `output` (second argument) or `-o` |
| `-SkipCloudOnly`   | `--skip-cloud-only`            |
| `-Verbose`         | `--verbose`                    |
| `-PassThru`        | `find_duplicate_files()`       |

## Running the tests

```sh
cd python
python -m unittest discover -s tests -t .
```

`tests/test_parity.py` runs both tools on the same folder tree and requires
identical reports. It needs PowerShell 7 (`pwsh`) and is skipped without it;
CI always runs it.
