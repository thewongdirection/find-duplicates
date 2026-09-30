# find-duplicates

A PowerShell tool that scans a folder and all of its sub folders for duplicate
files and saves every match to an Excel workbook.

## What counts as a duplicate

Two files are duplicates only when **all three** of these match:

| Check      | How it is compared                                          |
|------------|-------------------------------------------------------------|
| File name  | Case-insensitive (`Photo.JPG` = `photo.jpg`)                |
| Saved date | Last-modified time, to the whole second                     |
| MD5        | Hash of the file contents                                   |

For speed, the MD5 is **only** calculated for files whose name and saved date
already match another file. Files that differ in size are skipped too, because
they cannot have the same MD5. Every other file is treated as unique without
being read.

## Usage

```powershell
# Scan the current folder, save to .\duplicates.xlsx
.\Find-Duplicates.ps1

# Scan a folder, save to .\duplicates.xlsx
.\Find-Duplicates.ps1 -Path D:\Photos

# Choose the report name (".xlsx" is added when missing)
.\Find-Duplicates.ps1 -Path D:\Photos -OutputFile C:\Reports\photo-dupes

# Print every folder as it is scanned, and return the results as objects
.\Find-Duplicates.ps1 -Path D:\Photos -Verbose -PassThru
```

While it runs, a progress bar shows the folder currently being scanned and,
afterwards, the file currently being hashed.

If Windows blocks the script, run it with
`powershell -ExecutionPolicy Bypass -File .\Find-Duplicates.ps1 ...`.

## The report

One row per duplicated file, one column per copy:

| File Name  | Last Modified       | Size (bytes) | MD5     | Copies | Location 1     | Location 2        | Location 3 |
|------------|---------------------|--------------|---------|--------|----------------|-------------------|------------|
| report.doc | 2024-05-17 10:30:00 | 48128        | 9A0F... | 3      | D:\Docs\2024   | D:\Backup\Docs    | E:\Old     |

Each `Location` column holds the full folder path of one copy. The header row
is frozen and filtered. Microsoft Excel does **not** need to be installed to
create the file.

## Notes

- Works with Windows PowerShell 5.1 and PowerShell 7+ (Windows, Linux, macOS).
- Folders that cannot be read are reported as warnings and skipped.
- Folder links (symlinks and junctions) are not followed, to avoid loops.
- A report saved inside the scanned folder is not counted as a file.

## Running the tests

```powershell
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser
Invoke-Pester ./tests
```

The tests also run on every push via GitHub Actions (Windows, Linux, macOS,
and Windows PowerShell 5.1).
