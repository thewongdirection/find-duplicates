# find-duplicates

Finds duplicate files in a folder and all of its sub folders and records every
copy in an Excel workbook. It can later re-check that workbook against the disk
without rescanning.

The tool comes in two versions with the same features and the same output: a
**PowerShell** script (this folder) and a **Python** package
([`python/`](python/README.md)).

## Contents

- [Capabilities](#capabilities)
- [Deployment](#deployment)
- [What counts as a duplicate](#what-counts-as-a-duplicate)
- [Duplicate folders](#duplicate-folders)
- [Command reference](#command-reference)
- [The report](#the-report)
- [Network folders and cloud drives](#network-folders-and-cloud-drives)
- [Limits and edge cases](#limits-and-edge-cases)
- [Performance](#performance)
- [Using the functions directly](#using-the-functions-directly)
- [Running the tests](#running-the-tests)

## Capabilities

| Capability | Details |
|---|---|
| Recursive scan | Scans a folder and every sub folder; shows the folder currently being scanned. |
| Strict duplicate rule | Name **and** saved date **and** MD5 must all match. |
| Fast by design | MD5 is only calculated for files whose name, saved date and size already match another file. Everything else is never read. |
| Parallel scanning and hashing | `-ThrottleLimit N` lists up to N folders and hashes up to N files at the same time (1-64). |
| Every copy recorded | One row per duplicated file, one column per copy, with the full folder path of each. |
| Duplicate folders | `-IncludeFolders` also finds whole folders with the same name and identical contents (every file and sub folder), on a second sheet. |
| Excel output without Excel | Writes a real `.xlsx`: data sheets with a frozen, filterable header and real dates, plus a *Rules* sheet stating in plain words what counts as a match. Excel does not need to be installed. |
| Empty files optional | `-IgnoreEmptyFiles` leaves files of 0 bytes out of the file duplicates. |
| Leave things out | `-Exclude` skips files and folders by name (`Thumbs.db`, `.git`, `*.tmp`); `-MinimumSize` leaves small files out. Both are recorded in the report. |
| Any language | File and folder names in any script (Chinese, Arabic, Cyrillic, emoji ...) are matched and stored correctly; a name typed on a Mac matches the same name saved on Windows. |
| Validate without rescanning | `-Validate` re-checks every copy listed in an existing report (files with a quick lookup, folders by re-listing them, never reading contents) and removes those that are gone or changed. |
| Network shares | UNC paths (`\\server\share`) and mapped drives. Unreachable folders are skipped with a warning; an unreachable share never wipes report entries. |
| Cloud drives | OneDrive, Google Drive, Dropbox, iCloud, Box. Online-only files are downloaded only if they must be hashed, or never with `-SkipCloudOnly`. |
| Safe on any tree | Symbolic links and junctions are not followed (no loops); unreadable folders and files are warnings, not failures. |
| Preview | `-WhatIf` shows what would be saved or removed without touching the report. |
| No dependencies | PowerShell: Windows PowerShell 5.1 or PowerShell 7+ on Windows, Linux or macOS. Python: 3.9+, standard library only. No internet access needed. |

## Deployment

There is nothing to install beyond a PowerShell or Python that most computers
already have, and nothing runs as a service. Use either version; both write the
same report.

Get the files with `git clone https://github.com/thewongdirection/find-duplicates.git`
or, on GitHub, *Code* > *Download ZIP*.

### PowerShell version

**Needs:** Windows PowerShell 5.1 (part of Windows 10 and 11) or PowerShell 7+
on Windows, Linux or macOS. No Excel, no modules, no internet access.

**Copy these three files into one folder** (any folder):

```text
Find-Duplicates.ps1
DuplicateFinder.psm1
DuplicateFinder.cs
```

The script loads the module from its own folder, and the module compiles
`DuplicateFinder.cs` (its fast per-file code) the first time it is loaded in a
PowerShell session, which takes about half a second.

**On Windows:**

- **Blocked script.** Files downloaded from the internet are marked as such, and
  Windows may refuse to run them. Unblock them once:

  ```powershell
  Get-ChildItem -File .\find-duplicates | Unblock-File
  ```

  or start the script with the policy relaxed for that run only:
  `powershell -ExecutionPolicy Bypass -File .\Find-Duplicates.ps1 D:\Photos`.
- **`AllSigned` policy.** Where only signed scripts may run, sign
  `Find-Duplicates.ps1` and `DuplicateFinder.psm1` with your organisation's
  code-signing certificate (`Set-AuthenticodeSignature`).
- **Locked-down computers.** Where AppLocker or Windows Defender Application
  Control puts PowerShell in Constrained Language Mode, the module cannot
  compile its helpers and will not load; use the Python version there.
- **Temp folder.** Windows PowerShell 5.1 compiles in `%TEMP%`, which must be
  writable.

**Check it works:** `.\Find-Duplicates.ps1 -Path . -WhatIf` scans the current
folder and prints the totals without writing a report.

### Python version

**Needs:** Python 3.9 or later, standard library only. See
[python/README.md](python/README.md) for its options.

- **Without installing:** copy the `python/find_duplicates` folder and run
  `python -m find_duplicates D:\Photos` from the folder that contains it.
- **Installed:** from the repository root, `pip install ./python` adds a
  `find-duplicates` command; `pip uninstall find-duplicates` removes it. (pip
  fetches `setuptools` to build it; offline, copy the folder instead.)

### Access it needs

- **Read** access to every folder to scan (folders it cannot read are skipped
  with a warning), and **write** access to the folder of the report.
- **Network shares:** a UNC path (`\\server\share`) or a mapped drive, reachable
  with the account that runs the tool.
- **Cloud drives** (OneDrive, Google Drive, Dropbox ...): the provider's sync
  client. Add `-SkipCloudOnly` (`--skip-cloud-only`) to avoid downloading
  online-only files.
- **The report** opens in Excel, LibreOffice Calc or Google Sheets; none is
  needed to create it.

### Running it on a schedule

A scheduled scan reuses the previous report's hashes, so it only reads files
that are new or changed. On Windows, with Task Scheduler:

```powershell
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument '-NoProfile -ExecutionPolicy Bypass -File C:\Tools\find-duplicates\Find-Duplicates.ps1 -Path D:\Photos -OutputFile D:\Reports\photos.xlsx'
Register-ScheduledTask -TaskName 'Find duplicate photos' -Action $action `
    -Trigger (New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 3am)
```

On Linux or macOS, with cron (`crontab -e`):

```text
0 3 * * 0  cd /opt/find-duplicates/python && python3 -m find_duplicates /srv/photos /srv/reports/photos.xlsx
```

### Updating

Replace the files with the new version. Reports written by earlier versions are
still read, validated and reused.

## What counts as a duplicate

Two files are duplicates only when **all three** of these match:

| Check      | How it is compared                                          |
|------------|-------------------------------------------------------------|
| File name  | Case-insensitive (`Photo.JPG` = `photo.jpg`)                |
| Saved date | Last-modified time, to the whole second                     |
| MD5        | Hash of the file contents                                   |

The checks run from cheapest to most expensive. Files are grouped by name and
saved date; groups are split by size (files of different sizes cannot have the
same MD5); only files left in a group of two or more are read and hashed. A file
can have any number of duplicates, in any folders.

Files of 0 bytes all have the same contents, so every same-named empty file
with the same saved date is a duplicate. They are included by default; add
`-IgnoreEmptyFiles` to leave them out, or `-MinimumSize` to leave out every file
smaller than a given size. (Duplicate folders always compare every file, empty
and small ones included.)

`-Exclude` leaves files and folders out by name, as if they were not there:
left-out folders are not scanned at all, and folders are compared without the
left-out names, so two photo folders still match when only their `Thumbs.db`
differs. Patterns use `*` for any characters and `?` for any one character, and
ignore upper/lower case, like names do. They match whole names, not paths: a
pattern holding `/` or `\` is refused, and the folder being scanned is never left
out itself. Everything else in a pattern is literal. `-Validate` leaves out the
same names, read from the report. The smallest file size (`-MinimumSize`, or 1
byte with `-IgnoreEmptyFiles`) is recorded in the report too.

Names are compared the way people read them: ignoring upper/lower case, and
ignoring how accented letters are encoded. macOS often stores "é" as "e" plus a
separate accent, where Windows stores one character; both count as the same
name. Otherwise names must match exactly, in any language or script.

## Duplicate folders

With `-IncludeFolders`, whole folders are compared too. Two folders are
duplicates when:

- their **names** match (ignoring case), and
- they contain the **same tree** of file and sub folder names (empty sub
  folders included), and
- **every file** is a duplicate, by the rule above (name, saved date, MD5), of
  the file at the same place in the other folder.

Their total sizes therefore match too. The checks again run cheapest first:
folder trees are compared by names, sizes and saved dates from the scan
itself, and only files inside folders that still match are hashed. Hashes
already calculated for the file duplicates are reused, so files are not read again.

- **Top-most only.** When `D:\Photos` and `E:\Backup\Photos` are duplicates,
  their matching sub folders (`2024`, `2024\Jan`, ...) are not listed again. A
  nested set *is* listed when at least one of its copies is outside a
  duplicate folder (e.g. a third copy of `2024` somewhere else).
- A folder with an unreadable sub folder, a folder link (symlink or
  junction, which is not followed), the report file itself, or (with
  `-SkipCloudOnly`) an online-only file is never reported: a match must be
  proven.
- Folders with no files anywhere inside them are not reported.
- The file rows are unaffected: files inside duplicate folders are still
  listed on the *Duplicates* sheet.

Duplicate folders go on a second sheet, **Duplicate Folders**, one row per
duplicated folder:

| Folder Name | Files | Sub Folders | Size (bytes) | Copies | Location 1 | Location 2 |
|-------------|-------|-------------|--------------|--------|------------|------------|
| Photos      | 1250  | 14          | 5368709120   | 2      | D:\Photos  | E:\Backup\Photos |

Here each `Location` is the full path of the duplicate folder itself, from the least
nested copy (*Location 1*) to the most nested, as on the *Duplicates* sheet.

## Command reference

```text
Find-Duplicates.ps1 [[-Path] <folder>] [[-OutputFile] <report>] [-ThrottleLimit <1-64>]
                    [-IncludeFolders] [-IgnoreEmptyFiles] [-Exclude <pattern[]>] [-MinimumSize <bytes>]
                    [-SkipCloudOnly] [-Rehash] [-PassThru] [-WhatIf] [-Verbose]

Find-Duplicates.ps1 -Validate [[-OutputFile] <report>] [-ThrottleLimit <1-64>] [-PassThru] [-WhatIf] [-Verbose]
```

Built-in help: `Get-Help .\Find-Duplicates.ps1 -Full`.

If Windows blocks the script, run it as
`powershell -ExecutionPolicy Bypass -File .\Find-Duplicates.ps1 ...`.

### Scan (the default)

| Parameter | Default | What it does |
|---|---|---|
| `-Path <folder>` | current folder | Folder to scan, including all sub folders. Also the first positional argument. |
| `-OutputFile <report>` | `duplicates.xlsx` in the current folder | Report to write. `.xlsx` is added when there is no extension. An existing report is replaced, after its MD5 hashes are reused (see `-Rehash`). Also the second positional argument. |
| `-ThrottleLimit <1-64>` | `4` on a network share or drive, `1` otherwise | How many files to hash, and folders to list, at the same time. See [Performance](#performance). |
| `-IncludeFolders` | off | Also find [duplicate folders](#duplicate-folders) and save them on the *Duplicate Folders* sheet. |
| `-IgnoreEmptyFiles` | off | Leave files of 0 bytes out of the duplicate files. |
| `-Exclude <pattern[]>` | none | Leave files and folders with these names out: see [What counts as a duplicate](#what-counts-as-a-duplicate). Recorded on the *Rules* sheet. |
| `-MinimumSize <bytes>` | `0` | Leave files smaller than this out of the duplicate files. Takes PowerShell sizes such as `100KB` or `1.5MB`. Recorded on the *Rules* sheet. |
| `-SkipCloudOnly` | off | Never download online-only cloud files to hash them. Duplicates among such files are then not reported. |
| `-Rehash` | off | Read every candidate file again. Without it, when the report already exists (from an earlier scan), files it lists whose size and saved date have not changed keep the MD5 recorded there instead of being read again. |
| `-PassThru` | off | Also return the duplicates as PowerShell objects (for piping or scripting): file sets, then folder sets (which have a `FolderName` property). |
| `-WhatIf` | off | Scan and report the totals, but do not save the report. |
| `-Verbose` | off | Print every folder as it is scanned, and every link or online-only file skipped. |

```powershell
# Scan the current folder, save .\duplicates.xlsx
.\Find-Duplicates.ps1

# Scan a folder, save .\duplicates.xlsx
.\Find-Duplicates.ps1 -Path D:\Photos
.\Find-Duplicates.ps1 D:\Photos                                   # same, positional

# Choose where to save the report (".xlsx" is added when missing)
.\Find-Duplicates.ps1 -Path D:\Photos -OutputFile C:\Reports\photo-dupes
.\Find-Duplicates.ps1 D:\Photos C:\Reports\photo-dupes.xlsx       # same, positional

# Hash 8 files at a time (SSD, network share or cloud folder)
.\Find-Duplicates.ps1 -Path \\nas\photos -ThrottleLimit 8

# Also find duplicate folders (second sheet)
.\Find-Duplicates.ps1 -Path D:\Backups -IncludeFolders

# Leave out empty (0-byte) files
.\Find-Duplicates.ps1 -Path D:\Photos -IgnoreEmptyFiles

# Leave out thumbnail caches, Git folders, temporary files and files under 100 KB
.\Find-Duplicates.ps1 -Path D:\Photos -Exclude Thumbs.db, .git, *.tmp -MinimumSize 100KB

# OneDrive without downloading online-only files
.\Find-Duplicates.ps1 -Path "$env:OneDrive" -SkipCloudOnly

# See the totals without writing a report
.\Find-Duplicates.ps1 -Path D:\Photos -WhatIf

# Log every folder scanned
.\Find-Duplicates.ps1 -Path D:\Photos -Verbose

# Work with the results in PowerShell
.\Find-Duplicates.ps1 -Path D:\Photos -PassThru | Sort-Object Count -Descending | Select-Object -First 10
```

### Validate an existing report (`-Validate`)

Re-checks every copy listed in a report without rescanning, and updates the
report in place. Much faster than a full scan when you have been deleting
duplicates and want the report to catch up.

| Parameter | Default | What it does |
|---|---|---|
| `-Validate` | — | Switches to validation. Required for this mode. |
| `-OutputFile <report>` | `duplicates.xlsx` in the current folder | Report to check and update. Also the first positional argument in this mode. |
| `-ThrottleLimit <1-64>` | `1` | How many folders to check at the same time. Try 4-8 for network shares. |
| `-PassThru` | off | Also return the rows that remain. |
| `-WhatIf` | off | Report what would be removed, but do not change the report. |
| `-Verbose` | off | Print every copy that is removed and why (`Missing` or `Changed`). |

The report can be open in Excel while it is validated with `-WhatIf`; to save
changes, close it first.

For every file copy listed, a file lookup decides (copies in the same folder
are all checked from one listing of it, one round trip on a network share); for
every folder copy, the folder is listed again. No file contents are read. Each
drive or share that cannot be reached is tried once:

| Result | File copy | Folder copy | Action |
|---|---|---|---|
| Present | Still there, same size and saved date | Still there, same number of files and sub folders, same total size | Kept |
| Missing | No longer in that folder | No longer there | Removed |
| Changed | Size or saved date changed | Files, sub folders or total size changed | Removed |
| Unavailable | Its drive or network share cannot be reached | Same, or part of it cannot be read | **Kept**, with a warning |

Each drive or network share is checked once per run, so an offline share
costs one timeout, not one per copy. A report without a *Duplicate Folders*
sheet stays without one.

Rows left with fewer than two copies are removed. The report is only rewritten
when something changed. File contents are never read (so nothing is
downloaded from the cloud); a file edited without changing its size or saved
date is not detected. Run a full scan for that, and to find new duplicates.
File names containing control characters (possible on Linux) are stored with
a replacement character, so validation cannot find those copies and removes them.
Saved dates are compared as moments in time (using the report's *UTC Offset*
column), so a report can be validated on a computer in another time zone.
Reports made before that column existed are compared by local time: validate
those in the time zone of the scan. Rewriting such a report adds the column.

```powershell
# Check .\duplicates.xlsx
.\Find-Duplicates.ps1 -Validate

# Check a named report
.\Find-Duplicates.ps1 -Validate -OutputFile C:\Reports\photo-dupes.xlsx
.\Find-Duplicates.ps1 -Validate C:\Reports\photo-dupes.xlsx       # same, positional

# Preview what would be removed, listing each copy
.\Find-Duplicates.ps1 -Validate C:\Reports\photo-dupes.xlsx -WhatIf -Verbose
```

The report can have been opened and saved in Excel in the meantime; extra
formatting is dropped when it is rewritten.

## The report

The workbook has these sheets:

| Sheet | Contents |
|---|---|
| **Duplicates** | One row per duplicated file (below). |
| **Duplicate Folders** | Only with `-IncludeFolders`: one row per duplicated folder (see [Duplicate folders](#duplicate-folders)). |
| **Rules** | The matching rules behind the other sheets, in plain words, so anyone reviewing the data can check what a match means. Rewritten with the data, so it always matches the sheets present. When the scan used `-Exclude` or `-MinimumSize`, a *Scan settings* section records them; `-Validate` reads it back and keeps it. |

On the *Duplicates* sheet: one row per duplicated file, one column per copy:

| File Name  | Last Modified       | UTC Offset | Size (bytes) | MD5     | Copies | Location 1     | Location 2        | Location 3 |
|------------|---------------------|------------|--------------|---------|--------|----------------|-------------------|------------|
| report.doc | 2024-05-17 10:30:00 | +10:00     | 48128        | 9A0F... | 3      | E:\Old         | D:\Backup\Docs    | D:\Docs\2024 |

- Each `Location` column holds the full folder path of one copy; there are as
  many columns as the file with the most copies needs. *Location 1* is the least
  nested copy (fewest folders deep) and the last one the most nested, so the
  copy furthest right is usually the one to delete; copies equally deep are in
  alphabetical order.
- The header row is frozen and has filters; *Last Modified* is a real Excel
  date; *Size* and *Copies* are numbers.
- *Last Modified* is local time on the computer that ran the scan, and *UTC
  Offset* its difference from UTC at that date (daylight saving included), so
  the moment each file was saved is known in any time zone.
- Rows are sorted by file name, then saved date, then MD5; locations from the
  least to the most nested, as above. Where copies' names differ only in case, the row shows the name of
  the copy in the alphabetically first folder.
- A report saved inside the scanned folder is not counted as a file.

## Network folders and cloud drives

The tool needs no internet or network access of its own: it only reads the
folders you point it at, wherever they live.

- **Network shares**: use a UNC path (`\\server\share\folder`) or a mapped
  drive (`Z:\`). Folders that become unreachable mid-scan are reported as
  warnings and skipped; the rest of the scan carries on. When validating,
  copies on a share or drive that cannot be reached are kept, never removed.
- **OneDrive, Google Drive, Dropbox, iCloud, Box**: point it at the synced
  folder or drive letter, e.g. `-Path "$env:OneDrive"` or `-Path G:\`.
  Listing folders never downloads anything. Files that are only stored online
  are downloaded **only** when they have to be hashed, i.e. when another file
  already has the same name, saved date and size. Add `-SkipCloudOnly` to
  never download them. Validation never downloads anything.
- Folder links (symlinks and junctions) are not followed, to avoid loops.
  Cloud-synced folders are followed.

## Limits and edge cases

Each of these is covered by an automated test on Windows, Linux and macOS
(see [Running the tests](#running-the-tests)).

| Situation | What happens |
|---|---|
| Paths longer than 260 characters | Scanned and validated normally by PowerShell 7 and Python. Windows PowerShell 5.1 may be unable to reach them; it then says so in a warning and carries on. Python on Windows needs the system's *long paths* setting (`LongPathsEnabled`), which current Windows versions usually have on. |
| Files larger than 4 GB | Fully supported, including their sizes in the report. |
| Hidden and system files and folders | Included, like any other file. |
| Saved dates before 1970 or after 2038 | Fully supported. |
| Copies on FAT/exFAT drives (USB sticks, memory cards) | These store saved times in 2-second steps, so a copy of a file saved at 10:30:01 shows 10:30:02 there. Such copies are **not** matched: the saved date must agree to the second. |
| Daylight saving time | Makes no difference: dates are compared as instants. |
| Changing the computer's time zone between the scan and `-Validate` | Makes no difference: the *UTC Offset* column lets validation compare instants. (Reports made before that column are compared by local time; validate those in the time zone of the scan.) |
| A folder that is deleted, or cannot be read, during the scan | Reported as a warning and skipped; the scan carries on. |
| A file that cannot be read (permissions, locked by another program) | Reported as a warning and left out; its other copies are still matched. |
| The scanned folder is itself a symbolic link | Scanned through the link; locations show the link's path. |
| Two files in one folder whose names differ only in case (Linux, or case-sensitive folders on Windows) | Reported as duplicates; the row lists that folder twice. |
| Names Windows reserves or trims (`NUL`, `PRN.txt`, names ending in a dot or space) | Ordinary names on Linux and macOS. On Windows such files can only be made by special tools; the scan never fails on them, but may skip them with a warning, and `-Validate` keeps such copies rather than guess. |
| Excel limits | A file with more than 16,378 copies, more than 1,048,575 duplicated files, or a cell longer than 32,767 characters stops the tool with an error, rather than writing a report Excel would reject or cut short. |
| Reports opened and saved in Excel or LibreOffice | Still read and validated (shared strings and re-numbered sheet parts are handled). |

Some situations need real equipment and are checked by hand before a release:
see [tests/MANUAL-TESTS.md](tests/MANUAL-TESTS.md).

## Performance

Reading files is what takes the time, not the MD5 calculation: one CPU core
hashes faster than most disks and networks deliver data. The tool is designed
around reading as little as possible, and these options help further:

| What | Effect | When to use it |
|---|---|---|
| Built-in pre-filter | Files that differ in name, saved date or size are never read. The size and saved date of a file whose name no other file has are never even looked up (on Linux, macOS and network drives each lookup is a request). | Always on. |
| Large files compared by their start | Files of 16 MB or more that share a name, saved date and size are first compared by the MD5 of their first 1 MB; only those that still match are read in full. True duplicates cost at most 1/16 more reading. | Always on. |
| Duplicate folders narrowed by name | Only folders whose name another folder has (and the folders below them) are fingerprinted and have their files looked up. | Always on with `-IncludeFolders`. |
| `-ThrottleLimit 4` to `8` | Lists several folders and hashes several files at once, hiding per-request latency. Often 2-4x faster, more on slow networks. Scans of a network share or drive use 4 unless told otherwise. | SSDs, network shares, cloud folders. Keep `1` for a single spinning hard disk, where parallel reads cause seeking. |
| `-Exclude` / `-MinimumSize` | Left-out folders are never listed, and small files are never compared. | Caches, version-control folders, thumbnails, tiny files. |
| Rescanning to the same report | Files the report lists whose size and saved date have not changed keep their recorded MD5 instead of being read again (`-Rehash` to read them all). | Repeated scans of large libraries or shares. |
| `-Validate` instead of a rescan | Checks only the files already in the report, without reading them. | After deleting or moving duplicates. |
| `-SkipCloudOnly` | Avoids downloading online-only files. | Large cloud libraries on a slow connection. |
| Run it on the file server | Local disk reads instead of network transfers. | Very large network shares. |
| Scan the narrowest folder | Fewer files to list. | Always worthwhile. |

Also built in: each folder is listed once, and its files and sub folders are
sorted and split by .NET in one step; large reads with sequential-read
hints, into a buffer no larger than the file; plain loops
instead of per-file script blocks when grouping and sorting; a progress display
redrawn a few times a second rather than per file; parallel hashing through a
fixed set of workers rather than a new job per file (which matters when there
are many small files); and a report writer and reader that handle hundreds of
thousands of rows. A tree of 10,000 files, all duplicated, is scanned, saved
and validated in a few seconds.

The PowerShell module's per-file loops (grouping names, listing folders, hashing,
reading the report, checking copies) run as compiled .NET code, which the module
builds from `DuplicateFinder.cs` when it is imported (about half a second, once
per PowerShell session); no separate download or install is involved.

GPU/CUDA acceleration would not help: MD5 cannot be split across GPU cores
within a single file, the bottleneck is reading the data, and copying it to
the GPU adds overhead.

## Using the functions directly

`DuplicateFinder.psm1` exports the building blocks:

```powershell
Import-Module .\DuplicateFinder.psm1

$files = Get-FileInventory -Path D:\Photos                       # every file, recursively
$dupes = Find-DuplicateFile -File $files -ThrottleLimit 4        # name + date + MD5 sets
Export-DuplicateReport -DuplicateSet $dupes -Path .\dupes.xlsx   # write the workbook

# Duplicate folders need the folder records from the scan
$info    = [System.Collections.Generic.List[object]]::new()
$files   = Get-FileInventory -Path D:\Photos -FolderInfo $info
$folders = @(Find-DuplicateFolder -File $files -Folder $info.ToArray())   # @(): an empty result stays a list
Export-DuplicateReport -DuplicateSet @(Find-DuplicateFile -File $files) -FolderSet $folders -Path .\dupes.xlsx

$rows   = Import-DuplicateReport -Path .\dupes.xlsx              # read the file rows back
$frows  = Import-DuplicateFolderReport -Path .\dupes.xlsx        # read the folder rows back
$result = Update-DuplicateReport -Path .\dupes.xlsx -WhatIf      # validate (summary object)
```

## Running the tests

```powershell
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser
Invoke-Pester ./tests
```

The tests also run on every push via GitHub Actions (Windows, Linux, macOS,
and Windows PowerShell 5.1), together with the Python tests and a parity test
that runs both versions on the same folders, with and without parallel
hashing and validation, and requires identical reports.

The *Edge cases* tests cover the situations under
[Limits and edge cases](#limits-and-edge-cases). Some need particular
conditions and are skipped without them: a file system that keeps case
(Linux), symbolic links, a non-administrator account for unreadable files,
Linux or macOS for 4 GB sparse files and time zone changes, and LibreOffice
Calc (`soffice`) for the round trip through another spreadsheet program. CI
installs LibreOffice on Linux and sets `FIND_DUPLICATES_REQUIRE_LIBREOFFICE=1`
so that test cannot be skipped there.
