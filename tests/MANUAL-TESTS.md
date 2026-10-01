# Manual tests

These situations need real equipment, so CI cannot run them. Check them by hand
before a release, with both tools (`Find-Duplicates.ps1` and
`python -m find_duplicates`). Each check lists what to do and what must happen.

## Network share (SMB)

Needs a second computer or a NAS sharing a folder that holds some duplicates.

1. **Scan a share by UNC path.** Run a scan of `\\server\share\folder` with
   `-IncludeFolders` / `--folders`. The report lists the duplicates, with
   locations starting `\\server\share\`.
2. **Scan a mapped drive.** Map the share (`Z:`) and scan `Z:\folder`. Same
   result, with locations starting `Z:\`.
3. **Share drops mid-scan.** Start a scan of a large share with `-Verbose` /
   `--verbose`, then unplug the server's network cable (or stop sharing).
   Each folder that can no longer be read gives a warning; the scan finishes
   and saves a report of what it did read. Nothing hangs for more than the
   system's network timeout per folder.
4. **Validate with the share offline.** With the server still disconnected, run
   `-Validate` / `--validate` on that report. The copies on the share are
   reported as unavailable and kept; the report is not emptied. The wait is one
   timeout for the share, not one per file.
5. **Validate with the share back.** Reconnect and validate again: copies that
   are still there are kept, deleted ones are removed.
6. **Several at a time by default.** Scan the share without `-ThrottleLimit` /
   `-j`. The tool says it works on 4 folders and files at a time; with
   `-ThrottleLimit 1` it does not, and is slower.

## OneDrive (Files On-Demand) and other cloud drives

Needs a OneDrive folder with some files set to *online-only* (right-click,
*Free up space*): at least two online-only copies of one file, and an
online-only copy of a file that also exists locally.

1. **Listing downloads nothing.** Scan a folder with online-only files that have
   no duplicate candidates. No file changes to *locally available* (the cloud
   icon stays).
2. **`-SkipCloudOnly` / `--skip-cloud-only`.** Scan with the option. A warning
   says how many online-only files were not checked; none of them is
   downloaded; duplicates among local files are still found. With
   `-IncludeFolders`, folders holding skipped online-only files are not listed.
3. **Default: candidates are downloaded.** Scan without the option. Only the
   online-only files that share a name, saved date and size with another file
   are downloaded, and their duplicates are reported.
4. **Validation downloads nothing.** Run `-Validate` on the report. No
   online-only file is downloaded.
5. Repeat 1-4 with Google Drive (stream mode) or Dropbox (online-only files) if
   available.

## Microsoft Excel

Needs Excel on Windows or macOS.

1. **Opens cleanly.** Open a report made by a default scan (duplicate folders included). Excel opens
   it without a repair prompt. The *Duplicates*, *Duplicate Folders* and
   *Rules* sheets are present; header rows are frozen and have filters;
   *Last Modified* shows as a date and time.
2. **Filtering and sorting work.** Filter the *Duplicates* sheet by file name
   and sort it by size; no error.
3. **Round trip.** Delete one row, save in Excel (keeping .xlsx), close Excel,
   then run `-Validate` on the saved file. It is read without errors, and the
   deleted row stays deleted.
4. **Open while scanning or validating.** With the report open in Excel, run a
   scan to it, then `-Validate` / `--validate` on it. Each stops at once, before
   scanning or validating, with "The report '...' is locked by another program
   (is it open in Excel?). Close it or free whatever is locking it, then run the
   command again." The report is unchanged; after closing Excel both run.
5. **Unicode.** A report of files named in Japanese, Arabic and with emoji shows
   those names correctly.
6. **Edited scan settings.** Scan with `-Exclude Thumbs.db -MinimumSize 1KB`
   (duplicate folders included by default), open the report in Excel, add a pattern in a new cell
   of the *Names left out* row on the *Rules* sheet and type `abc` as the
   smallest size, and save. `-Validate` warns that it ignores the size, leaves
   out the added names too, and keeps the names row when it rewrites the report.

## Locked-down Windows

1. **Constrained Language Mode or AppLocker.** On a computer where PowerShell
   runs in Constrained Language Mode, importing the module fails with a clear
   error from `Add-Type` (its compiled helpers cannot be built there); use the
   Python tool on such computers.
2. **No writable temp folder.** Windows PowerShell 5.1 compiles the helpers in
   `%TEMP%`; with it read-only, importing fails with a clear error.

## Very long paths on Windows PowerShell 5.1

1. On Windows with the *long paths* setting on and off, scan a folder whose
   paths exceed 260 characters with `powershell.exe` (5.1). Either the
   duplicates are reported, or each unreachable folder or file gives a
   warning; the scan never stops with an error.
