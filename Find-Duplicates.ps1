#Requires -Version 5.1

<#
.SYNOPSIS
    Finds duplicate files in a folder and all of its sub folders and saves them to
    Excel, or re-checks an existing report without rescanning.

.DESCRIPTION
    SCAN (default)
    Two files are duplicates only when ALL three of these match:
      * file name (case-insensitive)
      * saved date (last modified time, to the whole second)
      * MD5 hash of the contents

    MD5 is only calculated for files whose name and saved date already match
    another file (and whose size matches too), so most files are never read. When
    the report already exists, unchanged files keep the MD5 recorded there.
    Files of 0 bytes are included unless -IgnoreEmptyFiles is used. -Exclude leaves
    files and folders out by name, and -MinimumSize leaves small files out.

    A "Rules" sheet in the report states the matching rules in plain words.
    The report has one row per duplicated file with the columns
    File Name | Last Modified | UTC Offset | Size (bytes) | MD5 | Copies | Location 1 | ...
    where each "Location" column holds the full folder path of one copy, from the least
    nested (Location 1) to the most nested. Last Modified
    is local time and UTC Offset its difference from UTC, so -Validate works in any
    time zone.

    DUPLICATE FOLDERS (-IncludeFolders)
    Also finds folders with the same name and exactly the same contents: the same tree
    of file and sub folder names, where every file is a duplicate (name, saved date,
    MD5) of the file at the same place in the other folder. Only the top-most
    duplicate folders are reported, on a second sheet, "Duplicate Folders":
    Folder Name | Files | Sub Folders | Size (bytes) | Copies | Location 1 | ...
    where each "Location" column holds the full path of one copy of the folder, from the
    least nested (Location 1) to the most nested.

    VALIDATE (-Validate)
    Re-checks every copy listed in an existing report with a quick file lookup
    (no contents are read) and removes copies that no longer exist or whose size
    or saved date changed. Rows left with fewer than two copies are removed. The
    report is updated in place. Copies on a drive or network share that cannot be
    reached are kept. Duplicate folders, when the report has them, are re-listed
    and must still have the same number of files and sub folders and total size.

    Local folders, network shares (\\server\share or mapped drives) and synced
    cloud folders (OneDrive, Google Drive, Dropbox ...) are all supported.
    Microsoft Excel does NOT need to be installed.

    Started without any parameters, the script shows this help and does nothing
    else. To scan the current folder, give it: -Path .

.PARAMETER Path
    Folder to scan. Defaults to the current folder when other parameters are given
    (without any parameters the script shows this help instead).

.PARAMETER OutputFile
    Report to write (scan) or to check and update (validate). Defaults to
    "duplicates.xlsx" in the current folder. ".xlsx" is appended when no
    extension is given.

.PARAMETER SkipCloudOnly
    Never download online-only cloud files to hash them. Duplicates among such
    files are then not reported.

.PARAMETER ThrottleLimit
    How many files to hash, and folders to list or check, at the same time (1-64).
    Default: 4 when scanning a network share or network drive, 1 otherwise. Try 4-8 for
    SSDs, network shares and cloud folders; keep 1 for a single spinning hard disk. Also
    speeds up -Validate (default 1).

.PARAMETER IgnoreEmptyFiles
    Leave files of 0 bytes out of the duplicate files (they all have the same
    contents). Off by default. Duplicate folders still compare every file.

.PARAMETER Exclude
    Names of files and folders to leave out, as wildcard patterns: * stands for any
    characters and ? for any one character; upper/lower case is ignored. Patterns match
    names, not paths. Left-out folders are not scanned, and duplicate folders are compared
    as if left-out files and folders were not there. Recorded in the report, so -Validate
    leaves the same names out. Example: -Exclude Thumbs.db, .DS_Store, *.tmp

.PARAMETER MinimumSize
    Leave files smaller than this many bytes out of the duplicate files, for example
    -MinimumSize 1MB or 1.5MB (KB, MB, GB, TB, PB: 1024-based). Duplicate folders still
    compare every file. Recorded in the report.

.PARAMETER IncludeFolders
    Also find duplicate folders and save them on the "Duplicate Folders" sheet.

.PARAMETER Rehash
    Read every candidate file again. Without it, when the report already exists (from
    an earlier scan), files it lists whose size and saved date have not changed keep the
    MD5 hash recorded there instead of being read again.

.PARAMETER Validate
    Re-check an existing report instead of scanning.

.PARAMETER PassThru
    Also return the duplicate sets (after validation, the rows that remain): file
    sets first, then folder sets (which have a FolderName property).

.PARAMETER WhatIf
    Show what would be saved or removed without changing the report.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path .

    Scans the current folder and saves duplicates.xlsx there.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path D:\Photos -OutputFile C:\Reports\photo-dupes

    Scans D:\Photos and saves C:\Reports\photo-dupes.xlsx. The same, positionally:
    .\Find-Duplicates.ps1 D:\Photos C:\Reports\photo-dupes

.EXAMPLE
    .\Find-Duplicates.ps1 -Path \\server\share -OutputFile C:\Reports\share-dupes.xlsx -ThrottleLimit 8

    Scans a network share, hashing 8 files at a time.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path "$env:OneDrive" -SkipCloudOnly

    Scans OneDrive without downloading online-only files.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path D:\Backups -IncludeFolders

    Also lists duplicate folders, on the "Duplicate Folders" sheet.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path D:\Photos -Exclude Thumbs.db, .git, *.tmp -MinimumSize 100KB -IgnoreEmptyFiles

    Leaves out thumbnail caches, Git folders, temporary files, and files under 100 KB.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path D:\Photos -Rehash

    Scans again, reading every candidate file instead of reusing the MD5 hashes in
    the existing duplicates.xlsx.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path D:\Photos -WhatIf -Verbose

    Prints every folder scanned and the totals, without saving a report.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path D:\Photos -PassThru | Sort-Object Count -Descending | Select-Object -First 10

    Also returns the duplicates as objects: here, the 10 files with the most copies.

.EXAMPLE
    .\Find-Duplicates.ps1 -Validate -OutputFile C:\Reports\share-dupes.xlsx -WhatIf

    Shows which copies listed in the report are gone or changed, without changing it.
    Leave out -WhatIf to remove them from the report. The same, positionally:
    .\Find-Duplicates.ps1 -Validate C:\Reports\share-dupes.xlsx -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Scan')]
param(
    [Parameter(ParameterSetName = 'Scan', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string] $Path = '.',

    [Parameter(ParameterSetName = 'Scan', Position = 1)]
    [Parameter(ParameterSetName = 'Validate', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string] $OutputFile = 'duplicates.xlsx',

    [Parameter(ParameterSetName = 'Scan')]
    [switch] $SkipCloudOnly,

    [Parameter(ParameterSetName = 'Scan')]
    [Parameter(ParameterSetName = 'Validate')]
    [ValidateRange(1, 64)]
    [int] $ThrottleLimit = 1,

    [Parameter(ParameterSetName = 'Scan')]
    [switch] $IgnoreEmptyFiles,

    [Parameter(ParameterSetName = 'Scan')]
    [AllowEmptyString()]
    [string[]] $Exclude = @(),

    [Parameter(ParameterSetName = 'Scan')]
    # Text, read by ConvertFrom-SizeText: Windows PowerShell 5.1 cannot turn '2KB' into a number.
    [string] $MinimumSize = '0',

    [Parameter(ParameterSetName = 'Scan')]
    [switch] $IncludeFolders,

    [Parameter(ParameterSetName = 'Scan')]
    [switch] $Rehash,

    [Parameter(ParameterSetName = 'Validate', Mandatory)]
    [switch] $Validate,

    [switch] $PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Started without parameters: show how to use them instead of scanning the current folder.
if ($PSBoundParameters.Count -eq 0) {
    Get-Help -Name $PSCommandPath -Full
    return
}

Import-Module (Join-Path $PSScriptRoot 'DuplicateFinder.psm1') -Force

# Preference variables such as -Verbose do not flow into module functions, so pass it on.
$verbose = @{ Verbose = $VerbosePreference -eq 'Continue' }

if (-not [System.IO.Path]::HasExtension($OutputFile)) {
    $OutputFile += '.xlsx'
}
$reportPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($OutputFile)

if ($Validate) {
    Write-Host "Validating '$reportPath' ..."
    # -WhatIf does not flow into module functions on its own, so pass it on.
    $result = Update-DuplicateReport -Path $reportPath -ThrottleLimit $ThrottleLimit -WhatIf:$WhatIfPreference @verbose
    Write-Host ("Checked {0} copies in {1} rows: {2} missing or changed, {3} unreachable (kept)." -f
        $result.CopiesChecked, $result.RowsChecked, $result.CopiesRemoved, $result.CopiesUnavailable)
    Write-Host "Removed $($result.RowsRemoved) rows that are no longer duplicates; $($result.RowsRemaining) remain."
    if ($null -ne $result.DuplicateFolderSet) {
        Write-Host ("Checked {0} folder copies in {1} rows: {2} missing or changed, {3} unreachable (kept)." -f
            $result.FolderCopiesChecked, $result.FolderRowsChecked, $result.FolderCopiesRemoved, $result.FolderCopiesUnavailable)
        Write-Host "Removed $($result.FolderRowsRemoved) folder rows that are no longer duplicates; $($result.FolderRowsRemaining) remain."
    }
    if ($result.Saved) { Write-Host "Report updated: '$reportPath'." }
    elseif ($result.CopiesRemoved -or $result.RowsRemoved -or $result.FolderCopiesRemoved -or $result.FolderRowsRemoved) {
        Write-Host 'Report not changed (-WhatIf).'
    }
    else { Write-Host 'Report is up to date; nothing to change.' }

    if ($PassThru) {
        $result.DuplicateSet
        if ($null -ne $result.DuplicateFolderSet) { $result.DuplicateFolderSet }
    }
    return
}

# Checked before anything is scanned.
Assert-NamePattern -Pattern $Exclude
$minimumBytes = ConvertFrom-SizeText -Text $MinimumSize -Name '-MinimumSize'
$scanRoot = (Resolve-Path -LiteralPath $Path).ProviderPath

# Resolve the report's folder like the scan root (e.g. Windows short names expanded), so the
# report is recognised and left out when it is saved inside the scanned folder.
$reportFolder = Split-Path -Parent $reportPath
if (Test-Path -LiteralPath $reportFolder -PathType Container) {
    $reportPath = Join-Path (Resolve-Path -LiteralPath $reportFolder).ProviderPath (Split-Path -Leaf $reportPath)
}

Write-Host "Scanning '$scanRoot' ..."
if (-not $PSBoundParameters.ContainsKey('ThrottleLimit')) {
    $ThrottleLimit = Get-DefaultThrottleLimit -Path $scanRoot
    if ($ThrottleLimit -gt 1) {
        Write-Host "On a network drive: working on $ThrottleLimit folders and files at a time (-ThrottleLimit to change)."
    }
}
if ($Exclude.Count) { Write-Host "Leaving out files and folders named: $($Exclude -join ', ')" }
$folderInfo = $null  # (not "= if ...": an empty list would be unrolled into $null)
if ($IncludeFolders) { $folderInfo = [System.Collections.Generic.List[object]]::new() }
$files = @(Get-FileInventory -Path $scanRoot -ExcludeFile $reportPath -FolderInfo $folderInfo -ThrottleLimit $ThrottleLimit -ExcludeName $Exclude @verbose)
Write-Host "Found $($files.Count) files. Checking for duplicates ..."

# Hashes are shared so that folder matching never reads a file twice.
$md5Cache = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)

# An earlier report's hashes are reused for files that have not changed since.
if (-not $Rehash -and [System.IO.File]::Exists($reportPath)) {
    try {
        $previous = Get-PreviousMd5 -Path $reportPath -File $files
        foreach ($entry in $previous.GetEnumerator()) { $md5Cache[$entry.Key] = $entry.Value }
        if ($previous.Count) { Write-Host "Reusing $($previous.Count) MD5 hashes from the previous report (-Rehash to read every file again)." }
    }
    catch { Write-Warning "Not reusing MD5 hashes from '$reportPath': $($_.Exception.Message)" }
}
$matchOptions = @{ SkipCloudOnly = $SkipCloudOnly; ThrottleLimit = $ThrottleLimit; Md5Cache = $md5Cache } + $verbose

$duplicates = @(Find-DuplicateFile -File $files -IgnoreEmptyFiles:$IgnoreEmptyFiles -MinimumSize $minimumBytes @matchOptions)
$copies = 0
foreach ($set in $duplicates) { $copies += $set.Count }
Write-Host "Found $($duplicates.Count) duplicated files ($copies copies in total)."

# The smallest file listed, recorded in the report: -IgnoreEmptyFiles means 1 byte.
$smallest = $minimumBytes
if ($IgnoreEmptyFiles -and $smallest -lt 1) { $smallest = 1 }
$export = @{ DuplicateSet = $duplicates; Path = $reportPath; ExcludeName = $Exclude; MinimumSize = $smallest }
if ($IncludeFolders) {
    Write-Host 'Checking for duplicate folders ...'
    $folderDuplicates = @(Find-DuplicateFolder -File $files -Folder $folderInfo.ToArray() @matchOptions)
    $folderCopies = 0
    foreach ($set in $folderDuplicates) { $folderCopies += $set.Count }
    Write-Host "Found $($folderDuplicates.Count) duplicated folders ($folderCopies copies in total)."
    $export.FolderSet = $folderDuplicates
}

if ($PSCmdlet.ShouldProcess($reportPath, 'Save duplicates report')) {
    Export-DuplicateReport @export
    Write-Host "Report saved to '$reportPath'."
}
else { Write-Host "Report not saved (-WhatIf): '$reportPath'." }

if ($PassThru) {
    $duplicates
    if ($IncludeFolders) { $folderDuplicates }
}
