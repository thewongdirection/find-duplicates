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
    another file (and whose size matches too), so most files are never read.

    The report has one row per duplicated file with the columns
    File Name | Last Modified | Size (bytes) | MD5 | Copies | Location 1 | Location 2 | ...
    where each "Location" column holds the full folder path of one copy.

    VALIDATE (-Validate)
    Re-checks every copy listed in an existing report with a quick file lookup
    (no contents are read) and removes copies that no longer exist or whose size
    or saved date changed. Rows left with fewer than two copies are removed. The
    report is updated in place. Copies on a drive or network share that cannot be
    reached are kept.

    Local folders, network shares (\\server\share or mapped drives) and synced
    cloud folders (OneDrive, Google Drive, Dropbox ...) are all supported.
    Microsoft Excel does NOT need to be installed.

.PARAMETER Path
    Folder to scan. Defaults to the current folder.

.PARAMETER OutputFile
    Report to write (scan) or to check and update (validate). Defaults to
    "duplicates.xlsx" in the current folder. ".xlsx" is appended when no
    extension is given.

.PARAMETER SkipCloudOnly
    Never download online-only cloud files to hash them. Duplicates among such
    files are then not reported.

.PARAMETER ThrottleLimit
    How many files to hash at the same time (1-64, default 1). Try 4-8 for SSDs,
    network shares and cloud folders; keep 1 for a single spinning hard disk.

.PARAMETER Validate
    Re-check an existing report instead of scanning.

.PARAMETER PassThru
    Also return the duplicate sets (after validation, the rows that remain).

.PARAMETER WhatIf
    Show what would be saved or removed without changing the report.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path D:\Photos

.EXAMPLE
    .\Find-Duplicates.ps1 -Path \\server\share -OutputFile C:\Reports\share-dupes.xlsx -ThrottleLimit 8

.EXAMPLE
    .\Find-Duplicates.ps1 -Path "$env:OneDrive" -SkipCloudOnly

.EXAMPLE
    .\Find-Duplicates.ps1 -Validate -OutputFile C:\Reports\share-dupes.xlsx -WhatIf
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
    [ValidateRange(1, 64)]
    [int] $ThrottleLimit = 1,

    [Parameter(ParameterSetName = 'Validate', Mandatory)]
    [switch] $Validate,

    [switch] $PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'src/DuplicateFinder.psm1') -Force

if (-not [System.IO.Path]::HasExtension($OutputFile)) {
    $OutputFile += '.xlsx'
}
$reportPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($OutputFile)

if ($Validate) {
    Write-Host "Validating '$reportPath' ..."
    # -WhatIf does not flow into module functions on its own, so pass it on.
    $result = Update-DuplicateReport -Path $reportPath -WhatIf:$WhatIfPreference
    Write-Host ("Checked {0} copies in {1} rows: {2} missing or changed, {3} unreachable (kept)." -f
        $result.CopiesChecked, $result.RowsChecked, $result.CopiesRemoved, $result.CopiesUnavailable)
    Write-Host "Removed $($result.RowsRemoved) rows that are no longer duplicates; $($result.RowsRemaining) remain."
    if ($result.Saved) { Write-Host "Report updated: '$reportPath'." }
    elseif ($result.CopiesRemoved -or $result.RowsRemoved) { Write-Host 'Report not changed (-WhatIf).' }
    else { Write-Host 'Report is up to date; nothing to change.' }

    if ($PassThru) { $result.DuplicateSet }
    return
}

$scanRoot = (Resolve-Path -LiteralPath $Path).ProviderPath

Write-Host "Scanning '$scanRoot' ..."
$files = @(Get-FileInventory -Path $scanRoot -ExcludeFile $reportPath)
Write-Host "Found $($files.Count) files. Checking for duplicates ..."

$duplicates = @(Find-DuplicateFile -File $files -SkipCloudOnly:$SkipCloudOnly -ThrottleLimit $ThrottleLimit)

$copies = 0
foreach ($set in $duplicates) { $copies += $set.Count }
Write-Host "Found $($duplicates.Count) duplicated files ($copies copies in total)."

if ($PSCmdlet.ShouldProcess($reportPath, 'Save duplicates report')) {
    Export-DuplicateReport -DuplicateSet $duplicates -Path $reportPath
    Write-Host "Report saved to '$reportPath'."
}
else { Write-Host "Report not saved (-WhatIf): '$reportPath'." }

if ($PassThru) { $duplicates }
