#Requires -Version 5.1
<#
.SYNOPSIS
    Finds duplicate files in a folder and all of its sub folders and saves them to Excel.

.DESCRIPTION
    Two files are duplicates only when ALL three of these match:
      * file name (case-insensitive)
      * saved date (last modified time, to the whole second)
      * MD5 hash of the contents

    MD5 is only calculated for files whose name and saved date already match
    another file (and whose size matches too), so most files are never read.

    The report has one row per duplicated file with the columns
    File Name | Last Modified | Size (bytes) | MD5 | Copies | Location 1 | Location 2 | ...
    where each "Location" column holds the full folder path of one copy.

    Microsoft Excel does NOT need to be installed.

.PARAMETER Path
    Folder to scan. Defaults to the current folder.

.PARAMETER OutputFile
    Workbook to write. Defaults to "duplicates.xlsx" in the current folder.
    ".xlsx" is appended when no extension is given.

.PARAMETER PassThru
    Also return the duplicate sets as objects.

.EXAMPLE
    .\Find-Duplicates.ps1 -Path D:\Photos

.EXAMPLE
    .\Find-Duplicates.ps1 -Path \\server\share -OutputFile C:\Reports\share-dupes.xlsx -Verbose
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string] $Path = '.',

    [Parameter(Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string] $OutputFile = 'duplicates.xlsx',

    [switch] $PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'src/DuplicateFinder.psm1') -Force

if (-not [System.IO.Path]::HasExtension($OutputFile)) {
    $OutputFile += '.xlsx'
}
$reportPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($OutputFile)
$scanRoot   = (Resolve-Path -LiteralPath $Path).ProviderPath

Write-Host "Scanning '$scanRoot' ..."
$files = @(Get-FileInventory -Path $scanRoot -ExcludeFile $reportPath)
Write-Host "Found $($files.Count) files. Checking for duplicates ..."

$duplicates = @(Find-DuplicateFile -File $files)
Export-DuplicateReport -DuplicateSet $duplicates -Path $reportPath

$copies = 0
foreach ($set in $duplicates) { $copies += $set.Count }
Write-Host "Found $($duplicates.Count) duplicated files ($copies copies in total)."
Write-Host "Report saved to '$reportPath'."

if ($PassThru) { $duplicates }
