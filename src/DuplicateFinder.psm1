#Requires -Version 5.1
<#
    DuplicateFinder module.

    A file is a duplicate of another when ALL of the following match:
      1. File name (case-insensitive)
      2. Saved date (LastWriteTime, compared to the whole second)
      3. MD5 hash of the contents

    MD5 is only computed for files whose name and saved date already match
    another file (and whose size matches, since files of different sizes can
    never share an MD5), so most files are never read.

    Works on local folders, network shares (UNC paths or mapped drives) and
    synced cloud folders (OneDrive, Google Drive, Dropbox ...). No network
    access is needed beyond reading the folders themselves.
#>
Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression

$script:ProgressIntervalMs = 250
$script:ExcelMaxRows       = 1048576
$script:ExcelMaxColumns    = 16384
$script:FixedColumns       = @(
    @{ Header = 'File Name';     Width = 40 }
    @{ Header = 'Last Modified'; Width = 20 }
    @{ Header = 'Size (bytes)';  Width = 14 }
    @{ Header = 'MD5';           Width = 34 }
    @{ Header = 'Copies';        Width = 8  }
)
$script:LocationColumnWidth = 60

# Attributes Windows sets on cloud placeholders (OneDrive "Files On-Demand" and other
# Cloud Files providers) whose contents are not stored locally. Reading them downloads them.
$script:CloudOnlyAttributes = 0x1000 -bor 0x40000 -bor 0x400000  # Offline | RecallOnOpen | RecallOnDataAccess

# Orderings shared with the Python port (python/find_duplicates) so both tools
# report the same "first" copy and sort rows and locations identically on every OS.
$script:ByDuplicateSet = [System.Comparison[object]] {
    param($x, $y)
    $order = [System.StringComparer]::OrdinalIgnoreCase.Compare($x.FileName, $y.FileName)
    if ($order -eq 0) { $order = $x.LastWriteTime.CompareTo($y.LastWriteTime) }
    if ($order -eq 0) { $order = [string]::CompareOrdinal($x.MD5, $y.MD5) }
    $order
}

#region Scanning

function Get-FileInventory {
    <#
    .SYNOPSIS
        Recursively lists every file below a folder, reporting the folder being scanned.
    .DESCRIPTION
        Walks the tree iteratively so that deep trees cannot overflow the call stack.
        Folders that cannot be read (permissions, dropped network connection) are
        reported as warnings and skipped. Symbolic links and junctions are not
        followed, which prevents infinite loops; cloud-synced folders are followed.
        Listing folders never downloads cloud files.
    #>
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        # Full paths of files to leave out of the inventory (e.g. the report itself).
        [string[]] $ExcludeFile = @()
    )

    $root = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if (-not $root.PSIsContainer) {
        throw "'$Path' is not a folder."
    }

    $excluded = [System.Collections.Generic.HashSet[string]]::new(
        [string[]] $ExcludeFile, [System.StringComparer]::OrdinalIgnoreCase)

    $pending = [System.Collections.Generic.Stack[System.IO.DirectoryInfo]]::new()
    $pending.Push([System.IO.DirectoryInfo] $root.FullName)

    $folderCount = 0
    $fileCount   = 0
    $timer       = [System.Diagnostics.Stopwatch]::StartNew()
    $lastShownMs = - $script:ProgressIntervalMs

    while ($pending.Count -gt 0) {
        $folder = $pending.Pop()
        $folderCount++

        Write-Verbose "Scanning $($folder.FullName)"
        if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
            $lastShownMs = $timer.ElapsedMilliseconds
            Write-Progress -Id 1 -Activity 'Scanning folders' `
                -Status "Folders: $folderCount   Files: $fileCount" `
                -CurrentOperation $folder.FullName
        }

        try {
            # File systems list entries in different orders (alphabetical on NTFS, arbitrary on ext4).
            $files      = Get-SortedByName -Item $folder.GetFiles()
            $subFolders = Get-SortedByName -Item $folder.GetDirectories()
        }
        catch [System.UnauthorizedAccessException], [System.IO.IOException], [System.Security.SecurityException] {
            Write-Warning "Skipping '$($folder.FullName)': $($_.Exception.Message)"
            continue
        }

        foreach ($file in $files) {
            if ($excluded.Contains($file.FullName)) { continue }
            $fileCount++
            $file
        }

        # Push in reverse so folders are visited in alphabetical order.
        for ($i = $subFolders.Count - 1; $i -ge 0; $i--) {
            $sub = $subFolders[$i]
            if (Test-FolderLink -Folder $sub) {
                Write-Verbose "Not following link '$($sub.FullName)'"
                continue
            }
            $pending.Push($sub)
        }
    }

    Write-Progress -Id 1 -Activity 'Scanning folders' -Completed
}

function Get-SortedByName {
    # Files or folders in ordinal name order, as a new array.
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Item)
    if ($Item.Count -gt 1) {
        $names = [string[]] $Item.Name
        [System.Array]::Sort($names, $Item, [System.StringComparer]::Ordinal)
    }
    , $Item
}

function Test-FolderLink {
    # True for symbolic links and junctions, which could loop back on the tree.
    # Other reparse points are ordinary folders to the user: OneDrive and other
    # cloud-synced folders, deduplicated volumes, DFS links. Those are scanned.
    # LinkType is added to DirectoryInfo by PowerShell on every platform.
    param([Parameter(Mandatory)] [object] $Folder)

    if (-not ($Folder.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) { return $false }
    $Folder.LinkType -in 'SymbolicLink', 'Junction'
}

function Test-CloudOnlyFile {
    # True for cloud placeholders whose contents would have to be downloaded to hash them.
    param([Parameter(Mandatory)] [object] $File)
    ([long] $File.Attributes -band $script:CloudOnlyAttributes) -ne 0
}

#endregion

#region Matching

# Computes the MD5 of one file as upper-case hex. Kept as a script block so the very
# same code runs in the current session and in the parallel runspaces. The large
# buffer and the sequential-scan hint make reads from disks and shares much faster.
$script:ComputeMd5 = {
    param([string] $Path)
    $ErrorActionPreference = 'Stop'
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        $stream = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
            [System.IO.FileShare] 'ReadWrite, Delete', 1MB, [System.IO.FileOptions]::SequentialScan)
        try { [System.BitConverter]::ToString($md5.ComputeHash($stream)).Replace('-', '') }
        finally { $stream.Dispose() }
    }
    finally { $md5.Dispose() }
}

function Get-FileMd5 {
    # MD5 of one file's contents as upper-case hex (the Get-FileHash format).
    param([Parameter(Mandatory)] [string] $Path)
    & $script:ComputeMd5 $Path
}

function Get-FileMd5Map {
    <#
        Hashes files, up to $ThrottleLimit at a time, returning a map of full path -> MD5.
        Files that cannot be read are reported as warnings and left out of the map.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]] $Path,
        [ValidateRange(1, 64)] [int] $ThrottleLimit = 1
    )

    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $activity = 'Comparing MD5 hashes'
    $done = 0

    if ($ThrottleLimit -eq 1) {
        foreach ($p in $Path) {
            $done++
            Write-Progress -Id 2 -Activity $activity -Status "File $done of $($Path.Count)" -CurrentOperation $p `
                -PercentComplete ([int] (100 * $done / $Path.Count))
            try { $map[$p] = Get-FileMd5 -Path $p }
            catch { Write-Warning "Could not hash '$p': $($_.Exception.Message)" }
        }
    }
    else {
        # A bounded window of jobs keeps memory flat however many files there are.
        $window = $ThrottleLimit * 4
        $inFlight = [System.Collections.Generic.Queue[object]]::new()
        $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit)
        $pool.Open()
        try {
            $next = 0
            while ($next -lt $Path.Count -or $inFlight.Count -gt 0) {
                while ($next -lt $Path.Count -and $inFlight.Count -lt $window) {
                    $shell = [System.Management.Automation.PowerShell]::Create()
                    $shell.RunspacePool = $pool
                    $null = $shell.AddScript($script:ComputeMd5.ToString()).AddArgument($Path[$next])
                    $inFlight.Enqueue([pscustomobject] @{ Path = $Path[$next]; Shell = $shell; Handle = $shell.BeginInvoke() })
                    $next++
                }

                $job = $inFlight.Dequeue()
                $done++
                Write-Progress -Id 2 -Activity $activity -Status "File $done of $($Path.Count) ($ThrottleLimit at a time)" `
                    -CurrentOperation $job.Path -PercentComplete ([int] (100 * $done / $Path.Count))
                try {
                    $output = $job.Shell.EndInvoke($job.Handle)
                    if ($job.Shell.Streams.Error.Count -gt 0) { throw $job.Shell.Streams.Error[0].Exception }
                    $map[$job.Path] = [string] $output[0]
                }
                catch {
                    $reason = $_.Exception
                    while ($reason.InnerException) { $reason = $reason.InnerException }
                    Write-Warning "Could not hash '$($job.Path)': $($reason.Message)"
                }
                finally { $job.Shell.Dispose() }
            }
        }
        finally {
            foreach ($job in $inFlight) { $job.Shell.Dispose() }
            $pool.Dispose()
        }
    }

    Write-Progress -Id 2 -Activity $activity -Completed
    $map
}

function Group-ByKey {
    # Groups items by the matching entry in $Key, returning only the groups that hold
    # more than one item. Keys are computed by the caller in plain loops, which is far
    # faster than invoking a script block per item on large trees.
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $InputItems,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]] $Key
    )

    $groups = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)

    for ($i = 0; $i -lt $InputItems.Count; $i++) {
        $list = $null
        if (-not $groups.TryGetValue($Key[$i], [ref] $list)) {
            $list = [System.Collections.Generic.List[object]]::new()
            $groups.Add($Key[$i], $list)
        }
        $list.Add($InputItems[$i])
    }

    foreach ($list in $groups.Values) {
        if ($list.Count -gt 1) { , $list.ToArray() }
    }
}

function Find-DuplicateFile {
    <#
    .SYNOPSIS
        Finds sets of files whose name, saved date and MD5 hash all match.
    .OUTPUTS
        One object per duplicate set: FileName, LastWriteTime, SizeBytes, MD5,
        Count and Folders (full folder path of every copy, sorted).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.IO.FileInfo[]] $File,

        # Do not hash cloud files that are not stored locally (hashing would download them).
        # Duplicates among such files are then not reported.
        [switch] $SkipCloudOnly,

        # How many files to hash at the same time.
        [ValidateRange(1, 64)]
        [int] $ThrottleLimit = 1
    )

    # Stage 1: name + saved date (UTC, whole second: copies made to network shares or
    # other file systems often lose sub-second precision). The date key is all digits,
    # so '|' is a safe separator.
    $ticksPerSecond = [System.TimeSpan]::TicksPerSecond
    $keys = [System.Collections.Generic.List[string]]::new($File.Count)
    foreach ($f in $File) {
        $ticks = $f.LastWriteTimeUtc.Ticks
        $keys.Add([string] ($ticks - ($ticks % $ticksPerSecond)) + '|' + $f.Name)
    }
    $nameDateGroups = @(Group-ByKey -InputItems $File -Key $keys.ToArray())

    # Stage 2: size. A cheap check that avoids hashing files that cannot match.
    # (Not $group.Length: on an array that is the array's own length.)
    $candidateGroups = @(foreach ($group in $nameDateGroups) {
            $sizes = [string[]] @(foreach ($f in $group) { $f.Length })
            Group-ByKey -InputItems $group -Key $sizes
        })

    if ($SkipCloudOnly) {
        $skipped = 0
        $candidateGroups = @(foreach ($group in $candidateGroups) {
                $local = @(foreach ($f in $group) {
                        if (Test-CloudOnlyFile -File $f) {
                            $skipped++
                            Write-Verbose "Not downloading online-only file '$($f.FullName)'"
                        }
                        else { $f }
                    })
                if ($local.Count -gt 1) { , $local }
            })
        if ($skipped) {
            Write-Warning "$skipped online-only cloud file(s) were not checked; duplicates among them are not reported."
        }
    }

    # Stage 3: MD5, only for files that already match on name, date and size.
    $candidates = [System.Collections.Generic.List[string]]::new()
    foreach ($group in $candidateGroups) { foreach ($f in $group) { $candidates.Add($f.FullName) } }
    $md5ByPath = Get-FileMd5Map -Path $candidates.ToArray() -ThrottleLimit $ThrottleLimit

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $candidateGroups) {
        $hashed = @($group | Where-Object { $md5ByPath.ContainsKey($_.FullName) })
        $md5s = [string[]] @($hashed | ForEach-Object { $md5ByPath[$_.FullName] })

        foreach ($set in @(Group-ByKey -InputItems $hashed -Key $md5s)) {
            $first = $set[0]
            $results.Add([pscustomobject] @{
                FileName      = $first.Name
                LastWriteTime = $first.LastWriteTime
                SizeBytes     = $first.Length
                MD5           = $md5ByPath[$first.FullName]
                Count         = $set.Count
                Folders       = Get-SortedFolder -Path @($set | ForEach-Object { $_.DirectoryName })
            })
        }
    }

    $results.Sort($script:ByDuplicateSet)
    $results
}

function Get-SortedFolder {
    # Folder paths in case-insensitive ordinal order.
    param([Parameter(Mandatory)] [string[]] $Path)
    $sorted = [string[]] $Path.Clone()
    [System.Array]::Sort($sorted, [System.StringComparer]::OrdinalIgnoreCase)
    , $sorted
}

#endregion

#region Excel output

function ConvertTo-ColumnName {
    # 1 -> A, 26 -> Z, 27 -> AA ...
    param([Parameter(Mandatory)] [ValidateRange(1, 16384)] [int] $Index)

    # Work on a copy: the ValidateRange attribute would reject assigning 0 to $Index.
    $remaining = $Index
    $name = ''
    while ($remaining -gt 0) {
        $digit     = ($remaining - 1) % 26
        $name      = "$([char] (65 + $digit))$name"
        $remaining = [int] [Math]::Floor(($remaining - 1) / 26)
    }
    $name
}

function ConvertTo-XmlSafeText {
    # File names may contain control characters (Linux) or unpaired surrogates (NTFS)
    # that XML 1.0 cannot represent; replace them with U+FFFD.
    param([AllowEmptyString()] [string] $Text)
    $invalid = '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]'
    [regex]::Replace($Text, $invalid, [string] [char] 0xFFFD)
}

function Write-ZipXmlEntry {
    # Creates a zip entry and passes an XmlWriter for it to $Body.
    param(
        [Parameter(Mandatory)] [System.IO.Compression.ZipArchive] $Archive,
        [Parameter(Mandatory)] [string] $EntryName,
        [Parameter(Mandatory)] [scriptblock] $Body
    )

    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)

    $stream = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal).Open()
    try {
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        try {
            $writer.WriteStartDocument($true)
            & $Body $writer
            $writer.WriteEndDocument()
        }
        finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Write-ZipTextEntry {
    # Writes a fixed XML part (no user data) with the standard declaration.
    param(
        [Parameter(Mandatory)] [System.IO.Compression.ZipArchive] $Archive,
        [Parameter(Mandatory)] [string] $EntryName,
        [Parameter(Mandatory)] [string] $Content
    )

    $stream = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal).Open()
    try {
        $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
        try {
            $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            $writer.Write($Content)
        }
        finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Write-Cell {
    # Writes one <c> element. Strings are inline, numbers/dates are numeric.
    param(
        [Parameter(Mandatory)] [System.Xml.XmlWriter] $Writer,
        [Parameter(Mandatory)] [string] $Reference,
        [Parameter(Mandatory)] [AllowNull()] [AllowEmptyString()] [object] $Value,
        [int] $Style = 0
    )

    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $Writer.WriteStartElement('c', $ns)
    $Writer.WriteAttributeString('r', $Reference)
    if ($Style) { $Writer.WriteAttributeString('s', [string] $Style) }

    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    if ($Value -is [datetime]) {
        $Writer.WriteElementString('v', $ns, $Value.ToOADate().ToString('R', $invariant))
    }
    elseif ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) {
        $Writer.WriteElementString('v', $ns, $Value.ToString($invariant))
    }
    else {
        $Writer.WriteAttributeString('t', 'inlineStr')
        $Writer.WriteStartElement('is', $ns)
        $Writer.WriteStartElement('t', $ns)
        $Writer.WriteAttributeString('xml', 'space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        $Writer.WriteString((ConvertTo-XmlSafeText ([string] $Value)))
        $Writer.WriteEndElement()
        $Writer.WriteEndElement()
    }
    $Writer.WriteEndElement()
}

function Write-WorksheetXml {
    param(
        [Parameter(Mandatory)] [System.Xml.XmlWriter] $Writer,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $DuplicateSet,
        [Parameter(Mandatory)] [int] $LocationColumns,
        [Parameter(Mandatory)] [string] $LastColumn
    )

    $ns       = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $lastRow  = $DuplicateSet.Count + 1
    $styleBold = 1
    $styleDate = 2

    $Writer.WriteStartElement('worksheet', $ns)

    # Frozen header row.
    $Writer.WriteStartElement('sheetViews', $ns)
    $Writer.WriteStartElement('sheetView', $ns)
    $Writer.WriteAttributeString('workbookViewId', '0')
    $Writer.WriteStartElement('pane', $ns)
    $Writer.WriteAttributeString('ySplit', '1')
    $Writer.WriteAttributeString('topLeftCell', 'A2')
    $Writer.WriteAttributeString('activePane', 'bottomLeft')
    $Writer.WriteAttributeString('state', 'frozen')
    $Writer.WriteEndElement()
    $Writer.WriteEndElement()
    $Writer.WriteEndElement()

    # Column widths.
    $widths = @($script:FixedColumns | ForEach-Object { $_.Width })
    $widths += @(1..$LocationColumns | ForEach-Object { $script:LocationColumnWidth })
    $Writer.WriteStartElement('cols', $ns)
    for ($i = 0; $i -lt $widths.Count; $i++) {
        $Writer.WriteStartElement('col', $ns)
        $Writer.WriteAttributeString('min', [string] ($i + 1))
        $Writer.WriteAttributeString('max', [string] ($i + 1))
        $Writer.WriteAttributeString('width', [string] $widths[$i])
        $Writer.WriteAttributeString('customWidth', '1')
        $Writer.WriteEndElement()
    }
    $Writer.WriteEndElement()

    $Writer.WriteStartElement('sheetData', $ns)

    # Header row.
    $headers = @($script:FixedColumns | ForEach-Object { $_.Header })
    $headers += @(1..$LocationColumns | ForEach-Object { "Location $_" })
    $Writer.WriteStartElement('row', $ns)
    $Writer.WriteAttributeString('r', '1')
    for ($c = 0; $c -lt $headers.Count; $c++) {
        Write-Cell -Writer $Writer -Reference "$(ConvertTo-ColumnName ($c + 1))1" -Value $headers[$c] -Style $styleBold
    }
    $Writer.WriteEndElement()

    # One row per duplicated file; one column per folder holding a copy.
    $rowNumber = 1
    foreach ($set in $DuplicateSet) {
        $rowNumber++
        $values = @($set.FileName, $set.LastWriteTime, [long] $set.SizeBytes, $set.MD5, [int] $set.Count) + @($set.Folders)

        $Writer.WriteStartElement('row', $ns)
        $Writer.WriteAttributeString('r', [string] $rowNumber)
        for ($c = 0; $c -lt $values.Count; $c++) {
            $style = if ($values[$c] -is [datetime]) { $styleDate } else { 0 }
            Write-Cell -Writer $Writer -Reference "$(ConvertTo-ColumnName ($c + 1))$rowNumber" -Value $values[$c] -Style $style
        }
        $Writer.WriteEndElement()
    }

    $Writer.WriteEndElement()  # sheetData

    $Writer.WriteStartElement('autoFilter', $ns)
    $Writer.WriteAttributeString('ref', "A1:$LastColumn$lastRow")
    $Writer.WriteEndElement()

    $Writer.WriteEndElement()  # worksheet
}

function Export-DuplicateReport {
    <#
    .SYNOPSIS
        Saves duplicate sets to an .xlsx workbook. Needs neither Excel nor extra modules.
    .DESCRIPTION
        One row per duplicated file. Columns: File Name, Last Modified, Size (bytes),
        MD5, Copies, then "Location 1..N" holding the full folder path of every copy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $DuplicateSet,

        [Parameter(Mandatory)]
        [string] $Path
    )

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)

    $maxCopies = 1
    foreach ($set in $DuplicateSet) { $maxCopies = [Math]::Max($maxCopies, $set.Folders.Count) }

    $columnCount = $script:FixedColumns.Count + $maxCopies
    if ($columnCount -gt $script:ExcelMaxColumns) {
        throw "A file has $maxCopies copies; Excel supports at most $($script:ExcelMaxColumns - $script:FixedColumns.Count) location columns."
    }
    if ($DuplicateSet.Count + 1 -gt $script:ExcelMaxRows) {
        throw "Found $($DuplicateSet.Count) duplicated files; Excel supports at most $($script:ExcelMaxRows - 1) rows."
    }
    $lastColumn = ConvertTo-ColumnName $columnCount
    $lastRow    = $DuplicateSet.Count + 1

    # Build next to the target, then swap in, so a failure never leaves a half-written report.
    $tempPath = "$Path.$([System.Guid]::NewGuid().ToString('N')).tmp"
    try {
        $fileStream = [System.IO.File]::Open($tempPath, [System.IO.FileMode]::CreateNew)
        try {
            $zip = [System.IO.Compression.ZipArchive]::new($fileStream, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                Write-ZipTextEntry -Archive $zip -EntryName '[Content_Types].xml' -Content (
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
                    '</Types>')

                Write-ZipTextEntry -Archive $zip -EntryName '_rels/.rels' -Content (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                    '</Relationships>')

                Write-ZipTextEntry -Archive $zip -EntryName 'xl/workbook.xml' -Content (
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                    '<sheets><sheet name="Duplicates" sheetId="1" r:id="rId1"/></sheets>' +
                    "<definedNames><definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"0`" hidden=`"1`">Duplicates!`$A`$1:`$$lastColumn`$$lastRow</definedName></definedNames>" +
                    '</workbook>')

                Write-ZipTextEntry -Archive $zip -EntryName 'xl/_rels/workbook.xml.rels' -Content (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
                    '</Relationships>')

                # Style 0 = default, 1 = bold header, 2 = date/time.
                Write-ZipTextEntry -Archive $zip -EntryName 'xl/styles.xml' -Content (
                    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>' +
                    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
                    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
                    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
                    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
                    '<cellXfs count="3">' +
                    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
                    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
                    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
                    '</cellXfs>' +
                    '</styleSheet>')

                Write-ZipXmlEntry -Archive $zip -EntryName 'xl/worksheets/sheet1.xml' -Body {
                    param($w)
                    Write-WorksheetXml -Writer $w -DuplicateSet $DuplicateSet `
                        -LocationColumns $maxCopies -LastColumn $lastColumn
                }
            }
            finally { $zip.Dispose() }
        }
        finally { $fileStream.Dispose() }

        Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
    }
}

#endregion

#region Reading and validating an existing report

function ConvertFrom-ColumnName {
    # A -> 1, Z -> 26, AA -> 27 ...
    param([Parameter(Mandatory)] [string] $Name)
    $index = 0
    foreach ($letter in $Name.ToUpperInvariant().ToCharArray()) { $index = $index * 26 + ([int] $letter - 64) }
    $index
}

function Read-ZipXml {
    # Parses one XML part of a zip package; $null when the part does not exist.
    param(
        [Parameter(Mandatory)] [System.IO.Compression.ZipArchive] $Archive,
        [Parameter(Mandatory)] [string] $EntryName
    )
    $entry = $Archive.GetEntry($EntryName)
    if (-not $entry) { return $null }

    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
        $xml = [System.Xml.XmlDocument]::new()
        $xml.XmlResolver = $null  # never resolve external entities
        $xml.LoadXml($reader.ReadToEnd())
    }
    finally { $reader.Dispose() }
    , $xml  # an XmlDocument would otherwise be enumerated into its child nodes
}

function Get-SpreadsheetNamespace {
    # A namespace manager for one parsed part (s: SpreadsheetML, p: package relationships).
    param([Parameter(Mandatory)] [System.Xml.XmlDocument] $Xml)
    $ns = [System.Xml.XmlNamespaceManager]::new($Xml.NameTable)
    $ns.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')
    , $ns  # a namespace manager would otherwise be enumerated into its prefixes
}

function Get-CellText {
    # Text of an inline or shared string, including rich-text runs (phonetic runs excluded).
    param([Parameter(Mandatory)] [System.Xml.XmlNode] $Node, [Parameter(Mandatory)] [System.Xml.XmlNamespaceManager] $Ns)
    -join @($Node.SelectNodes('s:t | s:r/s:t', $Ns) | ForEach-Object { $_.InnerText })
}

function Get-WorksheetRow {
    # The first worksheet as rows of cell text (a string[] per row, one entry per column).
    # Handles workbooks written by this tool and the same workbook after Excel saved it.
    param([Parameter(Mandatory)] [System.IO.Compression.ZipArchive] $Archive)

    $workbook = Read-ZipXml -Archive $Archive -EntryName 'xl/workbook.xml'
    $rels     = Read-ZipXml -Archive $Archive -EntryName 'xl/_rels/workbook.xml.rels'
    if (-not $workbook -or -not $rels) { throw 'The file is not an Excel workbook.' }

    $relNs  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $relId  = $workbook.SelectSingleNode('/s:workbook/s:sheets/s:sheet', (Get-SpreadsheetNamespace -Xml $workbook)).GetAttribute('id', $relNs)
    $target = @($rels.SelectNodes('/p:Relationships/p:Relationship', (Get-SpreadsheetNamespace -Xml $rels)) |
            Where-Object { $_.GetAttribute('Id') -eq $relId })[0].GetAttribute('Target')
    $sheetPath = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { "xl/$target" }

    $shared = [System.Collections.Generic.List[string]]::new()
    $sharedXml = Read-ZipXml -Archive $Archive -EntryName 'xl/sharedStrings.xml'
    if ($sharedXml) {
        $sharedNs = Get-SpreadsheetNamespace -Xml $sharedXml
        foreach ($item in $sharedXml.SelectNodes('/s:sst/s:si', $sharedNs)) { $shared.Add((Get-CellText -Node $item -Ns $sharedNs)) }
    }

    $sheet = Read-ZipXml -Archive $Archive -EntryName $sheetPath
    $sheetNs = Get-SpreadsheetNamespace -Xml $sheet

    foreach ($row in $sheet.SelectNodes('/s:worksheet/s:sheetData/s:row', $sheetNs)) {
        $cells = [System.Collections.Generic.Dictionary[int, string]]::new()
        $column = 0
        foreach ($cell in $row.SelectNodes('s:c', $sheetNs)) {
            $reference = $cell.GetAttribute('r')
            # Excel may leave out empty cells, so place each by its reference when present.
            $column = if ($reference) { ConvertFrom-ColumnName ($reference -replace '\d', '') } else { $column + 1 }
            $value = $cell.SelectSingleNode('s:v', $sheetNs)
            $cells[$column] = switch ($cell.GetAttribute('t')) {
                's'         { $shared[[int] $value.InnerText] }
                'inlineStr' { Get-CellText -Node $cell.SelectSingleNode('s:is', $sheetNs) -Ns $sheetNs }
                default     { if ($value) { $value.InnerText } else { '' } }
            }
        }

        $width = 0
        foreach ($c in $cells.Keys) { $width = [Math]::Max($width, $c) }
        $values = [string[]]::new($width)
        foreach ($c in $cells.Keys) { $values[$c - 1] = $cells[$c] }
        , $values
    }
}

function Import-DuplicateReport {
    <#
    .SYNOPSIS
        Reads a report written by Export-DuplicateReport back into duplicate sets.
    .OUTPUTS
        The same objects Find-DuplicateFile returns.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not [System.IO.File]::Exists($Path)) { throw "Report '$Path' was not found." }

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try { $rows = @(Get-WorksheetRow -Archive $zip) }
        finally { $zip.Dispose() }
    }
    catch [System.IO.InvalidDataException] { throw "'$Path' is not an Excel workbook." }
    finally { $stream.Dispose() }

    $expected = @($script:FixedColumns | ForEach-Object { $_.Header })
    if ($rows.Count -eq 0 -or $rows[0].Count -lt $expected.Count -or
        (Compare-Object $expected ($rows[0][0..($expected.Count - 1)]) -SyncWindow 0)) {
        throw "'$Path' is not a duplicates report: its header row is not '$($expected -join ', ')'."
    }

    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $firstLocation = $expected.Count
    for ($r = 1; $r -lt $rows.Count; $r++) {
        $row = $rows[$r]
        if ($row.Count -lt $firstLocation -or -not $row[0]) { continue }
        $folders = if ($row.Count -gt $firstLocation) { @($row[$firstLocation..($row.Count - 1)] | Where-Object { $_ }) } else { @() }
        [pscustomobject] @{
            FileName      = $row[0]
            LastWriteTime = [datetime]::FromOADate([double]::Parse($row[1], $invariant))
            SizeBytes     = [long] [double]::Parse($row[2], $invariant)
            MD5           = $row[3]
            Count         = $folders.Count
            Folders       = [string[]] $folders
        }
    }
}

function Find-FileIgnoringCase {
    # The file in $Folder whose name matches $FileName ignoring case, or $null.
    # Only needed on case-sensitive file systems (Linux, some macOS volumes).
    param([Parameter(Mandatory)] [string] $Folder, [Parameter(Mandatory)] [string] $FileName)
    if (-not [System.IO.Directory]::Exists($Folder)) { return $null }
    foreach ($candidate in [System.IO.Directory]::GetFiles($Folder)) {
        if ([System.StringComparer]::OrdinalIgnoreCase.Equals([System.IO.Path]::GetFileName($candidate), $FileName)) {
            return [System.IO.FileInfo] $candidate
        }
    }
    $null
}

function Test-DuplicateCopy {
    <#
        Checks one recorded copy without reading its contents:
          Present     - still there with the same size and saved date
          Missing     - no longer there
          Changed     - still there but its size or saved date changed
          Unavailable - its drive or network share cannot be reached (kept as is)
    #>
    param(
        [Parameter(Mandatory)] [string] $Folder,
        [Parameter(Mandatory)] [string] $FileName,
        [Parameter(Mandatory)] [long] $SizeBytes,
        [Parameter(Mandatory)] [datetime] $LastWriteTime
    )

    try {
        $path = [System.IO.Path]::Combine($Folder, $FileName)
        $file = if ([System.IO.File]::Exists($path)) { [System.IO.FileInfo] $path } else { Find-FileIgnoringCase -Folder $Folder -FileName $FileName }
        if (-not $file) {
            $root = [System.IO.Path]::GetPathRoot($Folder)
            if ($root -and -not [System.IO.Directory]::Exists($root)) { return 'Unavailable' }
            return 'Missing'
        }
        $size  = $file.Length
        $ticks = $file.LastWriteTime.Ticks
    }
    catch [System.IO.FileNotFoundException] { return 'Missing' }  # deleted while being checked
    catch [System.UnauthorizedAccessException], [System.IO.IOException], [System.Security.SecurityException] {
        return 'Unavailable'
    }

    # Whole seconds, in exact integer arithmetic (as when scanning).
    $ticksPerSecond = [System.TimeSpan]::TicksPerSecond
    $savedTicks = $LastWriteTime.Ticks
    if ($size -ne $SizeBytes -or
        ($ticks - ($ticks % $ticksPerSecond)) -ne ($savedTicks - ($savedTicks % $ticksPerSecond))) { return 'Changed' }
    'Present'
}

function Update-DuplicateReport {
    <#
    .SYNOPSIS
        Re-checks every copy listed in an existing report and removes the ones that
        no longer exist, without rescanning the folders.
    .DESCRIPTION
        Each copy is checked with a single file lookup (no contents are read or
        downloaded). Copies that are missing, or whose size or saved date changed,
        are removed; rows left with fewer than two copies are removed. Copies on a
        drive or network share that cannot be reached are kept. The report is
        rewritten in place only when something changed. Supports -WhatIf.
    .OUTPUTS
        A summary object; its DuplicateSet property holds the rows that remain.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)] [string] $Path)

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
    $sets = @(Import-DuplicateReport -Path $Path)

    $kept = [System.Collections.Generic.List[object]]::new()
    $checked = 0; $removed = 0; $unavailable = 0; $rowsRemoved = 0
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $lastShownMs = - $script:ProgressIntervalMs

    for ($i = 0; $i -lt $sets.Count; $i++) {
        $set = $sets[$i]
        if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
            $lastShownMs = $timer.ElapsedMilliseconds
            Write-Progress -Id 3 -Activity 'Validating report' -Status "Row $($i + 1) of $($sets.Count)" `
                -CurrentOperation $set.FileName -PercentComplete ([int] (100 * $i / $sets.Count))
        }

        $present = [System.Collections.Generic.List[string]]::new()
        foreach ($folder in $set.Folders) {
            $checked++
            $state = Test-DuplicateCopy -Folder $folder -FileName $set.FileName -SizeBytes $set.SizeBytes -LastWriteTime $set.LastWriteTime
            if ($state -eq 'Present') { $present.Add($folder); continue }
            if ($state -eq 'Unavailable') {
                $unavailable++
                $present.Add($folder)
                Write-Warning "Cannot reach '$folder'; keeping its copy of '$($set.FileName)'."
                continue
            }
            $removed++
            Write-Verbose "$state`: '$([System.IO.Path]::Combine($folder, $set.FileName))'"
        }

        if ($present.Count -ge 2) {
            $kept.Add([pscustomobject] @{
                FileName      = $set.FileName
                LastWriteTime = $set.LastWriteTime
                SizeBytes     = $set.SizeBytes
                MD5           = $set.MD5
                Count         = $present.Count
                Folders       = $present.ToArray()
            })
        }
        else { $rowsRemoved++ }
    }
    Write-Progress -Id 3 -Activity 'Validating report' -Completed

    $saved = $false
    if (($removed -gt 0 -or $rowsRemoved -gt 0) -and
        $PSCmdlet.ShouldProcess($Path, "Remove $removed copies and $rowsRemoved rows")) {
        Export-DuplicateReport -DuplicateSet $kept.ToArray() -Path $Path
        $saved = $true
    }

    [pscustomobject] @{
        Path                 = $Path
        RowsChecked          = $sets.Count
        CopiesChecked        = $checked
        CopiesRemoved        = $removed
        CopiesUnavailable    = $unavailable
        RowsRemoved          = $rowsRemoved
        RowsRemaining        = $kept.Count
        Saved                = $saved
        DuplicateSet         = $kept.ToArray()
    }
}

#endregion

Export-ModuleMember -Function Get-FileInventory, Find-DuplicateFile, Export-DuplicateReport, ConvertTo-ColumnName,
    Import-DuplicateReport, Update-DuplicateReport
