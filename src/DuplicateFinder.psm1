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

#region Scanning

function Get-FileInventory {
    <#
    .SYNOPSIS
        Recursively lists every file below a folder, reporting the folder being scanned.
    .DESCRIPTION
        Walks the tree iteratively so that deep trees cannot overflow the call stack.
        Folders that cannot be read are reported as warnings and skipped. Directory
        symlinks / junctions are not followed, which prevents infinite loops.
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
            $files      = $folder.GetFiles()
            $subFolders = $folder.GetDirectories()
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
            if ($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                Write-Verbose "Not following link '$($sub.FullName)'"
                continue
            }
            $pending.Push($sub)
        }
    }

    Write-Progress -Id 1 -Activity 'Scanning folders' -Completed
}

#endregion

#region Matching

function Get-SavedDateKey {
    # Saved date in UTC, truncated to the whole second. Copies made to network
    # shares or other file systems frequently lose sub-second precision.
    param([Parameter(Mandatory)] [System.IO.FileInfo] $File)

    $ticks = $File.LastWriteTimeUtc.Ticks
    $ticks - ($ticks % [System.TimeSpan]::TicksPerSecond)
}

function Group-ByKey {
    # Groups items into lists keyed by the result of a script block,
    # returning only the groups that hold more than one item.
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $InputItems,
        [Parameter(Mandatory)] [scriptblock] $KeySelector
    )

    $groups = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)

    foreach ($item in $InputItems) {
        $key = [string] (& $KeySelector $item)
        $list = $null
        if (-not $groups.TryGetValue($key, [ref] $list)) {
            $list = [System.Collections.Generic.List[object]]::new()
            $groups.Add($key, $list)
        }
        $list.Add($item)
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
        [System.IO.FileInfo[]] $File
    )

    # Stage 1: name + saved date. The date key is all digits, so '|' is a safe separator.
    $nameDateGroups = @(Group-ByKey -InputItems $File -KeySelector {
            param($f) '{0}|{1}' -f (Get-SavedDateKey -File $f), $f.Name
        })

    # Stage 2: size. A cheap check that avoids hashing files that cannot match.
    $candidateGroups = @(foreach ($group in $nameDateGroups) {
            Group-ByKey -InputItems $group -KeySelector { param($f) $f.Length }
        })

    $toHash = 0
    foreach ($group in $candidateGroups) { $toHash += $group.Count }
    $hashed = 0

    # Stage 3: MD5, only for files that already match on name, date and size.
    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $candidateGroups) {
        $hashedFiles = [System.Collections.Generic.List[object]]::new()
        foreach ($f in $group) {
            $hashed++
            Write-Progress -Id 2 -Activity 'Comparing MD5 hashes' `
                -Status "File $hashed of $toHash" -CurrentOperation $f.FullName `
                -PercentComplete ([int] (100 * $hashed / [Math]::Max($toHash, 1)))
            try {
                $md5 = (Get-FileHash -LiteralPath $f.FullName -Algorithm MD5 -ErrorAction Stop).Hash
                $hashedFiles.Add([pscustomobject] @{ File = $f; MD5 = $md5 })
            }
            catch {
                Write-Warning "Could not hash '$($f.FullName)': $($_.Exception.Message)"
            }
        }

        foreach ($set in @(Group-ByKey -InputItems $hashedFiles.ToArray() -KeySelector { param($h) $h.MD5 })) {
            $first = $set[0].File
            $results.Add([pscustomobject] @{
                FileName      = $first.Name
                LastWriteTime = $first.LastWriteTime
                SizeBytes     = $first.Length
                MD5           = $set[0].MD5
                Count         = $set.Count
                Folders       = [string[]] @($set | ForEach-Object { $_.File.DirectoryName } | Sort-Object)
            })
        }
    }

    Write-Progress -Id 2 -Activity 'Comparing MD5 hashes' -Completed
    $results | Sort-Object FileName, LastWriteTime, MD5
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

Export-ModuleMember -Function Get-FileInventory, Find-DuplicateFile, Export-DuplicateReport, ConvertTo-ColumnName
