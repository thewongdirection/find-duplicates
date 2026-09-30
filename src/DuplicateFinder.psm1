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
$script:FolderColumns      = @(
    @{ Header = 'Folder Name';  Width = 40 }
    @{ Header = 'Files';        Width = 10 }
    @{ Header = 'Sub Folders';  Width = 12 }
    @{ Header = 'Size (bytes)'; Width = 16 }
    @{ Header = 'Copies';       Width = 8  }
)
$script:FileSheetName   = 'Duplicates'
$script:FolderSheetName = 'Duplicate Folders'

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
        [string[]] $ExcludeFile = @(),

        # When given, receives one record per folder listed: Path and Readable.
        # Needed to compare folder trees, including empty and unreadable folders.
        [System.Collections.Generic.List[object]] $FolderInfo
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
            if ($null -ne $FolderInfo) { $FolderInfo.Add([pscustomobject] @{ Path = $folder.FullName; Readable = $false }) }
            continue
        }
        if ($null -ne $FolderInfo) { $FolderInfo.Add([pscustomobject] @{ Path = $folder.FullName; Readable = $true }) }

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

function Get-InnermostMessage {
    # The message of the innermost exception, so failures read the same however they surfaced.
    param([Parameter(Mandatory)] [System.Exception] $Exception)
    while ($Exception.InnerException) { $Exception = $Exception.InnerException }
    $Exception.Message
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
        With $Cache, files already in it are not read again and new hashes are added to it.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]] $Path,
        [ValidateRange(1, 64)] [int] $ThrottleLimit = 1,
        [System.Collections.Generic.Dictionary[string, string]] $Cache
    )

    $requested = $Path
    if ($null -ne $Cache) { $Path = [string[]] @($Path | Where-Object { -not $Cache.ContainsKey($_) }) }

    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $activity = 'Comparing MD5 hashes'
    $done = 0

    if ($ThrottleLimit -eq 1) {
        foreach ($p in $Path) {
            $done++
            Write-Progress -Id 2 -Activity $activity -Status "File $done of $($Path.Count)" -CurrentOperation $p `
                -PercentComplete ([int] (100 * $done / $Path.Count))
            try { $map[$p] = Get-FileMd5 -Path $p }
            catch { Write-Warning "Could not hash '$p': $(Get-InnermostMessage $_.Exception)" }
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
                catch { Write-Warning "Could not hash '$($job.Path)': $(Get-InnermostMessage $_.Exception)" }
                finally { $job.Shell.Dispose() }
            }
        }
        finally {
            foreach ($job in $inFlight) { $job.Shell.Dispose() }
            $pool.Dispose()
        }
    }

    Write-Progress -Id 2 -Activity $activity -Completed
    if ($null -ne $Cache) {
        foreach ($entry in $map.GetEnumerator()) { $Cache[$entry.Key] = $entry.Value }
        foreach ($p in $requested) { if ($Cache.ContainsKey($p)) { $map[$p] = $Cache[$p] } }
    }
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
        [int] $ThrottleLimit = 1,

        # Hashes shared with Find-DuplicateFolder so no file is read twice.
        [System.Collections.Generic.Dictionary[string, string]] $Md5Cache
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
    $md5ByPath = Get-FileMd5Map -Path $candidates.ToArray() -ThrottleLimit $ThrottleLimit -Cache $Md5Cache

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

#region Duplicate folders

# Same ordering as the Python port: folder name (ordinal, ignoring case), size, first location.
$script:ByFolderSet = [System.Comparison[object]] {
    param($x, $y)
    $order = [System.StringComparer]::OrdinalIgnoreCase.Compare($x.FolderName, $y.FolderName)
    if ($order -eq 0) { $order = $x.SizeBytes.CompareTo($y.SizeBytes) }
    if ($order -eq 0) { $order = [System.StringComparer]::OrdinalIgnoreCase.Compare($x.Folders[0], $y.Folders[0]) }
    $order
}

function Get-FolderTree {
    # Indexes a scan: the sub folders and files of every folder, and which folders were readable.
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.IO.FileInfo[]] $File,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Folder
    )

    $children = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]::new([System.StringComparer]::Ordinal)
    $filesIn  = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[System.IO.FileInfo]]]::new([System.StringComparer]::Ordinal)
    $readable = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

    foreach ($record in $Folder) {
        $children[$record.Path] = [System.Collections.Generic.List[string]]::new()
        $filesIn[$record.Path]  = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        if ($record.Readable) { $null = $readable.Add($record.Path) }
    }
    foreach ($record in $Folder) {
        $parent = [System.IO.Path]::GetDirectoryName($record.Path)
        if ($parent -and $children.ContainsKey($parent)) { $children[$parent].Add($record.Path) }
    }
    foreach ($f in $File) {
        if ($filesIn.ContainsKey($f.DirectoryName)) { $filesIn[$f.DirectoryName].Add($f) }
    }

    # Deepest first: a sub folder's path is always longer than its parent's.
    $order = [string[]] @($children.Keys)
    $lengths = [int[]] @($order | ForEach-Object { - $_.Length })
    [System.Array]::Sort($lengths, $order, [System.Collections.Comparer]::DefaultInvariant)

    [pscustomobject] @{ Children = $children; Files = $filesIn; Readable = $readable; DeepestFirst = $order }
}

function Get-FolderSignature {
    <#
        Fingerprints folder trees bottom-up. A folder's fingerprint covers the name (ignoring
        case), size and saved second of every file below it, the name of every sub folder
        (including empty ones), and with $Md5 each file's MD5. A folder whose tree could not
        be fully read, or has a file missing from $Md5, gets no fingerprint.
        Returns path -> @{ Signature; FileCount; FolderCount; SizeBytes }.
    #>
    param(
        [Parameter(Mandatory)] [object] $Tree,
        [System.Collections.Generic.Dictionary[string, string]] $Md5,
        [System.Collections.Generic.HashSet[string]] $Scope
    )

    $ticksPerSecond = [System.TimeSpan]::TicksPerSecond
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $result = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::Ordinal)
    try {
        foreach ($folder in $Tree.DeepestFirst) {
            if ($null -ne $Scope -and -not $Scope.Contains($folder)) { continue }
            $info = [pscustomobject] @{ Signature = $null; FileCount = 0; FolderCount = 0; SizeBytes = [long] 0 }
            $result[$folder] = $info
            if (-not $Tree.Readable.Contains($folder)) { continue }

            $lines = [System.Collections.Generic.List[string]]::new()
            $complete = $true
            foreach ($f in $Tree.Files[$folder]) {
                $ticks = $f.LastWriteTimeUtc.Ticks
                $line = "F|$($f.Name.ToUpperInvariant())|$($ticks - ($ticks % $ticksPerSecond))|$($f.Length)"
                if ($null -ne $Md5) {
                    if (-not $Md5.ContainsKey($f.FullName)) { $complete = $false; break }
                    $line += "|$($Md5[$f.FullName])"
                }
                $lines.Add($line)
                $info.FileCount++
                $info.SizeBytes += $f.Length
            }
            foreach ($sub in $Tree.Children[$folder]) {
                $child = $result[$sub]
                if (-not $child.Signature) { $complete = $false; break }
                $lines.Add("D|$([System.IO.Path]::GetFileName($sub).ToUpperInvariant())|$($child.Signature)")
                $info.FileCount   += $child.FileCount
                $info.FolderCount += $child.FolderCount + 1
                $info.SizeBytes   += $child.SizeBytes
            }
            if (-not $complete) { continue }

            $sorted = $lines.ToArray()
            [System.Array]::Sort($sorted, [System.StringComparer]::Ordinal)
            $bytes = [System.Text.Encoding]::UTF8.GetBytes(($sorted -join "`n"))
            $info.Signature = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
        }
    }
    finally { $sha.Dispose() }
    $result
}

function Find-DuplicateFolder {
    <#
    .SYNOPSIS
        Finds sets of folders that have the same name and exactly the same contents.
    .DESCRIPTION
        Two folders are duplicates when their names match (ignoring case), they contain the
        same tree of file and sub folder names, and every file at the same relative path is
        a duplicate by the file rule (name, saved date and MD5). Their total sizes therefore
        match too. Files are only hashed inside folders whose names, sizes and saved dates
        already match. Only the top-most duplicates are reported: a set is left out when
        every one of its folders sits inside a folder that is itself a reported duplicate.
        Folders with no files anywhere below them are not reported.
    .OUTPUTS
        One object per duplicate set: FolderName, FileCount, FolderCount, SizeBytes, Count
        and Folders (full path of every copy, sorted).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.IO.FileInfo[]] $File,

        # Folder records from Get-FileInventory -FolderInfo.
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Folder,

        [switch] $SkipCloudOnly,

        [ValidateRange(1, 64)]
        [int] $ThrottleLimit = 1,

        [System.Collections.Generic.Dictionary[string, string]] $Md5Cache
    )

    $tree = Get-FolderTree -File $File -Folder $Folder

    # Pass 1: names, sizes and saved dates only; nothing is read.
    $cheap = Get-FolderSignature -Tree $tree
    $candidates = [System.Collections.Generic.List[string]]::new()
    $keys = [System.Collections.Generic.List[string]]::new()
    foreach ($path in $tree.DeepestFirst) {
        $info = $cheap[$path]
        if ($info.Signature -and $info.FileCount -gt 0) {
            $candidates.Add($path)
            $keys.Add("$([System.IO.Path]::GetFileName($path).ToUpperInvariant())|$($info.Signature)")
        }
    }
    $candidateGroups = @(Group-ByKey -InputItems $candidates.ToArray() -Key $keys.ToArray())
    if ($candidateGroups.Count -eq 0) { return }

    # Pass 2: hash every file below the candidates (hashes from the file scan are reused).
    $scope = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    $pending = [System.Collections.Generic.Stack[string]]::new()
    foreach ($group in $candidateGroups) { foreach ($path in $group) { $pending.Push($path) } }
    while ($pending.Count -gt 0) {
        $path = $pending.Pop()
        if ($scope.Add($path)) { foreach ($sub in $tree.Children[$path]) { $pending.Push($sub) } }
    }

    $toHash = [System.Collections.Generic.List[string]]::new()
    $skipped = 0
    foreach ($path in $scope) {
        foreach ($f in $tree.Files[$path]) {
            if ($SkipCloudOnly -and (Test-CloudOnlyFile -File $f)) {
                $skipped++
                Write-Verbose "Not downloading online-only file '$($f.FullName)'"
            }
            else { $toHash.Add($f.FullName) }
        }
    }
    if ($skipped) {
        Write-Warning "$skipped online-only cloud file(s) were not checked; folders containing them are not reported."
    }
    $md5 = Get-FileMd5Map -Path $toHash.ToArray() -ThrottleLimit $ThrottleLimit -Cache $Md5Cache
    $full = Get-FolderSignature -Tree $tree -Md5 $md5 -Scope $scope

    # Group the candidates again, now by contents.
    $confirmed = [System.Collections.Generic.List[string]]::new()
    $keys = [System.Collections.Generic.List[string]]::new()
    foreach ($group in $candidateGroups) {
        foreach ($path in $group) {
            if ($full[$path].Signature) {
                $confirmed.Add($path)
                $keys.Add("$([System.IO.Path]::GetFileName($path).ToUpperInvariant())|$($full[$path].Signature)")
            }
        }
    }
    $sets = @(Group-ByKey -InputItems $confirmed.ToArray() -Key $keys.ToArray())

    # Report only the top-most duplicates.
    $duplicated = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($set in $sets) { foreach ($path in $set) { $null = $duplicated.Add($path) } }

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($set in $sets) {
        $implied = $true
        foreach ($path in $set) {
            if (-not $duplicated.Contains([System.IO.Path]::GetDirectoryName($path))) { $implied = $false; break }
        }
        if ($implied) { continue }

        $folders = Get-SortedFolder -Path ([string[]] $set)
        $info = $full[$folders[0]]
        $results.Add([pscustomobject] @{
            FolderName  = [System.IO.Path]::GetFileName($folders[0])
            FileCount   = $info.FileCount
            FolderCount = $info.FolderCount
            SizeBytes   = $info.SizeBytes
            Count       = $folders.Count
            Folders     = $folders
        })
    }

    $results.Sort($script:ByFolderSet)
    $results
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

function ConvertTo-WorksheetData {
    # Lays out one sheet: fixed columns, then as many "Location N" columns as the longest row needs.
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [object[]] $Column,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Row,
        [Parameter(Mandatory)] [string] $Noun
    )

    $maxCopies = 1
    foreach ($values in $Row) { $maxCopies = [Math]::Max($maxCopies, $values.Count - $Column.Count) }

    $columnCount = $Column.Count + $maxCopies
    if ($columnCount -gt $script:ExcelMaxColumns) {
        throw "A $Noun has $maxCopies copies; Excel supports at most $($script:ExcelMaxColumns - $Column.Count) location columns."
    }
    if ($Row.Count + 1 -gt $script:ExcelMaxRows) {
        throw "Found $($Row.Count) duplicated $($Noun)s; Excel supports at most $($script:ExcelMaxRows - 1) rows."
    }

    [pscustomobject] @{
        Name       = $Name
        Headers    = [string[]] (@($Column | ForEach-Object { $_.Header }) + @(1..$maxCopies | ForEach-Object { "Location $_" }))
        Widths     = [int[]] (@($Column | ForEach-Object { $_.Width }) + @(1..$maxCopies | ForEach-Object { $script:LocationColumnWidth }))
        Rows       = $Row
        LastColumn = ConvertTo-ColumnName $columnCount
        LastRow    = $Row.Count + 1
    }
}

function Write-WorksheetXml {
    param(
        [Parameter(Mandatory)] [System.Xml.XmlWriter] $Writer,
        [Parameter(Mandatory)] [object] $Sheet
    )

    $ns        = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
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
    $Writer.WriteStartElement('cols', $ns)
    for ($i = 0; $i -lt $Sheet.Widths.Count; $i++) {
        $Writer.WriteStartElement('col', $ns)
        $Writer.WriteAttributeString('min', [string] ($i + 1))
        $Writer.WriteAttributeString('max', [string] ($i + 1))
        $Writer.WriteAttributeString('width', [string] $Sheet.Widths[$i])
        $Writer.WriteAttributeString('customWidth', '1')
        $Writer.WriteEndElement()
    }
    $Writer.WriteEndElement()

    $Writer.WriteStartElement('sheetData', $ns)

    # Header row.
    $Writer.WriteStartElement('row', $ns)
    $Writer.WriteAttributeString('r', '1')
    for ($c = 0; $c -lt $Sheet.Headers.Count; $c++) {
        Write-Cell -Writer $Writer -Reference "$(ConvertTo-ColumnName ($c + 1))1" -Value $Sheet.Headers[$c] -Style $styleBold
    }
    $Writer.WriteEndElement()

    # One row per duplicated item; one column per copy.
    $rowNumber = 1
    foreach ($values in $Sheet.Rows) {
        $rowNumber++
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
    $Writer.WriteAttributeString('ref', "A1:$($Sheet.LastColumn)$($Sheet.LastRow)")
    $Writer.WriteEndElement()

    $Writer.WriteEndElement()  # worksheet
}

function Export-DuplicateReport {
    <#
    .SYNOPSIS
        Saves duplicate sets to an .xlsx workbook. Needs neither Excel nor extra modules.
    .DESCRIPTION
        Sheet "Duplicates": one row per duplicated file. Columns: File Name, Last Modified,
        Size (bytes), MD5, Copies, then "Location 1..N" holding the full folder path of
        every copy.
        Sheet "Duplicate Folders" (only when -FolderSet is given): one row per duplicated
        folder. Columns: Folder Name, Files, Sub Folders, Size (bytes), Copies, then
        "Location 1..N" holding the full path of every copy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $DuplicateSet,

        [Parameter(Mandatory)]
        [string] $Path,

        # Duplicate folder sets (from Find-DuplicateFolder); adds the "Duplicate Folders" sheet.
        [AllowEmptyCollection()]
        [object[]] $FolderSet
    )

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)

    $fileRows = @(foreach ($set in $DuplicateSet) {
            , (@($set.FileName, $set.LastWriteTime, [long] $set.SizeBytes, $set.MD5, [int] $set.Count) + @($set.Folders))
        })
    $sheets = [System.Collections.Generic.List[object]]::new()
    $sheets.Add((ConvertTo-WorksheetData -Name $script:FileSheetName -Column $script:FixedColumns -Row $fileRows -Noun 'file'))

    if ($PSBoundParameters.ContainsKey('FolderSet') -and $null -ne $FolderSet) {
        $folderRows = @(foreach ($set in $FolderSet) {
                , (@($set.FolderName, [int] $set.FileCount, [int] $set.FolderCount, [long] $set.SizeBytes, [int] $set.Count) + @($set.Folders))
            })
        $sheets.Add((ConvertTo-WorksheetData -Name $script:FolderSheetName -Column $script:FolderColumns -Row $folderRows -Noun 'folder'))
    }

    $sheetEntries = ''; $overrides = ''; $sheetRels = ''; $filters = ''
    for ($i = 1; $i -le $sheets.Count; $i++) {
        $sheet = $sheets[$i - 1]
        $sheetEntries += "<sheet name=`"$($sheet.Name)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
        $overrides += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $sheetRels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
        $filters += "<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$($i - 1)`" hidden=`"1`">'$($sheet.Name)'!`$A`$1:`$$($sheet.LastColumn)`$$($sheet.LastRow)</definedName>"
    }
    $stylesId = "rId$($sheets.Count + 1)"

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
                    $overrides +
                    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
                    '</Types>')

                Write-ZipTextEntry -Archive $zip -EntryName '_rels/.rels' -Content (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                    '</Relationships>')

                Write-ZipTextEntry -Archive $zip -EntryName 'xl/workbook.xml' -Content (
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                    "<sheets>$sheetEntries</sheets><definedNames>$filters</definedNames>" +
                    '</workbook>')

                Write-ZipTextEntry -Archive $zip -EntryName 'xl/_rels/workbook.xml.rels' -Content (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    $sheetRels +
                    "<Relationship Id=`"$stylesId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>" +
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

                for ($i = 1; $i -le $sheets.Count; $i++) {
                    $sheet = $sheets[$i - 1]
                    Write-ZipXmlEntry -Archive $zip -EntryName "xl/worksheets/sheet$i.xml" -Body {
                        param($w)
                        Write-WorksheetXml -Writer $w -Sheet $sheet
                    }
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

    # Reports never contain DTDs, so refuse them (no entity expansion, no external resolution).
    $settings = [System.Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.CloseInput = $true
    $reader = [System.Xml.XmlReader]::Create($entry.Open(), $settings)
    try {
        $xml = [System.Xml.XmlDocument]::new()
        $xml.XmlResolver = $null
        $xml.Load($reader)
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

function Get-WorkbookSheet {
    # The workbook's sheets in order, each with its Name and the Path of its part.
    param([Parameter(Mandatory)] [System.IO.Compression.ZipArchive] $Archive)

    $workbook = Read-ZipXml -Archive $Archive -EntryName 'xl/workbook.xml'
    $rels     = Read-ZipXml -Archive $Archive -EntryName 'xl/_rels/workbook.xml.rels'
    if (-not $workbook -or -not $rels) { throw 'The file is not an Excel workbook.' }

    $relNs   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $targets = @{}
    foreach ($rel in $rels.SelectNodes('/p:Relationships/p:Relationship', (Get-SpreadsheetNamespace -Xml $rels))) {
        $targets[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target')
    }
    foreach ($sheet in $workbook.SelectNodes('/s:workbook/s:sheets/s:sheet', (Get-SpreadsheetNamespace -Xml $workbook))) {
        $target = $targets[$sheet.GetAttribute('id', $relNs)]
        [pscustomobject] @{
            Name = $sheet.GetAttribute('name')
            Path = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { "xl/$target" }
        }
    }
}

function Get-SharedString {
    # The workbook's shared strings (Excel stores text there when it saves a workbook).
    param([Parameter(Mandatory)] [System.IO.Compression.ZipArchive] $Archive)

    $shared = [System.Collections.Generic.List[string]]::new()
    $sharedXml = Read-ZipXml -Archive $Archive -EntryName 'xl/sharedStrings.xml'
    if ($sharedXml) {
        $ns = Get-SpreadsheetNamespace -Xml $sharedXml
        foreach ($item in $sharedXml.SelectNodes('/s:sst/s:si', $ns)) { $shared.Add((Get-CellText -Node $item -Ns $ns)) }
    }
    , $shared
}

function Get-WorksheetRow {
    # One worksheet as rows of cell text (a string[] per row, one entry per column).
    # Handles workbooks written by this tool and the same workbook after Excel saved it.
    param(
        [Parameter(Mandatory)] [System.IO.Compression.ZipArchive] $Archive,
        [Parameter(Mandatory)] [string] $SheetPath,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.List[string]] $SharedString
    )

    $sheet = Read-ZipXml -Archive $Archive -EntryName $SheetPath
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
                's'         { $SharedString[[int] $value.InnerText] }
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

function Select-ReportRow {
    # Checks a sheet's header row and returns its data rows as (fixed values, folders) pairs.
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Row,
        [Parameter(Mandatory)] [object[]] $Column,
        [Parameter(Mandatory)] [string] $ErrorMessage
    )

    $expected = @($Column | ForEach-Object { $_.Header })
    $headerOk = $Row.Count -gt 0 -and $Row[0].Count -ge $expected.Count
    for ($c = 0; $headerOk -and $c -lt $expected.Count; $c++) {
        $headerOk = [System.StringComparer]::OrdinalIgnoreCase.Equals([string] $Row[0][$c], $expected[$c])
    }
    if (-not $headerOk) { throw "$ErrorMessage '$($expected -join ', ')'." }

    $firstLocation = $expected.Count
    for ($r = 1; $r -lt $Row.Count; $r++) {
        $values = $Row[$r]
        if ($values.Count -lt $firstLocation -or -not $values[0]) { continue }
        # (Not "$x = if ...": that would unroll an empty or one-item array.)
        $folders = [System.Collections.Generic.List[string]]::new()
        for ($c = $firstLocation; $c -lt $values.Count; $c++) { if ($values[$c]) { $folders.Add($values[$c]) } }
        [pscustomobject] @{ Values = $values; Folders = $folders.ToArray() }
    }
}

function Read-DuplicateWorkbook {
    # Reads both sheets of a report. Folders is $null when the report has no folder sheet.
    param([Parameter(Mandatory)] [string] $Path)

    if (-not [System.IO.File]::Exists($Path)) { throw "Report '$Path' was not found." }

    # ReadWrite sharing: the report can be read while Excel has it open.
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $sheets = @(Get-WorkbookSheet -Archive $zip)
            $shared = Get-SharedString -Archive $zip
            $fileRows = @(Get-WorksheetRow -Archive $zip -SheetPath $sheets[0].Path -SharedString $shared)
            $folderSheet = @($sheets | Where-Object { $_.Name -eq $script:FolderSheetName })
            $folderRows = $null
            if ($folderSheet) { $folderRows = @(Get-WorksheetRow -Archive $zip -SheetPath $folderSheet[0].Path -SharedString $shared) }
        }
        finally { $zip.Dispose() }
    }
    catch [System.IO.InvalidDataException] { throw "'$Path' is not an Excel workbook." }
    finally { $stream.Dispose() }

    # Numbers go through [double] then [long]/[int], which rounds half to even like Python's round().
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $files = @(foreach ($row in (Select-ReportRow -Row $fileRows -Column $script:FixedColumns `
                    -ErrorMessage "'$Path' is not a duplicates report: its header row is not")) {
            [pscustomobject] @{
                FileName      = $row.Values[0]
                LastWriteTime = [datetime]::FromOADate([double]::Parse($row.Values[1], $invariant))
                SizeBytes     = [long] [double]::Parse($row.Values[2], $invariant)
                MD5           = $row.Values[3]
                Count         = $row.Folders.Count
                Folders       = $row.Folders
            }
        })

    $folders = $null
    if ($folderSheet) {
        $folders = @(foreach ($row in (Select-ReportRow -Row $folderRows -Column $script:FolderColumns `
                        -ErrorMessage "'$Path' is not a duplicates report: the header row of sheet '$($script:FolderSheetName)' is not")) {
                [pscustomobject] @{
                    FolderName  = $row.Values[0]
                    FileCount   = [int] [double]::Parse($row.Values[1], $invariant)
                    FolderCount = [int] [double]::Parse($row.Values[2], $invariant)
                    SizeBytes   = [long] [double]::Parse($row.Values[3], $invariant)
                    Count       = $row.Folders.Count
                    Folders     = $row.Folders
                }
            })
    }

    [pscustomobject] @{ Files = $files; Folders = $folders }
}

function Import-DuplicateReport {
    <#
    .SYNOPSIS
        Reads the duplicate files back from a report written by Export-DuplicateReport.
    .OUTPUTS
        The same objects Find-DuplicateFile returns.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)

    (Read-DuplicateWorkbook -Path $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)).Files
}

function Import-DuplicateFolderReport {
    <#
    .SYNOPSIS
        Reads the duplicate folders back from a report; nothing when it has no folder sheet.
    .OUTPUTS
        The same objects Find-DuplicateFolder returns.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)

    $folders = (Read-DuplicateWorkbook -Path $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)).Folders
    if ($null -ne $folders) { $folders }
}

function Test-PathRootReachable {
    <#
        True when the drive or network share holding $Folder can be reached. Each root is
        checked once per $Cache: an offline share can take many seconds to time out, and a
        report may list thousands of copies on it.
    #>
    param(
        [Parameter(Mandatory)] [string] $Folder,
        [System.Collections.Generic.Dictionary[string, bool]] $Cache
    )
    $root = [System.IO.Path]::GetPathRoot($Folder)
    if (-not $root) { return $true }
    if ($null -ne $Cache -and $Cache.ContainsKey($root)) { return $Cache[$root] }
    $reachable = [System.IO.Directory]::Exists($root)
    if ($null -ne $Cache) { $Cache[$root] = $reachable }
    $reachable
}

function Find-FileIgnoringCase {
    # The file in $Folder whose name matches $FileName ignoring case, or $null.
    # Only needed on case-sensitive file systems (Linux, some macOS volumes); on Windows
    # File.Exists already ignores case, so the folder is not listed.
    param([Parameter(Mandatory)] [string] $Folder, [Parameter(Mandatory)] [string] $FileName)
    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) { return $null }
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
        [Parameter(Mandatory)] [datetime] $LastWriteTime,
        [System.Collections.Generic.Dictionary[string, bool]] $RootCache
    )

    try {
        if (-not (Test-PathRootReachable -Folder $Folder -Cache $RootCache)) { return 'Unavailable' }
        $path = [System.IO.Path]::Combine($Folder, $FileName)
        $file = if ([System.IO.File]::Exists($path)) { [System.IO.FileInfo] $path } else { Find-FileIgnoringCase -Folder $Folder -FileName $FileName }
        if (-not $file) { return 'Missing' }
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

function Test-DuplicateFolderCopy {
    <#
        Checks one recorded folder copy by listing its tree again (no file contents are read):
          Present     - still there with the same number of files and sub folders and total size
          Missing     - no longer there
          Changed     - still there but its files, sub folders or total size changed
          Unavailable - its drive or share cannot be reached, or part of it cannot be read (kept)
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [int] $FileCount,
        [Parameter(Mandatory)] [int] $FolderCount,
        [Parameter(Mandatory)] [long] $SizeBytes,
        [System.Collections.Generic.Dictionary[string, bool]] $RootCache
    )

    try {
        if (-not (Test-PathRootReachable -Folder $Path -Cache $RootCache)) { return 'Unavailable' }
        if (-not [System.IO.Directory]::Exists($Path)) { return 'Missing' }
        $folders = [System.Collections.Generic.List[object]]::new()
        $files = @(Get-FileInventory -Path $Path -FolderInfo $folders -Verbose:$false -WarningAction SilentlyContinue)
    }
    catch { return 'Unavailable' }

    if (@($folders | Where-Object { -not $_.Readable }).Count -gt 0) { return 'Unavailable' }
    $size = [long] 0
    foreach ($f in $files) { $size += $f.Length }
    if ($files.Count -ne $FileCount -or $folders.Count - 1 -ne $FolderCount -or $size -ne $SizeBytes) { return 'Changed' }
    'Present'
}

function Invoke-CopyCheck {
    <#
        Runs $TestCopy on every copy of every row. Keeps copies that are Present or
        Unavailable, drops Missing and Changed ones, and drops rows left with fewer than
        two copies. Returns the rows kept and the counts.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Set,
        [Parameter(Mandatory)] [scriptblock] $TestCopy,
        [Parameter(Mandatory)] [string] $NameProperty,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Noun,
        [switch] $LocationIsItem
    )

    $kept = [System.Collections.Generic.List[object]]::new()
    $checked = 0; $removed = 0; $unavailable = 0; $rowsRemoved = 0
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $lastShownMs = - $script:ProgressIntervalMs

    for ($i = 0; $i -lt $Set.Count; $i++) {
        $row = $Set[$i]
        $name = $row.$NameProperty
        if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
            $lastShownMs = $timer.ElapsedMilliseconds
            Write-Progress -Id 3 -Activity 'Validating report' -Status "Row $($i + 1) of $($Set.Count)" `
                -CurrentOperation $name -PercentComplete ([int] (100 * $i / $Set.Count))
        }

        $present = [System.Collections.Generic.List[string]]::new()
        foreach ($location in $row.Folders) {
            $checked++
            $state = & $TestCopy $row $location
            if ($state -eq 'Present') { $present.Add($location); continue }
            if ($state -eq 'Unavailable') {
                $unavailable++
                $present.Add($location)
                Write-Warning "Cannot reach '$location'; keeping its copy of $Noun'$name'."
                continue
            }
            $removed++
            $item = if ($LocationIsItem) { $location } else { [System.IO.Path]::Combine($location, $name) }
            Write-Verbose "$state`: '$item'"
        }

        if ($present.Count -ge 2) {
            $copy = $row.PSObject.Copy()
            $copy.Count = $present.Count
            $copy.Folders = $present.ToArray()
            $kept.Add($copy)
        }
        else { $rowsRemoved++ }
    }
    Write-Progress -Id 3 -Activity 'Validating report' -Completed

    [pscustomobject] @{
        Kept = $kept.ToArray(); Checked = $checked; Removed = $removed; Unavailable = $unavailable; RowsRemoved = $rowsRemoved
    }
}

function Update-DuplicateReport {
    <#
    .SYNOPSIS
        Re-checks every copy listed in an existing report and removes the ones that
        no longer exist, without rescanning the folders.
    .DESCRIPTION
        Each file copy is checked with a single file lookup, and each folder copy by
        listing its tree again; no contents are read or downloaded. Copies that are
        missing or changed are removed; rows left with fewer than two copies are
        removed. Copies on a drive or network share that cannot be reached are kept.
        The report is rewritten in place only when something changed. Supports -WhatIf.
    .OUTPUTS
        A summary object; DuplicateSet holds the file rows that remain and
        DuplicateFolderSet the folder rows ($null when the report has no folder sheet).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)] [string] $Path)

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
    $workbook = Read-DuplicateWorkbook -Path $Path
    $rootCache = [System.Collections.Generic.Dictionary[string, bool]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $files = Invoke-CopyCheck -Set $workbook.Files -NameProperty FileName -Noun '' -TestCopy {
        param($row, $location)
        Test-DuplicateCopy -Folder $location -FileName $row.FileName -SizeBytes $row.SizeBytes -LastWriteTime $row.LastWriteTime -RootCache $rootCache
    }

    $folders = $null
    if ($null -ne $workbook.Folders) {
        $folders = Invoke-CopyCheck -Set $workbook.Folders -NameProperty FolderName -Noun 'folder ' -LocationIsItem -TestCopy {
            param($row, $location)
            Test-DuplicateFolderCopy -Path $location -FileCount $row.FileCount -FolderCount $row.FolderCount -SizeBytes $row.SizeBytes -RootCache $rootCache
        }
    }

    $removed = $files.Removed + $files.RowsRemoved
    if ($folders) { $removed += $folders.Removed + $folders.RowsRemoved }

    $saved = $false
    if ($removed -gt 0 -and $PSCmdlet.ShouldProcess($Path, 'Remove missing or changed copies')) {
        $export = @{ DuplicateSet = $files.Kept; Path = $Path }
        if ($folders) { $export.FolderSet = $folders.Kept }
        Export-DuplicateReport @export
        $saved = $true
    }

    $summary = [pscustomobject] @{
        Path                    = $Path
        RowsChecked             = $workbook.Files.Count
        CopiesChecked           = $files.Checked
        CopiesRemoved           = $files.Removed
        CopiesUnavailable       = $files.Unavailable
        RowsRemoved             = $files.RowsRemoved
        RowsRemaining           = $files.Kept.Count
        FolderRowsChecked       = if ($folders) { $workbook.Folders.Count } else { 0 }
        FolderCopiesChecked     = if ($folders) { $folders.Checked } else { 0 }
        FolderCopiesRemoved     = if ($folders) { $folders.Removed } else { 0 }
        FolderCopiesUnavailable = if ($folders) { $folders.Unavailable } else { 0 }
        FolderRowsRemoved       = if ($folders) { $folders.RowsRemoved } else { 0 }
        FolderRowsRemaining     = if ($folders) { $folders.Kept.Count } else { 0 }
        Saved                   = $saved
        DuplicateSet            = $files.Kept
        DuplicateFolderSet      = $null
    }
    # Assigned separately: "= if ..." would unroll an empty array into $null.
    if ($folders) { $summary.DuplicateFolderSet = $folders.Kept }
    $summary
}

#endregion

Export-ModuleMember -Function Get-FileInventory, Find-DuplicateFile, Find-DuplicateFolder, Export-DuplicateReport,
    ConvertTo-ColumnName, Import-DuplicateReport, Import-DuplicateFolderReport, Update-DuplicateReport
