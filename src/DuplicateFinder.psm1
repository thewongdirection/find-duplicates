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
$script:ExcelMaxCellText   = 32767
$script:SpreadsheetMain    = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'  # the SpreadsheetML namespace
$script:FixedColumns       = @(
    @{ Header = 'File Name';     Width = 40 }
    @{ Header = 'Last Modified'; Width = 20 }
    @{ Header = 'UTC Offset';    Width = 11 }
    @{ Header = 'Size (bytes)';  Width = 14 }
    @{ Header = 'MD5';           Width = 34 }
    @{ Header = 'Copies';        Width = 8  }
)
# Reports written before the UTC Offset column; still read, and validated by local time.
$script:LegacyFixedColumns = @($script:FixedColumns | Where-Object { $_.Header -ne 'UTC Offset' })
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
$script:RulesSheetName  = 'Rules'
$script:RulesColumnWidth = 150

# The matching rules, written on the Rules sheet. Shared word for word with the Python port.
$script:RulesIntro = @(
    'Matching rules'
    'These rules decide what counts as a match on the other sheets of this workbook.'
)
$script:FileRulesTitle = "Sheet '$script:FileSheetName': duplicate files"
$script:FileRules = @(
    'A file is listed when another file has ALL of: the same name (ignoring upper/lower case), the same saved date (last modified, to the whole second) and the same contents (MD5 hash).'
    'Each row is one duplicated file. Each Location column is the full path of a folder that holds a copy.'
    'Last Modified is local time on the computer that ran the scan, and UTC Offset its difference from UTC then, so the report can be checked in any time zone.'
    'Files of 0 bytes are included unless the scan used -IgnoreEmptyFiles (Python: --ignore-empty-files).'
)
$script:FolderRulesTitle = "Sheet '$script:FolderSheetName': duplicate folders"
$script:FolderRules = @(
    'A folder is listed when another folder has ALL of: the same name (ignoring upper/lower case), the same tree of files and sub folders (empty sub folders included), and every file matching the file at the same place in the other folder (same name, saved date and MD5).'
    'Only the top-most duplicates are listed: a sub folder is listed on its own only when one of its copies is outside a duplicate folder. Folders that contain no files are not listed.'
    'Each row is one duplicated folder. Each Location column is the full path of one copy.'
)
# Written only when the scan left names or small files out. The label rows are read back, so
# validating leaves the same names out and rewriting the report keeps the settings.
$script:SettingsTitle = 'Scan settings'
$script:SettingsRules = @(
    'Files and folders whose names match a pattern below (* stands for any characters, ? for any one character, ignoring upper/lower case) were not scanned: folders were compared, and are validated, as if they were not there.'
    'Files smaller than the size below are not listed on the Duplicates sheet; folders are compared with all their files.'
)
$script:ExcludeNameLabel = 'Names left out (-Exclude; Python: --exclude)'
$script:MinimumSizeLabel = 'Smallest file listed, in bytes (-MinimumSize or -IgnoreEmptyFiles; Python: --minimum-size or --ignore-empty-files)'

# Attributes Windows sets on cloud placeholders (OneDrive "Files On-Demand" and other
# Cloud Files providers) whose contents are not stored locally. Reading them downloads them.
$script:CloudOnlyAttributes = 0x1000 -bor 0x40000 -bor 0x400000  # Offline | RecallOnOpen | RecallOnDataAccess

# Orderings shared with the Python port (python/find_duplicates) so both tools
# report the same "first" copy and sort rows and locations identically on every OS.
# Case-insensitive order compares upper-cased text by UTF-16 code units; unlike
# StringComparer.OrdinalIgnoreCase, that is the same on .NET Framework (Windows
# PowerShell 5.1) and on .NET (PowerShell 7) for characters beyond U+FFFF.
function Compare-IgnoringCase {
    param([AllowEmptyString()] [string] $X, [AllowEmptyString()] [string] $Y)
    [string]::CompareOrdinal($X.ToUpperInvariant(), $Y.ToUpperInvariant())
}

$script:ByPathIgnoringCase = [System.Comparison[string]] {
    param($x, $y)
    $order = Compare-IgnoringCase $x $y
    if ($order -eq 0) { $order = [string]::CompareOrdinal($x, $y) }  # a fixed order for names differing only in case
    $order
}

function Get-SortedDuplicateSet {
    # Duplicate sets ordered by file name (ignoring case, as Compare-IgnoringCase does), then
    # saved date, then MD5. Sorted on precomputed keys compared ordinally: a script block
    # comparer would run far too often for large results. The separator cannot occur in file
    # names and sorts before every other character, so a name sorts before its extensions.
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $DuplicateSet)
    if ($DuplicateSet.Count -gt 1) {
        $separator = [string] [char] 0
        $keys = [string[]] @(foreach ($set in $DuplicateSet) {
                $set.FileName.ToUpperInvariant() + $separator + $set.LastWriteTime.Ticks.ToString('D19') + $separator + $set.MD5
            })
        # The casts pick the non-generic overload: with the generic one, PowerShell passes a
        # copy of the items array and only the keys end up sorted.
        [System.Array]::Sort([System.Array] $keys, [System.Array] $DuplicateSet, [System.Collections.IComparer] [System.StringComparer]::Ordinal)
    }
    , $DuplicateSet
}

#region Parallel work

$script:ModulePath = $PSCommandPath

# Runs in each worker runspace: takes items from the pool's queue until it is completed and
# runs the work script block on each, in this module's scope, passing back the value or the
# failure. One runspace per worker, not per item, keeps the cost per item low.
$script:WorkerLoop = {
    param($Queue, $Results, [string] $Work, [string] $ModulePath)
    $module = Get-Module | Where-Object { $_.Path -eq $ModulePath } | Select-Object -First 1
    if (-not $module) { throw "The module '$ModulePath' did not load in a worker runspace." }
    $run = $module.NewBoundScriptBlock([scriptblock]::Create($Work))
    foreach ($item in $Queue.GetConsumingEnumerable()) {
        try { $Results.Add([pscustomobject] @{ Item = $item; Value = (& $run $item); Error = $null }) }
        catch { $Results.Add([pscustomobject] @{ Item = $item; Value = $null; Error = $_.Exception }) }
    }
}

function Open-WorkerPool {
    <#
        Starts $ThrottleLimit worker runspaces, each with this module loaded, that run $Work
        (a script block taking one item) on every item added to the pool's Queue. Take the
        results with Receive-WorkerResult; always finish with Close-WorkerPool.
    #>
    param(
        [Parameter(Mandatory)] [scriptblock] $Work,
        [Parameter(Mandatory)] [ValidateRange(2, 64)] [int] $ThrottleLimit
    )

    $state = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $state.ImportPSModule([string[]] @($script:ModulePath))
    $runspaces = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $state, $Host)
    $runspaces.Open()
    $pool = [pscustomobject] @{
        Queue     = [System.Collections.Concurrent.BlockingCollection[object]]::new()
        Results   = [System.Collections.Concurrent.BlockingCollection[object]]::new()
        Workers   = [System.Collections.Generic.List[object]]::new()
        Runspaces = $runspaces
    }
    for ($w = 0; $w -lt $ThrottleLimit; $w++) {
        $shell = [System.Management.Automation.PowerShell]::Create()
        $shell.RunspacePool = $runspaces
        $null = $shell.AddScript($script:WorkerLoop.ToString()).AddArgument($pool.Queue).AddArgument($pool.Results).AddArgument($Work.ToString()).AddArgument($script:ModulePath)
        $pool.Workers.Add([pscustomobject] @{ Shell = $shell; Handle = $shell.BeginInvoke() })
    }
    $pool
}

function Receive-WorkerResult {
    # The pool's next result: Item, Value, and Error (the exception when the work failed on
    # that item). Throws when every worker has stopped without finishing its items.
    param([Parameter(Mandatory)] [object] $Pool)
    $result = $null
    while (-not $Pool.Results.TryTake([ref] $result, 50)) {
        $running = $false
        foreach ($worker in $Pool.Workers) { if (-not $worker.Handle.IsCompleted) { $running = $true } }
        if (-not $running -and $Pool.Results.Count -eq 0) {
            foreach ($worker in $Pool.Workers) { $null = $worker.Shell.EndInvoke($worker.Handle) }  # rethrows why
            throw 'The parallel workers stopped before finishing their work.'
        }
    }
    $result
}

function Close-WorkerPool {
    # Drops any items still queued (after an error or Ctrl+C), lets each worker finish its
    # current item, and frees the runspaces.
    param([Parameter(Mandatory)] [object] $Pool)
    $Pool.Queue.CompleteAdding()
    $unused = $null
    while ($Pool.Queue.TryTake([ref] $unused)) { $unused = $null }
    foreach ($worker in $Pool.Workers) { $worker.Shell.Dispose() }
    $Pool.Runspaces.Dispose()
    $Pool.Queue.Dispose()
    $Pool.Results.Dispose()
}

#endregion

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
        Listing folders never downloads cloud files, and each folder is listed once.
        With -ThrottleLimit above 1, that many folders are listed at the same time
        (much faster on network shares); the files come out in the same order.
    #>
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        # Full paths of files to leave out of the inventory (e.g. the report itself).
        [string[]] $ExcludeFile = @(),

        # When given, receives one record per folder listed: Path and Readable.
        # Needed to compare folder trees, including empty and unreadable folders. A
        # folder that holds an excluded file, and each folder link that is not followed,
        # is recorded as not readable: its contents are not fully known, so it can never
        # be proven identical to another folder.
        [System.Collections.Generic.List[object]] $FolderInfo,

        # How many folders to list at the same time.
        [ValidateRange(1, 64)]
        [int] $ThrottleLimit = 1,

        # Wildcard patterns (* and ?) of file and folder names to leave out, ignoring case.
        # Left-out folders are not scanned; a folder is recorded (-FolderInfo) as if the
        # left-out files and folders were not there.
        [string[]] $ExcludeName = @()
    )

    $root = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if (-not $root.PSIsContainer) {
        throw "'$Path' is not a folder."
    }

    $excluded = [System.Collections.Generic.HashSet[string]]::new(
        [string[]] $ExcludeFile, [System.StringComparer]::OrdinalIgnoreCase)
    # Folders holding an excluded file: only their files are checked one by one.
    $excludedIn = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($path in $ExcludeFile) { $null = $excludedIn.Add([System.IO.Path]::GetDirectoryName($path)) }

    # Listing several folders at a time lists the whole tree first; it is then walked below
    # exactly as when listing one folder at a time, so the output is the same.
    $nameFilter = ConvertTo-NameFilter -Pattern $ExcludeName

    $listings = $null
    if ($ThrottleLimit -gt 1) { $listings = Get-TreeListing -Path $root.FullName -ThrottleLimit $ThrottleLimit -NameFilter $nameFilter }

    $pending = [System.Collections.Generic.Stack[string]]::new()
    $pending.Push($root.FullName)

    $folderCount = 0
    $fileCount   = 0
    $timer       = [System.Diagnostics.Stopwatch]::StartNew()
    $lastShownMs = - $script:ProgressIntervalMs

    while ($pending.Count -gt 0) {
        $folder = $pending.Pop()
        $folderCount++

        Write-Verbose "Scanning $folder"
        if ($null -ne $listings) { $listing = $listings[$folder] }
        else {
            if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
                $lastShownMs = $timer.ElapsedMilliseconds
                Write-Progress -Id 1 -Activity 'Scanning folders' `
                    -Status "Folders: $folderCount   Files: $fileCount" -CurrentOperation $folder
            }
            $listing = Get-FolderListing -Path $folder
        }

        if ($null -ne $listing.Error) {
            Write-Warning "Skipping '$folder': $($listing.Error)"
            if ($null -ne $FolderInfo) { $FolderInfo.Add([pscustomobject] @{ Path = $folder; Readable = $false }) }
            continue
        }
        if ($null -ne $FolderInfo) { $FolderInfo.Add([pscustomobject] @{ Path = $folder; Readable = $true }) }

        $keptFiles = $listing.Files
        $subFolders = $listing.Folders
        if ($null -ne $nameFilter) {
            $keptFiles = Select-UnmatchedName -Item $keptFiles -Filter $nameFilter
            $subFolders = Select-UnmatchedName -Item $subFolders -Filter $nameFilter
        }

        # (By a file's own folder, not $folder: a scan root given with a trailing separator
        # keeps it in $folder.)
        if ($keptFiles.Count -and $excludedIn.Contains([System.IO.Path]::GetDirectoryName($keptFiles[0].FullName))) {
            foreach ($file in $keptFiles) {
                if ($excluded.Contains($file.FullName)) {
                    if ($null -ne $FolderInfo) { $FolderInfo.Add([pscustomobject] @{ Path = $folder; Readable = $false }) }
                    continue
                }
                $fileCount++
                $file
            }
        }
        else {
            # The whole listing at once: a statement per file costs more than the listing itself.
            $fileCount += $keptFiles.Count
            $keptFiles
        }

        # Push in reverse so folders are visited in alphabetical order.
        for ($i = $subFolders.Count - 1; $i -ge 0; $i--) {
            $sub = $subFolders[$i]
            if (Test-FolderLink -Folder $sub) {
                Write-Verbose "Not following link '$($sub.FullName)'"
                if ($null -ne $FolderInfo) { $FolderInfo.Add([pscustomobject] @{ Path = $sub.FullName; Readable = $false }) }
                continue
            }
            $pending.Push($sub.FullName)
        }
    }

    Write-Progress -Id 1 -Activity 'Scanning folders' -Completed
}

# .NET sorts a listing and filters it into files and folders far faster than PowerShell:
# names are read through a compiled delegate rather than PowerShell's member access.
$script:NameOf = [System.Delegate]::CreateDelegate([Func[System.IO.FileSystemInfo, string]],
    [System.IO.FileSystemInfo].GetProperty('Name').GetGetMethod())
$script:SelectName = @([System.Linq.Enumerable].GetMethods() | Where-Object {
        $_.Name -eq 'Select' -and $_.GetParameters()[1].ParameterType.GetGenericArguments().Count -eq 2  # Func<T, TResult>
    })[0].MakeGenericMethod([System.IO.FileSystemInfo], [string])
$script:StringArray  = [System.Linq.Enumerable].GetMethod('ToArray').MakeGenericMethod([string])
$script:OfFileType   = [System.Linq.Enumerable].GetMethod('OfType').MakeGenericMethod([System.IO.FileInfo])
$script:OfFolderType = [System.Linq.Enumerable].GetMethod('OfType').MakeGenericMethod([System.IO.DirectoryInfo])
$script:FileArray    = [System.Linq.Enumerable].GetMethod('ToArray').MakeGenericMethod([System.IO.FileInfo])
$script:FolderArray  = [System.Linq.Enumerable].GetMethod('ToArray').MakeGenericMethod([System.IO.DirectoryInfo])

function Get-FolderListing {
    <#
        One folder's files and sub folders, each in ordinal name order, from a single listing
        of the folder (over a network every listing is a round trip). Error holds the reason
        when the folder cannot be read.
    #>
    # A simple function (no parameter validation): it runs once per folder.
    param([string] $Path)
    try {
        $entries = Get-SortedByName -Item ([System.IO.DirectoryInfo] $Path).GetFileSystemInfos()
        [pscustomobject] @{
            Path    = $Path
            Files   = $script:FileArray.Invoke($null, @(, $script:OfFileType.Invoke($null, @(, $entries))))
            Folders = $script:FolderArray.Invoke($null, @(, $script:OfFolderType.Invoke($null, @(, $entries))))
            Error   = $null
        }
    }
    catch [System.UnauthorizedAccessException], [System.IO.IOException], [System.Security.SecurityException] {
        [pscustomobject] @{ Path = $Path; Files = @(); Folders = @(); Error = $_.Exception.Message }
    }
}

function Get-TreeListing {
    <#
        Lists every folder below $Path (folder links are not followed), $ThrottleLimit folders
        at a time, showing progress, leaving out folders whose names match $NameFilter (see
        ConvertTo-NameFilter). Returns a map of folder path -> Get-FolderListing result.
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [ValidateRange(2, 64)] [int] $ThrottleLimit,
        [regex] $NameFilter
    )

    $listings = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::Ordinal)
    $fileCount = 0
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $lastShownMs = - $script:ProgressIntervalMs
    $pool = $null
    try {
        $pool = Open-WorkerPool -Work { param($Folder) Get-FolderListing -Path $Folder } -ThrottleLimit $ThrottleLimit
        $pool.Queue.Add($Path)
        $outstanding = 1
        while ($outstanding -gt 0) {
            $result = Receive-WorkerResult -Pool $pool
            $outstanding--
            if ($null -ne $result.Error) { throw $result.Error }  # Get-FolderListing reports the expected failures itself
            $listing = $result.Value
            $listings[$listing.Path] = $listing
            $fileCount += $listing.Files.Count
            if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
                $lastShownMs = $timer.ElapsedMilliseconds
                Write-Progress -Id 1 -Activity 'Scanning folders' `
                    -Status "Folders: $($listings.Count)   Files: $fileCount ($ThrottleLimit at a time)" -CurrentOperation $listing.Path
            }
            $subFolders = $listing.Folders
            if ($null -ne $NameFilter) { $subFolders = Select-UnmatchedName -Item $subFolders -Filter $NameFilter }
            foreach ($sub in $subFolders) {
                if (Test-FolderLink -Folder $sub) { continue }
                $pool.Queue.Add($sub.FullName)
                $outstanding++
            }
        }
    }
    finally {
        if ($null -ne $pool) { Close-WorkerPool -Pool $pool }
    }
    , $listings
}

function Get-SortedByName {
    # Files and folders (FileSystemInfo) sorted in place into ordinal name order; returns them.
    param([System.IO.FileSystemInfo[]] $Item)
    if ($Item.Count -gt 1) {
        $names = [string[]] $script:StringArray.Invoke($null, @(, $script:SelectName.Invoke($null, @($Item, $script:NameOf))))
        [System.Array]::Sort([System.Array] $names, [System.Array] $Item, [System.Collections.IComparer] [System.StringComparer]::Ordinal)
    }
    , $Item
}

# File systems reached over a network, as Linux names them in /proc/self/mounts. Shared with
# the Python port (scanner.NETWORK_FILE_SYSTEMS).
$script:NetworkFileSystems = [System.Collections.Generic.HashSet[string]]::new([string[]] @(
        '9p', 'afs', 'ceph', 'cifs', 'davfs', 'fuse.davfs2', 'fuse.gcsfuse', 'fuse.glusterfs', 'fuse.rclone',
        'fuse.s3fs', 'fuse.sshfs', 'glusterfs', 'gpfs', 'lustre', 'ncpfs', 'nfs', 'nfs4', 'smb3', 'smbfs'),
    [System.StringComparer]::Ordinal)

# How many folders and files to list and hash at a time on a network drive when the
# command line does not say (-ThrottleLimit): each request waits on the server.
$script:NetworkThrottleLimit = 4

function Test-NetworkDrive {
    <#
    .SYNOPSIS
        Whether a path is on a network share: a UNC path or network drive on Windows, a
        network file system (NFS, SMB ...) on Linux. False when unknown (macOS).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)] [string] $Path)

    $Path = [System.IO.Path]::GetFullPath($Path)
    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
        # Long-path forms: \\?\UNC\server\share is a share, \\?\C:\... a drive.
        foreach ($prefix in '\\?\', '\\.\') {
            if ($Path.StartsWith($prefix + 'UNC\', [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
            if ($Path.StartsWith($prefix)) { $Path = $Path.Substring($prefix.Length) }
        }
        if ($Path.StartsWith('\\') -or $Path.StartsWith('//')) { return $true }
        try { return ([System.IO.DriveInfo]::new([System.IO.Path]::GetPathRoot($Path))).DriveType -eq [System.IO.DriveType]::Network }
        catch { return $false }
    }
    $mountsFile = '/proc/self/mounts'
    if (-not [System.IO.File]::Exists($mountsFile)) { return $false }
    # (mount point, file system) pairs; mount points write spaces and some other characters
    # as octal (\040).
    $mounts = @(foreach ($line in [System.IO.File]::ReadAllLines($mountsFile)) {
            $fields = $line.Split(' ')
            if ($fields.Count -ge 3) {
                $point = [regex]::Replace($fields[1], '\\([0-7]{3})', { param($m) [string] [char] [Convert]::ToInt32($m.Groups[1].Value, 8) })
                , @($point, $fields[2])
            }
        })
    $script:NetworkFileSystems.Contains((Get-MountType -Path $Path -Mount $mounts))
}

function Get-MountType {
    # The file system of the mount holding $Path, given (mount point, file system) pairs:
    # the longest mount point that is the path or one of its parent folders.
    param([Parameter(Mandatory)] [string] $Path, [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Mount)
    $best = ''; $kind = ''
    foreach ($pair in $Mount) {
        $point = $pair[0]
        $inside = $Path -ceq $point -or $Path.StartsWith($point.TrimEnd('/') + '/', [System.StringComparison]::Ordinal)
        if ($inside -and $point.Length -gt $best.Length) { $best = $point; $kind = $pair[1] }
    }
    $kind
}

function Get-DefaultThrottleLimit {
    <#
    .SYNOPSIS
        The -ThrottleLimit to use for a scan of $Path when none was given: several at a time
        on a network drive, one otherwise.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)] [string] $Path)
    if (Test-NetworkDrive -Path $Path) { return $script:NetworkThrottleLimit }
    1
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
# same code runs in the current session and in the parallel runspaces. Reads go straight
# from the file in chunks of up to 1 MB (much faster from disks and shares) into a buffer
# no larger than the file, so small files, the most common, cost no large allocation.
$script:ComputeMd5 = {
    # With $Limit above 0, only the first $Limit bytes are hashed.
    param([string] $Path, [long] $Limit = 0)
    $ErrorActionPreference = 'Stop'
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        $stream = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
            [System.IO.FileShare] 'ReadWrite, Delete', 1, [System.IO.FileOptions]::SequentialScan)
        try {
            $remaining = [long]::MaxValue
            if ($Limit -gt 0) { $remaining = $Limit }
            # [long]: with an [int] first argument PowerShell picks Math.Min(int, int), which
            # overflows for files over 2 GB.
            $buffer = [byte[]]::new([int] [Math]::Max([long] 1, [Math]::Min([Math]::Min([long] 1MB, $stream.Length), $remaining)))
            while ($remaining -gt 0 -and ($read = $stream.Read($buffer, 0, [int] [Math]::Min([long] $buffer.Length, $remaining))) -gt 0) {
                $null = $md5.TransformBlock($buffer, 0, $read, $null, 0)
                $remaining -= $read
            }
            $null = $md5.TransformFinalBlock($buffer, 0, 0)
            [System.BitConverter]::ToString($md5.Hash).Replace('-', '')
        }
        finally { $stream.Dispose() }
    }
    finally { $md5.Dispose() }
}

function ConvertTo-NameKey {
    <#
        The form in which names are compared: Unicode-normalised (NFC, so "e + accent"
        as macOS often stores it equals the single character Windows stores) and upper
        case (so the comparison ignores case). Names that are not valid UTF-16 cannot
        be normalised and are compared as they are.
        Inlined, and to be kept in step, where it would run per file: Find-DuplicateFile and
        Select-UnmatchedName.
    #>
    # A simple function (no parameter validation): it runs for every file in some loops.
    param([string] $Name)
    try { if (-not $Name.IsNormalized()) { $Name = $Name.Normalize() } }
    catch [System.ArgumentException] { Write-Debug "Cannot normalise '$Name'; comparing it as it is." }
    $Name.ToUpperInvariant()
}

# One character for a ? wildcard: a surrogate pair (a character beyond U+FFFF, such as an
# emoji) or a single UTF-16 unit, as Python's regular expressions match one code point.
# Atomic, so that it never falls back to half of a pair.
$script:AnyCharacterPattern = '(?>[\uD800-\uDBFF][\uDC00-\uDFFF]|.)'

function Get-NamePatternProblem {
    # Why an exclusion pattern can never match a name, or $null when it can.
    param([AllowEmptyString()] [string] $Pattern)
    if (-not $Pattern) { return 'Exclusion patterns cannot be empty.' }
    if ($Pattern.IndexOfAny([char[]] '/\') -ge 0) {
        return "Exclusion pattern '$Pattern' contains / or \: patterns match file and folder names, not paths."
    }
    $null
}

function ConvertFrom-SizeText {
    <#
    .SYNOPSIS
        A size in bytes from text such as 1500, 2KB or 1.5MB: a number, optionally followed by
        KB, MB, GB, TB or PB (1024-based, any case), rounded to a whole number of bytes. Read
        the same way as the Python port's --minimum-size.
    #>
    [CmdletBinding()]
    [OutputType([long])]
    param([Parameter(Mandatory)] [AllowEmptyString()] [string] $Text, [string] $Name = 'The size')
    $match = [regex]::Match($Text, '\A([0-9]+(?:\.[0-9]+)?)([KMGTP]B)?\z', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { throw "$Name must be a number of bytes, optionally followed by KB, MB, GB, TB or PB." }
    $power = 0
    if ($match.Groups[2].Success) { $power = 'KMGTP'.IndexOf([char]::ToUpperInvariant($match.Groups[2].Value[0])) + 1 }
    $bytes = [Math]::Round([double]::Parse($match.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) * [Math]::Pow(1024, $power))
    if ($bytes -ge [Math]::Pow(2, 63)) { throw "$Name must be at most $([long]::MaxValue) bytes." }
    [long] $bytes
}

function Assert-NamePattern {
    <#
    .SYNOPSIS
        Throws when an exclusion pattern could never match a file or folder name (it is
        empty or holds a path separator), so a scan stops before it starts.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [AllowEmptyString()] [string[]] $Pattern)
    foreach ($p in $Pattern) {
        $problem = Get-NamePatternProblem -Pattern $p
        if ($problem) { throw $problem }
    }
}

function ConvertTo-NameFilter {
    <#
        One regular expression that matches the name key (see ConvertTo-NameKey) of any name
        matching one of the wildcard patterns, so names match whatever their case or Unicode
        form. * stands for any run of characters and ? for any one character; everything
        else is literal. $null when there are no patterns.
        Each pattern is split at its stars: the first part must start the name and the last
        end it; each part between is taken at its first place after the one before, in an
        atomic group. That is always right for wildcards and never backtracks, so a pattern
        such as *a*a*a*a*b cannot take exponential time (as name_filter in Python).
    #>
    param([AllowEmptyCollection()] [string[]] $Pattern)
    if (-not $Pattern) { return $null }
    $alternatives = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $Pattern) {
        $problem = Get-NamePatternProblem -Pattern $p
        if ($problem) { throw $problem }
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($part in (ConvertTo-NameKey $p).Split('*')) {
            $parts.Add((-join @(foreach ($piece in [regex]::Split($part, '(\?)')) {
                            if ($piece -eq '?') { $script:AnyCharacterPattern } else { [regex]::Escape($piece) }
                        })))
        }
        $expression = '\A' + $parts[0]
        if ($parts.Count -gt 1) {
            for ($i = 1; $i -lt $parts.Count - 1; $i++) { if ($parts[$i]) { $expression += "(?>.*?$($parts[$i]))" } }
            $expression += '.*' + $parts[$parts.Count - 1]
        }
        $alternatives.Add($expression + '\z')
    }
    [regex]::new('(?:' + ($alternatives -join '|') + ')', [System.Text.RegularExpressions.RegexOptions]::Singleline)
}

function Select-UnmatchedName {
    <#
        The files or folders whose names do not match $Filter (from ConvertTo-NameFilter), in
        order. Called once per folder listing, with the name key inlined (see
        ConvertTo-NameKey): a call per item would slow scanning several times over.
    #>
    param([AllowEmptyCollection()] [object[]] $Item, [regex] $Filter)
    , [object[]] @(foreach ($i in $Item) {
            $name = $i.Name
            try { if (-not $name.IsNormalized()) { $name = $name.Normalize() } }
            catch [System.ArgumentException] { Write-Debug "Cannot normalise '$name'; comparing it as it is." }
            if (-not $Filter.IsMatch($name.ToUpperInvariant())) { $i }
        })
}

function Get-InnermostMessage {
    # The message of the innermost exception, so failures read the same however they surfaced.
    param([Parameter(Mandatory)] [System.Exception] $Exception)
    while ($Exception.InnerException) { $Exception = $Exception.InnerException }
    $Exception.Message
}

function ConvertTo-UtcOffsetText {
    # A UTC offset as text: +10:00, -04:30, or with seconds (+00:19:32) for historic local times.
    param([Parameter(Mandatory)] [TimeSpan] $Offset)
    $sign = if ($Offset -lt [TimeSpan]::Zero) { '-' } else { '+' }
    $size = $Offset.Duration()
    $text = $sign + ([int] [Math]::Floor($size.TotalHours)).ToString('00') + ':' + $size.Minutes.ToString('00')
    if ($size.Seconds) { $text += ':' + $size.Seconds.ToString('00') }
    $text
}

function ConvertFrom-UtcOffsetText {
    # The reverse of ConvertTo-UtcOffsetText, or a clear error naming the report.
    param([AllowNull()] [AllowEmptyString()] [string] $Text, [Parameter(Mandatory)] [string] $Path)
    $parts = [regex]::Match([string] $Text, '^([+-])(\d{2}):(\d{2})(?::(\d{2}))?$')
    if (-not $parts.Success) { throw "'$Path' could not be read as a duplicates report: '$Text' is not a UTC offset." }
    $seconds = [int] $parts.Groups[2].Value * 3600 + [int] $parts.Groups[3].Value * 60
    if ($parts.Groups[4].Success) { $seconds += [int] $parts.Groups[4].Value }
    if ($parts.Groups[1].Value -eq '-') { $seconds = - $seconds }
    [TimeSpan]::FromSeconds($seconds)
}

function Get-DuplicateSetUtcOffset {
    # A duplicate set's UTC offset; for sets without one (read from a report made before the
    # UTC Offset column, or built by hand), this computer's offset at its saved date.
    param([Parameter(Mandatory)] [object] $Set)
    $offset = $Set.PSObject.Properties['UtcOffset']
    if ($offset -and $null -ne $offset.Value) { return [TimeSpan] $offset.Value }
    [System.TimeZoneInfo]::Local.GetUtcOffset([datetime]::SpecifyKind($Set.LastWriteTime, [System.DateTimeKind]::Unspecified))
}

function Get-FileMd5 {
    # MD5 of one file's contents as upper-case hex (the Get-FileHash format); with $Limit,
    # of its first $Limit bytes only.
    param([Parameter(Mandatory)] [string] $Path, [long] $Limit = 0)
    & $script:ComputeMd5 $Path $Limit
}

# Large candidates are first compared by the MD5 of their start (see Split-ByStartHash):
# files that share a name, saved date and size but not their contents are then rarely read
# in full. Only for files this large, so that true duplicates cost at most 1/16 more reading.
$script:FirstBytesToHash = 1MB
$script:FirstBytesMinSize = 16MB

function Get-FileMd5Map {
    <#
        Hashes files, up to $ThrottleLimit at a time, returning a map of full path -> MD5.
        Files that cannot be read are reported as warnings and left out of the map.
        With $Cache, files already in it are not read again and new hashes are added to it.
        With -FirstBytes, only the first $script:FirstBytesToHash bytes of each file are
        hashed (and $Cache, which holds whole-file hashes, must not be given).
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]] $Path,
        [ValidateRange(1, 64)] [int] $ThrottleLimit = 1,
        [System.Collections.Generic.Dictionary[string, string]] $Cache,
        [switch] $FirstBytes
    )

    $requested = $Path
    if ($null -ne $Cache) { $Path = [string[]] @($Path | Where-Object { -not $Cache.ContainsKey($_) }) }
    $limit = [long] 0
    if ($FirstBytes) { $limit = $script:FirstBytesToHash }

    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $activity = 'Comparing MD5 hashes'
    if ($FirstBytes) { $activity = 'Comparing the start of large files' }
    $status = if ($ThrottleLimit -eq 1) { '' } else { " ($ThrottleLimit at a time)" }
    $done = 0
    # Progress is shown a few times a second, not per file: drawing it costs more than
    # hashing a small file.
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $lastShownMs = - $script:ProgressIntervalMs

    if ($ThrottleLimit -eq 1 -or $Path.Count -lt 2) {
        foreach ($p in $Path) {
            $done++
            if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
                $lastShownMs = $timer.ElapsedMilliseconds
                Write-Progress -Id 2 -Activity $activity -Status "File $done of $($Path.Count)$status" -CurrentOperation $p `
                    -PercentComplete ([int] (100 * $done / $Path.Count))
            }
            try { $map[$p] = Get-FileMd5 -Path $p -Limit $limit }
            catch { Write-Warning "Could not hash '$p': $(Get-InnermostMessage $_.Exception)" }
        }
    }
    else {
        $pool = $null
        try {
            $work = { param($File) & $script:ComputeMd5 $File }
            if ($FirstBytes) { $work = { param($File) & $script:ComputeMd5 $File $script:FirstBytesToHash } }
            $pool = Open-WorkerPool -Work $work -ThrottleLimit ([Math]::Min($ThrottleLimit, $Path.Count))
            foreach ($p in $Path) { $pool.Queue.Add($p) }
            $result = $null
            while ($done -lt $Path.Count) {
                # A result that is already waiting is taken directly: a function call per file
                # would cost more than hashing a small one.
                if (-not $pool.Results.TryTake([ref] $result)) { $result = Receive-WorkerResult -Pool $pool }
                $done++
                if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
                    $lastShownMs = $timer.ElapsedMilliseconds
                    Write-Progress -Id 2 -Activity $activity -Status "File $done of $($Path.Count)$status" -CurrentOperation $result.Item `
                        -PercentComplete ([int] (100 * $done / $Path.Count))
                }
                if ($null -eq $result.Error) { $map[$result.Item] = [string] $result.Value }
                else { Write-Warning "Could not hash '$($result.Item)': $(Get-InnermostMessage $result.Error)" }
            }
        }
        finally {
            if ($null -ne $pool) { Close-WorkerPool -Pool $pool }
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
        One object per duplicate set: FileName, LastWriteTime (local), UtcOffset (local
        time minus UTC at that date), SizeBytes, MD5, Count and Folders (full folder path
        of every copy, sorted).
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
        [System.Collections.Generic.Dictionary[string, string]] $Md5Cache,

        # Leave files of 0 bytes out (they all share one MD5).
        [switch] $IgnoreEmptyFiles,

        # Leave files smaller than this many bytes out.
        [ValidateRange(0, [long]::MaxValue)]
        [long] $MinimumSize = 0
    )

    # Stage 1: name. Grouped first, so the size and saved date of a file whose name no other
    # file has are never read: on Linux, macOS and network drives each is a request per file.
    # The dictionary in Group-ByKey ignores case; only Unicode normalisation is needed here.
    $names = [System.Collections.Generic.List[string]]::new($File.Count)
    foreach ($f in $File) {
        # Inline rather than ConvertTo-NameKey: this loop runs once per file.
        $name = $f.Name
        try { if (-not $name.IsNormalized()) { $name = $name.Normalize() } }
        catch [System.ArgumentException] { Write-Debug "Cannot normalise '$name'; comparing it as it is." }
        $names.Add($name)
    }
    $nameGroups = @(Group-ByKey -InputItems $File -Key $names.ToArray())

    # Stage 2: saved date (UTC, whole second: copies made to network shares or other file
    # systems often lose sub-second precision) and size, a cheap check that avoids hashing
    # files that cannot match. Both keys are digits, so '|' is a safe separator.
    $minimumSize = $MinimumSize
    if ($IgnoreEmptyFiles -and $minimumSize -lt 1) { $minimumSize = 1 }
    $ticksPerSecond = [System.TimeSpan]::TicksPerSecond
    $candidateGroups = @(foreach ($group in $nameGroups) {
            $kept = [System.Collections.Generic.List[object]]::new()
            $keys = [System.Collections.Generic.List[string]]::new()
            foreach ($f in $group) {
                # Getter methods, not properties: PowerShell turns a failing property into $null.
                try { $size = $f.get_Length(); $ticks = $f.get_LastWriteTimeUtc().Ticks }
                catch {
                    Write-Warning "Skipping '$($f.FullName)': $(Get-InnermostMessage $_.Exception)"  # gone or unreadable since the scan
                    continue
                }
                if ($size -lt $minimumSize) { continue }
                $kept.Add($f)
                $keys.Add([string] ($ticks - ($ticks % $ticksPerSecond)) + '|' + $size)
            }
            Group-ByKey -InputItems $kept.ToArray() -Key $keys.ToArray()
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

    # Stage 3: for large files, the MD5 of their start.
    $candidateGroups = @(Split-ByStartHash -Group $candidateGroups -ThrottleLimit $ThrottleLimit -Md5Cache $Md5Cache)

    # Stage 4: MD5, only for files that already match on name, date and size.
    $candidates = [System.Collections.Generic.List[string]]::new()
    foreach ($group in $candidateGroups) { foreach ($f in $group) { $candidates.Add($f.FullName) } }
    $md5ByPath = Get-FileMd5Map -Path $candidates.ToArray() -ThrottleLimit $ThrottleLimit -Cache $Md5Cache

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $candidateGroups) {
        # Loops rather than pipelines: large trees have thousands of groups.
        $hashed = @(foreach ($f in $group) { if ($md5ByPath.ContainsKey($f.FullName)) { $f } })
        $md5s = [string[]] @(foreach ($f in $hashed) { $md5ByPath[$f.FullName] })

        foreach ($set in @(Group-ByKey -InputItems $hashed -Key $md5s)) {
            $first = $set[0]
            $results.Add([pscustomobject] @{
                FileName      = $first.Name
                LastWriteTime = $first.LastWriteTime
                UtcOffset     = $first.LastWriteTime - $first.LastWriteTimeUtc
                SizeBytes     = $first.Length
                MD5           = $md5ByPath[$first.FullName]
                Count         = $set.Count
                Folders       = Get-SortedFolder -Path @(foreach ($f in $set) { $f.DirectoryName })
            })
        }
    }

    $sorted = Get-SortedDuplicateSet -DuplicateSet $results.ToArray()
    $sorted
}

function Split-ByStartHash {
    <#
        Splits each group of files of $script:FirstBytesMinSize or more by the MD5 of their
        first $script:FirstBytesToHash bytes, keeping the groups that still hold more than one
        file. Other groups, and groups a file of which already has its whole-file MD5 in
        $Md5Cache (from a previous report), are kept as they are.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Group,
        [ValidateRange(1, 64)] [int] $ThrottleLimit = 1,
        [System.Collections.Generic.Dictionary[string, string]] $Md5Cache
    )
    $large = [System.Collections.Generic.List[object]]::new()
    foreach ($files in $Group) {
        # The files share their size (stage 2); a file gone since then is left to stage 4.
        $isLarge = $false
        try { $isLarge = $files[0].get_Length() -ge $script:FirstBytesMinSize } catch { $isLarge = $false }
        if ($isLarge -and $null -ne $Md5Cache) { foreach ($f in $files) { if ($Md5Cache.ContainsKey($f.FullName)) { $isLarge = $false } } }
        if ($isLarge) { $large.Add($files) } else { , $files }
    }
    if ($large.Count -eq 0) { return }

    $paths = [string[]] @(foreach ($files in $large) { foreach ($f in $files) { $f.FullName } })
    $starts = Get-FileMd5Map -Path $paths -ThrottleLimit $ThrottleLimit -FirstBytes
    foreach ($files in $large) {
        $read = @(foreach ($f in $files) { if ($starts.ContainsKey($f.FullName)) { $f } })
        $keys = [string[]] @(foreach ($f in $read) { $starts[$f.FullName] })
        Group-ByKey -InputItems $read -Key $keys
    }
}

function Get-SortedFolder {
    # Folder paths in case-insensitive order (see Compare-IgnoringCase), paths differing only
    # in case in a fixed order: $script:ByPathIgnoringCase's order. Sorted on precomputed keys
    # compared ordinally (upper case, a separator that sorts first, then the path as it is):
    # a script block comparer would be slow for files with thousands of copies.
    param([Parameter(Mandatory)] [string[]] $Path)
    $sorted = [string[]] $Path.Clone()
    if ($sorted.Count -gt 1) {
        $separator = [string] [char] 0
        $keys = [string[]] @(foreach ($p in $sorted) { $p.ToUpperInvariant() + $separator + $p })
        [System.Array]::Sort([System.Array] $keys, [System.Array] $sorted, [System.Collections.IComparer] [System.StringComparer]::Ordinal)
    }
    , $sorted
}

#endregion

#region Duplicate folders

# Same ordering as the Python port: folder name (ordinal, ignoring case), size, first location.
$script:ByFolderSet = [System.Comparison[object]] {
    param($x, $y)
    $order = Compare-IgnoringCase $x.FolderName $y.FolderName
    if ($order -eq 0) { $order = $x.SizeBytes.CompareTo($y.SizeBytes) }
    if ($order -eq 0) { $order = $script:ByPathIgnoringCase.Invoke($x.Folders[0], $y.Folders[0]) }
    $order
}

function Get-FolderTree {
    <#
        Indexes a scan: the sub folders and files of every folder, and which folders were
        readable. File details are not read here (see Confirm-TreeFile): only the folders
        that could be duplicates need them.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.IO.FileInfo[]] $File,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Folder
    )

    $children = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]::new([System.StringComparer]::Ordinal)
    $filesIn  = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[System.IO.FileInfo]]]::new([System.StringComparer]::Ordinal)
    $readable = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

    # A folder may be recorded more than once; recorded as not readable anywhere (an
    # excluded file, a skipped link) means not readable.
    $unreadable = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($record in $Folder) {
        $children[$record.Path] = [System.Collections.Generic.List[string]]::new()
        $filesIn[$record.Path]  = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        if (-not $record.Readable) { $null = $unreadable.Add($record.Path) }
    }
    foreach ($f in $File) {
        if ($filesIn.ContainsKey($f.DirectoryName)) { $filesIn[$f.DirectoryName].Add($f) }
    }
    foreach ($path in @($children.Keys)) {
        if (-not $unreadable.Contains($path)) { $null = $readable.Add($path) }
        $parent = [System.IO.Path]::GetDirectoryName($path)
        if ($parent -and $children.ContainsKey($parent)) { $children[$parent].Add($path) }
    }

    # Deepest first: a sub folder's path is always longer than its parent's.
    $order = [string[]] @($children.Keys)
    $lengths = [int[]] @($order | ForEach-Object { - $_.Length })
    [System.Array]::Sort([System.Array] $lengths, [System.Array] $order, [System.Collections.IComparer] [System.Collections.Comparer]::DefaultInvariant)

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
                $line = "F|$(ConvertTo-NameKey $f.Name)|$($ticks - ($ticks % $ticksPerSecond))|$($f.Length)"
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
                $lines.Add("D|$(ConvertTo-NameKey ([System.IO.Path]::GetFileName($sub)))|$($child.Signature)")
                $info.FileCount   += $child.FileCount
                $info.FolderCount += $child.FolderCount + 1
                $info.SizeBytes   += $child.SizeBytes
            }
            if (-not $complete) { continue }

            $sorted = $lines.ToArray()
            [System.Array]::Sort([System.Array] $sorted, [System.Collections.IComparer] [System.StringComparer]::Ordinal)
            $bytes = [System.Text.Encoding]::UTF8.GetBytes(($sorted -join "`n"))
            $info.Signature = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
        }
    }
    finally { $sha.Dispose() }
    $result
}

function Confirm-TreeFile {
    <#
        Reads the size and saved date of every file in the $Scope folders of $Tree (in scan
        order). A file gone or unreadable since the scan is dropped with a warning, and
        leaves its folder's contents unknown: the folder is no longer readable.
    #>
    param(
        [Parameter(Mandatory)] [object] $Tree,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.IO.FileInfo[]] $File,
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]] $Scope
    )
    foreach ($f in $File) {
        $folder = $f.DirectoryName
        if (-not $Scope.Contains($folder) -or -not $Tree.Files.ContainsKey($folder)) { continue }
        # (A getter method, not the property: PowerShell turns a failing property into $null.
        # Reading the size fails for a missing file and loads the saved date with it.)
        try { $null = $f.get_Length() }
        catch {
            Write-Warning "Skipping '$($f.FullName)': $(Get-InnermostMessage $_.Exception)"
            $null = $Tree.Files[$folder].Remove($f)
            $null = $Tree.Readable.Remove($folder)
        }
    }
}

function Get-FolderScope {
    # The given folders and every folder below them.
    param([Parameter(Mandatory)] [object] $Tree, [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]] $Folder)
    $scope = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    $pending = [System.Collections.Generic.Stack[string]]::new([string[]] $Folder)
    while ($pending.Count -gt 0) {
        $path = $pending.Pop()
        if ($scope.Add($path)) { foreach ($sub in $Tree.Children[$path]) { $pending.Push($sub) } }
    }
    , $scope
}

function Get-RepeatedNameScope {
    # The folders whose name (ignoring case) another folder has, and every folder below them.
    param([Parameter(Mandatory)] [object] $Tree)
    $paths = [string[]] $Tree.DeepestFirst
    $keys = [string[]] @(foreach ($path in $paths) { ConvertTo-NameKey ([System.IO.Path]::GetFileName($path)) })
    $repeated = @(foreach ($group in (Group-ByKey -InputItems $paths -Key $keys)) { $group })
    Get-FolderScope -Tree $Tree -Folder $repeated
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
        every one of its folders sits inside a folder that is itself a duplicate.
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

    # Pass 1: names, sizes and saved dates only; nothing is read. Only folders whose name
    # another folder has can be duplicates, so only they (and the folders below them, part
    # of their fingerprint) are fingerprinted.
    $named = Get-RepeatedNameScope -Tree $tree
    if ($named.Count -eq 0) { return }
    Confirm-TreeFile -Tree $tree -File $File -Scope $named
    $cheap = Get-FolderSignature -Tree $tree -Scope $named
    $candidates = [System.Collections.Generic.List[string]]::new()
    $keys = [System.Collections.Generic.List[string]]::new()
    foreach ($path in $tree.DeepestFirst) {
        if (-not $named.Contains($path)) { continue }
        $info = $cheap[$path]
        if ($info.Signature -and $info.FileCount -gt 0) {
            $candidates.Add($path)
            $keys.Add("$(ConvertTo-NameKey ([System.IO.Path]::GetFileName($path)))|$($info.Signature)")
        }
    }
    $candidateGroups = @(Group-ByKey -InputItems $candidates.ToArray() -Key $keys.ToArray())
    if ($candidateGroups.Count -eq 0) { return }

    # Pass 2: hash every file below the candidates (hashes from the file scan are reused).
    $scope = Get-FolderScope -Tree $tree -Folder @(foreach ($group in $candidateGroups) { $group })

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
                $keys.Add("$(ConvertTo-NameKey ([System.IO.Path]::GetFileName($path)))|$($full[$path].Signature)")
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

# File names may contain control characters (Linux) or unpaired surrogates (NTFS) that
# XML 1.0 cannot represent; Write-RowXml replaces them with U+FFFD. Built once: it runs per cell.
$script:InvalidXmlChars = [regex]::new(
    '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]')

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

# Cell styles, as defined in xl/styles.xml.
$script:StyleBold = 1
$script:StyleDate = 2

function Write-RowXml {
    # Writes one <row> element and its cells: strings inline, numbers and dates numeric.
    # Called once per row and writing the cells itself, with no parameter validation:
    # it runs for every row of reports that can hold millions of cells.
    param([System.Xml.XmlWriter] $Writer, [int] $Number, [object[]] $Values, [string[]] $Letters, [switch] $Bold)

    $ns = $script:SpreadsheetMain
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $Writer.WriteStartElement('row', $ns)
    $Writer.WriteAttributeString('r', [string] $Number)
    for ($c = 0; $c -lt $Values.Count; $c++) {
        $value = $Values[$c]
        $Writer.WriteStartElement('c', $ns)
        $Writer.WriteAttributeString('r', $Letters[$c] + $Number)
        if ($value -is [datetime]) {
            $Writer.WriteAttributeString('s', [string] $script:StyleDate)
            $Writer.WriteElementString('v', $ns, $value.ToOADate().ToString('R', $invariant))
        }
        else {
            if ($Bold) { $Writer.WriteAttributeString('s', [string] $script:StyleBold) }
            if ($value -is [int] -or $value -is [long] -or $value -is [double]) {
                $Writer.WriteElementString('v', $ns, $value.ToString($invariant))
            }
            else {
                $text = [string] $value
                if ($text.Length -gt $script:ExcelMaxCellText) {
                    throw "Cell $($Letters[$c])$Number would hold $($text.Length) characters; Excel allows at most $($script:ExcelMaxCellText)."
                }
                $Writer.WriteAttributeString('t', 'inlineStr')
                $Writer.WriteStartElement('is', $ns)
                $Writer.WriteStartElement('t', $ns)
                $Writer.WriteAttributeString('xml', 'space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
                $Writer.WriteString($script:InvalidXmlChars.Replace($text, [string] [char] 0xFFFD))
                $Writer.WriteEndElement()
                $Writer.WriteEndElement()
            }
        }
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

    # The table's header row (the rules are on their own sheet).
    $headerRow = 1

    $maxCopies = 1
    foreach ($values in $Row) { $maxCopies = [Math]::Max($maxCopies, $values.Count - $Column.Count) }

    $columnCount = $Column.Count + $maxCopies
    if ($columnCount -gt $script:ExcelMaxColumns) {
        throw "A $Noun has $maxCopies copies; Excel supports at most $($script:ExcelMaxColumns - $Column.Count) location columns."
    }
    if ($Row.Count + $headerRow -gt $script:ExcelMaxRows) {
        throw "Found $($Row.Count) duplicated $($Noun)s; Excel supports at most $($script:ExcelMaxRows - $headerRow) rows."
    }

    [pscustomobject] @{
        Kind       = 'Table'
        Name       = $Name
        HeaderRow  = $headerRow
        Headers    = [string[]] (@($Column | ForEach-Object { $_.Header }) + @(1..$maxCopies | ForEach-Object { "Location $_" }))
        Widths     = [int[]] (@($Column | ForEach-Object { $_.Width }) + @(1..$maxCopies | ForEach-Object { $script:LocationColumnWidth }))
        Rows       = $Row
        LastColumn = ConvertTo-ColumnName $columnCount
        LastRow    = $Row.Count + $headerRow
    }
}

function Write-WorksheetXml {
    param(
        [Parameter(Mandatory)] [System.Xml.XmlWriter] $Writer,
        [Parameter(Mandatory)] [object] $Sheet
    )

    $ns = $script:SpreadsheetMain

    $Writer.WriteStartElement('worksheet', $ns)

    # Frozen header row.
    $Writer.WriteStartElement('sheetViews', $ns)
    $Writer.WriteStartElement('sheetView', $ns)
    $Writer.WriteAttributeString('workbookViewId', '0')
    $Writer.WriteStartElement('pane', $ns)
    $Writer.WriteAttributeString('ySplit', [string] $Sheet.HeaderRow)
    $Writer.WriteAttributeString('topLeftCell', "A$($Sheet.HeaderRow + 1)")
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

    # Column letters, worked out once rather than per cell (large reports have millions of cells).
    $letters = [string[]] @(for ($c = 1; $c -le $Sheet.Headers.Count; $c++) { ConvertTo-ColumnName $c })
    Write-RowXml -Writer $Writer -Number $Sheet.HeaderRow -Values $Sheet.Headers -Letters $letters -Bold

    # One row per duplicated item; one column per copy.
    $rowNumber = $Sheet.HeaderRow
    foreach ($values in $Sheet.Rows) {
        $rowNumber++
        Write-RowXml -Writer $Writer -Number $rowNumber -Values $values -Letters $letters
    }

    $Writer.WriteEndElement()  # sheetData

    $Writer.WriteStartElement('autoFilter', $ns)
    $Writer.WriteAttributeString('ref', "A$($Sheet.HeaderRow):$($Sheet.LastColumn)$($Sheet.LastRow)")
    $Writer.WriteEndElement()

    $Writer.WriteEndElement()  # worksheet
}

function Get-RulesSheetData {
    <#
        The Rules sheet: a title, a section per data sheet, the scan settings when the scan
        left anything out, and blank rows ($null) between. Each line is a row of cells.
    #>
    param([switch] $IncludeFolders, [string[]] $ExcludeName = @(), [long] $MinimumSize = 0)

    $lines = [System.Collections.Generic.List[object]]::new()
    $lines.Add([pscustomobject] @{ Cells = @($script:RulesIntro[0]); Bold = $true })
    foreach ($text in $script:RulesIntro[1..($script:RulesIntro.Count - 1)]) { $lines.Add([pscustomobject] @{ Cells = @($text); Bold = $false }) }
    $sections = @(@{ Title = $script:FileRulesTitle; Rules = $script:FileRules; Values = @() })
    if ($IncludeFolders) { $sections += @{ Title = $script:FolderRulesTitle; Rules = $script:FolderRules; Values = @() } }
    if ($ExcludeName.Count -or $MinimumSize -gt 0) {
        $values = @()
        if ($ExcludeName.Count) { $values += , (@($script:ExcludeNameLabel) + $ExcludeName) }
        if ($MinimumSize -gt 0) { $values += , @($script:MinimumSizeLabel, $MinimumSize) }
        $sections += @{ Title = $script:SettingsTitle; Rules = $script:SettingsRules; Values = $values }
    }
    foreach ($section in $sections) {
        $lines.Add($null)
        $lines.Add([pscustomobject] @{ Cells = @($section.Title); Bold = $true })
        foreach ($text in $section.Rules) { $lines.Add([pscustomobject] @{ Cells = @($text); Bold = $false }) }
        foreach ($cells in $section.Values) { $lines.Add([pscustomobject] @{ Cells = $cells; Bold = $false }) }
    }
    [pscustomobject] @{ Kind = 'Rules'; Name = $script:RulesSheetName; Lines = $lines.ToArray() }
}

function Write-RulesSheetXml {
    param(
        [Parameter(Mandatory)] [System.Xml.XmlWriter] $Writer,
        [Parameter(Mandatory)] [object] $Sheet
    )

    $ns = $script:SpreadsheetMain
    $Writer.WriteStartElement('worksheet', $ns)
    $Writer.WriteStartElement('cols', $ns)
    $Writer.WriteStartElement('col', $ns)
    $Writer.WriteAttributeString('min', '1')
    $Writer.WriteAttributeString('max', '1')
    $Writer.WriteAttributeString('width', [string] $script:RulesColumnWidth)
    $Writer.WriteAttributeString('customWidth', '1')
    $Writer.WriteEndElement()
    $Writer.WriteEndElement()

    $Writer.WriteStartElement('sheetData', $ns)
    for ($r = 0; $r -lt $Sheet.Lines.Count; $r++) {
        $line = $Sheet.Lines[$r]
        if ($null -eq $line) { continue }  # a blank row
        $letters = [string[]] @(for ($c = 1; $c -le $line.Cells.Count; $c++) { ConvertTo-ColumnName $c })
        Write-RowXml -Writer $Writer -Number ($r + 1) -Values $line.Cells -Letters $letters -Bold:$line.Bold
    }
    $Writer.WriteEndElement()  # sheetData
    $Writer.WriteEndElement()  # worksheet
}

function Export-DuplicateReport {
    <#
    .SYNOPSIS
        Saves duplicate sets to an .xlsx workbook. Needs neither Excel nor extra modules.
    .DESCRIPTION
        Sheet "Duplicates": one row per duplicated file. Columns: File Name, Last Modified,
        UTC Offset, Size (bytes), MD5, Copies, then "Location 1..N" holding the full folder
        path of every copy.
        Sheet "Duplicate Folders" (only when -FolderSet is given): one row per duplicated
        folder. Columns: Folder Name, Files, Sub Folders, Size (bytes), Copies, then
        "Location 1..N" holding the full path of every copy.
        Sheet "Rules": the matching rules behind the other sheets, in plain words, and the
        scan settings (-ExcludeName, -MinimumSize) when the scan left anything out.
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
        [object[]] $FolderSet,

        # The name patterns and smallest file size the scan used, recorded on the Rules sheet.
        [string[]] $ExcludeName = @(),
        [ValidateRange(0, [long]::MaxValue)]
        [long] $MinimumSize = 0
    )

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)

    $fileRows = @(foreach ($set in $DuplicateSet) {
            $offset = ConvertTo-UtcOffsetText (Get-DuplicateSetUtcOffset $set)
            , (@($set.FileName, $set.LastWriteTime, $offset, [long] $set.SizeBytes, $set.MD5, [int] $set.Count) + @($set.Folders))
        })
    $sheets = [System.Collections.Generic.List[object]]::new()
    $sheets.Add((ConvertTo-WorksheetData -Name $script:FileSheetName -Column $script:FixedColumns -Row $fileRows -Noun 'file'))

    if ($PSBoundParameters.ContainsKey('FolderSet') -and $null -ne $FolderSet) {
        $folderRows = @(foreach ($set in $FolderSet) {
                , (@($set.FolderName, [int] $set.FileCount, [int] $set.FolderCount, [long] $set.SizeBytes, [int] $set.Count) + @($set.Folders))
            })
        $sheets.Add((ConvertTo-WorksheetData -Name $script:FolderSheetName -Column $script:FolderColumns -Row $folderRows -Noun 'folder'))
    }
    $sheets.Add((Get-RulesSheetData -IncludeFolders:($sheets.Count -gt 1) -ExcludeName $ExcludeName -MinimumSize $MinimumSize))

    $sheetEntries = ''; $overrides = ''; $sheetRels = ''; $filters = ''
    for ($i = 1; $i -le $sheets.Count; $i++) {
        $sheet = $sheets[$i - 1]
        $sheetEntries += "<sheet name=`"$($sheet.Name)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
        $overrides += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $sheetRels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
        if ($sheet.Kind -eq 'Table') {
            $filters += "<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$($i - 1)`" hidden=`"1`">'$($sheet.Name)'!`$A`$$($sheet.HeaderRow):`$$($sheet.LastColumn)`$$($sheet.LastRow)</definedName>"
        }
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

                # Style 0 = default, 1 = bold ($script:StyleBold), 2 = date/time ($script:StyleDate).
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
                        if ($sheet.Kind -eq 'Rules') { Write-RulesSheetXml -Writer $w -Sheet $sheet }
                        else { Write-WorksheetXml -Writer $w -Sheet $sheet }
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
    $ns.AddNamespace('s', $script:SpreadsheetMain)
    $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')
    , $ns  # a namespace manager would otherwise be enumerated into its prefixes
}

function Get-CellText {
    # Text of an inline (<is>) or shared (<si>) string, including rich-text runs (<r>);
    # phonetic runs (<rPh>) are left out. Walks child nodes rather than running XPath
    # queries, and has no parameter validation: this runs for every text cell.
    param([System.Xml.XmlNode] $Node)
    $text = ''
    if ($null -eq $Node) { return $text }
    foreach ($child in $Node.ChildNodes) {
        if ($child.NamespaceURI -ne $script:SpreadsheetMain) { continue }
        if ($child.LocalName -eq 't') { $text += $child.InnerText }
        elseif ($child.LocalName -eq 'r') {
            $run = $child.Item('t', $script:SpreadsheetMain)
            if ($run) { $text += $run.InnerText }
        }
    }
    $text
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
        if (-not $target) { throw "The workbook part for sheet '$($sheet.GetAttribute('name'))' is missing." }
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
        foreach ($item in $sharedXml.SelectNodes('/s:sst/s:si', $ns)) { $shared.Add((Get-CellText -Node $item)) }
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

    # Column numbers by letters, worked out once per sheet: a sheet has few distinct columns
    # but may have millions of cells.
    $columnOf = [System.Collections.Generic.Dictionary[string, int]]::new()
    $digits = '0123456789'.ToCharArray()

    foreach ($row in $sheet.SelectNodes('/s:worksheet/s:sheetData/s:row', $sheetNs)) {
        $cells = [System.Collections.Generic.Dictionary[int, string]]::new()
        $column = 0
        foreach ($cell in $row.ChildNodes) {
            if ($cell.LocalName -ne 'c' -or $cell.NamespaceURI -ne $script:SpreadsheetMain) { continue }
            $reference = $cell.GetAttribute('r')
            # Excel may leave out empty cells, so place each by its reference when present.
            if ($reference) {
                $letters = $reference.TrimEnd($digits)
                if (-not $columnOf.ContainsKey($letters)) { $columnOf[$letters] = ConvertFrom-ColumnName $letters }
                $column = $columnOf[$letters]
            }
            else { $column++ }
            $value = $cell.Item('v', $script:SpreadsheetMain)
            $type = $cell.GetAttribute('t')
            if ($type -eq 's') { $cells[$column] = $SharedString[[int] $value.InnerText] }
            elseif ($type -eq 'inlineStr') {
                # Plain text is a lone <t>; take it directly (a function call per cell is slow).
                $inline = $cell.Item('is', $script:SpreadsheetMain)
                $only = $null
                if ($null -ne $inline -and $inline.ChildNodes.Count -eq 1) { $only = $inline.FirstChild }
                if ($null -ne $only -and $only.LocalName -eq 't') { $cells[$column] = $only.InnerText }
                else { $cells[$column] = Get-CellText -Node $inline }
            }
            elseif ($value) { $cells[$column] = $value.InnerText }
            else { $cells[$column] = '' }
        }

        $width = 0
        foreach ($c in $cells.Keys) { $width = [Math]::Max($width, $c) }
        $values = [string[]]::new($width)
        foreach ($c in $cells.Keys) { $values[$c - 1] = $cells[$c] }
        , $values
    }
}

function Test-ReportHeader {
    # True when a row starts with the given column headers (ignoring case).
    param([object[]] $Values, [string[]] $Header)
    if ($null -eq $Values -or $Values.Count -lt $Header.Count) { return $false }
    for ($c = 0; $c -lt $Header.Count; $c++) {
        if (-not [System.StringComparer]::OrdinalIgnoreCase.Equals([string] $Values[$c], $Header[$c])) { return $false }
    }
    $true
}

function Select-ReportRow {
    # Finds a sheet's table header row (row 1; lower in reports that had the rules above the
    # table) and returns the data rows under it: fixed values, folders, and whether the
    # header was $LegacyColumn's (a report made before a column was added).
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Row,
        [Parameter(Mandatory)] [object[]] $Column,
        [object[]] $LegacyColumn = @(),
        [Parameter(Mandatory)] [string] $ErrorMessage
    )

    $expected = [string[]] @($Column | ForEach-Object { $_.Header })
    $legacy = [string[]] @($LegacyColumn | ForEach-Object { $_.Header })
    $headerAt = -1
    $isLegacy = $false
    for ($r = 0; $headerAt -lt 0 -and $r -lt $Row.Count; $r++) {
        if (Test-ReportHeader -Values $Row[$r] -Header $expected) { $headerAt = $r }
        elseif ($legacy.Count -and (Test-ReportHeader -Values $Row[$r] -Header $legacy)) { $headerAt = $r; $isLegacy = $true }
    }
    if ($headerAt -lt 0) { throw "$ErrorMessage '$($expected -join ', ')'." }

    $firstLocation = $expected.Count
    if ($isLegacy) { $firstLocation = $legacy.Count }
    for ($r = $headerAt + 1; $r -lt $Row.Count; $r++) {
        $values = $Row[$r]
        if ($values.Count -lt $firstLocation -or -not $values[0]) { continue }
        # (Not "$x = if ...": that would unroll an empty or one-item array.)
        $folders = [System.Collections.Generic.List[string]]::new()
        for ($c = $firstLocation; $c -lt $values.Count; $c++) { if ($values[$c]) { $folders.Add($values[$c]) } }
        [pscustomobject] @{ Values = $values; Folders = $folders.ToArray(); Legacy = $isLegacy }
    }
}

function ConvertFrom-CellNumber {
    # A numeric cell's text as a number, or a clear error naming the report. A simple
    # function (no parameter validation): it runs for several cells of every row.
    param([string] $Text, [string] $Path)
    $number = 0.0
    if (-not [double]::TryParse($Text, [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref] $number)) {
        throw "'$Path' could not be read as a duplicates report: '$Text' is not a number."
    }
    $number
}

function Read-DuplicateWorkbook {
    <#
        Reads a report: Files, Folders ($null when the report has no folder sheet) and the
        scan settings recorded on the Rules sheet (ExcludeName, MinimumSize).
    #>
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
            $rulesSheet = @($sheets | Where-Object { $_.Name -eq $script:RulesSheetName })
            $rulesRows = @()
            if ($rulesSheet) { $rulesRows = @(Get-WorksheetRow -Archive $zip -SheetPath $rulesSheet[0].Path -SharedString $shared) }
        }
        finally { $zip.Dispose() }
    }
    catch [System.IO.InvalidDataException] { throw "'$Path' is not an Excel workbook." }
    finally { $stream.Dispose() }

    # Numbers go through [double] then [long]/[int], which rounds half to even like Python's round().
    $files = @(foreach ($row in (Select-ReportRow -Row $fileRows -Column $script:FixedColumns -LegacyColumn $script:LegacyFixedColumns `
                    -ErrorMessage "'$Path' is not a duplicates report: it has no header row")) {
            # Reports made before the UTC Offset column have no offset: they are checked by local time.
            $offset = $null
            $at = 2
            if (-not $row.Legacy) { $offset = ConvertFrom-UtcOffsetText $row.Values[2] -Path $Path; $at = 3 }
            [pscustomobject] @{
                FileName      = $row.Values[0]
                LastWriteTime = [datetime]::FromOADate((ConvertFrom-CellNumber $row.Values[1] -Path $Path))
                UtcOffset     = $offset
                SizeBytes     = [long] (ConvertFrom-CellNumber $row.Values[$at] -Path $Path)
                MD5           = $row.Values[$at + 1]
                Count         = $row.Folders.Count
                Folders       = $row.Folders
            }
        })

    $folders = $null
    if ($folderSheet) {
        $folders = @(foreach ($row in (Select-ReportRow -Row $folderRows -Column $script:FolderColumns `
                        -ErrorMessage "'$Path' is not a duplicates report: sheet '$($script:FolderSheetName)' has no header row")) {
                [pscustomobject] @{
                    FolderName  = $row.Values[0]
                    FileCount   = [int] (ConvertFrom-CellNumber $row.Values[1] -Path $Path)
                    FolderCount = [int] (ConvertFrom-CellNumber $row.Values[2] -Path $Path)
                    SizeBytes   = [long] (ConvertFrom-CellNumber $row.Values[3] -Path $Path)
                    Count       = $row.Folders.Count
                    Folders     = $row.Folders
                }
            })
    }

    $settings = Get-ScanSetting -Row $rulesRows -Path $Path
    [pscustomobject] @{ Files = $files; Folders = $folders; ExcludeName = $settings.ExcludeName; MinimumSize = $settings.MinimumSize }
}

function Get-ScanSetting {
    <#
        The scan settings from the Rules sheet's rows (see Get-RulesSheetData); none when absent.
        A smallest size that is not a whole number of bytes (the report was edited) is ignored
        with a warning: it only ever narrowed the scan, so the rest of the report still holds.
    #>
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Row, [Parameter(Mandatory)] [string] $Path)
    $excludeName = [System.Collections.Generic.List[string]]::new()
    $minimumSize = [long] 0
    foreach ($values in $Row) {
        if ($null -eq $values -or $values.Count -lt 2) { continue }
        $cells = [string[]] @(for ($c = 1; $c -lt $values.Count; $c++) { if ($values[$c]) { $values[$c] } })
        if ($values[0] -ceq $script:ExcludeNameLabel) {
            foreach ($pattern in $cells) {
                $problem = Get-NamePatternProblem -Pattern $pattern
                if ($problem) { Write-Warning "Ignoring an exclusion pattern in '$Path': $problem" }
                else { $excludeName.Add($pattern) }
            }
        }
        elseif ($values[0] -ceq $script:MinimumSizeLabel -and $cells.Count) {
            $number = 0.0
            if ([double]::TryParse($cells[0], [System.Globalization.NumberStyles]::Float,
                    [System.Globalization.CultureInfo]::InvariantCulture, [ref] $number) -and
                $number -ge 0 -and $number -lt [long]::MaxValue -and $number -eq [Math]::Floor($number)) {
                $minimumSize = [long] $number
            }
            else { Write-Warning "Ignoring the smallest file size in '$Path': '$($cells[0])' is not a whole number of bytes." }
        }
    }
    [pscustomobject] @{ ExcludeName = $excludeName.ToArray(); MinimumSize = $minimumSize }
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

function Find-FileByNameKey {
    # The file in $Folder whose name matches $FileName ignoring case and Unicode form, or
    # $null. Used when the plain lookup fails: case-sensitive file systems (Linux), and
    # names stored in another Unicode form (Windows and Linux keep both forms apart).
    param([Parameter(Mandatory)] [string] $Folder, [Parameter(Mandatory)] [string] $FileName)
    if (-not [System.IO.Directory]::Exists($Folder)) { return $null }
    $wanted = ConvertTo-NameKey $FileName
    foreach ($candidate in [System.IO.Directory]::GetFiles($Folder)) {
        if ((ConvertTo-NameKey ([System.IO.Path]::GetFileName($candidate))) -ceq $wanted) {
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
          Unavailable - its drive or network share cannot be reached, or its size and saved
                        date cannot be read, as for names Windows reserves for devices (kept as is)
    #>
    param(
        [Parameter(Mandatory)] [string] $Folder,
        [Parameter(Mandatory)] [string] $FileName,
        [Parameter(Mandatory)] [long] $SizeBytes,
        [Parameter(Mandatory)] [datetime] $LastWriteTime,
        # The report's UTC offset for $LastWriteTime; without one (older reports) the saved
        # date is compared as local time on this computer.
        [AllowNull()] [object] $UtcOffset,
        [System.Collections.Generic.Dictionary[string, bool]] $RootCache
    )

    try {
        if (-not (Test-PathRootReachable -Folder $Folder -Cache $RootCache)) { return 'Unavailable' }
        $file = Find-CopyFile -Folder $Folder -FileName $FileName
    }
    catch { return 'Unavailable' }  # something that cannot be checked is kept, never removed
    if (-not $file) { return 'Missing' }
    Get-CopyState -File $file -SizeBytes $SizeBytes -LastWriteTime $LastWriteTime -UtcOffset $UtcOffset
}

function Find-CopyFile {
    # A recorded copy's file: by its exact name, else ignoring case and Unicode form; $null
    # when there is none.
    param([Parameter(Mandatory)] [string] $Folder, [Parameter(Mandatory)] [string] $FileName)
    $path = [System.IO.Path]::Combine($Folder, $FileName)
    if ([System.IO.File]::Exists($path)) { return [System.IO.FileInfo] $path }
    Find-FileByNameKey -Folder $Folder -FileName $FileName
}

function Get-CopyState {
    # Present or Changed for a copy's file, by its size and saved date (see Test-DuplicateCopy);
    # Missing when it was deleted while being checked; Unavailable when its details cannot be read.
    # A simple function (no parameter validation): it runs for every copy in a report.
    param([object] $File, [long] $SizeBytes, [datetime] $LastWriteTime, [object] $UtcOffset)
    try {
        # (Getter methods, not properties: PowerShell turns a failing property into $null.)
        $size  = $File.get_Length()
        $local = $File.get_LastWriteTime()
        $utc   = $File.get_LastWriteTimeUtc()
    }
    catch {
        $reason = $_.Exception
        while ($reason.InnerException) { $reason = $reason.InnerException }
        if ($reason -is [System.IO.FileNotFoundException]) { return 'Missing' }
        return 'Unavailable'
    }
    if ($size -ne $SizeBytes -or -not (Test-SameSavedDate -Local $local -Utc $utc -LastWriteTime $LastWriteTime -UtcOffset $UtcOffset)) {
        return 'Changed'
    }
    'Present'
}

function Test-CopyInFolder {
    <#
        Checks the recorded copies in one folder (see Test-DuplicateCopy), whose drive or
        share is reachable. Several copies are checked from a single listing of the folder
        rather than a lookup each: on a network share every lookup is a round trip.
        $Check holds Key, FileName, SizeBytes, LastWriteTime and UtcOffset; returns Key and
        State for each.
    #>
    param([Parameter(Mandatory)] [string] $Folder, [Parameter(Mandatory)] [object[]] $Check)

    if ($Check.Count -eq 1) {
        $c = $Check[0]
        $state = 'Missing'
        try { $file = Find-CopyFile -Folder $Folder -FileName $c.FileName }
        catch { $file = $null; $state = 'Unavailable' }
        if ($file) { $state = Get-CopyState -File $file -SizeBytes $c.SizeBytes -LastWriteTime $c.LastWriteTime -UtcOffset $c.UtcOffset }
        return [pscustomobject] @{ Key = $c.Key; State = $state }
    }

    $failure = $null
    $files = @()
    try { $files = ([System.IO.DirectoryInfo] $Folder).GetFiles() }
    catch {
        $reason = $_.Exception
        while ($reason.InnerException) { $reason = $reason.InnerException }
        $failure = 'Unavailable'
        if ($reason -is [System.IO.DirectoryNotFoundException]) { $failure = 'Missing' }
    }
    # Names searched by .NET (Array.IndexOf) rather than indexed in a PowerShell loop:
    # folders can hold thousands of files and only a few copies.
    # (A loop, not $files.Name: under strict mode that fails when the folder is empty.)
    $names = [string[]] @(foreach ($file in $files) { $file.Name })

    $keys = $null  # name keys, worked out only when a name is not found as it is
    foreach ($c in $Check) {
        $state = $failure
        if (-not $state) {
            $file = $null
            $at = [System.Array]::IndexOf($names, $c.FileName)
            if ($at -lt 0) {
                if ($null -eq $keys) { $keys = [string[]] @(foreach ($name in $names) { ConvertTo-NameKey $name }) }
                $at = [System.Array]::IndexOf($keys, (ConvertTo-NameKey $c.FileName))
            }
            if ($at -ge 0) { $file = $files[$at] }
            $state = 'Missing'
            if ($file) { $state = Get-CopyState -File $file -SizeBytes $c.SizeBytes -LastWriteTime $c.LastWriteTime -UtcOffset $c.UtcOffset }
        }
        [pscustomobject] @{ Key = $c.Key; State = $state }
    }
}

function Invoke-WorkItem {
    # Runs $Work (taking one item) on each item, $ThrottleLimit at a time, showing progress,
    # and returns everything the work outputs. A failure on any item is rethrown.
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Item,
        [Parameter(Mandatory)] [scriptblock] $Work,
        [ValidateRange(1, 64)] [int] $ThrottleLimit = 1,
        [Parameter(Mandatory)] [string] $Activity
    )
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $lastShownMs = - $script:ProgressIntervalMs
    if ($ThrottleLimit -eq 1 -or $Item.Count -lt 2) {
        for ($i = 0; $i -lt $Item.Count; $i++) {
            if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
                $lastShownMs = $timer.ElapsedMilliseconds
                Write-Progress -Id 3 -Activity $Activity -Status "$($i + 1) of $($Item.Count)" -PercentComplete ([int] (100 * $i / $Item.Count))
            }
            & $Work $Item[$i]
        }
    }
    else {
        $pool = $null
        try {
            $pool = Open-WorkerPool -Work $Work -ThrottleLimit ([Math]::Min($ThrottleLimit, $Item.Count))
            foreach ($one in $Item) { $pool.Queue.Add($one) }
            for ($done = 1; $done -le $Item.Count; $done++) {
                $result = Receive-WorkerResult -Pool $pool
                if ($null -ne $result.Error) { throw $result.Error }
                if ($timer.ElapsedMilliseconds - $lastShownMs -ge $script:ProgressIntervalMs) {
                    $lastShownMs = $timer.ElapsedMilliseconds
                    Write-Progress -Id 3 -Activity $Activity -Status "$done of $($Item.Count) ($ThrottleLimit at a time)" `
                        -PercentComplete ([int] (100 * $done / $Item.Count))
                }
                $result.Value
            }
        }
        finally {
            if ($null -ne $pool) { Close-WorkerPool -Pool $pool }
        }
    }
    Write-Progress -Id 3 -Activity $Activity -Completed
}

function Get-FileCopyState {
    <#
        Checks every copy of every file row; returns "row number|folder" -> state (see
        Test-DuplicateCopy). Copies on a drive or share that cannot be reached are
        Unavailable without further checks (each drive or share is tried once). The others
        are grouped by folder and each folder checked with Test-CopyInFolder, $ThrottleLimit
        folders at a time.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Row,
        [ValidateRange(1, 64)] [int] $ThrottleLimit = 1,
        [System.Collections.Generic.Dictionary[string, bool]] $RootCache
    )
    $states = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $byFolder = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]]::new([System.StringComparer]::Ordinal)
    for ($i = 0; $i -lt $Row.Count; $i++) {
        foreach ($folder in $Row[$i].Folders) {
            $key = "$i|$folder"
            if (-not (Test-PathRootReachable -Folder $folder -Cache $RootCache)) { $states[$key] = 'Unavailable'; continue }
            $checks = $null
            if (-not $byFolder.TryGetValue($folder, [ref] $checks)) {
                $checks = [System.Collections.Generic.List[object]]::new()
                $byFolder[$folder] = $checks
            }
            $checks.Add([pscustomobject] @{
                    Key = $key; FileName = $Row[$i].FileName; SizeBytes = $Row[$i].SizeBytes
                    LastWriteTime = $Row[$i].LastWriteTime; UtcOffset = $Row[$i].UtcOffset
                })
        }
    }
    $jobs = @(foreach ($entry in $byFolder.GetEnumerator()) { [pscustomobject] @{ Folder = $entry.Key; Check = $entry.Value.ToArray() } })
    $work = { param($Job) Test-CopyInFolder -Folder $Job.Folder -Check $Job.Check }
    foreach ($result in (Invoke-WorkItem -Item $jobs -Work $work -ThrottleLimit $ThrottleLimit -Activity 'Validating file copies')) {
        $states[$result.Key] = $result.State
    }
    , $states
}

function Get-FolderCopyState {
    <#
        Checks every copy of every folder row with Test-DuplicateFolderCopy, $ThrottleLimit
        at a time; returns "row number|folder" -> state. Copies on a drive or share that
        cannot be reached are Unavailable without further checks.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Row,
        [ValidateRange(1, 64)] [int] $ThrottleLimit = 1,
        [System.Collections.Generic.Dictionary[string, bool]] $RootCache,
        [string[]] $ExcludeName = @()
    )
    $states = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $jobs = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Row.Count; $i++) {
        foreach ($folder in $Row[$i].Folders) {
            $key = "$i|$folder"
            if (-not (Test-PathRootReachable -Folder $folder -Cache $RootCache)) { $states[$key] = 'Unavailable'; continue }
            $jobs.Add([pscustomobject] @{
                    Key = $key; Path = $folder; FileCount = $Row[$i].FileCount; FolderCount = $Row[$i].FolderCount; SizeBytes = $Row[$i].SizeBytes
                    ExcludeName = $ExcludeName
                })
        }
    }
    $work = {
        param($Job)
        $state = Test-DuplicateFolderCopy -Path $Job.Path -FileCount $Job.FileCount -FolderCount $Job.FolderCount -SizeBytes $Job.SizeBytes `
            -ExcludeName $Job.ExcludeName
        [pscustomobject] @{ Key = $Job.Key; State = $state }
    }
    foreach ($result in (Invoke-WorkItem -Item $jobs.ToArray() -Work $work -ThrottleLimit $ThrottleLimit -Activity 'Validating folder copies')) {
        $states[$result.Key] = $result.State
    }
    , $states
}

function Test-SameSavedDate {
    <#
        True when a file's saved date (as $Local and $Utc) is a report's, to the whole second,
        in exact integer arithmetic as when scanning: compared as instants when the report has
        the UTC offset, as local times otherwise (reports made before the UTC Offset column).
    #>
    # A simple function (no parameter validation): it runs for every copy in a report.
    param([datetime] $Local, [datetime] $Utc, [datetime] $LastWriteTime, [object] $UtcOffset)
    $ticks = $Local.Ticks
    $saved = $LastWriteTime.Ticks
    if ($null -ne $UtcOffset) { $ticks = $Utc.Ticks; $saved -= ([TimeSpan] $UtcOffset).Ticks }
    $ticksPerSecond = [System.TimeSpan]::TicksPerSecond
    ($ticks - ($ticks % $ticksPerSecond)) -eq ($saved - ($saved % $ticksPerSecond))
}

function Get-PreviousMd5 {
    <#
    .SYNOPSIS
        MD5 hashes to take from an earlier report instead of reading the files again.
    .DESCRIPTION
        Returns full path -> MD5 for each scanned file that the report lists (same folder
        and name, ignoring case) whose size and saved date are still the report's. Only
        those files' details are looked up. MD5s that are not 32 hexadecimal digits (an
        edited report) are ignored.
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.IO.FileInfo[]] $File
    )

    $rows = @((Read-DuplicateWorkbook -Path $Path).Files)
    $names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $rows) { $null = $names.Add($row.FileName) }

    # Folder + separator + name key -> scanned file, for the names the report lists.
    $separator = [string] [char] 0
    $scanned = [System.Collections.Generic.Dictionary[string, System.IO.FileInfo]]::new([System.StringComparer]::Ordinal)
    foreach ($f in $File) {
        if (-not $names.Contains($f.Name)) { continue }
        $key = $f.DirectoryName + $separator + (ConvertTo-NameKey $f.Name)
        if (-not $scanned.ContainsKey($key)) { $scanned[$key] = $f }
    }

    $previous = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    foreach ($row in $rows) {
        if ([string] $row.MD5 -notmatch '^[0-9A-Fa-f]{32}$') { continue }
        $name = $separator + (ConvertTo-NameKey $row.FileName)
        foreach ($folder in $row.Folders) {
            $f = $null
            if (-not $scanned.TryGetValue($folder + $name, [ref] $f)) { continue }
            # A file whose details cannot be read is simply hashed (and reported on) as usual.
            try { $size = $f.get_Length(); $local = $f.get_LastWriteTime(); $utc = $f.get_LastWriteTimeUtc() }
            catch { continue }
            if ($size -eq $row.SizeBytes -and (Test-SameSavedDate -Local $local -Utc $utc -LastWriteTime $row.LastWriteTime -UtcOffset $row.UtcOffset)) {
                $previous[$f.FullName] = ([string] $row.MD5).ToUpperInvariant()
            }
        }
    }
    , $previous
}

function Test-DuplicateFolderCopy {
    <#
        Checks one recorded folder copy by listing its tree again (no file contents are read):
          Present     - still there with the same number of files and sub folders and total size
          Missing     - no longer there
          Changed     - still there but its files, sub folders or total size changed
          Unavailable - its drive or share cannot be reached, or part of it cannot be read (kept)
        Files and folders whose names match $ExcludeName (the scan's -Exclude) are left out.
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [int] $FileCount,
        [Parameter(Mandatory)] [int] $FolderCount,
        [Parameter(Mandatory)] [long] $SizeBytes,
        [System.Collections.Generic.Dictionary[string, bool]] $RootCache,
        [string[]] $ExcludeName = @()
    )

    try {
        if (-not (Test-PathRootReachable -Folder $Path -Cache $RootCache)) { return 'Unavailable' }
        if (-not [System.IO.Directory]::Exists($Path)) { return 'Missing' }
        $folders = [System.Collections.Generic.List[object]]::new()
        $files = @(Get-FileInventory -Path $Path -FolderInfo $folders -ExcludeName $ExcludeName -Verbose:$false -WarningAction SilentlyContinue)
        if (@($folders | Where-Object { -not $_.Readable }).Count -gt 0) { return 'Unavailable' }
        $size = [long] 0
        foreach ($f in $files) { $size += $f.get_Length() }  # a file gone since the listing throws here
    }
    catch { return 'Unavailable' }

    if ($files.Count -ne $FileCount -or $folders.Count - 1 -ne $FolderCount -or $size -ne $SizeBytes) { return 'Changed' }
    'Present'
}

function Invoke-CopyCheck {
    <#
        Runs $TestCopy (location, row number) on every copy of every row. Keeps copies that are Present or
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
            $state = & $TestCopy $location $i
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
        Each file copy is checked with a file lookup (copies in the same folder from one
        listing of it), and each folder copy by listing its tree again; no contents are
        read or downloaded. -ThrottleLimit checks that many folders at the same time. Copies that are
        missing or changed are removed; rows left with fewer than two copies are
        removed. Copies on a drive or network share that cannot be reached are kept.
        The report is rewritten in place only when something changed. Supports -WhatIf.
    .OUTPUTS
        A summary object; DuplicateSet holds the file rows that remain and
        DuplicateFolderSet the folder rows ($null when the report has no folder sheet).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string] $Path,

        # How many folders to check at the same time.
        [ValidateRange(1, 64)]
        [int] $ThrottleLimit = 1
    )

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
    $workbook = Read-DuplicateWorkbook -Path $Path
    $rootCache = [System.Collections.Generic.Dictionary[string, bool]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $fileStates = Get-FileCopyState -Row $workbook.Files -ThrottleLimit $ThrottleLimit -RootCache $rootCache
    $files = Invoke-CopyCheck -Set $workbook.Files -NameProperty FileName -Noun '' -TestCopy {
        param($location, $number)
        $fileStates["$number|$location"]
    }

    $folders = $null
    if ($null -ne $workbook.Folders) {
        $folderStates = Get-FolderCopyState -Row $workbook.Folders -ThrottleLimit $ThrottleLimit -RootCache $rootCache `
            -ExcludeName $workbook.ExcludeName
        $folders = Invoke-CopyCheck -Set $workbook.Folders -NameProperty FolderName -Noun 'folder ' -LocationIsItem -TestCopy {
            param($location, $number)
            $folderStates["$number|$location"]
        }
    }

    $removed = $files.Removed + $files.RowsRemoved
    if ($folders) { $removed += $folders.Removed + $folders.RowsRemoved }

    $saved = $false
    if ($removed -gt 0 -and $PSCmdlet.ShouldProcess($Path, 'Remove missing or changed copies')) {
        $export = @{ DuplicateSet = $files.Kept; Path = $Path; ExcludeName = $workbook.ExcludeName; MinimumSize = $workbook.MinimumSize }
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
    ConvertTo-ColumnName, Import-DuplicateReport, Import-DuplicateFolderReport, Update-DuplicateReport, Get-PreviousMd5,
    Test-NetworkDrive, Get-DefaultThrottleLimit, Assert-NamePattern, ConvertFrom-SizeText
