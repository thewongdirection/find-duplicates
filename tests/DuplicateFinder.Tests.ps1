#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.3.0' }
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingBrokenHashAlgorithms', '',
    Justification = 'MD5 is part of the duplicate definition and is not used for security.')]
param()

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $script:ScriptPath = Join-Path $repoRoot 'Find-Duplicates.ps1'
    Import-Module (Join-Path $repoRoot 'src/DuplicateFinder.psm1') -Force

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $script:Saved = [datetime]::new(2024, 5, 17, 10, 30, 0, [System.DateTimeKind]::Utc)
    $script:OnWindows = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT

    function Add-TestFile {
        param(
            [string] $Root,
            [string] $RelativePath,
            [string] $Content = 'same content',
            [datetime] $SavedUtc = $script:Saved
        )
        $full = Join-Path $Root $RelativePath
        $null = New-Item -ItemType Directory -Force -Path (Split-Path -Parent $full)
        [System.IO.File]::WriteAllText($full, $Content)
        [System.IO.File]::SetLastWriteTimeUtc($full, $SavedUtc)
        Get-Item -LiteralPath $full -Force  # -Force: names starting with a dot are hidden on Linux and macOS
    }

    function Add-TestRoot {
        $root = Join-Path $TestDrive ([System.Guid]::NewGuid().ToString('N'))
        (New-Item -ItemType Directory -Path $root).FullName
    }

    function Get-SheetName {
        # The workbook's sheet names, in order.
        param([string] $Path)
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $reader = [System.IO.StreamReader]::new($zip.GetEntry('xl/workbook.xml').Open())
            try { [xml] $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        finally { $zip.Dispose() }
        @($xml.workbook.sheets.sheet | ForEach-Object { $_.name })
    }

    function Read-Worksheet {
        # Returns a worksheet as a list of rows, each row a list of cell texts: from the
        # table's header row down, or with -AllRows every row (for the Rules sheet).
        param([string] $Path, [string] $Part = 'xl/worksheets/sheet1.xml', [switch] $AllRows)
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $reader = [System.IO.StreamReader]::new($zip.GetEntry($Part).Open())
            try { [xml] $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        finally { $zip.Dispose() }

        $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $inTable = [bool] $AllRows
        foreach ($row in $xml.SelectNodes('//s:sheetData/s:row', $ns)) {
            $cells = @(foreach ($cell in $row.SelectNodes('s:c', $ns)) {
                    $text = $cell.SelectSingleNode('s:is/s:t', $ns)
                    if ($text) { $text.InnerText } else { $cell.SelectSingleNode('s:v', $ns).InnerText }
                })
            if (-not $inTable -and $cells[0] -in 'File Name', 'Folder Name') { $inTable = $true }
            if ($inTable) { , $cells }
        }
    }
}

Describe 'ConvertTo-ColumnName' {
    It 'converts <Index> to <Expected>' -ForEach @(
        @{ Index = 1; Expected = 'A' }
        @{ Index = 26; Expected = 'Z' }
        @{ Index = 27; Expected = 'AA' }
        @{ Index = 52; Expected = 'AZ' }
        @{ Index = 702; Expected = 'ZZ' }
        @{ Index = 703; Expected = 'AAA' }
        @{ Index = 16384; Expected = 'XFD' }
    ) {
        ConvertTo-ColumnName $Index | Should -BeExactly $Expected
    }
}

Describe 'Get-FileInventory' {
    It 'finds files in the folder and every sub folder' {
        $root = Add-TestRoot
        $null = Add-TestFile $root 'a.txt'
        $null = Add-TestFile $root 'one/b.txt'
        $null = Add-TestFile $root 'one/two/three/c.txt'

        $names = Get-FileInventory -Path $root | ForEach-Object Name | Sort-Object
        $names | Should -Be @('a.txt', 'b.txt', 'c.txt')
    }

    It 'leaves out excluded files' {
        $root = Add-TestRoot
        $keep = Add-TestFile $root 'keep.txt'
        $skip = Add-TestFile $root 'duplicates.xlsx'

        $found = @(Get-FileInventory -Path $root -ExcludeFile $skip.FullName)
        $found.FullName | Should -Be @($keep.FullName)
    }

    It 'lists folders and files in name order whatever order they were created in' {
        $root = Add-TestRoot
        foreach ($name in 'z', 'a', 'm') { $null = Add-TestFile $root "$name/$name.txt" }

        (Get-FileInventory -Path $root).Name | Should -Be @('a.txt', 'm.txt', 'z.txt')
    }

    It 'handles folder and file names containing wildcard characters' {
        $root = Add-TestRoot
        $null = Add-TestFile $root '[set]/file[1].txt'

        (Get-FileInventory -Path $root).Name | Should -Be 'file[1].txt'
    }

    It 'does not follow a folder link that loops back to the root' {
        $root = Add-TestRoot
        $null = Add-TestFile $root 'sub/a.txt'
        $link = Join-Path $root 'sub/loop'
        try { $null = New-Item -ItemType SymbolicLink -Path $link -Target $root -ErrorAction Stop }
        catch { Set-ItResult -Skipped -Because "symbolic links cannot be created here: $_"; return }

        try {
            $found = @(Get-FileInventory -Path $root)
            $found.Name | Should -Be @('a.txt')
        }
        finally {
            # Remove the loop ourselves: Pester's TestDrive cleanup follows links and
            # would recurse forever. Deleting a link never touches its target.
            [System.IO.Directory]::Delete($link)
        }
    }

    It 'rejects a path that is not a folder' {
        $root = Add-TestRoot
        $file = Add-TestFile $root 'a.txt'
        { Get-FileInventory -Path $file.FullName } | Should -Throw '*is not a folder*'
    }

    It 'reports the folder being scanned' {
        $root = Add-TestRoot
        $null = Add-TestFile $root 'sub/a.txt'

        $verbose = Get-FileInventory -Path $root -Verbose 4>&1 |
            Where-Object { $_ -is [System.Management.Automation.VerboseRecord] }
        $verbose.Message | Should -Contain "Scanning $(Join-Path $root 'sub')"
    }

    It 'lists several folders at a time and returns the same files in the same order' {
        $root = Add-TestRoot
        foreach ($path in 'b/x.txt', 'a/y.txt', 'a/deep/er/z.txt', 'c/w.txt', 'a/b.txt', 'top.txt') { $null = Add-TestFile $root $path }
        $null = New-Item -ItemType Directory -Path (Join-Path $root 'a/empty')
        $link = Join-Path $root 'a/loop'
        $linked = $true
        try { $null = New-Item -ItemType SymbolicLink -Path $link -Target $root -ErrorAction Stop }
        catch { $linked = $false }  # the rest is still worth checking

        try {
            $one = [System.Collections.Generic.List[object]]::new()
            $many = [System.Collections.Generic.List[object]]::new()
            $sequential = @(Get-FileInventory -Path $root -FolderInfo $one)
            $parallel = @(Get-FileInventory -Path $root -FolderInfo $many -ThrottleLimit 4)

            $parallel.FullName | Should -Be $sequential.FullName
            @($many | ForEach-Object { "$($_.Path)|$($_.Readable)" }) | Should -Be @($one | ForEach-Object { "$($_.Path)|$($_.Readable)" })
        }
        finally {
            if ($linked) { [System.IO.Directory]::Delete($link) }  # the link only; Pester's cleanup would loop
        }
    }
}

Describe 'Folder and cloud file detection' {
    It 'follows a <Case>' -ForEach @(
        @{ Case = 'plain folder'; Attributes = [System.IO.FileAttributes]::Directory; LinkType = $null }
        @{ Case = 'cloud-synced (OneDrive) folder'
           Attributes = [System.IO.FileAttributes] 'Directory, ReparsePoint'; LinkType = $null }
    ) {
        InModuleScope DuplicateFinder -Parameters $_ {
            Test-FolderLink -Folder ([pscustomobject] @{ Attributes = $Attributes; LinkType = $LinkType }) |
                Should -BeFalse
        }
    }

    It 'does not follow a <LinkType>' -ForEach @(
        @{ LinkType = 'SymbolicLink' }
        @{ LinkType = 'Junction' }
    ) {
        InModuleScope DuplicateFinder -Parameters $_ {
            $folder = [pscustomobject] @{ Attributes = [System.IO.FileAttributes] 'Directory, ReparsePoint'; LinkType = $LinkType }
            Test-FolderLink -Folder $folder | Should -BeTrue
        }
    }

    It 'treats attributes 0x<Hex> as online-only: <Expected>' -ForEach @(
        @{ Hex = '20'; Expected = $false }       # Archive: a normal local file
        @{ Hex = '420'; Expected = $false }      # Archive + ReparsePoint: pinned / locally available
        @{ Hex = '1020'; Expected = $true }      # Offline
        @{ Hex = '40020'; Expected = $true }     # RecallOnOpen
        @{ Hex = '400420'; Expected = $true }    # RecallOnDataAccess (OneDrive Files On-Demand)
    ) {
        InModuleScope DuplicateFinder -Parameters $_ {
            # A plain number: .NET Framework's FileAttributes enum does not define the cloud bits.
            $file = [pscustomobject] @{ Attributes = [Convert]::ToInt32($Hex, 16) }
            Test-CloudOnlyFile -File $file | Should -Be $Expected
        }
    }
}

Describe 'Find-DuplicateFile' {
    It 'records every copy when a file is duplicated in many folders' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/report.doc'
            Add-TestFile $root 'b/report.doc'
            Add-TestFile $root 'b/c/d/report.doc'
        )

        $result = @(Find-DuplicateFile -File $files)

        $result.Count | Should -Be 1
        $result[0].FileName | Should -Be 'report.doc'
        $result[0].Count | Should -Be 3
        $result[0].Folders | Should -Be @($files.DirectoryName | Sort-Object)
        $result[0].MD5 | Should -Be (Get-FileHash -LiteralPath $files[0].FullName -Algorithm MD5).Hash
    }

    It 'ignores files whose contents differ' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/x.txt' -Content 'aaaa'
            Add-TestFile $root 'b/x.txt' -Content 'bbbb'
        )
        @(Find-DuplicateFile -File $files).Count | Should -Be 0
    }

    It 'ignores files whose saved dates differ' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/x.txt'
            Add-TestFile $root 'b/x.txt' -SavedUtc $script:Saved.AddMinutes(1)
        )
        @(Find-DuplicateFile -File $files).Count | Should -Be 0
    }

    It 'ignores files whose names differ' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/x.txt'
            Add-TestFile $root 'b/y.txt'
        )
        @(Find-DuplicateFile -File $files).Count | Should -Be 0
    }

    It 'treats names that differ only by case as the same name' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/Photo.JPG'
            Add-TestFile $root 'b/photo.jpg'
        )
        @(Find-DuplicateFile -File $files).Count | Should -Be 1
    }

    It 'does not match names that differ beyond case' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root "a/stra$([char] 0xDF)e.txt"
            Add-TestFile $root 'b/STRASSE.txt'
        )
        @(Find-DuplicateFile -File $files).Count | Should -Be 0
    }

    It 'ignores sub-second differences in the saved date' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/x.txt'
            Add-TestFile $root 'b/x.txt' -SavedUtc $script:Saved.AddMilliseconds(400)
        )
        @(Find-DuplicateFile -File $files).Count | Should -Be 1
    }

    It 'splits same-name, same-date files into separate sets by content' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/x.txt' -Content 'first'
            Add-TestFile $root 'b/x.txt' -Content 'first'
            Add-TestFile $root 'c/x.txt' -Content 'other'
            Add-TestFile $root 'd/x.txt' -Content 'other'
            Add-TestFile $root 'e/x.txt' -Content 'unique!'
        )

        $result = @(Find-DuplicateFile -File $files)

        $result.Count | Should -Be 2
        $result | ForEach-Object { $_.Count | Should -Be 2 }
        ($result.Folders | Sort-Object) | Should -Be @($files[0..3].DirectoryName | Sort-Object)
    }

    It 'names a duplicate after the copy in the first folder by name' {
        $root = Add-TestRoot
        $null = Add-TestFile $root 'b/photo.jpg'
        $null = Add-TestFile $root 'a/Photo.JPG'
        $scanned = @(Get-FileInventory -Path $root)

        (Find-DuplicateFile -File $scanned).FileName | Should -BeExactly 'Photo.JPG'
    }

    It 'includes files of 0 bytes by default' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/empty.txt' -Content ''
            Add-TestFile $root 'b/empty.txt' -Content ''
        )
        @(Find-DuplicateFile -File $files).Count | Should -Be 1
    }

    It 'leaves files of 0 bytes out with -IgnoreEmptyFiles' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/empty.txt' -Content ''
            Add-TestFile $root 'b/empty.txt' -Content ''
            Add-TestFile $root 'a/full.txt'
            Add-TestFile $root 'b/full.txt'
        )
        (Find-DuplicateFile -File $files -IgnoreEmptyFiles).FileName | Should -Be @('full.txt')
    }

    It 'returns nothing for an empty list' {
        @(Find-DuplicateFile -File @()).Count | Should -Be 0
    }

    It 'never reads the size or saved date of a file whose name no other file has' {
        $root = Add-TestRoot
        $files = @(Add-TestFile $root 'a/x.txt'; Add-TestFile $root 'b/x.txt')
        # Never created: reading its size or saved date would fail with a warning.
        $unique = [System.IO.FileInfo] (Join-Path $root 'c/unique.txt')

        $result = @(Find-DuplicateFile -File ($files + $unique) -WarningVariable warnings -WarningAction SilentlyContinue)

        @($warnings).Count | Should -Be 0
        $result.Count | Should -Be 1
    }

    It 'skips a file that has gone since the scan, with a warning' {
        $root = Add-TestRoot
        $files = @(Add-TestFile $root 'a/x.txt'; Add-TestFile $root 'b/x.txt')
        $gone = [System.IO.FileInfo] (Join-Path $root 'c/x.txt')  # listed by the scan, deleted before it was compared

        $result = @(Find-DuplicateFile -File ($files + $gone) -WarningVariable warnings -WarningAction SilentlyContinue)

        "$warnings" | Should -BeLike "Skipping '$($gone.FullName)'*"
        $result[0].Count | Should -Be 2
    }

    Context 'MD5 is only calculated when name and saved date already match' {
        BeforeEach {
            Mock -ModuleName DuplicateFinder Get-FileMd5 { 'ABC' }
        }

        It 'does not hash files with different names' {
            $root = Add-TestRoot
            $null = Find-DuplicateFile -File @((Add-TestFile $root 'a/x.txt'), (Add-TestFile $root 'b/y.txt'))
            Should -Invoke -ModuleName DuplicateFinder Get-FileMd5 -Times 0 -Exactly
        }

        It 'does not hash files with different saved dates' {
            $root = Add-TestRoot
            $null = Find-DuplicateFile -File @(
                (Add-TestFile $root 'a/x.txt'),
                (Add-TestFile $root 'b/x.txt' -SavedUtc $script:Saved.AddDays(1)))
            Should -Invoke -ModuleName DuplicateFinder Get-FileMd5 -Times 0 -Exactly
        }

        It 'does not hash files of different sizes' {
            $root = Add-TestRoot
            $null = Find-DuplicateFile -File @(
                (Add-TestFile $root 'a/x.txt' -Content 'short'),
                (Add-TestFile $root 'b/x.txt' -Content 'much longer'))
            Should -Invoke -ModuleName DuplicateFinder Get-FileMd5 -Times 0 -Exactly
        }

        It 'hashes only the matching candidates' {
            $root = Add-TestRoot
            $null = Find-DuplicateFile -File @(
                (Add-TestFile $root 'a/x.txt'),
                (Add-TestFile $root 'b/x.txt'),
                (Add-TestFile $root 'c/y.txt'))
            Should -Invoke -ModuleName DuplicateFinder Get-FileMd5 -Times 2 -Exactly
        }
    }

    It 'downloads (hashes) online-only cloud files by default' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'cloud/x.txt'
            Add-TestFile $root 'local/x.txt'
        )
        Mock -ModuleName DuplicateFinder Test-CloudOnlyFile { $true }

        @(Find-DuplicateFile -File $files).Count | Should -Be 1
    }

    It 'does not download online-only cloud files with -SkipCloudOnly' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'cloud/x.txt'
            Add-TestFile $root 'local1/x.txt'
            Add-TestFile $root 'local2/x.txt'
        )
        Mock -ModuleName DuplicateFinder Test-CloudOnlyFile { "$($File.FullName)" -like '*cloud*' }
        Mock -ModuleName DuplicateFinder Get-FileMd5 { 'SAME' }

        $result = @(Find-DuplicateFile -File $files -SkipCloudOnly -WarningVariable warnings -WarningAction SilentlyContinue)

        Should -Invoke -ModuleName DuplicateFinder Get-FileMd5 -Times 2 -Exactly
        Should -Invoke -ModuleName DuplicateFinder Get-FileMd5 -Times 0 -Exactly -ParameterFilter { $Path -like '*cloud*' }
        $result[0].Folders | Should -Be @($files[1..2].DirectoryName | Sort-Object)
        "$($warnings[0])" | Should -BeLike '1 online-only*'
    }

    It 'skips a file it cannot hash and keeps the rest' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'locked/x.txt'
            Add-TestFile $root 'b/x.txt'
            Add-TestFile $root 'c/x.txt'
        )
        # Decide by path alone so the mock needs no captured variables.
        Mock -ModuleName DuplicateFinder Get-FileMd5 {
            if ($Path -like '*locked*') { throw 'file is locked' }
            'SAME'
        }

        $result = @(Find-DuplicateFile -File $files -WarningVariable warnings -WarningAction SilentlyContinue)

        @($warnings).Count | Should -Be 1
        "$($warnings[0])" | Should -BeLike '*locked*'
        $result.Count | Should -Be 1
        $result[0].Folders | Should -Be @($files[1..2].DirectoryName | Sort-Object)
    }

    It 'computes the MD5 of a file' {
        $path = (Add-TestFile (Add-TestRoot) 'abc.txt' -Content 'abc').FullName
        InModuleScope DuplicateFinder -Parameters @{ Path = $path } {
            Get-FileMd5 -Path $Path | Should -BeExactly '900150983CD24FB0D6963F7D28E17F72'
        }
    }

    Context 'Hashing several files at a time (-ThrottleLimit)' {
        BeforeAll {
            $script:ParallelRoot = Add-TestRoot
            foreach ($folder in 'a', 'b', 'c', 'd') {
                $null = Add-TestFile $script:ParallelRoot "$folder/same.txt" -Content 'same'
                $null = Add-TestFile $script:ParallelRoot "$folder/split.txt" -Content "half $([int] ($folder -in 'a', 'b'))"
            }
            $script:ParallelFiles = @(Get-FileInventory -Path $script:ParallelRoot)
            $script:Sequential = @(Find-DuplicateFile -File $script:ParallelFiles)
        }

        It 'finds the same duplicates hashing <Limit> files at a time' -ForEach @(
            @{ Limit = 2 }
            @{ Limit = 8 }
        ) {
            $parallel = @(Find-DuplicateFile -File $script:ParallelFiles -ThrottleLimit $Limit)

            $parallel.Count | Should -Be 3
            for ($i = 0; $i -lt $parallel.Count; $i++) {
                $parallel[$i].FileName | Should -Be $script:Sequential[$i].FileName
                $parallel[$i].MD5 | Should -Be $script:Sequential[$i].MD5
                $parallel[$i].Folders | Should -Be $script:Sequential[$i].Folders
            }
        }

        It 'reports a file it cannot read and keeps the rest (<Limit> at a time)' -ForEach @(
            @{ Limit = 1 }
            @{ Limit = 4 }
        ) {
            $root = Add-TestRoot
            $null = Add-TestFile $root 'a/x.txt'
            $null = Add-TestFile $root 'b/x.txt'
            $null = Add-TestFile $root 'c/x.txt'
            $files = @(Get-FileInventory -Path $root)
            Remove-Item -LiteralPath $files[0].FullName  # gone before it could be hashed

            $result = @(Find-DuplicateFile -File $files -ThrottleLimit $Limit -WarningVariable warnings -WarningAction SilentlyContinue)

            @($warnings).Count | Should -Be 1
            "$($warnings[0])" | Should -BeLike "*$($files[0].Name)*"
            $result.Count | Should -Be 1
            $result[0].Folders | Should -Be @($files[1..2].DirectoryName | Sort-Object)
        }

        It 'rejects a throttle limit of <Limit>' -ForEach @(
            @{ Limit = 0 }
            @{ Limit = 65 }
        ) {
            { Find-DuplicateFile -File @() -ThrottleLimit $Limit } | Should -Throw
        }
    }
}

Describe 'Export-DuplicateReport' {
    BeforeAll {
        $script:Sets = @(
            [pscustomobject] @{
                FileName = 'a & b <1>.txt'; LastWriteTime = [datetime]::new(2024, 1, 2, 3, 4, 5)
                UtcOffset = [TimeSpan]::FromMinutes(330)
                SizeBytes = 1234; MD5 = 'AAAA'; Count = 3
                Folders = [string[]] @('C:\one', 'C:\two & more', 'D:\three')
            }
            [pscustomobject] @{
                FileName = 'z.txt'; LastWriteTime = [datetime]::new(2023, 6, 7, 8, 9, 10)
                SizeBytes = 5; MD5 = 'BBBB'; Count = 2
                Folders = [string[]] @('C:\x', "C:\bad$([char] 1)name")
            }
        )
    }

    It 'writes one row per duplicated file with a column per location' {
        $out = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet $script:Sets -Path $out

        $rows = @(Read-Worksheet $out)
        $rows.Count | Should -Be 3
        $rows[0] | Should -Be @('File Name', 'Last Modified', 'UTC Offset', 'Size (bytes)', 'MD5', 'Copies', 'Location 1', 'Location 2', 'Location 3')
        $rows[1][0] | Should -BeExactly 'a & b <1>.txt'
        $rows[1][2] | Should -Be '+05:30'
        $rows[1][3] | Should -Be '1234'
        $rows[1][4] | Should -Be 'AAAA'
        $rows[1][5] | Should -Be '3'
        $rows[1][6..8] | Should -Be @('C:\one', 'C:\two & more', 'D:\three')
        $rows[2][2] | Should -Match '^[+-]\d{2}:\d{2}' -Because "a set without an offset gets this computer's"
        $rows[2][6..7] | Should -Be @('C:\x', "C:\bad$([char] 0xFFFD)name")
    }

    It 'stores the saved date as a real Excel date' {
        $out = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet $script:Sets -Path $out

        $serial = [double]::Parse((Read-Worksheet $out)[1][1], [System.Globalization.CultureInfo]::InvariantCulture)
        [datetime]::FromOADate($serial) | Should -Be $script:Sets[0].LastWriteTime
    }

    It 'produces a package with every required part' {
        $out = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet $script:Sets -Path $out

        $zip = [System.IO.Compression.ZipFile]::OpenRead($out)
        try {
            $expected = '[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels',
                'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'
            ($zip.Entries.FullName | Sort-Object) | Should -Be ($expected | Sort-Object)
            foreach ($entry in $zip.Entries) {
                $reader = [System.IO.StreamReader]::new($entry.Open())
                try { { [xml] $reader.ReadToEnd() } | Should -Not -Throw -Because $entry.FullName }
                finally { $reader.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }

    It 'writes the matching rules on a Rules sheet' {
        $out = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet $script:Sets -Path $out

        Get-SheetName $out | Should -Be @('Duplicates', 'Rules')
        $rules = @(Read-Worksheet $out -Part 'xl/worksheets/sheet2.xml' -AllRows | ForEach-Object { $_[0] })
        $rules[0] | Should -Be 'Matching rules'
        $rules | Should -Contain "Sheet 'Duplicates': duplicate files"
        @($rules | Where-Object { $_ -like 'A file is listed when another file has ALL of*' }).Count | Should -Be 1
        $rules | Should -Not -Contain "Sheet 'Duplicate Folders': duplicate folders" -Because 'no folder sheet, no folder rules'
    }

    It 'starts the table on row 1' {
        $out = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet $script:Sets -Path $out

        (Read-Worksheet $out -AllRows)[0][0] | Should -Be 'File Name'
        $zip = [System.IO.Compression.ZipFile]::OpenRead($out)
        try {
            $reader = [System.IO.StreamReader]::new($zip.GetEntry('xl/worksheets/sheet1.xml').Open())
            try { $sheetXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        finally { $zip.Dispose() }
        $sheetXml | Should -Match '<pane ySplit="1" topLeftCell="A2"'
        $sheetXml | Should -Match '<autoFilter ref="A1:I3"'
    }

    It 'writes a header-only workbook when there are no duplicates' {
        $out = Join-Path (Add-TestRoot) 'empty.xlsx'
        Export-DuplicateReport -DuplicateSet @() -Path $out

        $rows = @(Read-Worksheet $out)
        $rows.Count | Should -Be 1
        $rows[0][-1] | Should -Be 'Location 1'
    }

    It 'overwrites an existing report and leaves no temporary files' {
        $dir = Add-TestRoot
        $out = Join-Path $dir 'report.xlsx'
        Set-Content -LiteralPath $out -Value 'old'

        Export-DuplicateReport -DuplicateSet $script:Sets -Path $out

        @(Read-Worksheet $out).Count | Should -Be 3
        @(Get-ChildItem -LiteralPath $dir).Name | Should -Be @('report.xlsx')
    }

    It 'resolves a relative path against the current PowerShell location' {
        $dir = Add-TestRoot
        Push-Location $dir
        try { Export-DuplicateReport -DuplicateSet $script:Sets -Path 'relative.xlsx' }
        finally { Pop-Location }

        Join-Path $dir 'relative.xlsx' | Should -Exist
    }
}

Describe 'Import-DuplicateReport' {
    BeforeAll {
        $script:RoundTrip = @(
            [pscustomobject] @{
                FileName = 'a & b.txt'; LastWriteTime = [datetime]::new(2024, 1, 2, 3, 4, 5, 678); UtcOffset = [TimeSpan]::FromHours(10)
                SizeBytes = 1234; MD5 = 'AAAA'; Count = 3; Folders = [string[]] @('C:\one', 'C:\two', 'D:\three')
            }
            [pscustomobject] @{
                FileName = 'z.txt'; LastWriteTime = [datetime]::new(2023, 6, 7, 8, 9, 10); UtcOffset = [TimeSpan]::new(-4, -30, 0)
                SizeBytes = 5; MD5 = 'BBBB'; Count = 2; Folders = [string[]] @('C:\x', 'C:\y')
            }
        )

        function Write-ExcelSavedWorkbook {
            # A workbook shaped like one Excel has re-saved: shared strings, a
            # renamed worksheet part, and cells without explicit types.
            param([string] $Path, [object[][]] $Rows)
            $strings = [System.Collections.Generic.List[string]]::new()
            $sheetRows = for ($r = 0; $r -lt $Rows.Count; $r++) {
                $cells = for ($c = 0; $c -lt $Rows[$r].Count; $c++) {
                    $ref = "$(ConvertTo-ColumnName ($c + 1))$($r + 1)"
                    $value = $Rows[$r][$c]
                    if ($null -eq $value) { continue }  # Excel leaves blank cells out
                    if ($value -is [string]) {
                        $strings.Add($value)
                        "<c r=`"$ref`" t=`"s`"><v>$($strings.Count - 1)</v></c>"
                    }
                    else { "<c r=`"$ref`"><v>$([System.Convert]::ToString($value, [System.Globalization.CultureInfo]::InvariantCulture))</v></c>" }
                }
                "<row r=`"$($r + 1)`">$(-join $cells)</row>"
            }
            $sst = -join ($strings | ForEach-Object { "<si><t>$([System.Security.SecurityElement]::Escape($_))</t></si>" })
            $parts = [ordered] @{
                '[Content_Types].xml'        = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
                'xl/workbook.xml'            = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Duplicates" sheetId="1" r:id="rId7"/></sheets></workbook>'
                'xl/_rels/workbook.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/data.xml"/></Relationships>'
                'xl/sharedStrings.xml'       = "<sst xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`">$sst</sst>"
                'xl/worksheets/data.xml'     = "<worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><sheetData>$(-join $sheetRows)</sheetData></worksheet>"
            }
            $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($name in $parts.Keys) {
                    $writer = [System.IO.StreamWriter]::new($zip.CreateEntry($name).Open())
                    try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
                }
            }
            finally { $zip.Dispose() }
        }
    }

    It 'validates a report made before the UTC Offset column, and adds the column when it rewrites it' {
        $root = Add-TestRoot
        $files = @(foreach ($folder in 'a', 'b', 'c') { Add-TestFile $root "$folder/x.txt" })
        $path = Join-Path $root 'old.xlsx'
        $header = [object[]] @('File Name', 'Last Modified', 'Size (bytes)', 'MD5', 'Copies', 'Location 1', 'Location 2', 'Location 3')
        $row = [object[]] (@('x.txt', $files[0].LastWriteTime.ToOADate(), [double] $files[0].Length, 'ABC', [double] 3) + @($files.DirectoryName))
        Write-ExcelSavedWorkbook -Path $path -Rows ([object[][]] @($header, $row))
        Remove-Item -LiteralPath $files[2].FullName

        $result = Update-DuplicateReport -Path $path

        $result.CopiesRemoved | Should -Be 1 -Because 'the other two copies are found by local time'
        (Read-Worksheet $path)[0][2] | Should -Be 'UTC Offset'
        (Import-DuplicateReport -Path $path).UtcOffset | Should -Be ($files[0].LastWriteTime - $files[0].LastWriteTimeUtc)
    }

    It 'writes the UTC offset <Text> and reads it back' -ForEach @(
        @{ Seconds = 19800; Text = '+05:30' }
        @{ Seconds = -16200; Text = '-04:30' }
        @{ Seconds = 0; Text = '+00:00' }
        @{ Seconds = 50400; Text = '+14:00' }
        @{ Seconds = 1172; Text = '+00:19:32' }   # local mean time, as some zones used before 1900
    ) {
        InModuleScope DuplicateFinder -Parameters $_ {
            ConvertTo-UtcOffsetText ([TimeSpan]::FromSeconds($Seconds)) | Should -BeExactly $Text
            ConvertFrom-UtcOffsetText $Text -Path 'report.xlsx' | Should -Be ([TimeSpan]::FromSeconds($Seconds))
        }
    }

    It 'rejects a UTC offset it cannot read' {
        InModuleScope DuplicateFinder {
            { ConvertFrom-UtcOffsetText '10:00' -Path 'report.xlsx' } | Should -Throw "*'10:00' is not a UTC offset*"
        }
    }

    It 'reads back what Export-DuplicateReport wrote' {
        $path = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet $script:RoundTrip -Path $path

        $read = @(Import-DuplicateReport -Path $path)

        $read.Count | Should -Be 2
        for ($i = 0; $i -lt 2; $i++) {
            foreach ($property in 'FileName', 'LastWriteTime', 'UtcOffset', 'SizeBytes', 'MD5', 'Count') {
                $read[$i].$property | Should -Be $script:RoundTrip[$i].$property -Because $property
            }
            $read[$i].Folders | Should -Be $script:RoundTrip[$i].Folders
        }
    }

    It 'reads a report after Excel has saved it (shared strings, renamed sheet part)' {
        $path = Join-Path (Add-TestRoot) 'excel.xlsx'
        Write-ExcelSavedWorkbook -Path $path -Rows @(
            , @('File Name', 'Last Modified', 'Size (bytes)', 'MD5', 'Copies', 'Location 1', 'Location 2')
            , @('x.txt', 45292.5, 10, 'CCCC', 2, 'C:\a', 'C:\b')
        )

        $read = @(Import-DuplicateReport -Path $path)

        $read.Count | Should -Be 1
        $read[0].FileName | Should -Be 'x.txt'
        $read[0].LastWriteTime | Should -Be ([datetime]::new(2024, 1, 1, 12, 0, 0))
        $read[0].SizeBytes | Should -Be 10
        $read[0].Folders | Should -Be @('C:\a', 'C:\b')
    }

    It 'reads a report that had the rules above the table' {
        $path = Join-Path (Add-TestRoot) 'older.xlsx'
        Write-ExcelSavedWorkbook -Path $path -Rows @(
            , @('Duplicate files')
            , @('Some rule.')
            , @($null)
            , @('File Name', 'Last Modified', 'Size (bytes)', 'MD5', 'Copies', 'Location 1', 'Location 2')
            , @('x.txt', 45292.5, 10, 'CCCC', 2, 'C:\a', 'C:\b')
        )
        @(Import-DuplicateReport -Path $path).FileName | Should -Be @('x.txt')
    }

    It 'rejects a workbook that is not a duplicates report' {
        $path = Join-Path (Add-TestRoot) 'other.xlsx'
        Write-ExcelSavedWorkbook -Path $path -Rows @(, @('Name', 'Amount'))
        { Import-DuplicateReport -Path $path } | Should -Throw '*is not a duplicates report*'
    }

    It 'rejects a file that is not a workbook' {
        $path = Join-Path (Add-TestRoot) 'notes.xlsx'
        Set-Content -LiteralPath $path -Value 'not a zip'
        { Import-DuplicateReport -Path $path } | Should -Throw '*is not an Excel workbook*'
    }

    It 'rejects a header row with a blank cell' {
        $path = Join-Path (Add-TestRoot) 'gap.xlsx'
        Write-ExcelSavedWorkbook -Path $path -Rows @(, @('File Name', $null, 'Size (bytes)', 'MD5', 'Copies'))
        { Import-DuplicateReport -Path $path } | Should -Throw '*is not a duplicates report*'
    }

    It 'reports a blank number cell as a clear error' {
        $path = Join-Path (Add-TestRoot) 'blank.xlsx'
        Write-ExcelSavedWorkbook -Path $path -Rows @(
            , @('File Name', 'Last Modified', 'Size (bytes)', 'MD5', 'Copies', 'Location 1', 'Location 2')
            , @('x.txt', $null, 10, 'CCCC', 2, 'C:\a', 'C:\b')
        )
        { Import-DuplicateReport -Path $path } | Should -Throw '*could not be read as a duplicates report*'
    }

    It 'rejects a report containing a DTD' {
        $path = Join-Path (Add-TestRoot) 'dtd.xlsx'
        Export-DuplicateReport -DuplicateSet $script:RoundTrip -Path $path
        $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.GetEntry('xl/workbook.xml')
            $reader = [System.IO.StreamReader]::new($entry.Open())
            try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $entry.Delete()
            $writer = [System.IO.StreamWriter]::new($zip.CreateEntry('xl/workbook.xml').Open())
            try { $writer.Write($xml.Replace('<workbook', '<!DOCTYPE workbook [<!ENTITY x "x">]><workbook')) } finally { $writer.Dispose() }
        }
        finally { $zip.Dispose() }

        { Import-DuplicateReport -Path $path } | Should -Throw '*DTD*'
    }

    It 'reads a report that another program has open' {
        $path = Join-Path (Add-TestRoot) 'open.xlsx'
        Export-DuplicateReport -DuplicateSet $script:RoundTrip -Path $path
        # Like Excel: open for writing, letting others read.
        $lock = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        try { @(Import-DuplicateReport -Path $path).Count | Should -Be 2 }
        finally { $lock.Dispose() }
    }

    It 'reports a missing report' {
        { Import-DuplicateReport -Path (Join-Path (Add-TestRoot) 'missing.xlsx') } | Should -Throw '*was not found*'
    }
}

Describe 'Update-DuplicateReport' {
    BeforeAll {
        function Export-ScannedReport {
            # Scans $Root into a report next to it and returns the report path.
            param([string] $Root)
            $path = "$Root.xlsx"
            $sets = @(Find-DuplicateFile -File @(Get-FileInventory -Path $Root))
            Export-DuplicateReport -DuplicateSet $sets -Path $path
            $path
        }
    }

    It 'keeps every copy when validating in another time zone than the scan' {
        if ($script:OnWindows) { Set-ItResult -Skipped -Because 'the TZ variable sets the time zone only on Linux and macOS'; return }
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b') { $null = Add-TestFile $root "$folder/x.txt" }
        $previous = $env:TZ
        try {
            $env:TZ = 'Asia/Kolkata'
            [System.TimeZoneInfo]::ClearCachedData()
            $report = Export-ScannedReport $root
            $env:TZ = 'America/New_York'
            [System.TimeZoneInfo]::ClearCachedData()

            (Update-DuplicateReport -Path $report -WhatIf).CopiesRemoved | Should -Be 0
        }
        finally {
            $env:TZ = $previous
            [System.TimeZoneInfo]::ClearCachedData()
        }
    }

    It 'leaves the report untouched when every copy still exists' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b') { $null = Add-TestFile $root "$folder/x.txt" }
        $report = Export-ScannedReport $root
        $before = (Get-Item -LiteralPath $report).LastWriteTimeUtc

        $result = Update-DuplicateReport -Path $report

        $result.Saved | Should -BeFalse
        $result.CopiesChecked | Should -Be 2
        $result.RowsRemaining | Should -Be 1
        (Get-Item -LiteralPath $report).LastWriteTimeUtc | Should -Be $before
    }

    It 'removes a copy that no longer exists' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b', 'c') { $null = Add-TestFile $root "$folder/x.txt" }
        $report = Export-ScannedReport $root
        Remove-Item -LiteralPath (Join-Path $root 'b/x.txt')

        $result = Update-DuplicateReport -Path $report

        $result.CopiesRemoved | Should -Be 1
        $result.Saved | Should -BeTrue
        $read = @(Import-DuplicateReport -Path $report)
        $read.Count | Should -Be 1
        $read[0].Count | Should -Be 2
        $read[0].Folders | Should -Be @((Join-Path $root 'a'), (Join-Path $root 'c'))
    }

    It 'removes a row left with fewer than two copies' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b') {
            $null = Add-TestFile $root "$folder/gone.txt" -Content 'gone'
            $null = Add-TestFile $root "$folder/kept.txt" -Content 'kept'
        }
        $report = Export-ScannedReport $root
        Remove-Item -LiteralPath (Join-Path $root 'a/gone.txt')

        $result = Update-DuplicateReport -Path $report

        $result.RowsRemoved | Should -Be 1
        @(Import-DuplicateReport -Path $report).FileName | Should -Be @('kept.txt')
    }

    It 'removes a copy that was changed since the scan' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b', 'c') { $null = Add-TestFile $root "$folder/x.txt" }
        $report = Export-ScannedReport $root
        $null = Add-TestFile $root 'c/x.txt' -Content 'edited' -SavedUtc $script:Saved.AddHours(1)

        $result = Update-DuplicateReport -Path $report

        $result.CopiesRemoved | Should -Be 1
        (Import-DuplicateReport -Path $report).Folders | Should -Be @((Join-Path $root 'a'), (Join-Path $root 'b'))
    }

    It 'finds a copy whose name differs only by case' {
        $root = Add-TestRoot
        $null = Add-TestFile $root 'a/Photo.JPG'
        $null = Add-TestFile $root 'b/photo.jpg'
        $report = Export-ScannedReport $root

        $result = Update-DuplicateReport -Path $report

        $result.CopiesRemoved | Should -Be 0
        $result.RowsRemaining | Should -Be 1
    }

    It 'keeps copies on a drive or share that cannot be reached' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b') { $null = Add-TestFile $root "$folder/x.txt" }
        $report = Export-ScannedReport $root
        Mock -ModuleName DuplicateFinder Test-DuplicateCopy { 'Unavailable' }

        $result = Update-DuplicateReport -Path $report -WarningVariable warnings -WarningAction SilentlyContinue

        $result.CopiesUnavailable | Should -Be 2
        $result.Saved | Should -BeFalse
        @($warnings).Count | Should -Be 2
        $result.RowsRemaining | Should -Be 1
    }

    It 'treats a copy on a missing drive letter as unreachable' {
        if (-not $script:OnWindows) { Set-ItResult -Skipped -Because 'drive letters are Windows-only'; return }
        $used = [System.IO.DriveInfo]::GetDrives().Name | ForEach-Object { $_.Substring(0, 1) }
        $free = [char[]] ([int][char] 'Q'..[int][char] 'Z') | Where-Object { "$_" -notin $used } | Select-Object -First 1
        if (-not $free) { Set-ItResult -Skipped -Because 'no free drive letter'; return }

        InModuleScope DuplicateFinder -Parameters @{ Folder = "${free}:\photos" } {
            Test-DuplicateCopy -Folder $Folder -FileName 'x.txt' -SizeBytes 1 -LastWriteTime ([datetime]::Now) |
                Should -Be 'Unavailable'
        }
    }

    It 'keeps a copy whose size and saved date cannot be read' {
        # As for a file named NUL on Windows, which the system treats as a device.
        $folder = Join-Path (Add-TestRoot) 'a'
        InModuleScope DuplicateFinder -Parameters @{ Folder = $folder; Saved = $script:Saved } {
            Mock Find-FileByNameKey { [pscustomobject] @{ Length = [long] 1; LastWriteTime = $null } }
            Test-DuplicateCopy -Folder $Folder -FileName 'NUL' -SizeBytes 1 -LastWriteTime $Saved | Should -Be 'Unavailable'
        }
    }

    It 'checks each drive or share only once' {
        InModuleScope DuplicateFinder -Parameters @{ Folder = (Add-TestRoot) } {
            $cache = [System.Collections.Generic.Dictionary[string, bool]]::new()
            Test-PathRootReachable -Folder (Join-Path $Folder 'a') -Cache $cache | Should -BeTrue
            Test-PathRootReachable -Folder (Join-Path $Folder 'b') -Cache $cache | Should -BeTrue
            $cache.Count | Should -Be 1
        }
    }

    It 'changes nothing with -WhatIf' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b', 'c') { $null = Add-TestFile $root "$folder/x.txt" }
        $report = Export-ScannedReport $root
        Remove-Item -LiteralPath (Join-Path $root 'a/x.txt')

        $result = Update-DuplicateReport -Path $report -WhatIf

        $result.CopiesRemoved | Should -Be 1
        $result.Saved | Should -BeFalse
        $rows = @(Import-DuplicateReport -Path $report)
        $rows.Count | Should -Be 1
        $rows[0].Count | Should -Be 3 -Because 'the removed copy is still listed'
    }
}

Describe 'Find-DuplicateFolder' {
    BeforeAll {
        function Get-FolderScan {
            # Files and folder records for a tree, as the script collects them.
            param([string] $Root)
            $info = [System.Collections.Generic.List[object]]::new()
            $files = @(Get-FileInventory -Path $Root -FolderInfo $info)
            [pscustomobject] @{ Files = $files; Folders = $info.ToArray() }
        }

        function Add-PhotoFolder {
            # A small tree: two files, one in a sub folder, and an empty sub folder.
            param([string] $Root, [string] $Folder, [string] $Content = 'photo')
            $null = Add-TestFile $Root "$Folder/a.jpg" -Content "a $Content"
            $null = Add-TestFile $Root "$Folder/sub/b.jpg" -Content "b $Content"
            $null = New-Item -ItemType Directory -Force -Path (Join-Path $Root "$Folder/empty")
        }

        function Find-InTree {
            param([string] $Root, [hashtable] $Options = @{})
            $scan = Get-FolderScan $Root
            # ", @(...)": returned bare, a single result would be unrolled into the set
            # object itself, whose Count property is its number of copies.
            , @(Find-DuplicateFolder -File $scan.Files -Folder $scan.Folders @Options)
        }
    }

    It 'finds folders with the same name and the same contents' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/Photos'

        $result = Find-InTree $root

        $result.Count | Should -Be 1
        $result[0].FolderName | Should -Be 'Photos'
        $result[0].FileCount | Should -Be 2
        $result[0].FolderCount | Should -Be 2
        $result[0].SizeBytes | Should -Be ('a photo'.Length + 'b photo'.Length)
        $result[0].Count | Should -Be 2
        $result[0].Folders | Should -Be @((Join-Path $root 'one/Photos'), (Join-Path $root 'two/Photos') | ForEach-Object { [System.IO.Path]::GetFullPath($_) })
    }

    It 'matches folder names that differ only by case' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/photos'
        (Find-InTree $root).Count | Should -Be 1
    }

    It 'ignores folders whose <Case>, but still finds their identical sub folders' -ForEach @(
        @{ Case = 'names differ'; Second = 'two/Pictures'; Change = {} }
        @{ Case = 'file contents differ at the same size and date'; Second = 'two/Photos'; Change = {
                param($r) $null = Add-TestFile $r 'two/Photos/a.jpg' -Content 'a PHOTO' } }
        @{ Case = 'file saved dates differ'; Second = 'two/Photos'; Change = {
                param($r) $null = Add-TestFile $r 'two/Photos/a.jpg' -Content 'a photo' -SavedUtc $script:Saved.AddMinutes(1) } }
        @{ Case = 'files differ (an extra file)'; Second = 'two/Photos'; Change = {
                param($r) $null = Add-TestFile $r 'two/Photos/extra.txt' } }
        @{ Case = 'sub folders differ (an extra empty folder)'; Second = 'two/Photos'; Change = {
                param($r) $null = New-Item -ItemType Directory -Path (Join-Path $r 'two/Photos/more') } }
        @{ Case = 'file names differ'; Second = 'two/Photos'; Change = {
                param($r) Rename-Item -LiteralPath (Join-Path $r 'two/Photos/a.jpg') -NewName 'c.jpg' } }
    ) {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root $Second
        & $Change $root

        # The Photos folders differ, but their untouched "sub" folders are still duplicates.
        (Find-InTree $root).FolderName | Should -Be @('sub')
    }

    It 'reports only the top-most duplicate folders' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos/2024'
        Add-PhotoFolder $root 'two/Photos/2024'

        $result = Find-InTree $root

        $result.Count | Should -Be 1
        $result[0].FolderName | Should -Be 'Photos'
    }

    It 'keeps a nested set that also has a copy somewhere else' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos/2024'
        Add-PhotoFolder $root 'two/Photos/2024'
        Add-PhotoFolder $root 'three/2024'

        $result = Find-InTree $root

        $result.FolderName | Should -Be @('2024', 'Photos')
        $result[0].Count | Should -Be 3
    }

    It 'does not report folders that contain no files' {
        $root = Add-TestRoot
        foreach ($folder in 'one/Empty/inner', 'two/Empty/inner') { $null = New-Item -ItemType Directory -Force -Path (Join-Path $root $folder) }
        (Find-InTree $root).Count | Should -Be 0
    }

    It 'does not report a folder that has an unreadable sub folder' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/Photos'
        $scan = Get-FolderScan $root
        ($scan.Folders | Where-Object { $_.Path -eq [System.IO.Path]::GetFullPath((Join-Path $root 'two/Photos/sub')) }).Readable = $false

        @(Find-DuplicateFolder -File $scan.Files -Folder $scan.Folders).Count | Should -Be 0
    }

    It 'does not report a folder that holds the excluded report' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/Photos'
        $report = Add-TestFile $root 'one/Photos/dupes.xlsx'
        $info = [System.Collections.Generic.List[object]]::new()
        $files = @(Get-FileInventory -Path $root -ExcludeFile $report.FullName -FolderInfo $info)

        # Without the report the two trees look identical, but one really holds an extra file.
        (Find-DuplicateFolder -File $files -Folder $info.ToArray()).FolderName | Should -Be @('sub')
    }

    It 'does not report a folder that holds a folder link' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/Photos'
        $link = Join-Path $root 'one/Photos/link'
        try { $null = New-Item -ItemType SymbolicLink -Path $link -Target (Join-Path $root 'two') -ErrorAction Stop }
        catch { Set-ItResult -Skipped -Because "symbolic links cannot be created here: $_"; return }

        try { (Find-InTree $root).FolderName | Should -Be @('sub') }
        finally { [System.IO.Directory]::Delete($link) }
    }

    It 'does not report a folder whose file has gone since the scan' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/Photos'
        $scan = Get-FolderScan $root
        $gone = Join-Path $root 'two/Photos/a.jpg'
        # A fresh FileInfo reads nothing up front, as the scan's do on Linux and macOS.
        $files = @(foreach ($f in $scan.Files) { if ($f.FullName -eq $gone) { [System.IO.FileInfo] $gone } else { $f } })
        Remove-Item -LiteralPath $gone

        $result = @(Find-DuplicateFolder -File $files -Folder $scan.Folders -WarningVariable warnings -WarningAction SilentlyContinue)

        "$warnings" | Should -BeLike "*$gone*"
        $result.FolderName | Should -Be @('sub') -Because 'the two sub folders are still identical'
    }

    It 'does not read files again that the file scan already hashed' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/Photos'
        $scan = Get-FolderScan $root
        $cache = [System.Collections.Generic.Dictionary[string, string]]::new()
        $null = Find-DuplicateFile -File $scan.Files -Md5Cache $cache
        Mock -ModuleName DuplicateFinder Get-FileMd5 { throw 'should not be called' }

        $result = @(Find-DuplicateFolder -File $scan.Files -Folder $scan.Folders -Md5Cache $cache)

        Should -Invoke -ModuleName DuplicateFinder Get-FileMd5 -Times 0 -Exactly
        $result.Count | Should -Be 1
    }

    It 'does not report folders holding online-only files with -SkipCloudOnly' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/Photos'
        Add-PhotoFolder $root 'cloud/Photos'
        Mock -ModuleName DuplicateFinder Test-CloudOnlyFile { "$($File.FullName)" -like '*cloud*' }

        # Called directly: -WarningVariable creates the variable in the caller's scope.
        $scan = Get-FolderScan $root
        $result = @(Find-DuplicateFolder -File $scan.Files -Folder $scan.Folders -SkipCloudOnly -WarningVariable warnings -WarningAction SilentlyContinue)

        $result.Count | Should -Be 1
        $result[0].Folders | Should -Not -BeLike '*cloud*'
        "$($warnings[0])" | Should -BeLike '2 online-only*folders*'
    }

    It 'finds the same folders hashing several files at a time' {
        $root = Add-TestRoot
        Add-PhotoFolder $root 'one/Photos'
        Add-PhotoFolder $root 'two/Photos'
        Add-PhotoFolder $root 'one/Other' -Content 'other'
        Add-PhotoFolder $root 'two/Other' -Content 'other'

        $result = Find-InTree $root @{ ThrottleLimit = 4 }

        $result.FolderName | Should -Be @('Other', 'Photos')
    }
}

Describe 'Unicode names' {
    BeforeAll {
        # Built from code points so this file stays ASCII (Windows PowerShell 5.1 reads
        # BOM-less scripts in the ANSI code page).
        function ConvertFrom-CodePoint { param([int[]] $CodePoint) -join ($CodePoint | ForEach-Object { [char]::ConvertFromUtf32($_) }) }
        $script:Names = @{
            Japanese = ConvertFrom-CodePoint 0x65E5, 0x672C, 0x8A9E
            Arabic   = ConvertFrom-CodePoint 0x0645, 0x0644, 0x0641
            Cyrillic = ConvertFrom-CodePoint 0x0444, 0x0430, 0x0439, 0x043B
            Emoji    = ConvertFrom-CodePoint 0x1F600, 0x1F4F7
        }
        $script:Composed   = ConvertFrom-CodePoint 0x63, 0x61, 0x66, 0xE9           # cafe with e-acute as one character
        $script:Decomposed = ConvertFrom-CodePoint 0x63, 0x61, 0x66, 0x65, 0x301    # e followed by a combining accent
    }

    It 'finds, saves, reads back and validates duplicates named in <_>' -ForEach @('Japanese', 'Arabic', 'Cyrillic', 'Emoji') {
        $name = $script:Names[$_]
        $root = Add-TestRoot
        $null = Add-TestFile $root "$name/$name.txt"
        $null = Add-TestFile $root "copy/$name.txt"
        $report = Join-Path (Add-TestRoot) "$name.xlsx"

        $found = @(Find-DuplicateFile -File @(Get-FileInventory -Path $root))
        Export-DuplicateReport -DuplicateSet $found -Path $report
        $read = @(Import-DuplicateReport -Path $report)
        $result = Update-DuplicateReport -Path $report

        $found.Count | Should -Be 1
        $read[0].FileName | Should -BeExactly "$name.txt"
        $read[0].Folders | Should -Contain ([System.IO.Path]::GetFullPath((Join-Path $root $name)))
        $result.CopiesRemoved | Should -Be 0
    }

    It 'matches names stored in different Unicode forms' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root "a/$($script:Composed).txt"
            Add-TestFile $root "b/$($script:Decomposed).txt"
        )
        if ($files[0].Name -eq $files[1].Name) { Set-ItResult -Skipped -Because 'this file system normalises names itself'; return }

        @(Find-DuplicateFile -File $files).Count | Should -Be 1
    }

    It 'keeps a copy whose name is stored in another Unicode form when validating' {
        $root = Add-TestRoot
        $null = Add-TestFile $root "a/$($script:Composed).txt"
        $second = Add-TestFile $root "b/$($script:Decomposed).txt"
        if ($second.Name -eq "$($script:Composed).txt") { Set-ItResult -Skipped -Because 'this file system normalises names itself'; return }
        $report = "$root.xlsx"
        Export-DuplicateReport -DuplicateSet @(Find-DuplicateFile -File @(Get-FileInventory -Path $root)) -Path $report

        (Update-DuplicateReport -Path $report).CopiesRemoved | Should -Be 0
    }

    It 'orders names the same way on Windows PowerShell 5.1 and PowerShell 7' {
        # In UTF-16 order an emoji comes before U+FF21 (full-width A); in code-point order
        # it comes after. OrdinalIgnoreCase differs between .NET versions here.
        $emoji = [char]::ConvertFromUtf32(0x1F600)
        $fullWidthA = [string] [char] 0xFF21
        InModuleScope DuplicateFinder -Parameters @{ Emoji = $emoji; FullWidthA = $fullWidthA } {
            # Assigned first, as the module's callers do: the function returns its array as one object.
            $sorted = Get-SortedFolder -Path @("/x/$FullWidthA", "/x/$Emoji")
            $sorted | Should -Be @("/x/$Emoji", "/x/$FullWidthA")
            $sorted = Get-SortedFolder -Path @('/x/photos', '/x/Photos')
            $sorted | Should -Be @('/x/Photos', '/x/photos') -Because 'names differing only in case keep a fixed order'
        }
    }

    It 'finds duplicate folders with Unicode names' {
        $root = Add-TestRoot
        foreach ($parent in 'one', 'two') {
            $null = Add-TestFile $root "$parent/$($script:Names.Japanese)/$($script:Names.Emoji).jpg"
        }
        $info = [System.Collections.Generic.List[object]]::new()
        $files = @(Get-FileInventory -Path $root -FolderInfo $info)

        $result = @(Find-DuplicateFolder -File $files -Folder $info.ToArray())

        $result.Count | Should -Be 1
        $result[0].FolderName | Should -BeExactly $script:Names.Japanese
    }

    It 'handles Unicode names end to end through the script' {
        $root = Add-TestRoot
        foreach ($parent in 'one', 'two') { $null = Add-TestFile $root "$parent/$($script:Names.Arabic)/$($script:Names.Cyrillic).txt" }
        $out = Join-Path (Add-TestRoot) "$($script:Names.Emoji).xlsx"

        $result = @(& $script:ScriptPath -Path $root -OutputFile $out -IncludeFolders -PassThru 6>$null)

        $out | Should -Exist
        @($result | Where-Object { $_.PSObject.Properties['FileName'] }).FileName | Should -Be "$($script:Names.Cyrillic).txt"
        @(Import-DuplicateFolderReport -Path $out).FolderName | Should -Be $script:Names.Arabic
    }
}

Describe 'Duplicate folders in the report' {
    BeforeAll {
        $script:FolderSets = @(
            [pscustomobject] @{
                FolderName = 'Photos'; FileCount = 12; FolderCount = 3; SizeBytes = 123456; Count = 2
                Folders = [string[]] @('C:\one\Photos', 'D:\two\Photos')
            }
        )
    }

    It 'writes and reads back a Duplicate Folders sheet' {
        $path = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet @() -FolderSet $script:FolderSets -Path $path

        $read = @(Import-DuplicateFolderReport -Path $path)

        $read.Count | Should -Be 1
        foreach ($property in 'FolderName', 'FileCount', 'FolderCount', 'SizeBytes', 'Count') {
            $read[0].$property | Should -Be $script:FolderSets[0].$property -Because $property
        }
        $read[0].Folders | Should -Be $script:FolderSets[0].Folders
        Get-SheetName $path | Should -Be @('Duplicates', 'Duplicate Folders', 'Rules')
        $rules = @(Read-Worksheet $path -Part 'xl/worksheets/sheet3.xml' -AllRows | ForEach-Object { $_[0] })
        $rules | Should -Contain "Sheet 'Duplicate Folders': duplicate folders"
    }

    It 'writes no folder sheet unless folder sets are given' {
        $path = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet @() -Path $path

        @(Import-DuplicateFolderReport -Path $path).Count | Should -Be 0
        Get-SheetName $path | Should -Not -Contain 'Duplicate Folders'
    }

    It 'writes an empty folder sheet when no duplicate folders were found' {
        $path = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet @() -FolderSet @() -Path $path

        Get-SheetName $path | Should -Contain 'Duplicate Folders'
        @(Import-DuplicateFolderReport -Path $path).Count | Should -Be 0
    }
}

Describe 'Validating duplicate folders' {
    BeforeAll {
        function Export-FolderReport {
            # Scans $Root (files and folders) into a report next to it and returns its path.
            param([string] $Root)
            $info = [System.Collections.Generic.List[object]]::new()
            $files = @(Get-FileInventory -Path $Root -FolderInfo $info)
            $path = "$Root.xlsx"
            Export-DuplicateReport -Path $path -DuplicateSet @(Find-DuplicateFile -File $files) `
                -FolderSet @(Find-DuplicateFolder -File $files -Folder $info.ToArray())
            $path
        }

        function Add-CopiedTree {
            param([string] $Root)
            foreach ($folder in 'one', 'two', 'three') {
                $null = Add-TestFile $Root "$folder/Photos/a.jpg" -Content 'a'
                $null = Add-TestFile $Root "$folder/Photos/sub/b.jpg" -Content 'b'
            }
        }
    }

    It 'removes a folder copy that no longer exists' {
        $root = Add-TestRoot
        Add-CopiedTree $root
        $report = Export-FolderReport $root
        Remove-Item -LiteralPath (Join-Path $root 'two/Photos') -Recurse

        $result = Update-DuplicateReport -Path $report

        $result.FolderCopiesRemoved | Should -Be 1
        $read = @(Import-DuplicateFolderReport -Path $report)
        $read[0].Count | Should -Be 2
        $read[0].Folders | Should -Be @((Join-Path $root 'one/Photos'), (Join-Path $root 'three/Photos') | ForEach-Object { [System.IO.Path]::GetFullPath($_) })
    }

    It 'removes a folder copy whose contents changed' {
        $root = Add-TestRoot
        Add-CopiedTree $root
        $report = Export-FolderReport $root
        $null = Add-TestFile $root 'three/Photos/sub/new.jpg'

        $result = Update-DuplicateReport -Path $report

        $result.FolderCopiesRemoved | Should -Be 1
        $result.FolderRowsRemaining | Should -Be 1
    }

    It 'keeps an empty folder sheet when every folder row is removed' {
        $root = Add-TestRoot
        Add-CopiedTree $root
        $report = Export-FolderReport $root
        foreach ($folder in 'one', 'two') { Remove-Item -LiteralPath (Join-Path $root "$folder/Photos") -Recurse }

        $result = Update-DuplicateReport -Path $report

        $result.FolderRowsRemoved | Should -Be 1
        $result.DuplicateFolderSet.Count | Should -Be 0
        Get-SheetName $report | Should -Contain 'Duplicate Folders'
    }

    It 'keeps folder copies that cannot be reached' {
        $root = Add-TestRoot
        Add-CopiedTree $root
        $report = Export-FolderReport $root
        Mock -ModuleName DuplicateFinder Test-DuplicateFolderCopy { 'Unavailable' }

        $result = Update-DuplicateReport -Path $report -WarningAction SilentlyContinue

        $result.FolderCopiesUnavailable | Should -Be 3
        $result.FolderRowsRemaining | Should -Be 1
    }

    It 'leaves reports without a folder sheet without one' {
        $root = Add-TestRoot
        Add-CopiedTree $root
        $report = "$root.xlsx"
        Export-DuplicateReport -DuplicateSet @(Find-DuplicateFile -File @(Get-FileInventory -Path $root)) -Path $report
        Remove-Item -LiteralPath (Join-Path $root 'one/Photos/a.jpg')

        $result = Update-DuplicateReport -Path $report

        $result.Saved | Should -BeTrue
        $result.DuplicateFolderSet | Should -BeNullOrEmpty
        @(Import-DuplicateFolderReport -Path $report).Count | Should -Be 0
        Get-SheetName $report | Should -Be @('Duplicates', 'Rules')
    }
}

Describe 'Find-Duplicates.ps1' {
    BeforeAll {
        $script:Root = Add-TestRoot
        $null = Add-TestFile $script:Root 'data/2023/invoice.pdf' -Content 'invoice'
        $null = Add-TestFile $script:Root 'data/backup/invoice.pdf' -Content 'invoice'
        $null = Add-TestFile $script:Root 'data/old/copy/invoice.pdf' -Content 'invoice'
        $null = Add-TestFile $script:Root 'data/notes.txt' -Content 'one'
        $null = Add-TestFile $script:Root 'data/other/notes.txt' -Content 'two'
    }

    It 'saves to duplicates.xlsx in the current folder by default' {
        $workDir = Add-TestRoot
        Push-Location $workDir
        try {
            $result = @(& $script:ScriptPath -Path (Join-Path $script:Root 'data') -PassThru 6>$null)
        }
        finally { Pop-Location }

        $report = Join-Path $workDir 'duplicates.xlsx'
        $report | Should -Exist
        $result.Count | Should -Be 1
        $result[0].Count | Should -Be 3

        $rows = @(Read-Worksheet $report)
        $rows.Count | Should -Be 2
        $rows[1][0] | Should -Be 'invoice.pdf'
        $expected = 'data/2023', 'data/backup', 'data/old/copy' |
            ForEach-Object { [System.IO.Path]::GetFullPath((Join-Path $script:Root $_)) }
        $rows[1][6..8] | Should -Be $expected -Because 'the full folder path of every copy is recorded'
    }

    It 'uses the given output file name and adds .xlsx when missing' {
        $workDir = Add-TestRoot
        $null = & $script:ScriptPath -Path $script:Root -OutputFile (Join-Path $workDir 'my-report') 6>$null
        Join-Path $workDir 'my-report.xlsx' | Should -Exist
    }

    It 'reports long folder names when given a short path' {
        if (-not $script:OnWindows) { Set-ItResult -Skipped -Because 'short (8.3) names are Windows-only'; return }
        $root = (Get-Item -LiteralPath (Add-TestRoot)).FullName
        $null = Add-TestFile $root 'a long folder name/x.txt'
        $null = Add-TestFile $root 'another long name/x.txt'
        $short = (New-Object -ComObject Scripting.FileSystemObject).GetFolder($root).ShortPath
        if ($short -eq $root) { Set-ItResult -Skipped -Because 'short names are disabled on this volume'; return }

        $result = @(& $script:ScriptPath -Path $short -OutputFile (Join-Path (Add-TestRoot) 'short.xlsx') -PassThru 6>$null)

        $result[0].Folders | ForEach-Object { $_ | Should -Not -BeLike '*~*' }
    }

    It 'scans and validates a network share given as a UNC path' {
        if (-not $script:OnWindows) { Set-ItResult -Skipped -Because 'UNC paths are Windows-only'; return }
        $root = Add-TestRoot
        $null = Add-TestFile $root 'a/x.txt'
        $null = Add-TestFile $root 'b/x.txt'
        # Reach the local test folder through the administrative share, e.g. \\localhost\C$\...
        $unc = '\\localhost\' + $root.Substring(0, 1) + '$' + $root.Substring(2)
        if (-not (Test-Path -LiteralPath $unc)) { Set-ItResult -Skipped -Because 'the administrative share is not available'; return }

        $out = Join-Path (Add-TestRoot) 'unc.xlsx'
        $result = @(& $script:ScriptPath -Path $unc -OutputFile $out -PassThru 6>$null)

        $result.Count | Should -Be 1
        $result[0].Folders | Should -Be @("$unc\a", "$unc\b")
        $out | Should -Exist

        # Validation over the network path too.
        Remove-Item -LiteralPath (Join-Path $root 'a/x.txt')
        $null = & $script:ScriptPath -Validate $out 6>$null
        @(Import-DuplicateReport -Path $out).Count | Should -Be 0
    }

    It 'hashes several files at a time with -ThrottleLimit' {
        $out = Join-Path (Add-TestRoot) 'parallel.xlsx'
        $result = @(& $script:ScriptPath -Path (Join-Path $script:Root 'data') -OutputFile $out -ThrottleLimit 4 -PassThru 6>$null)

        $result.Count | Should -Be 1
        $result[0].Count | Should -Be 3
    }

    It 'prints every folder with -Verbose' {
        $out = Join-Path (Add-TestRoot) 'verbose.xlsx'
        $verbose = & $script:ScriptPath -Path (Join-Path $script:Root 'data') -OutputFile $out -Verbose 4>&1 6>$null |
            Where-Object { $_ -is [System.Management.Automation.VerboseRecord] }

        $verbose.Message | Should -Contain "Scanning $(Join-Path $script:Root 'data/backup' | ForEach-Object { [System.IO.Path]::GetFullPath($_) })"
    }

    It 'leaves files of 0 bytes out with -IgnoreEmptyFiles' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b') { $null = Add-TestFile $root "$folder/empty.txt" -Content '' }
        $out = Join-Path (Add-TestRoot) 'empty.xlsx'

        @(& $script:ScriptPath -Path $root -OutputFile $out -PassThru 6>$null).Count | Should -Be 1
        @(& $script:ScriptPath -Path $root -OutputFile $out -IgnoreEmptyFiles -PassThru 6>$null).Count | Should -Be 0
    }

    It 'saves no report with -WhatIf' {
        $out = Join-Path (Add-TestRoot) 'whatif.xlsx'
        $result = @(& $script:ScriptPath -Path $script:Root -OutputFile $out -WhatIf -PassThru 6>$null)

        $result.Count | Should -Be 1
        $out | Should -Not -Exist
    }

    It 'validates duplicates.xlsx in the current folder with -Validate' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b', 'c') { $null = Add-TestFile $root "$folder/x.txt" }
        $workDir = Add-TestRoot
        Push-Location $workDir
        try {
            $null = & $script:ScriptPath -Path $root 6>$null
            Remove-Item -LiteralPath (Join-Path $root 'b/x.txt')
            $remaining = @(& $script:ScriptPath -Validate -PassThru 6>$null)
        }
        finally { Pop-Location }

        $remaining.Count | Should -Be 1
        $remaining[0].Folders | Should -Be @((Join-Path $root 'a'), (Join-Path $root 'c'))
        $rows = @(Import-DuplicateReport -Path (Join-Path $workDir 'duplicates.xlsx'))
        $rows.Count | Should -Be 1
        $rows[0].Count | Should -Be 2
    }

    It 'takes the report to validate as its first argument' {
        $root = Add-TestRoot
        foreach ($folder in 'a', 'b') { $null = Add-TestFile $root "$folder/x.txt" }
        $report = Join-Path (Add-TestRoot) 'named'
        $null = & $script:ScriptPath -Path $root -OutputFile $report 6>$null
        Remove-Item -LiteralPath (Join-Path $root 'a/x.txt')

        $remaining = @(& $script:ScriptPath -Validate $report -PassThru 6>$null)

        $remaining.Count | Should -Be 0
        @(Import-DuplicateReport -Path "$report.xlsx").Count | Should -Be 0
    }

    It 'adds duplicate folders to the report with -IncludeFolders' {
        $root = Add-TestRoot
        foreach ($folder in 'one', 'two') {
            $null = Add-TestFile $root "$folder/Photos/a.jpg" -Content 'a'
            $null = Add-TestFile $root "$folder/Photos/b.jpg" -Content 'b'
        }
        $out = Join-Path (Add-TestRoot) 'folders.xlsx'

        $result = @(& $script:ScriptPath -Path $root -OutputFile $out -IncludeFolders -PassThru 6>$null)

        @($result | Where-Object { $_.PSObject.Properties['FolderName'] }).Count | Should -Be 1
        @($result | Where-Object { $_.PSObject.Properties['FileName'] }).Count | Should -Be 2
        @(Import-DuplicateFolderReport -Path $out)[0].FolderName | Should -Be 'Photos'
    }

    It 're-checks duplicate folders with -Validate' {
        $root = Add-TestRoot
        foreach ($folder in 'one', 'two', 'three') { $null = Add-TestFile $root "$folder/Photos/a.jpg" -Content 'a' }
        $out = Join-Path (Add-TestRoot) 'folders.xlsx'
        $null = & $script:ScriptPath -Path $root -OutputFile $out -IncludeFolders 6>$null
        Remove-Item -LiteralPath (Join-Path $root 'one/Photos') -Recurse

        $null = & $script:ScriptPath -Validate $out 6>$null

        @(Import-DuplicateFolderReport -Path $out)[0].Count | Should -Be 2
    }

    It 'does not scan its own report when it is saved inside the scanned folder' {
        $root = Add-TestRoot
        $null = Add-TestFile $root 'a/x.txt'
        $null = Add-TestFile $root 'b/x.txt'
        $report = Join-Path $root 'duplicates.xlsx'

        $null = & $script:ScriptPath -Path $root -OutputFile $report 6>$null
        $second = @(& $script:ScriptPath -Path $root -OutputFile $report -PassThru 6>$null)

        $second.Count | Should -Be 1
        $second[0].FileName | Should -Be 'x.txt'
    }
}

Describe 'Edge cases' {
    BeforeAll {
        function ConvertTo-RawPath {
            # On Windows, the \\?\ form of a full path: only that form can create names Windows
            # otherwise reserves or trims, and paths longer than 260 characters on every .NET.
            param([string] $Path)
            if ($script:OnWindows) { "\\?\$Path" } else { $Path }
        }

        function Add-RawFile {
            # Creates a file through its raw path; returns its plain full path.
            param([string] $Folder, [string] $Name, [string] $Content = 'same content')
            $null = [System.IO.Directory]::CreateDirectory((ConvertTo-RawPath $Folder))
            $path = [System.IO.Path]::Combine($Folder, $Name)
            [System.IO.File]::WriteAllText((ConvertTo-RawPath $path), $Content)
            [System.IO.File]::SetLastWriteTimeUtc((ConvertTo-RawPath $path), $script:Saved)
            $path
        }

        function Test-RunningAsRoot {
            (-not $script:OnWindows) -and ((id -u) -eq '0')
        }

        function Save-WithLibreOffice {
            # Opens a report in LibreOffice Calc and saves it again as .xlsx; returns the new
            # file's path, or $null when LibreOffice is not installed or cannot convert.
            param([string] $Path, [string] $OutFolder)
            $soffice = Get-Command soffice -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $soffice) { return $null }
            $userProfile = [System.Uri]::new((Join-Path $OutFolder 'profile')).AbsoluteUri
            $null = & $soffice.Source "-env:UserInstallation=$userProfile" --headless --norestore `
                --convert-to 'xlsx:Calc MS Excel 2007 XML' --outdir $OutFolder $Path 2>&1
            $saved = Join-Path $OutFolder ([System.IO.Path]::GetFileName($Path))
            if (Test-Path -LiteralPath $saved) { $saved } else { $null }
        }
    }

    It 'finds duplicates in folders whose full path is longer than 260 characters' {
        $root = Add-TestRoot
        $deep = [string]::Join([System.IO.Path]::DirectorySeparatorChar, @(1..5 | ForEach-Object { 'd' * 50 }))
        foreach ($copy in 'a', 'b') { $null = Add-RawFile ([System.IO.Path]::Combine($root, $copy, $deep)) 'x.txt' }

        try {
            $files = @(Get-FileInventory -Path $root -WarningVariable scanWarnings -WarningAction SilentlyContinue)
            $result = @(Find-DuplicateFile -File $files -WarningVariable hashWarnings -WarningAction SilentlyContinue)
            if ($result.Count -eq 0 -and $PSVersionTable.PSEdition -eq 'Desktop') {
                # Windows PowerShell 5.1 may be unable to reach such paths; it must then say so.
                @($scanWarnings).Count + @($hashWarnings).Count | Should -BeGreaterThan 0
                return
            }
            $result.Count | Should -Be 1
            $result[0].Count | Should -Be 2
            foreach ($folder in $result[0].Folders) { $folder.Length | Should -BeGreaterThan 260 }

            $report = Join-Path $root 'long.xlsx'
            Export-DuplicateReport -DuplicateSet $result -Path $report
            (Update-DuplicateReport -Path $report).CopiesRemoved | Should -Be 0
        }
        finally {
            # Removed here: the test drive cleanup may not reach long paths on every platform.
            foreach ($copy in 'a', 'b') { [System.IO.Directory]::Delete((ConvertTo-RawPath (Join-Path $root $copy)), $true) }
        }
    }

    It 'handles files larger than 4 GB' {
        if ($script:OnWindows) { Set-ItResult -Skipped -Because 'sparse test files need extra set-up on Windows'; return }
        $root = Add-TestRoot
        $size = 4GB + 1
        foreach ($copy in 'one', 'two') {
            $folder = [System.IO.Directory]::CreateDirectory([System.IO.Path]::Combine($root, $copy, 'big')).FullName
            $path = Join-Path $folder 'big.bin'
            # Extending a file without writing to it makes a sparse file: no disk space is used.
            $stream = [System.IO.File]::Create($path)
            try { $stream.SetLength($size) } finally { $stream.Dispose() }
            [System.IO.File]::SetLastWriteTimeUtc($path, $script:Saved)
        }

        $folders = [System.Collections.Generic.List[object]]::new()
        $files = @(Get-FileInventory -Path $root -FolderInfo $folders)
        $cache = [System.Collections.Generic.Dictionary[string, string]]::new()
        $result = @(Find-DuplicateFile -File $files -Md5Cache $cache -ThrottleLimit 2)
        $folderSets = @(Find-DuplicateFolder -File $files -Folder $folders -Md5Cache $cache)
        $report = Join-Path (Add-TestRoot) 'big.xlsx'
        Export-DuplicateReport -DuplicateSet $result -FolderSet $folderSets -Path $report

        $result.Count | Should -Be 1
        $result[0].MD5 | Should -Be 'F18C798FF5D450DFE4D3ACDC12B621FF'
        (Import-DuplicateReport -Path $report).SizeBytes | Should -Be $size
        (Import-DuplicateFolderReport -Path $report).SizeBytes | Should -Be $size
        $summary = Update-DuplicateReport -Path $report
        $summary.CopiesRemoved + $summary.FolderCopiesRemoved | Should -Be 0
    }

    It 'finds every duplicate in a tree of 10,000 files' {
        $root = Add-TestRoot
        # 50 folders in each of two trees, 100 files in each; every file has one copy in the
        # other tree, and many files share a name, date and size but not their contents.
        foreach ($copy in 'a', 'b') {
            for ($f = 0; $f -lt 50; $f++) {
                $folder = [System.IO.Directory]::CreateDirectory([System.IO.Path]::Combine($root, $copy, "folder$f")).FullName
                for ($i = 0; $i -lt 100; $i++) {
                    $path = [System.IO.Path]::Combine($folder, "file$i.txt")
                    [System.IO.File]::WriteAllText($path, "$f-$i")
                    [System.IO.File]::SetLastWriteTimeUtc($path, $script:Saved)
                }
            }
        }

        $folders = [System.Collections.Generic.List[object]]::new()
        $files = @(Get-FileInventory -Path $root -FolderInfo $folders)
        $cache = [System.Collections.Generic.Dictionary[string, string]]::new()
        $result = @(Find-DuplicateFile -File $files -Md5Cache $cache -ThrottleLimit 4)
        $folderSets = @(Find-DuplicateFolder -File $files -Folder $folders -Md5Cache $cache)
        $report = Join-Path (Add-TestRoot) 'many.xlsx'
        Export-DuplicateReport -DuplicateSet $result -FolderSet $folderSets -Path $report

        $files.Count | Should -Be 10000
        $result.Count | Should -Be 5000
        $folderSets.Count | Should -Be 50
        @(Import-DuplicateReport -Path $report).Count | Should -Be 5000
        (Update-DuplicateReport -Path $report).CopiesRemoved | Should -Be 0
    }

    It 'does not match a copy whose saved time a FAT drive rounded to 2 seconds' {
        # FAT and exFAT (USB sticks, memory cards) store saved times to 2 seconds, so a copy of
        # a file saved at 10:30:01 reads 10:30:02. The saved date must match to the second.
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'disk/x.txt' -SavedUtc $script:Saved.AddSeconds(1)
            Add-TestFile $root 'usb/x.txt' -SavedUtc $script:Saved.AddSeconds(2)
        )
        @(Find-DuplicateFile -File $files).Count | Should -Be 0
    }

    It 'finds, saves and validates duplicates saved before 1970 and after 2038 (<Year>)' -ForEach @(
        @{ Year = 1960 }   # before 1970: a negative time on Linux and macOS
        @{ Year = 2040 }   # after January 2038, when 32-bit Unix time runs out
    ) {
        $saved = [datetime]::new($Year, 1, 15, 8, 0, 0, [System.DateTimeKind]::Utc)
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'a/dated.txt' -SavedUtc $saved
            Add-TestFile $root 'b/dated.txt' -SavedUtc $saved
        )
        $report = Join-Path $root 'dates.xlsx'
        Export-DuplicateReport -DuplicateSet @(Find-DuplicateFile -File $files) -Path $report

        (Import-DuplicateReport -Path $report).LastWriteTime | Should -Be $saved.ToLocalTime()
        (Update-DuplicateReport -Path $report).CopiesRemoved | Should -Be 0
    }

    It 'keeps copies saved in winter and in summer when validating in a time zone with daylight saving' {
        if ($script:OnWindows) { Set-ItResult -Skipped -Because 'the TZ variable sets the time zone only on Linux and macOS'; return }
        $previous = $env:TZ
        $env:TZ = 'America/New_York'
        [System.TimeZoneInfo]::ClearCachedData()
        try {
            $root = Add-TestRoot
            $files = foreach ($season in @{ Name = 'summer'; Month = 7 }, @{ Name = 'winter'; Month = 1 }) {
                $saved = [datetime]::new(2024, $season.Month, 15, 12, 0, 0, [System.DateTimeKind]::Utc)
                Add-TestFile $root "a/$($season.Name).txt" -SavedUtc $saved
                Add-TestFile $root "b/$($season.Name).txt" -SavedUtc $saved
            }
            $report = Join-Path $root 'seasons.xlsx'
            Export-DuplicateReport -DuplicateSet @(Find-DuplicateFile -File $files) -Path $report

            (Import-DuplicateReport -Path $report).LastWriteTime.Hour | Should -Be @(8, 7) -Because 'noon UTC is 8:00 in summer (UTC-4) and 7:00 in winter (UTC-5)'
            (Update-DuplicateReport -Path $report).CopiesRemoved | Should -Be 0
        }
        finally {
            $env:TZ = $previous
            [System.TimeZoneInfo]::ClearCachedData()
        }
    }

    It 'scans a folder given as a symbolic link' {
        $root = Add-TestRoot
        $null = Add-TestFile $root 'real/a/x.txt'
        $null = Add-TestFile $root 'real/b/x.txt'
        $link = Join-Path $root 'link'
        try { $null = New-Item -ItemType SymbolicLink -Path $link -Target (Join-Path $root 'real') -ErrorAction Stop }
        catch { Set-ItResult -Skipped -Because "symbolic links cannot be created here: $_"; return }

        try {
            $result = @(Find-DuplicateFile -File @(Get-FileInventory -Path $link))
            $result.Count | Should -Be 1
            $result[0].Folders | Should -Be @((Join-Path $link 'a'), (Join-Path $link 'b'))
        }
        finally { [System.IO.Directory]::Delete($link) }  # the link only, never its target
    }

    It 'skips a folder deleted during the scan and carries on' {
        $root = Add-TestRoot
        foreach ($folder in 'first', 'second', 'third') { $null = Add-TestFile $root "$folder/x.txt" }
        # Delete "second" while "first" is being scanned, after the root's listing included it.
        Mock -ModuleName DuplicateFinder Write-Verbose {
            $scanning = [regex]::Match($Message, '^Scanning (.+)[\\/]first$')
            if ($scanning.Success) { [System.IO.Directory]::Delete([System.IO.Path]::Combine($scanning.Groups[1].Value, 'second'), $true) }
        }

        $files = @(Get-FileInventory -Path $root -WarningVariable warnings -WarningAction SilentlyContinue)

        $files.DirectoryName | Should -Be @((Join-Path $root 'first'), (Join-Path $root 'third'))
        "$warnings" | Should -BeLike "*$(Join-Path $root 'second')*"
        @(Find-DuplicateFile -File $files).Count | Should -Be 1
    }

    It 'skips a file the operating system will not let it read' {
        if (Test-RunningAsRoot) { Set-ItResult -Skipped -Because 'root can read every file'; return }
        $root = Add-TestRoot
        $locked = (Add-TestFile $root 'locked/x.txt').FullName
        $null = Add-TestFile $root 'b/x.txt'
        $null = Add-TestFile $root 'c/x.txt'
        $stream = $null
        # Windows: held open by another handle that shares nothing. Elsewhere: no read permission.
        if ($script:OnWindows) { $stream = [System.IO.File]::Open($locked, 'Open', 'ReadWrite', 'None') }
        else { chmod 000 $locked }

        try {
            $result = @(Find-DuplicateFile -File @(Get-FileInventory -Path $root) -WarningVariable warnings -WarningAction SilentlyContinue)
            "$warnings" | Should -BeLike '*locked*'
            $result.Count | Should -Be 1
            $result[0].Count | Should -Be 2
        }
        finally {
            if ($stream) { $stream.Dispose() } else { chmod 644 $locked }
        }
    }

    It 'reports two files in one folder whose names differ only in case' {
        $root = Add-TestRoot
        $folder = [System.IO.Directory]::CreateDirectory((Join-Path $root 'photos')).FullName
        # Windows folders can be made case-sensitive (as WSL does); that needs no admin rights.
        if ($script:OnWindows) { $null = fsutil.exe file setCaseSensitiveInfo $folder enable 2>&1 }
        $null = Add-TestFile $root 'photos/IMG.JPG'
        if ([System.IO.File]::Exists((Join-Path $folder 'img.jpg'))) {
            Set-ItResult -Skipped -Because 'this folder ignores case'; return
        }
        $null = Add-TestFile $root 'photos/img.jpg'

        $result = @(Find-DuplicateFile -File @(Get-FileInventory -Path $root))
        $report = Join-Path (Add-TestRoot) 'case.xlsx'
        Export-DuplicateReport -DuplicateSet $result -Path $report

        $result.Count | Should -Be 1
        $result[0].FileName | Should -BeExactly 'IMG.JPG'
        $result[0].Folders | Should -Be @($folder, $folder)
        (Update-DuplicateReport -Path $report).CopiesRemoved | Should -Be 0
    }

    It 'writes a file with as many copies as Excel has location columns' {
        $set = [pscustomobject] @{
            FileName = 'x.txt'; LastWriteTime = $script:Saved; SizeBytes = 1; MD5 = 'A'; Count = 16378
            Folders = [string[]] @(1..16378 | ForEach-Object { "/copy$_" })
        }
        $report = Join-Path (Add-TestRoot) 'wide.xlsx'
        Export-DuplicateReport -DuplicateSet @($set) -Path $report
        (Import-DuplicateReport -Path $report).Count | Should -Be 16378
    }

    It 'refuses a file with more copies than Excel has location columns' {
        $set = [pscustomobject] @{
            FileName = 'x.txt'; LastWriteTime = $script:Saved; SizeBytes = 1; MD5 = 'A'; Count = 16379
            Folders = [string[]] @(1..16379 | ForEach-Object { "/copy$_" })
        }
        $report = Join-Path (Add-TestRoot) 'too-wide.xlsx'
        { Export-DuplicateReport -DuplicateSet @($set) -Path $report } |
            Should -Throw 'A file has 16379 copies; Excel supports at most 16378 location columns.'
        Test-Path -LiteralPath $report | Should -BeFalse
    }

    It 'refuses more duplicated files than Excel has rows' {
        $report = Join-Path (Add-TestRoot) 'too-long.xlsx'
        InModuleScope DuplicateFinder -Parameters @{ Report = $report; Saved = $script:Saved } {
            $sets = @(1..3 | ForEach-Object {
                    [pscustomobject] @{ FileName = "$_.txt"; LastWriteTime = $Saved; SizeBytes = 1; MD5 = 'A'; Count = 2; Folders = @('/a', '/b') }
                })
            $script:ExcelMaxRows = 3  # instead of writing a million rows
            try { { Export-DuplicateReport -DuplicateSet $sets -Path $Report } | Should -Throw 'Found 3 duplicated files; Excel supports at most 2 rows.' }
            finally { $script:ExcelMaxRows = 1048576 }
        }
    }

    It 'refuses text longer than an Excel cell can hold' {
        $set = [pscustomobject] @{
            FileName = 'x.txt'; LastWriteTime = $script:Saved; SizeBytes = 1; MD5 = 'A'; Count = 2
            Folders = @(('/' + ('x' * 32767)), '/b')
        }
        $report = Join-Path (Add-TestRoot) 'long-cell.xlsx'
        { Export-DuplicateReport -DuplicateSet @($set) -Path $report } |
            Should -Throw 'Cell G2 would hold 32768 characters; Excel allows at most 32767.'
        Test-Path -LiteralPath $report | Should -BeFalse
    }

    It 'reads and validates a report after LibreOffice has saved it' {
        $root = Add-TestRoot
        $data = Join-Path $root 'data'
        foreach ($copy in 'a', 'b') {
            $null = Add-TestFile $data "$copy/x & y.txt"
            $null = Add-TestFile $data "$copy/Holiday/p.jpg" -Content 'photo'
        }
        $report = Join-Path $root 'report.xlsx'
        $null = & $script:ScriptPath -Path $data -OutputFile $report -IncludeFolders 6>$null
        $saved = Save-WithLibreOffice -Path $report -OutFolder (Join-Path $root 'resaved')
        if (-not $saved) {
            if ($env:FIND_DUPLICATES_REQUIRE_LIBREOFFICE) { throw 'LibreOffice Calc is required for this test but could not convert the report.' }
            Set-ItResult -Skipped -Because 'LibreOffice Calc is not installed'; return
        }

        $before = @(Import-DuplicateReport -Path $report)
        $after = @(Import-DuplicateReport -Path $saved)
        $after.Count | Should -Be $before.Count
        for ($i = 0; $i -lt $before.Count; $i++) {
            foreach ($property in 'FileName', 'LastWriteTime', 'SizeBytes', 'MD5', 'Count') {
                $after[$i].$property | Should -Be $before[$i].$property
            }
            $after[$i].Folders | Should -Be $before[$i].Folders
        }
        (Import-DuplicateFolderReport -Path $saved).FolderName | Should -Be 'Holiday'
        $summary = Update-DuplicateReport -Path $saved
        $summary.CopiesRemoved + $summary.FolderCopiesRemoved | Should -Be 0
    }

    It 'copes with names Windows reserves or trims (<Name>)' -ForEach @(
        @{ Name = 'NUL' }                # reserved on Windows (a device)
        @{ Name = 'PRN.txt' }            # reserved on Windows, even with an extension
        @{ Name = 'ends in a dot.' }     # Windows trims a trailing dot from ordinary paths
        @{ Name = 'ends in a space ' }   # ... and a trailing space
    ) {
        $root = Add-TestRoot
        foreach ($copy in 'a', 'b') { $null = Add-RawFile (Join-Path $root $copy) $Name }

        try {
            # Windows: it must finish without failing (such copies may be reported or skipped with
            # a warning). Elsewhere these are ordinary names and must be matched.
            $result = @(Find-DuplicateFile -File @(Get-FileInventory -Path $root -WarningAction SilentlyContinue) -WarningAction SilentlyContinue)
            $report = Join-Path (Add-TestRoot) 'names.xlsx'
            Export-DuplicateReport -DuplicateSet $result -Path $report
            $summary = Update-DuplicateReport -Path $report -WarningAction SilentlyContinue
            if (-not $script:OnWindows) {
                $result.Count | Should -Be 1
                $result[0].FileName | Should -BeExactly $Name
                $summary.CopiesRemoved | Should -Be 0
            }
        }
        finally {
            # Removed here: ordinary deletion cannot reach these names on Windows.
            foreach ($copy in 'a', 'b') { [System.IO.Directory]::Delete((ConvertTo-RawPath (Join-Path $root $copy)), $true) }
        }
    }

    It 'includes hidden and system files and folders' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root '.hidden-folder/.hidden'
            Add-TestFile $root 'b/.hidden'
        )
        if ($script:OnWindows) {
            foreach ($file in $files) { $file.Attributes = [System.IO.FileAttributes] 'Hidden, System, Archive' }
            $folder = [System.IO.DirectoryInfo] (Join-Path $root '.hidden-folder')
            $folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::Hidden
        }

        $result = @(Find-DuplicateFile -File @(Get-FileInventory -Path $root))
        $result.Count | Should -Be 1
        $result[0].FileName | Should -Be '.hidden'
    }

    It 'orders rows by file name ignoring case, then saved date, then MD5' {
        $root = Add-TestRoot
        $files = @(
            Add-TestFile $root 'one/B.txt' -Content 'b'
            Add-TestFile $root 'two/b.txt' -Content 'b'
            Add-TestFile $root 'one/a.txt' -Content 'x' -SavedUtc $script:Saved.AddHours(1)
            Add-TestFile $root 'two/a.txt' -Content 'x' -SavedUtc $script:Saved.AddHours(1)
            Add-TestFile $root 'three/a.txt' -Content 'x'
            Add-TestFile $root 'four/a.txt' -Content 'x'
            Add-TestFile $root 'five/A.txt' -Content 'y'
            Add-TestFile $root 'six/A.txt' -Content 'y'
            Add-TestFile $root 'one/a' -Content 'z'
            Add-TestFile $root 'two/a' -Content 'z'
        )
        $md5x = (Get-FileHash -LiteralPath $files[4].FullName -Algorithm MD5).Hash
        $md5y = (Get-FileHash -LiteralPath $files[6].FullName -Algorithm MD5).Hash
        $sameDate = @($md5x, $md5y)
        if ([string]::CompareOrdinal($md5x, $md5y) -gt 0) { $sameDate = @($md5y, $md5x) }

        $result = @(Find-DuplicateFile -File $files)

        $result.FileName.ToUpperInvariant() | Should -Be @('A', 'A.TXT', 'A.TXT', 'A.TXT', 'B.TXT') -Because 'a name sorts before the longer names it starts'
        $result[1..2].MD5 | Should -Be $sameDate
        $result[3].LastWriteTime | Should -Be $script:Saved.AddHours(1).ToLocalTime()
    }
}
