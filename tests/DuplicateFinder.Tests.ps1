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
        Get-Item -LiteralPath $full
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
        $rows[0] | Should -Be @('File Name', 'Last Modified', 'Size (bytes)', 'MD5', 'Copies', 'Location 1', 'Location 2', 'Location 3')
        $rows[1][0] | Should -BeExactly 'a & b <1>.txt'
        $rows[1][2] | Should -Be '1234'
        $rows[1][3] | Should -Be 'AAAA'
        $rows[1][4] | Should -Be '3'
        $rows[1][5..7] | Should -Be @('C:\one', 'C:\two & more', 'D:\three')
        $rows[2][5..6] | Should -Be @('C:\x', "C:\bad$([char] 0xFFFD)name")
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
        $sheetXml | Should -Match '<autoFilter ref="A1:H3"'
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
                FileName = 'a & b.txt'; LastWriteTime = [datetime]::new(2024, 1, 2, 3, 4, 5, 678)
                SizeBytes = 1234; MD5 = 'AAAA'; Count = 3; Folders = [string[]] @('C:\one', 'C:\two', 'D:\three')
            }
            [pscustomobject] @{
                FileName = 'z.txt'; LastWriteTime = [datetime]::new(2023, 6, 7, 8, 9, 10)
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

    It 'reads back what Export-DuplicateReport wrote' {
        $path = Join-Path (Add-TestRoot) 'report.xlsx'
        Export-DuplicateReport -DuplicateSet $script:RoundTrip -Path $path

        $read = @(Import-DuplicateReport -Path $path)

        $read.Count | Should -Be 2
        for ($i = 0; $i -lt 2; $i++) {
            foreach ($property in 'FileName', 'LastWriteTime', 'SizeBytes', 'MD5', 'Count') {
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
        $rows[1][5..7] | Should -Be $expected -Because 'the full folder path of every copy is recorded'
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
