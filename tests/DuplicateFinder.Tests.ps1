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

    function Read-Worksheet {
        # Returns the worksheet as a list of rows, each row a list of cell texts.
        param([string] $Path)
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $reader = [System.IO.StreamReader]::new($zip.GetEntry('xl/worksheets/sheet1.xml').Open())
            try { [xml] $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        finally { $zip.Dispose() }

        $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        foreach ($row in $xml.SelectNodes('//s:sheetData/s:row', $ns)) {
            , @(foreach ($cell in $row.SelectNodes('s:c', $ns)) {
                    $text = $cell.SelectSingleNode('s:is/s:t', $ns)
                    if ($text) { $text.InnerText } else { $cell.SelectSingleNode('s:v', $ns).InnerText }
                })
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
        $files = @(
            Add-TestFile $root 'b/photo.jpg'
            Add-TestFile $root 'a/Photo.JPG'
        )
        $scanned = @(Get-FileInventory -Path $root)

        (Find-DuplicateFile -File $scanned).FileName | Should -BeExactly 'Photo.JPG'
    }

    It 'returns nothing for an empty list' {
        @(Find-DuplicateFile -File @()).Count | Should -Be 0
    }

    Context 'MD5 is only calculated when name and saved date already match' {
        BeforeEach {
            Mock -ModuleName DuplicateFinder Get-FileHash { [pscustomobject] @{ Hash = 'ABC' } }
        }

        It 'does not hash files with different names' {
            $root = Add-TestRoot
            $null = Find-DuplicateFile -File @((Add-TestFile $root 'a/x.txt'), (Add-TestFile $root 'b/y.txt'))
            Should -Invoke -ModuleName DuplicateFinder Get-FileHash -Times 0 -Exactly
        }

        It 'does not hash files with different saved dates' {
            $root = Add-TestRoot
            $null = Find-DuplicateFile -File @(
                (Add-TestFile $root 'a/x.txt'),
                (Add-TestFile $root 'b/x.txt' -SavedUtc $script:Saved.AddDays(1)))
            Should -Invoke -ModuleName DuplicateFinder Get-FileHash -Times 0 -Exactly
        }

        It 'does not hash files of different sizes' {
            $root = Add-TestRoot
            $null = Find-DuplicateFile -File @(
                (Add-TestFile $root 'a/x.txt' -Content 'short'),
                (Add-TestFile $root 'b/x.txt' -Content 'much longer'))
            Should -Invoke -ModuleName DuplicateFinder Get-FileHash -Times 0 -Exactly
        }

        It 'hashes only the matching candidates' {
            $root = Add-TestRoot
            $null = Find-DuplicateFile -File @(
                (Add-TestFile $root 'a/x.txt'),
                (Add-TestFile $root 'b/x.txt'),
                (Add-TestFile $root 'c/y.txt'))
            Should -Invoke -ModuleName DuplicateFinder Get-FileHash -Times 2 -Exactly
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
        Mock -ModuleName DuplicateFinder Get-FileHash { [pscustomobject] @{ Hash = 'SAME' } }

        $result = @(Find-DuplicateFile -File $files -SkipCloudOnly -WarningVariable warnings -WarningAction SilentlyContinue)

        Should -Invoke -ModuleName DuplicateFinder Get-FileHash -Times 2 -Exactly
        Should -Invoke -ModuleName DuplicateFinder Get-FileHash -Times 0 -Exactly -ParameterFilter { "$LiteralPath" -like '*cloud*' }
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
        Mock -ModuleName DuplicateFinder Get-FileHash {
            if ("$LiteralPath" -like '*locked*') { throw 'file is locked' }
            [pscustomobject] @{ Hash = 'SAME' }
        }

        $result = @(Find-DuplicateFile -File $files -WarningVariable warnings -WarningAction SilentlyContinue)

        @($warnings).Count | Should -Be 1
        "$($warnings[0])" | Should -BeLike '*locked*'
        $result.Count | Should -Be 1
        $result[0].Folders | Should -Be @($files[1..2].DirectoryName | Sort-Object)
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
                'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml'
            ($zip.Entries.FullName | Sort-Object) | Should -Be ($expected | Sort-Object)
            foreach ($entry in $zip.Entries) {
                $reader = [System.IO.StreamReader]::new($entry.Open())
                try { { [xml] $reader.ReadToEnd() } | Should -Not -Throw -Because $entry.FullName }
                finally { $reader.Dispose() }
            }
        }
        finally { $zip.Dispose() }
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

    It 'scans a network share given as a UNC path' {
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
