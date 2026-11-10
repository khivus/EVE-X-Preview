param([string]$RequestPath)

# Windows PowerShell 5.1 is available on supported Windows installations.
# The same implementation is embedded in the updater and used by the release builder.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$script:EveManifestName = 'EVE-X-Preview.package.json'

function Get-EveRelativePath([string]$Path, [switch]$Directory) {
    if ($Directory) { $Path = $Path.TrimEnd('/') }
    if (-not $Path -or $Path -match '[\\:<>"|?*\x00-\x1f]' -or $Path.StartsWith('/')) { throw "Unsafe package path: $Path" }
    foreach ($part in $Path.Split('/')) {
        if (-not $part -or $part -in @('.', '..') -or $part -match '[. ]$' -or $part -match '^(?i:con|prn|aux|nul|com[1-9]|lpt[1-9])(\.|$)') {
            throw "Unsafe package path: $Path"
        }
    }
    return $Path
}
function Assert-EvePackageFile([string]$Path) {
    $null = Get-EveRelativePath $Path
    if ($Path -notmatch '(?i)\.dll$' -and $Path -notin @('EVE-X-Preview.exe', 'EVE-X-Preview-Updater.exe',
        'EVE-X-Preview-Settings.exe', 'EVE-X-Preview-SettingsHost.exe', 'EVE-X-Preview-Settings.deps.json',
        'EVE-X-Preview-Settings.runtimeconfig.json')) { throw "Not a distributable package file: $Path" }
}
function Get-EveChildPath([string]$Root, [string]$RelativePath) {
    $null = Get-EveRelativePath $RelativePath
    $prefix = [IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    $path = [IO.Path]::GetFullPath((Join-Path $prefix $RelativePath.Replace('/', '\')))
    if (-not $path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Package path escapes its root.' }
    return $path
}
function Assert-EveNoReparse([string]$Path) {
    $current = [IO.Path]::GetFullPath($Path)
    while ($current) {
        if (Test-Path -LiteralPath $current) {
            if ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Refusing a reparse point or junction: $current"
            }
        }
        $parent = Split-Path -Parent $current
        if ($parent -eq $current) { break }
        $current = $parent
    }
}
function Read-EveManifest($Manifest, [string]$Version) {
    if ($Manifest.schemaVersion -ne 1 -or $Manifest.application -cne 'EVE-X-Preview' -or $Manifest.architecture -cne 'win-x64') {
        throw 'Unsupported package manifest.'
    }
    if ($Manifest.version -notmatch '^\d+(\.\d+){1,3}$' -or ($Version -and [version]$Manifest.version -ne [version]$Version)) {
        throw 'Package version does not match the selected release version.'
    }
    $files = @{}
    if (@($Manifest.files).Count -lt 2 -or @($Manifest.files).Count -gt 4096) { throw 'Invalid package file count.' }
    foreach ($file in $Manifest.files) {
        Assert-EvePackageFile $file.path
        if ($files.ContainsKey($file.path)) { throw "Duplicate manifest path: $($file.path)" }
        if ($file.sha256 -notmatch '^[a-fA-F0-9]{64}$' -or $file.size -isnot [ValueType] -or $file.size -lt 0 -or $file.size -gt 536870912) {
            throw "Invalid package size or hash: $($file.path)"
        }
        $files[$file.path] = $file
    }
    foreach ($required in @('EVE-X-Preview.exe', 'EVE-X-Preview-Updater.exe')) {
        if (-not $files.ContainsKey($required)) { throw "Missing required package file: $required" }
    }
    return $files
}
function Wait-EvePackageExit([string]$Directory, [string]$MainPath, [int]$TimeoutSeconds = 30) {
    $paths = @($MainPath)
    foreach ($name in @('EVE-X-Preview.exe', 'EVE-X-Preview-Settings.exe', 'EVE-X-Preview-SettingsHost.exe', 'EVE-X-Preview-Updater.exe')) { $paths += Join-Path $Directory $name }
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $running = @(Get-Process | Where-Object { try { $_.Path -in $paths } catch { $false } })
        if (-not $running.Count) { return }
        if ([DateTime]::UtcNow -ge $deadline) { throw 'Close EVE-X-Preview and its settings editor, then retry. No process was forcibly terminated.' }
        Start-Sleep -Milliseconds 200
    } while ($true)
}
function Remove-EveTransaction([string]$Directory, [string]$Transaction) {
    $expected = Join-Path ([IO.Path]::GetFullPath($Directory)) '.EVE-X-Preview-update'
    if ([IO.Path]::GetFullPath($Transaction) -ne $expected) { throw 'Unexpected transaction cleanup path.' }
    Assert-EveNoReparse $Transaction
    $pending = [Collections.Generic.Stack[string]]::new()
    $pending.Push($Transaction)
    while ($pending.Count) {
        foreach ($child in Get-ChildItem -LiteralPath $pending.Pop() -Force) {
            if ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Refusing cleanup through a reparse point: $($child.FullName)" }
            if ($child.PSIsContainer) { $pending.Push($child.FullName) }
        }
    }
    Remove-Item -LiteralPath $Transaction -Recurse -Force
}
function Write-EveJournal([string]$Transaction, $Journal) {
    $path = Join-Path $Transaction 'journal.json'
    $temporary = $path + '.next'
    $bytes = [Text.Encoding]::UTF8.GetBytes(($Journal | ConvertTo-Json -Depth 8))
    $stream = [IO.File]::Open($temporary, 'Create', 'Write', 'None')
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
    if (Test-Path -LiteralPath $path) { [IO.File]::Replace($temporary, $path, [NullString]::Value) }
    else { [IO.File]::Move($temporary, $path) }
}
function Restore-EveTransaction([string]$Directory, [string]$Transaction, $Journal) {
    if ($Journal.state -eq 'committed') { return }
    # Backups are completed before the journal is published. Restoring all paths
    # also handles an interrupted process between a rename and journal update.
    foreach ($entry in $Journal.operations) {
        if ($entry.path -eq $Journal.mainFile) {
            $null = Get-EveRelativePath $entry.path
            if ($entry.path.Contains('/') -or [IO.Path]::GetExtension($entry.path) -ne '.exe') { throw 'Invalid recovery main executable.' }
        } elseif ($entry.path -ne $script:EveManifestName) { Assert-EvePackageFile $entry.path }
        $target = Get-EveChildPath $Directory $entry.path
        Assert-EveNoReparse $target
        if ($entry.existed) {
            $backup = Get-EveChildPath (Join-Path $Transaction 'backup') $entry.path
            Assert-EveNoReparse $backup
            if (-not (Test-Path -LiteralPath $backup -PathType Leaf)) { throw "Missing recovery backup: $backup" }
            # Avoid replacing an unchanged file which may be locked by the cause of failure.
            if ((Test-Path -LiteralPath $target -PathType Leaf) -and
                (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $backup -Algorithm SHA256).Hash) { continue }
            New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
            [IO.File]::Copy($backup, $target, $true)
        } elseif (Test-Path -LiteralPath $target -PathType Leaf) { [IO.File]::Delete($target) }
    }
}
function Invoke-EvePackageUpdate {
    param([Parameter(Mandatory)][string]$ArchivePath, [Parameter(Mandatory)][string]$MainPath,
        [Parameter(Mandatory)][string]$Version, [bool]$IncludeSettings = $true, [string]$ExpectedArchiveHash)
    if ($Version -notmatch '^\d+(\.\d+){1,3}$') { throw 'Invalid release version.' }
    $MainPath = [IO.Path]::GetFullPath($MainPath)
    if ([IO.Path]::GetExtension($MainPath) -ne '.exe') { throw 'Only compiled installations can be updated.' }
    $directory = Split-Path -Parent $MainPath
    if ($directory -eq [IO.Path]::GetPathRoot($directory)) { throw 'Choose an application directory, not a drive root.' }
    Assert-EveNoReparse $MainPath
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { throw 'The installation directory does not exist.' }
    Wait-EvePackageExit $directory $MainPath
    # An exclusive lock prevents two updater instances from sharing a transaction.
    $lockPath = Join-Path $directory '.EVE-X-Preview-update.lock'
    Assert-EveNoReparse $lockPath
    $lock = [IO.File]::Open($lockPath, 'OpenOrCreate', 'ReadWrite', 'None')
    $transaction = Join-Path $directory '.EVE-X-Preview-update'
    $journal = $null; $ownsTransaction = $false; $keepRecovery = $false
    try {
        Assert-EveNoReparse $transaction
        if (Test-Path -LiteralPath $transaction) {
            $journalPath = Join-Path $transaction 'journal.json'
            if (Test-Path -LiteralPath $journalPath -PathType Leaf) {
                Assert-EveNoReparse $journalPath
                $previous = Get-Content -LiteralPath $journalPath -Raw | ConvertFrom-Json
                if ($previous.installation -ne $directory -or $previous.schemaVersion -ne 1) { throw 'Invalid interrupted update journal.' }
                Restore-EveTransaction $directory $transaction $previous
            }
            # No published journal means installation replacement never began.
            Remove-EveTransaction $directory $transaction
        }
        if ($ExpectedArchiveHash -and ($ExpectedArchiveHash -notmatch '^[a-fA-F0-9]{64}$' -or
            (Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256).Hash -ne $ExpectedArchiveHash)) { throw 'Release archive checksum mismatch.' }
        $archive = [IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($ArchivePath))
        try {
            $entries = @{}; [long]$total = 0
            foreach ($entry in $archive.Entries) {
                $isDirectory = $entry.FullName.EndsWith('/')
                $path = Get-EveRelativePath $entry.FullName -Directory:$isDirectory
                if ($entries.ContainsKey($path)) { throw "Duplicate archive path: $path" }
                # Reject Windows reparse points and Unix symbolic links.
                if (($entry.ExternalAttributes -band 0x400) -or (($entry.ExternalAttributes -shr 16) -band 0xF000) -eq 0xA000) { throw 'Archive contains a symbolic link or reparse point.' }
                if ($isDirectory) { continue }
                $total += $entry.Length
                if ($entry.Length -gt 536870912 -or $total -gt 1073741824 -or $entries.Count -ge 4097) { throw 'Release archive exceeds package limits.' }
                $entries[$path] = $entry
            }
            if (-not $entries.ContainsKey($script:EveManifestName)) { throw 'Release has no package manifest. Install the full portable package manually.' }
            if ($entries[$script:EveManifestName].Length -gt 1048576) { throw 'Package manifest is too large.' }
            $reader = [IO.StreamReader]::new($entries[$script:EveManifestName].Open())
            try { $manifest = $reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose() }
            $files = Read-EveManifest $manifest $Version
            if ($entries.Count -ne $files.Count + 1) { throw 'Archive contains unlisted package files.' }
            if ($IncludeSettings -and -not $files.ContainsKey('EVE-X-Preview-Settings.exe')) { throw 'Missing required settings executable.' }
            foreach ($file in $files.Values) {
                if (-not $entries.ContainsKey($file.path) -or $entries[$file.path].Length -ne $file.size) { throw "Package file size mismatch: $($file.path)" }
                $target = Get-EveChildPath $directory $file.path
                Assert-EveNoReparse $target
            }
            New-Item -ItemType Directory -Path $transaction | Out-Null
            $ownsTransaction = $true
            $stage = Join-Path $transaction 'stage'
            New-Item -ItemType Directory -Path $stage | Out-Null
            foreach ($file in $files.Values) {
                $path = Get-EveChildPath $stage $file.path
                New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
                [IO.Compression.ZipFileExtensions]::ExtractToFile($entries[$file.path], $path)
                if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $file.sha256) { throw "Package file hash mismatch: $($file.path)" }
            }
        } finally { $archive.Dispose() }
        $selected = @($manifest.files | Where-Object { $IncludeSettings -or $_.path -in @('EVE-X-Preview.exe', 'EVE-X-Preview-Updater.exe') })
        $installedManifestPath = Join-Path $directory $script:EveManifestName
        Assert-EveNoReparse $installedManifestPath
        $oldFiles = @{}
        if (Test-Path -LiteralPath $installedManifestPath -PathType Leaf) {
            $oldManifest = Get-Content -LiteralPath $installedManifestPath -Raw | ConvertFrom-Json
            $oldFiles = Read-EveManifest $oldManifest ''
        }
        $newManifest = [ordered]@{ schemaVersion=1; application='EVE-X-Preview'; version=$Version; architecture='win-x64'; files=$selected }
        $newManifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $stage $script:EveManifestName) -Encoding UTF8
        $desired = @{}; $operations = @()
        foreach ($file in $selected) {
            $path = $file.path
            if ($path -eq 'EVE-X-Preview.exe') { $path = Split-Path -Leaf $MainPath }
            $desired[$file.path] = $true
            $operations += @{ path=$path; source=$file.path; remove=$false }
        }
        foreach ($file in $oldFiles.Values) {
            if ($desired.ContainsKey($file.path)) { continue }
            $path = Get-EveChildPath $directory $file.path
            Assert-EveNoReparse $path
            if ((Test-Path -LiteralPath $path -PathType Leaf) -and (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $file.sha256) {
                throw "An obsolete package file was modified; preserve or remove it before retrying: $path"
            }
            $operations += @{ path=$file.path; source=''; remove=$true }
        }
        # Commit the manifest last so it never describes a partially installed package.
        $operations += @{ path=$script:EveManifestName; source=$script:EveManifestName; remove=$false }
        foreach ($operation in $operations) {
            $target = Get-EveChildPath $directory $operation.path
            Assert-EveNoReparse $target
            if (Test-Path -LiteralPath $target -PathType Container) { throw "A directory occupies a package file path: $target" }
            $operation.existed = Test-Path -LiteralPath $target -PathType Leaf
            if ($operation.existed) {
                $backup = Get-EveChildPath (Join-Path $transaction 'backup') $operation.path
                New-Item -ItemType Directory -Path (Split-Path -Parent $backup) -Force | Out-Null
                [IO.File]::Copy($target, $backup)
            }
        }
        $journal = @{ schemaVersion=1; installation=$directory; mainFile=(Split-Path -Leaf $MainPath); state='installing'; operations=$operations }
        Write-EveJournal $transaction $journal
        Wait-EvePackageExit $directory $MainPath
        foreach ($operation in $operations) {
            $target = Get-EveChildPath $directory $operation.path
            Assert-EveNoReparse $target
            if ($operation.remove) { if ($operation.existed) { [IO.File]::Delete($target) }; continue }
            New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
            $source = Get-EveChildPath $stage $operation.source
            if ($operation.existed) { [IO.File]::Replace($source, $target, [NullString]::Value) }
            else { [IO.File]::Move($source, $target) }
        }
        $journal.state = 'committed'
        Write-EveJournal $transaction $journal
    } catch {
        $failure = $_
        if ($journal -and $journal.state -ne 'committed') {
            try { Restore-EveTransaction $directory $transaction $journal }
            catch { $keepRecovery = $true; throw "Update failed and rollback could not finish. Recovery files retained at $transaction. $failure; rollback: $_" }
            throw "Update failed; the previous package was restored. $failure"
        }
        throw
    } finally {
        if ($ownsTransaction -and -not $keepRecovery) { Remove-EveTransaction $directory $transaction }
        $lock.Dispose()
        # Keep the zero-byte lock file: deleting a shared lock name creates a race.
    }
}
function New-EveReleasePackage {
    param([Parameter(Mandatory)][string]$Directory, [Parameter(Mandatory)][string]$Version, [Parameter(Mandatory)][string]$ArchivePath)
    $Directory = [IO.Path]::GetFullPath($Directory)
    $ArchivePath = [IO.Path]::GetFullPath($ArchivePath)
    if ($ArchivePath.StartsWith($Directory.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'The release archive must be outside the package directory.' }
    if (Test-Path -LiteralPath $ArchivePath) { throw 'Choose an unused archive path.' }
    $files = @()
    foreach ($file in Get-ChildItem -LiteralPath $Directory -File -Recurse | Sort-Object FullName) {
        Assert-EveNoReparse $file.FullName
        $relative = $file.FullName.Substring($Directory.TrimEnd('\').Length + 1).Replace('\', '/')
        if ($relative -eq $script:EveManifestName) { continue }
        Assert-EvePackageFile $relative
        $files += @{ path=$relative; size=$file.Length; sha256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
    }
    $manifest = @{ schemaVersion=1; application='EVE-X-Preview'; version=$Version; architecture='win-x64'; files=$files }
    $null = Read-EveManifest $manifest $Version
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $Directory $script:EveManifestName) -Encoding UTF8
    [IO.Compression.ZipFile]::CreateFromDirectory($Directory, [IO.Path]::GetFullPath($ArchivePath), 'Optimal', $false)
    $digest = (Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText($ArchivePath + '.sha256', $digest + '  ' + [IO.Path]::GetFileName($ArchivePath) + "`n", [Text.Encoding]::ASCII)
}

if ($RequestPath) {
    $result = $null
    try {
        $request = Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json
        Invoke-EvePackageUpdate -ArchivePath $request.archive -MainPath $request.main -Version $request.version -IncludeSettings ([bool]$request.includeSettings) -ExpectedArchiveHash $request.sha256
        $result = @{ ok=$true; message='Complete release package installed.' }
    } catch { $result = @{ ok=$false; message=$_.ToString() } }
    $result | ConvertTo-Json -Compress | Set-Content -LiteralPath ($RequestPath + '.result') -Encoding UTF8
    if (-not $result.ok) { exit 1 }
}
