param([string]$RequestPath)
$ErrorActionPreference='Stop'

function Assert-EveNoReparse([string]$Path) {
    $current=[IO.Path]::GetFullPath($Path)
    while ($current) {
        if ((Test-Path -LiteralPath $current) -and ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw "Refusing a reparse point or junction: $current" }
        $parent=Split-Path -Parent $current
        if ($parent -eq $current) { break }
        $current=$parent
    }
}
function Get-EveInstallation([string]$MainPath) {
    $main=[IO.Path]::GetFullPath($MainPath)
    $directory=Split-Path -Parent $main
    $name=Split-Path -Leaf $main
    if ([IO.Path]::GetExtension($main) -ne '.exe' -or $name -in @('EVE-X-Preview-Settings.exe','EVE-X-Preview-Updater.exe','EVE-X-Preview-SettingsHost.exe')) { throw 'Invalid main executable target.' }
    if ($directory -eq [IO.Path]::GetPathRoot($directory) -or -not (Test-Path -LiteralPath $directory -PathType Container)) { throw 'Choose an existing application directory, not a drive root.' }
    Assert-EveNoReparse $main
    return @{Main=$main;Directory=$directory;Name=$name}
}
function Get-EveUpdatePath([string]$Directory,[string]$Name) {
    if (-not $Name -or $Name -match '[\\/:<>"|?*\x00-\x1f]' -or $Name -match '[. ]$' -or $Name -in @('.','..')) { throw "Unsafe update path: $Name" }
    return Join-Path $Directory $Name
}
function Wait-EveUpdateExit($Installation,[int]$TimeoutSeconds=30) {
    $paths=@($Installation.Main)
    foreach ($name in @('EVE-X-Preview.exe','EVE-X-Preview-Settings.exe','EVE-X-Preview-SettingsHost.exe','EVE-X-Preview-Updater.exe')) { $paths+=Join-Path $Installation.Directory $name }
    $deadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $running=@(Get-Process | Where-Object { try { $_.Path -in $paths } catch { $false } })
        if (-not $running.Count) { return }
        if ([DateTime]::UtcNow -ge $deadline) { throw 'Close EVE-X-Preview and its settings editor, then retry. No process was forcibly terminated.' }
        Start-Sleep -Milliseconds 200
    } while ($true)
}
function Assert-EveReleaseExe([string]$Path,[string]$Version) {
    Assert-EveNoReparse $Path
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing release executable: $Path" }
    $stream=[IO.File]::Open($Path,'Open','Read','Read')
    $reader=[IO.BinaryReader]::new($stream)
    try {
        if ($stream.Length -lt 64 -or $stream.Length -gt 536870912 -or $reader.ReadUInt16() -ne 0x5A4D) { throw 'Invalid or truncated executable.' }
        $stream.Position=0x3C; $header=$reader.ReadUInt32()
        if ($header -gt $stream.Length-24) { throw 'Truncated PE header.' }
        $stream.Position=$header
        if ($reader.ReadUInt32() -ne 0x4550 -or $reader.ReadUInt16() -ne 0x8664) { throw 'Release must be a Windows x64 PE executable.' }
        $sections=$reader.ReadUInt16(); $stream.Position=$header+20
        $optionalSize=$reader.ReadUInt16(); $characteristics=$reader.ReadUInt16()
        if (($characteristics -band 0x2000) -or $sections -lt 1 -or $sections -gt 96 -or $header+24+$optionalSize+40*$sections -gt $stream.Length) { throw 'Invalid PE executable section table.' }
        for ($index=0;$index -lt $sections;$index++) {
            $stream.Position=$header+24+$optionalSize+40*$index+16
            [long]$size=$reader.ReadUInt32(); [long]$offset=$reader.ReadUInt32()
            if ($offset+$size -gt $stream.Length) { throw 'Truncated executable section.' }
        }
    } finally { $reader.Dispose() }
    $info=[Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
    $expected=[version]$Version
    $expected=[version]::new($expected.Major,$expected.Minor,[Math]::Max(0,$expected.Build),[Math]::Max(0,$expected.Revision))
    $actual=[version]::new($info.FileMajorPart,$info.FileMinorPart,$info.FileBuildPart,$info.FilePrivatePart)
    if (-not $info.FileVersion -or $actual -ne $expected) { throw "Executable version $actual does not match release $Version : $Path" }
}
function Write-EveUpdateJournal([string]$Transaction,$Journal) {
    $path=Join-Path $Transaction 'journal.json'; $temporary=$path+'.next'
    $bytes=[Text.Encoding]::UTF8.GetBytes(($Journal | ConvertTo-Json -Depth 8))
    $stream=[IO.File]::Open($temporary,'Create','Write','None')
    try { $stream.Write($bytes,0,$bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
    if (Test-Path -LiteralPath $path) { [IO.File]::Replace($temporary,$path,[NullString]::Value) } else { [IO.File]::Move($temporary,$path) }
}
function Assert-EveSettingsProbe([string]$Path) {
    # A self-contained .NET bundle extends beyond its PE sections. Exercise the
    # existing no-UI probe to verify the bundled runtime and application can load.
    $result=Join-Path (Split-Path -Parent $Path) ('probe-'+[guid]::NewGuid().ToString('N')+'.json')
    $process=$null
    try {
        $process=Start-Process -FilePath $Path -ArgumentList '--probe',('"'+$result+'"') -WorkingDirectory (Split-Path -Parent $Path) -WindowStyle Hidden -PassThru -RedirectStandardOutput ($result+'.stdout') -RedirectStandardError ($result+'.stderr')
        $null=$process.Handle
        if (-not $process.WaitForExit(15000)) { Stop-Process -Id $process.Id -Force; $process.WaitForExit(); throw 'Settings executable probe timed out.' }
        if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $result -PathType Leaf) -or (Get-Item -LiteralPath $result).Length -gt 4096) { throw 'Settings executable probe failed. The download may be incomplete.' }
        $probe=Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
        if ($probe.application -cne 'EVE-X-Preview-Settings' -or $probe.bridgeVersion -ne 1 -or $probe.schemaVersion -ne 1) { throw 'Settings executable probe returned an incompatible response.' }
    } finally { if ($process) { $process.Dispose() } }
}
function Remove-EveUpdateTransaction($Installation,[string]$Transaction) {
    $expected=Join-Path $Installation.Directory '.EVE-X-Preview-update'
    if ([IO.Path]::GetFullPath($Transaction) -ne $expected) { throw 'Unsafe recovery cleanup path.' }
    Assert-EveNoReparse $Transaction
    $pending=[Collections.Generic.Stack[string]]::new(); $pending.Push($Transaction)
    while ($pending.Count) {
        foreach ($child in Get-ChildItem -LiteralPath $pending.Pop() -Force) {
            if ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Refusing recovery cleanup through a junction.' }
            if ($child.PSIsContainer) { $pending.Push($child.FullName) }
        }
    }
    Remove-Item -LiteralPath $Transaction -Recurse -Force
}
function Restore-EveExecutables($Installation,[string]$Transaction,$Journal) {
    if ($Journal.schemaVersion -ne 2 -or $Journal.installation -ne $Installation.Directory -or $Journal.mainFile -ne $Installation.Name -or $Journal.state -notin @('installing','committed')) { throw 'Invalid executable recovery journal.' }
    $seen=@{}
    if (@($Journal.operations).Count -lt 1 -or @($Journal.operations).Count -gt 2) { throw 'Invalid recovery operation count.' }
    # Validate every target and backup before changing any installed file.
    foreach ($entry in $Journal.operations) {
        $target=Get-EveUpdatePath $Installation.Directory $entry.path
        if ($entry.path -notin @($Installation.Name,'EVE-X-Preview-Settings.exe') -or $seen.ContainsKey($entry.path) -or $entry.existed -isnot [bool]) { throw 'Invalid recovery target in journal.' }
        $seen[$entry.path]=$true
        Assert-EveNoReparse $target
        if ($Journal.state -eq 'committed') { continue }
        if ($entry.existed) {
            $backup=Get-EveUpdatePath (Join-Path $Transaction 'backup') $entry.path
            Assert-EveNoReparse $backup
            if ($entry.sha256 -notmatch '^[a-fA-F0-9]{64}$' -or -not (Test-Path -LiteralPath $backup -PathType Leaf) -or (Get-FileHash -LiteralPath $backup -Algorithm SHA256).Hash -ne $entry.sha256) { throw "Missing or damaged recovery backup: $backup" }
        }
    }
    if ($Journal.state -eq 'committed') { return }
    foreach ($entry in $Journal.operations) {
        $target=Get-EveUpdatePath $Installation.Directory $entry.path
        if ($entry.existed) {
            $backup=Get-EveUpdatePath (Join-Path $Transaction 'backup') $entry.path
            if ((Test-Path -LiteralPath $target -PathType Leaf) -and (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -eq $entry.sha256) { continue }
            [IO.File]::Copy($backup,$target,$true)
        } elseif (Test-Path -LiteralPath $target -PathType Leaf) { [IO.File]::Delete($target) }
    }
}
function Recover-EveExecutablesUnlocked($Installation,[string]$Transaction) {
    Assert-EveNoReparse $Transaction
    if (-not (Test-Path -LiteralPath $Transaction)) { return }
    $path=Join-Path $Transaction 'journal.json'
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        Assert-EveNoReparse $path
        $journal=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        Restore-EveExecutables $Installation $Transaction $journal
    }
    # No published journal means replacement never started.
    Remove-EveUpdateTransaction $Installation $Transaction
}
function Open-EveUpdateLock($Installation) {
    $path=Join-Path $Installation.Directory '.EVE-X-Preview-update.lock'
    Assert-EveNoReparse $path
    return [IO.File]::Open($path,'OpenOrCreate','ReadWrite','None')
}
function Recover-EveExeUpdate([string]$MainPath) {
    $installation=Get-EveInstallation $MainPath
    Wait-EveUpdateExit $installation
    $lock=Open-EveUpdateLock $installation
    try { Recover-EveExecutablesUnlocked $installation (Join-Path $installation.Directory '.EVE-X-Preview-update') }
    finally { $lock.Dispose() }
}
function Invoke-EveExeUpdate {
    param([Parameter(Mandatory)][string]$SourceDirectory,[Parameter(Mandatory)][string]$MainPath,[Parameter(Mandatory)][string]$Version)
    if ($Version -notmatch '^\d+(\.\d+){1,3}$') { throw 'Invalid release version.' }
    $installation=Get-EveInstallation $MainPath
    Wait-EveUpdateExit $installation
    $lock=Open-EveUpdateLock $installation
    $transaction=Join-Path $installation.Directory '.EVE-X-Preview-update'
    $ownsTransaction=$false; $retain=$false; $journal=$null
    try {
        Recover-EveExecutablesUnlocked $installation $transaction
        $files=@(@{asset='EVE-X-Preview.exe';path=$installation.Name})
        if (Test-Path -LiteralPath (Join-Path $installation.Directory 'EVE-X-Preview-Settings.exe') -PathType Leaf) { $files+=@{asset='EVE-X-Preview-Settings.exe';path='EVE-X-Preview-Settings.exe'} }
        foreach ($file in $files) {
            $source=Get-EveUpdatePath ([IO.Path]::GetFullPath($SourceDirectory)) $file.asset
            Assert-EveReleaseExe $source $Version
            Assert-EveNoReparse (Get-EveUpdatePath $installation.Directory $file.path)
        }
        New-Item -ItemType Directory -Path $transaction | Out-Null; $ownsTransaction=$true
        $stage=Join-Path $transaction 'stage'; $backupDirectory=Join-Path $transaction 'backup'
        New-Item -ItemType Directory -Path $stage,$backupDirectory | Out-Null
        $operations=@()
        foreach ($file in $files) {
            $staged=Join-Path $stage $file.path
            [IO.File]::Copy((Join-Path $SourceDirectory $file.asset),$staged)
            Assert-EveReleaseExe $staged $Version
            if ($file.asset -eq 'EVE-X-Preview-Settings.exe') { Assert-EveSettingsProbe $staged }
            $target=Get-EveUpdatePath $installation.Directory $file.path
            if (Test-Path -LiteralPath $target -PathType Container) { throw "A directory occupies an executable path: $target" }
            $existed=Test-Path -LiteralPath $target -PathType Leaf
            $hash=''
            if ($existed) {
                [IO.File]::Copy($target,(Join-Path $backupDirectory $file.path))
                $hash=(Get-FileHash -LiteralPath (Join-Path $backupDirectory $file.path) -Algorithm SHA256).Hash
            }
            $operations+=@{path=$file.path;existed=$existed;sha256=$hash}
        }
        $journal=@{schemaVersion=2;installation=$installation.Directory;mainFile=$installation.Name;state='installing';operations=$operations}
        Write-EveUpdateJournal $transaction $journal
        Wait-EveUpdateExit $installation
        foreach ($operation in $operations) {
            $target=Get-EveUpdatePath $installation.Directory $operation.path
            Assert-EveNoReparse $target
            $source=Join-Path $stage $operation.path
            if ($operation.existed) { [IO.File]::Replace($source,$target,[NullString]::Value) } else { [IO.File]::Move($source,$target) }
        }
        $journal.state='committed'; Write-EveUpdateJournal $transaction $journal
    } catch {
        $failure=$_
        if ($journal -and $journal.state -ne 'committed') {
            try { Restore-EveExecutables $installation $transaction $journal }
            catch { $retain=$true; throw "Update failed and recovery could not finish. Backups retained at $transaction. $failure; recovery: $_" }
            throw "Update failed; previous executables were restored. $failure"
        }
        throw
    } finally {
        try { if ($ownsTransaction -and -not $retain) { Remove-EveUpdateTransaction $installation $transaction } }
        finally { $lock.Dispose() }
    }
}

if ($RequestPath) {
    try {
        $request=Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json
        if ($request.action -eq 'recover') { Recover-EveExeUpdate -MainPath $request.main }
        elseif ($request.action -eq 'install') { Invoke-EveExeUpdate -SourceDirectory $request.source -MainPath $request.main -Version $request.version }
        else { throw 'Unknown updater action.' }
        $result=@{ok=$true;message='Executable update transaction completed.'}
    } catch { $result=@{ok=$false;message=$_.ToString()} }
    $result | ConvertTo-Json -Compress | Set-Content -LiteralPath ($RequestPath+'.result') -Encoding UTF8
    if (-not $result.ok) { exit 1 }
}
