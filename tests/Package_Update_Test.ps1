param([string]$OutputDirectory = (Join-Path $env:TEMP ('evex-package-tests-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
if ($env:USERNAME -ne 'WDAGUtilityAccount') { throw 'Run package installation tests inside Windows Sandbox.' }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
$script:passed = 0
$script:failed = 0
function Assert($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Test([string]$Name, [scriptblock]$Body) {
    try { & $Body; $script:passed++; "PASS: $Name" }
    catch { $script:failed++; "FAIL: $Name -- $_" }
}
function Write-Text([string]$Path, [string]$Text) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    [IO.File]::WriteAllText($Path, $Text)
}
function Fixture([string]$Name) {
    $root = Join-Path $OutputDirectory $Name
    $install = Join-Path $root 'installation with spaces'
    New-Item -ItemType Directory -Path $install -Force | Out-Null
    Write-Text (Join-Path $install 'EVE-X-Preview.exe') 'old main'
    Write-Text (Join-Path $install 'EVE-X-Preview-Settings.exe') 'old settings'
    Write-Text (Join-Path $install 'old-runtime.dll') 'obsolete runtime'
    Write-Text (Join-Path $install 'EVE-X-Preview.json') '{"personal":"settings"}'
    Write-Text (Join-Path $install 'Sounds\alert.wav') 'personal sound'
    Write-Text (Join-Path $install 'notes.txt') 'personal notes'
    return @{ Root = $root; Install = $install }
}
# Fixtures deliberately build the release format independently of the production packager.
function Archive($Fixture, [hashtable]$Overrides = @{}, [string[]]$Omit = @(), [string]$Version = '9.8.7.6') {
    $content = [ordered]@{
        'EVE-X-Preview.exe' = 'new main'
        'EVE-X-Preview-Updater.exe' = 'new updater'
        'EVE-X-Preview-Settings.exe' = 'new settings'
        'EVE-X-Preview-SettingsHost.exe' = 'new settings host'
        'coreclr.dll' = 'new runtime'
        'runtimes/win-x64/native/libSkiaSharp.dll' = 'new native dependency'
        'EVE-X-Preview-Settings.runtimeconfig.json' = '{"runtimeOptions":{}}'
    }
    foreach ($key in $Overrides.Keys) { $content[$key] = $Overrides[$key] }
    $files = @()
    $zipPath = Join-Path $Fixture.Root ('release-' + [guid]::NewGuid().ToString('N') + '.zip')
    $zip = [IO.Compression.ZipFile]::Open($zipPath, 'Create')
    try {
        foreach ($path in $content.Keys) {
            if ($path -in $Omit) { continue }
            $bytes = [Text.Encoding]::UTF8.GetBytes($content[$path])
            $hash = [Security.Cryptography.SHA256]::Create()
            try { $digest = [BitConverter]::ToString($hash.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant() } finally { $hash.Dispose() }
            $files += @{ path = $path; size = $bytes.Length; sha256 = $digest }
            $stream = $zip.CreateEntry($path).Open()
            try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
        }
        $manifest = @{ schemaVersion=1; application='EVE-X-Preview'; version=$Version; architecture='win-x64'; files=$files }
        $writer = [IO.StreamWriter]::new($zip.CreateEntry('EVE-X-Preview.package.json').Open())
        try { $writer.Write(($manifest | ConvertTo-Json -Depth 8)) } finally { $writer.Dispose() }
    } finally { $zip.Dispose() }
    return $zipPath
}
function Install($Fixture, [string]$ArchivePath, [bool]$IncludeSettings = $true) {
    Invoke-EvePackageUpdate -ArchivePath $ArchivePath -MainPath (Join-Path $Fixture.Install 'EVE-X-Preview.exe') -Version '9.8.7.6' -IncludeSettings $IncludeSettings
}
function Must-Reject([scriptblock]$Action, [string]$Pattern) {
    $caught = $false
    try { & $Action | Out-Null } catch { $caught = $true; Assert ($_ -match $Pattern) "Unexpected failure: $_" }
    Assert $caught 'The unsafe update was accepted.'
}
$engine = Join-Path (Split-Path $PSScriptRoot -Parent) 'src\PackageUpdate.ps1'
if (-not (Test-Path -LiteralPath $engine)) { throw 'Missing package updater: the existing EXE-only installer cannot update runtime dependencies.' }
. $engine

Test 'Updates the main app, settings app, host and nested dependencies together without replacing user data' {
    $f = Fixture 'complete'
    Install $f (Archive $f)
    Assert ((Get-Content (Join-Path $f.Install 'coreclr.dll') -Raw) -eq 'new runtime') 'Runtime dependency was not updated.'
    Assert ((Get-Content (Join-Path $f.Install 'runtimes\win-x64\native\libSkiaSharp.dll') -Raw) -eq 'new native dependency') 'Nested native dependency was not installed.'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview-SettingsHost.exe') -Raw) -eq 'new settings host') 'Settings host was not updated.'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview-Settings.exe') -Raw) -eq 'new settings') 'Settings app was not updated.'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.json') -Raw) -eq '{"personal":"settings"}') 'Settings changed.'
    Assert ((Get-Content (Join-Path $f.Install 'Sounds\alert.wav') -Raw) -eq 'personal sound') 'Sound changed.'
    Assert ((Get-Content (Join-Path $f.Install 'notes.txt') -Raw) -eq 'personal notes') 'User file changed.'
    Assert ((Get-Content (Join-Path $f.Install 'old-runtime.dll') -Raw) -eq 'obsolete runtime') 'An unowned legacy file was removed.'
}
Test 'Removes obsolete files owned by the previous manifest and preserves unowned files' {
    $f = Fixture 'obsolete'
    Install $f (Archive $f @{ 'obsolete.dll' = 'old managed dependency' })
    Install $f (Archive $f)
    Assert (-not (Test-Path (Join-Path $f.Install 'obsolete.dll'))) 'Obsolete dependency survived.'
    Assert (Test-Path (Join-Path $f.Install 'old-runtime.dll')) 'Unowned file was deleted.'
}
Test 'Preserves a built-in-only installation without implicitly adding the optional editor' {
    $f = Fixture 'builtin'
    Remove-Item -LiteralPath (Join-Path $f.Install 'EVE-X-Preview-Settings.exe')
    Install $f (Archive $f) $false
    Assert (-not (Test-Path (Join-Path $f.Install 'EVE-X-Preview-Settings.exe'))) 'Optional editor was implicitly installed.'
    Assert (-not (Test-Path (Join-Path $f.Install 'coreclr.dll'))) 'Optional runtime was implicitly installed.'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'new main') 'Main did not update.'
}
Test 'Rejects a release version mismatch before changing installed files' {
    $f = Fixture 'wrong-version'
    Must-Reject { Install $f (Archive $f @{} @() '1.0.0') } 'version'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'old main') 'Main changed after validation failed.'
}
Test 'Rejects a corrupt archive checksum before extraction' {
    $f = Fixture 'wrong-checksum'; $zip = Archive $f
    Must-Reject { Invoke-EvePackageUpdate -ArchivePath $zip -MainPath (Join-Path $f.Install 'EVE-X-Preview.exe') -Version '9.8.7.6' -ExpectedArchiveHash ('0' * 64) } 'checksum'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'old main') 'Main changed.'
}
Test 'Rejects archive traversal and never writes outside the installation' {
    $f = Fixture 'traversal'
    Must-Reject { Install $f (Archive $f @{ '../escape.dll' = 'bad' }) } 'path'
    Assert (-not (Test-Path (Join-Path $f.Root 'escape.dll'))) 'Archive escaped.'
}
Test 'Rejects package entries that try to replace application settings' {
    $f = Fixture 'config'
    Must-Reject { Install $f (Archive $f @{ 'EVE-X-Preview.json' = 'bad' }) } 'package file'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.json') -Raw) -eq '{"personal":"settings"}') 'Settings replaced.'
}
Test 'Rejects a missing settings executable without partially updating the main app' {
    $f = Fixture 'missing-editor'
    Must-Reject { Install $f (Archive $f @{} @('EVE-X-Preview-Settings.exe')) } 'required'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'old main') 'Main partially updated.'
}
Test 'Rejects a file whose content differs from the manifest hash' {
    $f = Fixture 'bad-file'; $zipPath = Archive $f
    $zip = [IO.Compression.ZipFile]::Open($zipPath, 'Update')
    try {
        $zip.GetEntry('coreclr.dll').Delete()
        $writer = [IO.StreamWriter]::new($zip.CreateEntry('coreclr.dll').Open())
        try { $writer.Write('tampered') } finally { $writer.Dispose() }
    } finally { $zip.Dispose() }
    Must-Reject { Install $f $zipPath } 'hash|size'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'old main') 'Main changed.'
}
Test 'Rolls back files replaced before a real sharing violation and removes newly introduced files' {
    $f = Fixture 'rollback'; $zip = Archive $f
    # The lock permits backup reads, but denies delete/rename of the last replacement.
    $locked = [IO.File]::Open((Join-Path $f.Install 'EVE-X-Preview-Settings.exe'), 'Open', 'Read', 'Read')
    try { Must-Reject { Install $f $zip } 'restored|rollback' } finally { $locked.Dispose() }
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'old main') 'Main was not restored.'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview-Settings.exe') -Raw) -eq 'old settings') 'Settings was not restored.'
    Assert (-not (Test-Path (Join-Path $f.Install 'coreclr.dll'))) 'New runtime file survived rollback.'
    Assert (-not (Test-Path (Join-Path $f.Install 'EVE-X-Preview-Updater.exe'))) 'New updater file survived rollback.'
    Assert (-not (Test-Path (Join-Path $f.Install 'EVE-X-Preview.package.json'))) 'Manifest committed after rollback.'
}
Test 'Rejects case-insensitive duplicate archive entries' {
    $f = Fixture 'duplicates'
    $zipPath = Archive $f
    $zip = [IO.Compression.ZipFile]::Open($zipPath, 'Update')
    try { $null = $zip.CreateEntry('CORECLR.DLL') } finally { $zip.Dispose() }
    Must-Reject { Install $f $zipPath } 'duplicate'
}
Test 'Rejects a destination junction instead of modifying its target' {
    $f = Fixture 'junction'
    $outside = Join-Path $f.Root 'outside'
    New-Item -ItemType Directory -Path $outside | Out-Null
    New-Item -ItemType Junction -Path (Join-Path $f.Install 'runtimes') -Target $outside | Out-Null
    Must-Reject { Install $f (Archive $f) } 'reparse|junction'
    Assert (-not (Test-Path (Join-Path $outside 'win-x64'))) 'Junction target was modified.'
}
Test 'Rolls back a renamed main executable after a later replacement fails' {
    $f = Fixture 'renamed'; $zip = Archive $f
    $renamed = Join-Path $f.Install 'My Preview.exe'
    Move-Item -LiteralPath (Join-Path $f.Install 'EVE-X-Preview.exe') -Destination $renamed
    $locked = [IO.File]::Open((Join-Path $f.Install 'EVE-X-Preview-Settings.exe'), 'Open', 'Read', 'Read')
    try { Must-Reject { Invoke-EvePackageUpdate -ArchivePath $zip -MainPath $renamed -Version '9.8.7.6' } 'previous package was restored' } finally { $locked.Dispose() }
    Assert ((Get-Content -LiteralPath $renamed -Raw) -eq 'old main') 'Renamed main was not restored.'
    Assert (-not (Test-Path (Join-Path $f.Install '.EVE-X-Preview-update'))) 'Rollback left a broken transaction.'
}
Test 'Discards an interrupted staging directory before retrying without touching user data' {
    $f = Fixture 'staging-interruption'
    Write-Text (Join-Path $f.Install '.EVE-X-Preview-update\stage\coreclr.dll') 'partially downloaded runtime'
    Install $f (Archive $f)
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'new main') 'Retry did not finish.'
}
Test 'Recovers an interrupted replacement before rejecting an unrelated bad release' {
    $f = Fixture 'commit-interruption'
    $transaction = Join-Path $f.Install '.EVE-X-Preview-update'
    Write-Text (Join-Path $transaction 'backup\EVE-X-Preview.exe') 'old main'
    Write-Text (Join-Path $f.Install 'EVE-X-Preview.exe') 'interrupted replacement'
    Write-Text (Join-Path $f.Install 'new.dll') 'partially installed dependency'
    $journal = @{ schemaVersion=1; installation=$f.Install; state='installing'; mainFile='EVE-X-Preview.exe'; operations=@(
        @{path='EVE-X-Preview.exe';existed=$true}, @{path='new.dll';existed=$false}) }
    Write-Text (Join-Path $transaction 'journal.json') ($journal | ConvertTo-Json -Depth 8)
    Must-Reject { Install $f (Archive $f @{} @() '1.0.0') } 'version'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'old main') 'Interrupted update was not recovered.'
    Assert (-not (Test-Path (Join-Path $f.Install 'new.dll'))) 'Interrupted dependency was not removed.'
}
Test 'Repackages an existing build output and produces a checksum matching the resulting ZIP' {
    $f = Fixture 'repackage'
    $package = Join-Path $f.Root 'payload'
    Write-Text (Join-Path $package 'EVE-X-Preview.exe') 'main'
    Write-Text (Join-Path $package 'EVE-X-Preview-Updater.exe') 'updater'
    $first = Join-Path $f.Root 'first.zip'; $second = Join-Path $f.Root 'second.zip'
    New-EveReleasePackage -Directory $package -Version '9.8.7.6' -ArchivePath $first
    New-EveReleasePackage -Directory $package -Version '9.8.7.6' -ArchivePath $second
    Assert ((Get-Content -LiteralPath ($second + '.sha256') -Raw).StartsWith((Get-FileHash -LiteralPath $second -Algorithm SHA256).Hash.ToLowerInvariant())) 'Archive checksum does not match.'
}
Test 'Recovers after a real installer process is terminated during replacement' {
    $f = Fixture 'killed-installer'
    $extra = @{}
    for ($index=0; $index -lt 700; $index++) { $extra[('dependency-{0:0000}.dll' -f $index)] = ('payload-' + $index) }
    $zip = Archive $f $extra
    $request = Join-Path $f.Root 'request.json'
    Write-Text $request (@{archive=$zip;main=(Join-Path $f.Install 'EVE-X-Preview.exe');version='9.8.7.6';includeSettings=$true} | ConvertTo-Json)
    $process = Start-Process powershell.exe -ArgumentList '-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',('"' + $engine + '"'),'-RequestPath',('"' + $request + '"') -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $f.Root 'stdout.log') -RedirectStandardError (Join-Path $f.Root 'stderr.log')
    $null = $process.Handle
    $killed = $false; $deadline = [DateTime]::UtcNow.AddSeconds(60)
    while (-not $process.HasExited -and [DateTime]::UtcNow -lt $deadline) {
        $observer = [IO.File]::Open((Join-Path $f.Install 'EVE-X-Preview.exe'), 'Open', 'Read', ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $reader = [IO.StreamReader]::new($observer)
        try { $observed = $reader.ReadToEnd() } finally { $reader.Dispose() }
        if ($observed -eq 'new main') {
            Stop-Process -Id $process.Id -Force
            $process.WaitForExit(); $killed = $true; break
        }
        Start-Sleep -Milliseconds 1
        $process.Refresh()
    }
    if (-not $killed -and -not $process.HasExited) { Stop-Process -Id $process.Id -Force }
    Assert $killed 'The installer did not reach the observed replacement boundary.'
    $journal = Get-Content -LiteralPath (Join-Path $f.Install '.EVE-X-Preview-update\journal.json') -Raw | ConvertFrom-Json
    Assert ($journal.state -eq 'installing') 'Test did not interrupt a live transaction.'
    Must-Reject { Install $f (Archive $f @{} @() '1.0.0') } 'version'
    Assert ((Get-Content (Join-Path $f.Install 'EVE-X-Preview.exe') -Raw) -eq 'old main') 'Killed transaction did not restore main.'
    Assert (-not (Test-Path (Join-Path $f.Install 'EVE-X-Preview-Updater.exe'))) 'Killed transaction left a new file.'
    Assert (-not (Test-Path (Join-Path $f.Install '.EVE-X-Preview-update'))) 'Recovery did not complete.'
}
"Package update tests: $script:passed passed, $script:failed failed."
if ($script:failed) { exit 1 }
