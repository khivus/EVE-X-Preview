param([string]$OutputDirectory = (Join-Path $env:TEMP ('evex-exe-tests-' + [guid]::NewGuid().ToString('N'))),
    [string]$ValidExecutable = 'C:\Runtime\AutoHotkey64.exe', [string]$BundledSettingsExecutable, [string]$MainReleaseExecutable)
$ErrorActionPreference='Stop'
if ($env:USERNAME -ne 'WDAGUtilityAccount') { throw 'Run executable update tests in Windows Sandbox.' }
$version=[Diagnostics.FileVersionInfo]::GetVersionInfo($ValidExecutable).FileVersion
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
. (Join-Path $PSScriptRoot 'Exe_Update_Fixtures.ps1')
$settingsFixture=Join-Path $OutputDirectory 'settings-probe.exe'
New-EveProbeFixture $settingsFixture $version $ValidExecutable
$script:passed=0; $script:failed=0
function Assert($Condition,[string]$Message) { if (-not $Condition) { throw $Message } }
function Test([string]$Name,[scriptblock]$Body) {
    try { & $Body; $script:passed++; "PASS: $Name" }
    catch { $script:failed++; "FAIL: $Name -- $_" }
}
function Text([string]$Path,[string]$Value) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    [IO.File]::WriteAllText($Path,$Value)
}
function Fixture([string]$Name) {
    $root=Join-Path $OutputDirectory $Name
    $install=Join-Path $root 'installation with spaces'
    $release=Join-Path $root 'separate release EXEs'
    New-Item -ItemType Directory -Path $install,$release -Force | Out-Null
    Text (Join-Path $install 'EVE-X-Preview.exe') 'old main'
    Text (Join-Path $install 'EVE-X-Preview-Settings.exe') 'old settings'
    Text (Join-Path $install 'EVE-X-Preview-Updater.exe') 'keep recovery updater'
    Text (Join-Path $install 'EVE-X-Preview.json') '{"personal":"keep me"}'
    Text (Join-Path $install 'Sounds\alert.wav') 'personal sound'
    Copy-Item -LiteralPath $ValidExecutable -Destination (Join-Path $release 'EVE-X-Preview.exe')
    Copy-Item -LiteralPath $settingsFixture -Destination (Join-Path $release 'EVE-X-Preview-Settings.exe')
    return @{Root=$root;Install=$install;Release=$release;Main=(Join-Path $install 'EVE-X-Preview.exe')}
}
function Install($Fixture) { Invoke-EveExeUpdate -MainPath $Fixture.Main -SourceDirectory $Fixture.Release -Version $version }
function Reject([scriptblock]$Action,[string]$Pattern) {
    $caught=$false
    try { & $Action | Out-Null } catch { $caught=$true; Assert ($_ -match $Pattern) "Unexpected rejection: $_" }
    Assert $caught 'An unsafe update was accepted.'
}
$engine=Join-Path (Split-Path $PSScriptRoot -Parent) 'src\ExeUpdate.ps1'
if (-not (Test-Path -LiteralPath $engine)) { throw 'EXE-only recovery installer is missing.' }
. $engine
Test 'Updates both installed EXEs and preserves settings, sounds and the recovery updater' {
    $f=Fixture 'installed'; Install $f
    foreach ($name in @('EVE-X-Preview.exe','EVE-X-Preview-Settings.exe')) {
        $expected=if ($name -eq 'EVE-X-Preview.exe') { $ValidExecutable } else { $settingsFixture }
        Assert ((Get-FileHash -LiteralPath (Join-Path $f.Install $name)).Hash -eq (Get-FileHash -LiteralPath $expected).Hash) "$name was not updated."
    }
    Assert ((Get-Content -LiteralPath (Join-Path $f.Install 'EVE-X-Preview.json') -Raw) -eq '{"personal":"keep me"}') 'Settings changed.'
    Assert ((Get-Content -LiteralPath (Join-Path $f.Install 'Sounds\alert.wav') -Raw) -eq 'personal sound') 'Sound changed.'
    Assert ((Get-Content -LiteralPath (Join-Path $f.Install 'EVE-X-Preview-Updater.exe') -Raw) -eq 'keep recovery updater') 'Recovery updater was downgraded.'
    Assert (-not (Test-Path (Join-Path $f.Install 'EVE-X-Preview.package.json'))) 'An unnecessary release manifest was installed.'
}
Test 'Built-in-only update succeeds with only the main EXE available' {
    $f=Fixture 'builtin'
    Remove-Item -LiteralPath (Join-Path $f.Install 'EVE-X-Preview-Settings.exe'),(Join-Path $f.Release 'EVE-X-Preview-Settings.exe')
    Install $f
    Assert (-not (Test-Path (Join-Path $f.Install 'EVE-X-Preview-Settings.exe'))) 'Optional editor was installed.'
    Assert ((Get-FileHash -LiteralPath $f.Main).Hash -eq (Get-FileHash -LiteralPath $ValidExecutable).Hash) 'Main did not update.'
}
Test 'Missing installed-settings download leaves both installed EXEs unchanged' {
    $f=Fixture 'missing-settings'; Remove-Item -LiteralPath (Join-Path $f.Release 'EVE-X-Preview-Settings.exe')
    Reject { Install $f } 'missing'
    Assert ((Get-Content -LiteralPath $f.Main -Raw) -eq 'old main') 'Main changed before downloads were complete.'
}
Test 'Wrong release version is rejected before replacement' {
    $f=Fixture 'wrong-version'
    Reject { Invoke-EveExeUpdate -MainPath $f.Main -SourceDirectory $f.Release -Version '99.99.99.99' } 'version'
    Assert ((Get-Content -LiteralPath $f.Main -Raw) -eq 'old main') 'Main changed.'
}
Test 'Truncated or non-PE settings EXE is rejected before replacement' {
    $f=Fixture 'truncated'; Text (Join-Path $f.Release 'EVE-X-Preview-Settings.exe') 'not an executable'
    Reject { Install $f } 'executable|PE|truncated'
    Assert ((Get-Content -LiteralPath $f.Main -Raw) -eq 'old main') 'Main changed.'
}
Test 'A sharing violation after replacing main rolls the original files back' {
    $f=Fixture 'locked'
    $locked=[IO.File]::Open((Join-Path $f.Install 'EVE-X-Preview-Settings.exe'),'Open','Read','Read')
    try { Reject { Install $f } 'previous.*restored' } finally { $locked.Dispose() }
    Assert ((Get-Content -LiteralPath $f.Main -Raw) -eq 'old main') 'Main was not restored.'
    Assert ((Get-Content -LiteralPath (Join-Path $f.Install 'EVE-X-Preview-Settings.exe') -Raw) -eq 'old settings') 'Settings was not restored.'
    Assert (-not (Test-Path (Join-Path $f.Install '.EVE-X-Preview-update'))) 'Successful rollback retained a journal.'
}
Test 'Renamed main executable is restored after a later failure' {
    $f=Fixture 'renamed'; $renamed=Join-Path $f.Install 'My Preview.exe'
    Move-Item -LiteralPath $f.Main -Destination $renamed; $f.Main=$renamed
    $locked=[IO.File]::Open((Join-Path $f.Install 'EVE-X-Preview-Settings.exe'),'Open','Read','Read')
    try { Reject { Install $f } 'previous.*restored' } finally { $locked.Dispose() }
    Assert ((Get-Content -LiteralPath $renamed -Raw) -eq 'old main') 'Renamed main was not restored.'
}
Test 'Interrupted replacement is recovered without a network connection or release files' {
    $f=Fixture 'interrupted'; $transaction=Join-Path $f.Install '.EVE-X-Preview-update'
    Text (Join-Path $transaction 'backup\EVE-X-Preview.exe') 'old main'
    Text $f.Main 'interrupted main'
    $hash=(Get-FileHash -LiteralPath (Join-Path $transaction 'backup\EVE-X-Preview.exe')).Hash
    $journal=@{schemaVersion=2;installation=$f.Install;mainFile='EVE-X-Preview.exe';state='installing';operations=@(@{path='EVE-X-Preview.exe';existed=$true;sha256=$hash})}
    Text (Join-Path $transaction 'journal.json') ($journal | ConvertTo-Json -Depth 8)
    Recover-EveExeUpdate -MainPath $f.Main
    Assert ((Get-Content -LiteralPath $f.Main -Raw) -eq 'old main') 'Offline recovery did not restore main.'
    Assert (-not (Test-Path -LiteralPath $transaction)) 'Recovery left its transaction.'
}
Test 'Unpublished staging is discarded without changing installed files' {
    $f=Fixture 'staging'; Text (Join-Path $f.Install '.EVE-X-Preview-update\stage\EVE-X-Preview.exe') 'partial download'
    Recover-EveExeUpdate -MainPath $f.Main
    Assert ((Get-Content -LiteralPath $f.Main -Raw) -eq 'old main') 'Staging recovery changed main.'
}
Test 'A bad backup is retained rather than copied over the installed file' {
    $f=Fixture 'bad-backup'; $transaction=Join-Path $f.Install '.EVE-X-Preview-update'
    Text (Join-Path $transaction 'backup\EVE-X-Preview.exe') 'damaged backup'
    $journal=@{schemaVersion=2;installation=$f.Install;mainFile='EVE-X-Preview.exe';state='installing';operations=@(@{path='EVE-X-Preview.exe';existed=$true;sha256=('0'*64)})}
    Text (Join-Path $transaction 'journal.json') ($journal | ConvertTo-Json -Depth 8)
    Reject { Recover-EveExeUpdate -MainPath $f.Main } 'backup|hash'
    Assert ((Get-Content -LiteralPath $f.Main -Raw) -eq 'old main') 'Damaged backup was copied.'
    Assert (Test-Path -LiteralPath $transaction) 'Failed recovery deleted backups.'
}
Test 'A recovery journal cannot overwrite user settings or escape the installation' {
    $f=Fixture 'unsafe-journal'; $transaction=Join-Path $f.Install '.EVE-X-Preview-update'
    $journal=@{schemaVersion=2;installation=$f.Install;mainFile='EVE-X-Preview.exe';state='installing';operations=@(@{path='../escape.exe';existed=$false})}
    Text (Join-Path $transaction 'journal.json') ($journal | ConvertTo-Json -Depth 8)
    Reject { Recover-EveExeUpdate -MainPath $f.Main } 'path|journal|target'
    Assert (-not (Test-Path (Join-Path $f.Root 'escape.exe'))) 'Recovery escaped.'
}
Test 'Optional settings restored during recovery remain included in the update' {
    $f=Fixture 'restored-settings'; $transaction=Join-Path $f.Install '.EVE-X-Preview-update'
    $settings=Join-Path $f.Install 'EVE-X-Preview-Settings.exe'
    Text (Join-Path $transaction 'backup\EVE-X-Preview-Settings.exe') 'old settings'
    $hash=(Get-FileHash -LiteralPath (Join-Path $transaction 'backup\EVE-X-Preview-Settings.exe')).Hash
    Remove-Item -LiteralPath $settings
    $journal=@{schemaVersion=2;installation=$f.Install;mainFile='EVE-X-Preview.exe';state='installing';operations=@(@{path='EVE-X-Preview-Settings.exe';existed=$true;sha256=$hash})}
    Text (Join-Path $transaction 'journal.json') ($journal | ConvertTo-Json -Depth 8)
    Install $f
    Assert ((Get-FileHash -LiteralPath $settings).Hash -eq (Get-FileHash -LiteralPath $settingsFixture).Hash) 'Recovered installed editor was skipped.'
}
Test 'Destination junctions are rejected without modifying their targets' {
    $f=Fixture 'junction'; $outside=Join-Path $f.Root 'outside'
    New-Item -ItemType Directory -Path $outside | Out-Null
    New-Item -ItemType Junction -Path (Join-Path $f.Install '.EVE-X-Preview-update') -Target $outside | Out-Null
    Text (Join-Path $outside 'personal.txt') 'keep'
    Reject { Install $f } 'junction|reparse'
    Assert ((Get-Content -LiteralPath (Join-Path $outside 'personal.txt') -Raw) -eq 'keep') 'Junction target was changed.'
}
if ($BundledSettingsExecutable) {
    Test 'A settings EXE truncated after its PE sections is rejected with both originals intact' {
        $f=Fixture 'truncated-bundle'
        Copy-Item -LiteralPath $MainReleaseExecutable -Destination (Join-Path $f.Release 'EVE-X-Preview.exe') -Force
        $truncated=Join-Path $f.Release 'EVE-X-Preview-Settings.exe'
        Copy-Item -LiteralPath $BundledSettingsExecutable -Destination $truncated -Force
        $stream=[IO.File]::Open($truncated,'Open','ReadWrite','None')
        $reader=[IO.BinaryReader]::new($stream)
        try {
            $stream.Position=0x3C; $header=$reader.ReadUInt32(); $stream.Position=$header+6; $sections=$reader.ReadUInt16()
            $stream.Position=$header+20; $optional=$reader.ReadUInt16(); [long]$end=0
            for ($index=0;$index -lt $sections;$index++) {
                $stream.Position=$header+24+$optional+40*$index+16
                [long]$size=$reader.ReadUInt32(); [long]$offset=$reader.ReadUInt32(); $end=[Math]::Max($end,$offset+$size)
            }
            Assert ($stream.Length -gt $end) 'Fixture does not have an appended bundle.'
            $stream.SetLength($end)
        } finally { $reader.Dispose() }
        $bundleVersion=[Diagnostics.FileVersionInfo]::GetVersionInfo($MainReleaseExecutable).FileVersion
        Reject { Invoke-EveExeUpdate -MainPath $f.Main -SourceDirectory $f.Release -Version $bundleVersion } 'probe|settings executable'
        Assert ((Get-Content -LiteralPath $f.Main -Raw) -eq 'old main') 'Truncated bundle replaced main.'
        Assert ((Get-Content -LiteralPath (Join-Path $f.Install 'EVE-X-Preview-Settings.exe') -Raw) -eq 'old settings') 'Truncated bundle replaced settings.'
    }
}
"Executable update tests: $script:passed passed, $script:failed failed."
if ($script:failed) { exit 1 }
