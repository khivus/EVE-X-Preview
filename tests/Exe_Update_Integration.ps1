param([Parameter(Mandatory)][string]$ReleaseDirectory,
    [Parameter(Mandatory)][string]$LegacyArchive, [Parameter(Mandatory)][string]$ManagedTests,
    [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if ($env:USERNAME -ne 'WDAGUtilityAccount') { throw 'Run compiled updater integration inside Windows Sandbox.' }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
& (Join-Path $PSScriptRoot 'Update_Startup_Guard_Test.ps1') -OutputDirectory (Join-Path $OutputDirectory 'startup-guard-regressions')
$source=Split-Path $PSScriptRoot -Parent
$fixture=Join-Path $env:TEMP ('evex-compiled-update-' + [guid]::NewGuid().ToString('N'))
$install=Join-Path $fixture 'installation with spaces'
New-Item -ItemType Directory -Path $install -Force | Out-Null
Expand-Archive -LiteralPath $LegacyArchive -DestinationPath $install
$main=Join-Path $install 'EVE-X-Preview.exe'
$legacyMain=Join-Path $fixture 'legacy-main.exe'
Copy-Item -LiteralPath $main -Destination $legacyMain
$oldVersion=[Diagnostics.FileVersionInfo]::GetVersionInfo($main).FileVersion
$version=[Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $ReleaseDirectory 'EVE-X-Preview.exe')).FileVersion
if ([version]$oldVersion -ge [version]$version) { throw 'Integration needs an actual older portable package.' }
$config=Join-Path $install 'EVE-X-Preview.json'
$defaults=[regex]::Match((Get-Content -LiteralPath (Join-Path $source 'Lib\DefaultJSON.ahk') -Raw),'(?s)\{.*\}').Value | ConvertFrom-Json
$defaults | Add-Member -NotePropertyName personalSentinel -NotePropertyValue 'preserve-me'
$defaults | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath $config -Encoding UTF8
[IO.File]::WriteAllText((Join-Path $install 'personal.txt'),'keep this file')
Copy-Item -LiteralPath (Join-Path $ReleaseDirectory 'EVE-X-Preview-Updater.exe') -Destination (Join-Path $install 'EVE-X-Preview-Updater.exe') -Force
$protocol=Start-Process C:\Runtime\AutoHotkey64.exe -ArgumentList '/ErrorStdOut',('"'+(Join-Path $PSScriptRoot 'Updater_Protocol_Test.ahk')+'"'),('"'+$install+'"'),('"'+(Join-Path $fixture 'staged-updater.exe')+'"'),('"'+(Join-Path $OutputDirectory 'checks.log')+'"') -WorkingDirectory $source -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput (Join-Path $OutputDirectory 'protocol-stdout.log') -RedirectStandardError (Join-Path $OutputDirectory 'protocol-stderr.log')
if ($protocol.ExitCode) { throw 'Compiled updater protocol selection failed.' }

# The compiled production updater runs from the installed directory, relocates
# itself, validates a local release exactly as a download, and relaunches the app.
$updater=Start-Process (Join-Path $install 'EVE-X-Preview-Updater.exe') -ArgumentList ('"'+$main+'"'),$version,('"'+$ReleaseDirectory+'"') -WorkingDirectory $install -WindowStyle Hidden -PassThru
$null=$updater.Handle
if (-not $updater.WaitForExit(15000) -or $updater.ExitCode -ne 0) { throw 'Installed updater did not hand off successfully.' }
$deadline=[DateTime]::UtcNow.AddMinutes(3)
while ([DateTime]::UtcNow -lt $deadline) {
    $live=@(Get-Process -Name 'EVE-X-Preview' -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $main })
    if ($live.Count -and [Diagnostics.FileVersionInfo]::GetVersionInfo($main).FileVersion -eq $version) { break }
    Start-Sleep -Milliseconds 100
}
if (-not $live.Count) { throw 'Updated application did not relaunch.' }
if ([Diagnostics.FileVersionInfo]::GetVersionInfo($main).FileVersion -ne $version) { throw 'Main executable version did not update.' }
if ([Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $install 'EVE-X-Preview-Settings.exe')).FileVersion -ne $version) { throw 'Settings executable version did not update.' }
if (-not (Test-Path (Join-Path $install 'EVE-X-Preview-Updater.exe'))) { throw 'The installed recovery updater was removed.' }
$saved=Get-Content -LiteralPath $config -Raw | ConvertFrom-Json
if ($saved.personalSentinel -ne 'preserve-me') { throw 'Personal settings were lost.' }
if ((Get-Content -LiteralPath (Join-Path $install 'personal.txt') -Raw) -ne 'keep this file') { throw 'Personal file was lost.' }
"PASS: compiled updater upgraded $oldVersion to $version and relaunched main; personal data preserved" | Add-Content (Join-Path $OutputDirectory 'checks.log')

try {
    $p=Start-Process $ManagedTests -ArgumentList '--integration',('"'+$install+'"') -WorkingDirectory $install -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput (Join-Path $OutputDirectory 'runtime-integration.log') -RedirectStandardError (Join-Path $OutputDirectory 'runtime-integration-error.log')
    if ($p.ExitCode) { throw 'Updated compiled runtime bridge integration failed.' }
} finally {
    # Only disposable guest fixtures are eligible for test cleanup.
    Get-Process | Where-Object { try { $_.Path -eq $main } catch { $false } } | Stop-Process -Force
}
Start-Sleep -Seconds 1
$editor=Start-Process (Join-Path $install 'EVE-X-Preview-Settings.exe') -ArgumentList '--installation',('"'+$install+'"') -WorkingDirectory $install -WindowStyle Hidden -PassThru
$null=$editor.Handle
try {
    $deadline=[DateTime]::UtcNow.AddSeconds(20)
    do {
        $hosts=@(Get-CimInstance Win32_Process -Filter "Name='EVE-X-Preview-SettingsHost.exe'" | Where-Object { $_.CommandLine -match ('--parent\s+'+$editor.Id+'\b') })
        if ($hosts.Count) { break }
        if ($editor.HasExited -or [DateTime]::UtcNow -ge $deadline) { throw 'Production settings editor did not start its embedded host.' }
        Start-Sleep -Milliseconds 100
    } while ($true)
    if ([Diagnostics.FileVersionInfo]::GetVersionInfo($hosts[0].ExecutablePath).FileVersion -ne $version) { throw 'Production editor started an outdated host.' }
    Start-Sleep -Milliseconds 500
    $p=Start-Process $ManagedTests -ArgumentList '--inspect-host',('"'+$install+'"') -WorkingDirectory $install -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput (Join-Path $OutputDirectory 'offline-integration.log') -RedirectStandardError (Join-Path $OutputDirectory 'offline-integration-error.log')
    if ($p.ExitCode) { throw 'Updated production settings executable/embedded host inspection failed.' }
    if (-not $editor.CloseMainWindow() -or -not $editor.WaitForExit(15000) -or $editor.ExitCode -ne 0) { throw 'Production settings editor did not close cleanly.' }
} finally { if (-not $editor.HasExited) { Stop-Process -Id $editor.Id -Force } }

# Protect both entry points from a partial update, then exercise compiled offline recovery.
New-Item -ItemType Directory -Path (Join-Path $install '.EVE-X-Preview-update') | Out-Null
$guard=Start-Process C:\Runtime\AutoHotkey64.exe -ArgumentList '/ErrorStdOut',('"'+(Join-Path $PSScriptRoot 'Update_Startup_Guard.ahk')+'"'),('"'+$main+'"'),('"'+(Join-Path $OutputDirectory 'checks.log')+'"') -WorkingDirectory $source -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $OutputDirectory 'guard-stdout.log') -RedirectStandardError (Join-Path $OutputDirectory 'guard-stderr.log')
$null=$guard.Handle
try {
    if (-not $guard.WaitForExit(25000)) { throw 'Compiled main startup guard timed out.' }
    if ($guard.ExitCode) { throw 'Compiled main startup guard failed.' }
} finally {
    if (-not $guard.HasExited) { Stop-Process -Id $guard.Id -Force }
    Get-Process | Where-Object { try { $_.Path -eq $main } catch { $false } } | Stop-Process -Force
}
$blocked=Start-Process (Join-Path $install 'EVE-X-Preview-Settings.exe') -ArgumentList '--installation',('"'+$install+'"') -WindowStyle Hidden -PassThru
$null=$blocked.Handle
if (-not $blocked.WaitForExit(15000) -or $blocked.ExitCode -ne 1) { throw 'Settings startup guard failed.' }
'PASS: compiled settings refused an unfinished package transaction' | Add-Content (Join-Path $OutputDirectory 'checks.log')
. (Join-Path $PSScriptRoot 'Exe_Update_Fixtures.ps1')
New-EveInterruptedFixture $main
$recover=Start-Process (Join-Path $install 'EVE-X-Preview-Updater.exe') -ArgumentList '--recover',('"'+$main+'"') -WorkingDirectory $install -WindowStyle Hidden -PassThru
$null=$recover.Handle
if (-not $recover.WaitForExit(15000) -or $recover.ExitCode) { throw 'Offline recovery updater did not hand off.' }
$deadline=[DateTime]::UtcNow.AddSeconds(60)
do {
    $live=@(Get-Process -Name 'EVE-X-Preview' -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $main })
    if ($live.Count -and -not (Test-Path (Join-Path $install '.EVE-X-Preview-update')) -and [Diagnostics.FileVersionInfo]::GetVersionInfo($main).FileVersion -eq $version) { break }
    if ([DateTime]::UtcNow -ge $deadline) { throw 'Compiled offline recovery did not restore and relaunch main.' }
    Start-Sleep -Milliseconds 100
} while ($true)
$live | Stop-Process -Force
'PASS: compiled --recover restored interrupted main and relaunched without network or release files' | Add-Content (Join-Path $OutputDirectory 'checks.log')

# A built-in-only installation has no settings asset available even in its source.
$builtin=Join-Path $fixture 'built-in only'; $builtinSource=Join-Path $fixture 'main-only source'
New-Item -ItemType Directory -Path $builtin,$builtinSource | Out-Null
$builtinMain=Join-Path $builtin 'EVE-X-Preview.exe'
Copy-Item -LiteralPath $legacyMain -Destination $builtinMain
Copy-Item -LiteralPath (Join-Path $ReleaseDirectory 'EVE-X-Preview.exe') -Destination $builtinSource
Copy-Item -LiteralPath (Join-Path $ReleaseDirectory 'EVE-X-Preview-Updater.exe') -Destination $builtin
$defaults | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath (Join-Path $builtin 'EVE-X-Preview.json') -Encoding UTF8
$updater=Start-Process (Join-Path $builtin 'EVE-X-Preview-Updater.exe') -ArgumentList ('"'+$builtinMain+'"'),$version,('"'+$builtinSource+'"') -WorkingDirectory $builtin -WindowStyle Hidden -PassThru
$null=$updater.Handle
if (-not $updater.WaitForExit(15000) -or $updater.ExitCode) { throw 'Built-in-only updater did not hand off.' }
$deadline=[DateTime]::UtcNow.AddSeconds(60)
do {
    $live=@(Get-Process -Name 'EVE-X-Preview' -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $builtinMain })
    if ($live.Count -and [Diagnostics.FileVersionInfo]::GetVersionInfo($builtinMain).FileVersion -eq $version) { break }
    if ([DateTime]::UtcNow -ge $deadline) { throw 'Built-in-only update did not complete.' }
    Start-Sleep -Milliseconds 100
} while ($true)
$live | Stop-Process -Force
if (Test-Path (Join-Path $builtin 'EVE-X-Preview-Settings.exe')) { throw 'Optional settings was installed implicitly.' }
'PASS: compiled built-in-only update used main EXE alone and left optional settings absent' | Add-Content (Join-Path $OutputDirectory 'checks.log')
[pscustomobject]@{previousVersion=$oldVersion;installedVersion=$version;installation=$install;settingsPreserved=$true;runtimeBridge=$true;offlineHost=$true;startupGuards=$true;offlineRecovery=$true;settingsOptional=$true} | ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'result.json')
