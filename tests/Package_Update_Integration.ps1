param([Parameter(Mandatory)][string]$PackageDirectory, [Parameter(Mandatory)][string]$ArchivePath,
    [Parameter(Mandatory)][string]$LegacyArchive, [Parameter(Mandatory)][string]$ManagedTests,
    [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if ($env:USERNAME -ne 'WDAGUtilityAccount') { throw 'Run compiled updater integration inside Windows Sandbox.' }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$source=Split-Path $PSScriptRoot -Parent
. (Join-Path $source 'src\PackageUpdate.ps1')
$fixture=Join-Path $env:TEMP ('evex-compiled-update-' + [guid]::NewGuid().ToString('N'))
$install=Join-Path $fixture 'installation with spaces'
New-Item -ItemType Directory -Path $install -Force | Out-Null
Expand-Archive -LiteralPath $LegacyArchive -DestinationPath $install
$main=Join-Path $install 'EVE-X-Preview.exe'
$oldVersion=[Diagnostics.FileVersionInfo]::GetVersionInfo($main).FileVersion
$version=[Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $PackageDirectory 'EVE-X-Preview.exe')).FileVersion
if ([version]$oldVersion -ge [version]$version) { throw 'Integration needs an actual older portable package.' }
$config=Join-Path $install 'EVE-X-Preview.json'
$defaults=[regex]::Match((Get-Content -LiteralPath (Join-Path $source 'Lib\DefaultJSON.ahk') -Raw),'(?s)\{.*\}').Value | ConvertFrom-Json
$defaults | Add-Member -NotePropertyName personalSentinel -NotePropertyValue 'preserve-me'
$defaults | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath $config -Encoding UTF8
[IO.File]::WriteAllText((Join-Path $install 'personal.txt'),'keep this file')
Copy-Item -LiteralPath (Join-Path $PackageDirectory 'EVE-X-Preview-Updater.exe') -Destination (Join-Path $install 'EVE-X-Preview-Updater.exe') -Force
$protocol=Start-Process C:\Runtime\AutoHotkey64.exe -ArgumentList '/ErrorStdOut',('"'+(Join-Path $PSScriptRoot 'Package_Updater_Protocol_Test.ahk')+'"'),('"'+$install+'"'),('"'+(Join-Path $fixture 'staged-updater.exe')+'"'),('"'+(Join-Path $OutputDirectory 'checks.log')+'"') -WorkingDirectory $source -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput (Join-Path $OutputDirectory 'protocol-stdout.log') -RedirectStandardError (Join-Path $OutputDirectory 'protocol-stderr.log')
if ($protocol.ExitCode) { throw 'Compiled updater protocol selection failed.' }

# The compiled production updater runs from the installed directory, relocates
# itself, validates a local release exactly as a download, and relaunches the app.
$updater=Start-Process (Join-Path $install 'EVE-X-Preview-Updater.exe') -ArgumentList ('"'+$main+'"'),$version,('"'+$ArchivePath+'"') -WorkingDirectory $install -WindowStyle Hidden -PassThru
$null=$updater.Handle
if (-not $updater.WaitForExit(15000) -or $updater.ExitCode -ne 0) { throw 'Installed updater did not hand off successfully.' }
$deadline=[DateTime]::UtcNow.AddMinutes(3)
while ([DateTime]::UtcNow -lt $deadline) {
    $manifest=Join-Path $install 'EVE-X-Preview.package.json'
    $live=@(Get-Process -Name 'EVE-X-Preview' -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $main })
    if ((Test-Path -LiteralPath $manifest) -and $live.Count) { break }
    Start-Sleep -Milliseconds 100
}
if (-not $live.Count) { throw 'Updated application did not relaunch.' }
if ([Diagnostics.FileVersionInfo]::GetVersionInfo($main).FileVersion -ne $version) { throw 'Main executable version did not update.' }
if ([Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $install 'EVE-X-Preview-Settings.exe')).FileVersion -ne $version) { throw 'Settings executable version did not update.' }
if (-not (Test-Path (Join-Path $install 'EVE-X-Preview-Updater.exe'))) { throw 'The installed package-aware updater was removed.' }
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

# Protect both entry points from a partial package, then remove the fixture marker.
New-Item -ItemType Directory -Path (Join-Path $install '.EVE-X-Preview-update') | Out-Null
$guard=Start-Process C:\Runtime\AutoHotkey64.exe -ArgumentList '/ErrorStdOut',('"'+(Join-Path $PSScriptRoot 'Update_Startup_Guard.ahk')+'"'),('"'+$main+'"'),('"'+(Join-Path $OutputDirectory 'checks.log')+'"') -WorkingDirectory $source -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput (Join-Path $OutputDirectory 'guard-stdout.log') -RedirectStandardError (Join-Path $OutputDirectory 'guard-stderr.log')
if ($guard.ExitCode) { throw 'Compiled main startup guard failed.' }
$blocked=Start-Process (Join-Path $install 'EVE-X-Preview-Settings.exe') -ArgumentList '--installation',('"'+$install+'"') -WindowStyle Hidden -PassThru
$null=$blocked.Handle
if (-not $blocked.WaitForExit(15000) -or $blocked.ExitCode -ne 1) { throw 'Settings startup guard failed.' }
'PASS: compiled settings refused an unfinished package transaction' | Add-Content (Join-Path $OutputDirectory 'checks.log')
[pscustomobject]@{previousVersion=$oldVersion;installedVersion=$version;installation=$install;settingsPreserved=$true;runtimeBridge=$true;offlineHost=$true;startupGuards=$true} | ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'result.json')
