[CmdletBinding()]
param(
    [Parameter(Mandatory)][guid]$SandboxId,
    [ValidatePattern('^[a-z0-9-]+$')][string]$Label = ('check-' + (Get-Date -Format 'yyyyMMdd-HHmmss')),
    [ValidateSet('Ahk', 'Managed', 'Integration', 'RuntimeIntegration', 'HandoffIntegration')][string]$Mode = 'Ahk',
    [string]$Filter = ''
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$inputDirectory = Join-Path $root '.vs\isolated\input'
$cli = Get-Command wsb.exe -ErrorAction Stop
if (!(Test-Path $inputDirectory)) { throw 'Prepare the local Windows Sandbox mapped input folder first.' }
if ($Filter -match '["\r\n]') { throw 'Filter cannot contain quotes or line breaks.' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$checks = Join-Path $inputDirectory 'checks'
New-Item -ItemType Directory -Force $checks | Out-Null
$archivePath = Join-Path $checks ($Label + '.zip')
if (Test-Path $archivePath) { throw "Check label already exists: $Label" }
$paths = & git -C $root -c core.quotepath=false ls-files --cached --others --exclude-standard
if ($LASTEXITCODE -ne 0) { throw 'Could not list the source files.' }
$archive = [IO.Compression.ZipFile]::Open($archivePath, 'Create')
try {
    foreach ($path in $paths) {
        $full = Join-Path $root $path
        if (Test-Path -LiteralPath $full -PathType Leaf) {
            [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $full, $path, 'Optimal') | Out-Null
        }
    }
} finally { $archive.Dispose() }
# Only this guest bootstrap invokes test executables. The host packages source
# and requests execution through the official Sandbox CLI.
$bootstrap = @'
param([string]$Label, [string]$Mode, [string]$Filter = '')
$ErrorActionPreference = 'Stop'
if ($env:USERNAME -ne 'WDAGUtilityAccount' -or $Label -notmatch '^[a-z0-9-]+$') { throw 'Configured Sandbox guest only.' }
$out = 'C:\Output\' + $Label
$source = 'C:\Checks\' + $Label
New-Item -ItemType Directory -Path $out | Out-Null
try {
    Expand-Archive -LiteralPath ('C:\Input\checks\' + $Label + '.zip') -DestinationPath $source
    if ($Mode -eq 'Ahk') {
        $runner = Join-Path $source 'tests\TestRunner.ahk'
        [IO.File]::WriteAllText($runner, "#Warn All, StdOut`r`n" + [IO.File]::ReadAllText($runner), [Text.UTF8Encoding]::new($true))
        Remove-Item (Join-Path $source 'tests\test_results.log') -ErrorAction SilentlyContinue
        $exe = 'C:\Runtime\AutoHotkey64.exe'
        $arguments = @('/ErrorStdOut', ('"' + $runner + '"'), '--headless')
        if ($Filter) { $arguments += @('--filter', ('"' + $Filter + '"')) }
    } else {
        $exe = 'C:\Input\settings-tests\EVE-X-Preview.Settings.Tests.exe'
        $arguments = @()
        if ($Mode -eq 'Integration' -or $Mode -eq 'RuntimeIntegration' -or $Mode -eq 'HandoffIntegration') {
            Copy-Item 'C:\Runtime\AutoHotkey64.exe' (Join-Path $source 'AutoHotkey64.exe')
            $arguments = @('--integration', ('"' + $source + '"'))
            if ($Mode -eq 'HandoffIntegration') { $arguments = @('--handoff-integration', ('"' + $source + '"')) }
            if ($Mode -eq 'RuntimeIntegration') {
                $main = Join-Path $source 'Main.ahk'
                $runtime = Start-Process 'C:\Runtime\AutoHotkey64.exe' -ArgumentList '/ErrorStdOut',('"' + $main + '"') -WorkingDirectory $source -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $out 'runtime-stdout.log') -RedirectStandardError (Join-Path $out 'runtime-stderr.log')
                Start-Sleep -Seconds 2
            }
        }
    }
    'running' | Set-Content (Join-Path $out 'status.txt')
    $startArguments = @{FilePath=$exe; WorkingDirectory=$source; WindowStyle='Hidden'; PassThru=$true; RedirectStandardOutput=(Join-Path $out 'stdout.log'); RedirectStandardError=(Join-Path $out 'stderr.log')}
    if ($arguments.Count) { $startArguments.ArgumentList = $arguments }
    $process = Start-Process @startArguments
    $null = $process.Handle
    $deadline = (Get-Date).AddMinutes(12)
    $log = Join-Path $source 'tests\test_results.log'
    while (!$process.WaitForExit(3000)) {
        if (Test-Path $log) { Copy-Item $log (Join-Path $out 'test_results.log') -Force -ErrorAction SilentlyContinue }
        if ((Get-Date) -gt $deadline) { Stop-Process -Id $process.Id -Force; throw 'Test timeout after 12 minutes.' }
    }
    if (Test-Path $log) { Copy-Item $log (Join-Path $out 'test_results.log') -Force }
    if ($null -eq $process.ExitCode) { throw 'Test process exit code was not retained.' }
    $process.ExitCode | Set-Content (Join-Path $out 'exit-code.txt')
    'finished' | Set-Content (Join-Path $out 'status.txt')
    exit $process.ExitCode
} catch {
    $_ | Out-String | Set-Content (Join-Path $out 'bootstrap-error.log')
    2 | Set-Content (Join-Path $out 'exit-code.txt')
    exit 2
} finally {
    if ($Mode -eq 'RuntimeIntegration') {
        # Reload creates a new PID. Match only this fresh guest fixture's script
        # path, never other AHK instances or the user's baseline Sandbox app.
        Get-CimInstance Win32_Process -Filter "Name = 'AutoHotkey64.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine.Contains($source + '\Main.ahk') } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    }
}
'@
$bootstrap | Set-Content -LiteralPath (Join-Path $checks 'Run-Checks.ps1') -Encoding UTF8
$command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Input\checks\Run-Checks.ps1 -Label ' + $Label + ' -Mode ' + $Mode
if ($Filter) { $command += ' -Filter "' + $Filter + '"' }
Write-Output "Guest source: C:\Checks\$Label; guest results: C:\Output\$Label"
$resultText = & $cli.Source exec --id $SandboxId.ToString() --command $command --run-as ExistingLogin --raw
if ($LASTEXITCODE -ne 0) { throw "Sandbox CLI failed: $LASTEXITCODE" }
$resultText | Write-Output
$guestResult = ($resultText -join "`n") | ConvertFrom-Json
if ($guestResult.ExitCode -ne 0) { throw "Isolated checks failed with guest exit code $($guestResult.ExitCode); see C:\Output\$Label." }
