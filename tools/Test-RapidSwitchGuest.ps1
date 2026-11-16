param(
    [Parameter(Mandatory=$true)][string]$Runtime,
    [Parameter(Mandatory=$true)][string]$Compiler,
    [ValidatePattern('^[a-z0-9-]+$')][string]$Label='rapid-switch',
    [ValidateRange(12,32)][int]$Clients=16
)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
if($env:USERNAME -ne 'WDAGUtilityAccount' -or !$root.StartsWith('C:\Checks\',[StringComparison]::OrdinalIgnoreCase)){
    throw 'Run only from a disposable C:\Checks checkout inside Windows Sandbox.'
}
if((Get-Service Audiosrv).Status -ne 'Stopped'){
    throw 'Stop and disable Audiosrv inside the Sandbox before running this test.'
}
$out='C:\Output\'+$Label
if(Test-Path $out){throw 'Use a fresh result label.'}
New-Item -ItemType Directory $out|Out-Null
$process=$null
try {
    # Synthetic exefile.exe clients trigger Game Bar discovery in fresh guests.
    # Suppress its missing-protocol dialog only in this disposable guest account.
    foreach($setting in @(
        @('HKCU:\System\GameConfigStore','GameDVR_Enabled'),
        @('HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR','AppCaptureEnabled'),
        @('HKCU:\Software\Microsoft\GameBar','ShowStartupPanel'))){
        New-Item $setting[0] -Force|Out-Null
        New-ItemProperty $setting[0] -Name $setting[1] -Value 0 -PropertyType DWord -Force|Out-Null
    }
    $process=Start-Process $Compiler -ArgumentList '/in',('"'+$root+'\Main.ahk"'),'/out',('"'+$root+'\EVE-X-Preview.exe"'),'/base',('"'+$Runtime+'"') -WindowStyle Hidden -PassThru
    $null=$process.Handle
    if(!$process.WaitForExit(120000)){throw 'Compiler timeout'}
    if($process.ExitCode -ne 0){throw 'Compilation failed'}
    Copy-Item $Runtime (Join-Path $root 'exefile.exe')
    $process=Start-Process $Runtime -ArgumentList '/ErrorStdOut',('"'+$root+'\tests\Rapid_Switch_Acceptance.ahk"'),$out,$Clients -WindowStyle Hidden -PassThru -RedirectStandardOutput "$out\stdout.txt" -RedirectStandardError "$out\stderr.txt"
    $null=$process.Handle
    if(!$process.WaitForExit(600000)){throw 'Acceptance timeout'}
    if($process.ExitCode -ne 0){throw "Acceptance exited $($process.ExitCode)"}
    $ErrorActionPreference='Continue'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$PSScriptRoot\Measure-RapidSwitch.ps1" -Directory $out -AssertAll *> "$out\analysis.txt"
    $result=$LASTEXITCODE
    $ErrorActionPreference='Stop'
    $result|Set-Content "$out\exit-code.txt"
    exit $result
} catch {
    $_|Out-String|Set-Content "$out\error.log"
    2|Set-Content "$out\exit-code.txt"
    exit 2
} finally {
    if($process -and !$process.HasExited){Stop-Process -Id $process.Id -Force}
    # A forced driver timeout does not invoke its OnExit handler. Match only
    # this guest checkout, never unrelated applications or another test copy.
    Get-CimInstance Win32_Process | Where-Object {
        $_.ExecutablePath -and ([IO.Path]::GetFullPath($_.ExecutablePath).StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase))
    } | ForEach-Object {Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue}
}
