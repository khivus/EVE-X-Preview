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
$audio=Get-Service Audiosrv
if($audio.Status -ne 'Stopped' -or $audio.StartType -ne 'Disabled'){
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
    # Compile the receiver in the guest with the inbox .NET Framework compiler.
    # It observes the native message queue without AHK OnMessage thread limits.
    $receiverCompiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    & $receiverCompiler /nologo /target:winexe /platform:x64 /reference:System.Windows.Forms.dll /reference:System.Drawing.dll "/out:$root\exefile.exe" "$root\tests\fixtures\RapidSwitchClient.cs" *> "$out\receiver-compile.txt"
    if($LASTEXITCODE -ne 0){throw 'Receiver compilation failed'}
    $process=Start-Process $Runtime -ArgumentList '/ErrorStdOut',('"'+$root+'\tests\Receiver_Input_Acceptance.ahk"'),("$out\receiver") -WindowStyle Hidden -PassThru -RedirectStandardOutput "$out\receiver-stdout.txt" -RedirectStandardError "$out\receiver-stderr.txt"
    $null=$process.Handle
    if(!$process.WaitForExit(30000)){throw 'Receiver regression timeout'}
    if($process.ExitCode -ne 0){throw 'Receiver regression driver failed'}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$PSScriptRoot\Measure-RapidSwitch.ps1" -Directory "$out\receiver" -AssertCase receiver-slow-log *> "$out\receiver-analysis.txt"
    if($LASTEXITCODE -ne 0){throw 'Receiver lost input while logging was delayed'}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$PSScriptRoot\Measure-RapidSwitch.ps1" -Directory "$out\receiver" -AssertCase receiver-modifier-release *> "$out\modifier-release-analysis.txt"
    if($LASTEXITCODE -ne 1){throw 'Modifier release negative control was not rejected'}
    $release=@(Import-Csv "$out\receiver\delivery.csv" | Where-Object Case -eq 'receiver-modifier-release')
    if($release.Count -ne 1 -or $release[0].MissingShift -ne 'True' -or
        $release[0].Wrong -ne 'False' -or $release[0].Missing -ne 'False' -or
        $release[0].MissingUp -ne 'False' -or $release[0].Split -ne 'False'){
        throw 'Negative control must fail only because Shift was released'
    }
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
