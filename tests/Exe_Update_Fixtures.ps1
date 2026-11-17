param([ValidateSet('probe','interruption')][string]$Mode,[string]$Path,[string]$FixtureVersion,[string]$BaseExecutable)
$ErrorActionPreference='Stop'
function New-EveProbeFixture([string]$Path,[string]$Version,[string]$BaseExecutable) {
    $compiler=@('C:\Input\ahk-tools\Compiler\Ahk2Exe.exe',(Join-Path $env:ProgramFiles 'AutoHotkey\Compiler\Ahk2Exe.exe')) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $compiler) { throw 'Ahk2Exe is required for the executable update fixtures.' }
    $source=@'
#Requires AutoHotkey v2.0
#NoTrayIcon
;@Ahk2Exe-SetVersion __VERSION__
;@Ahk2Exe-SetFileVersion __VERSION__
if A_Args.Length = 2 && A_Args[1] = "--probe" {
    FileAppend('{"application":"EVE-X-Preview-Settings","bridgeVersion":1,"schemaVersion":1}', A_Args[2], "UTF-8-RAW")
    ExitApp(0)
}
ExitApp(2)
'@
    [IO.File]::WriteAllText($Path+'.ahk',$source.Replace('__VERSION__',$Version),[Text.UTF8Encoding]::new($true))
    $arguments='/in "{0}" /out "{1}" /base "{2}" /silent' -f ($Path+'.ahk'),$Path,$BaseExecutable
    $process=Start-Process $compiler -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $null=$process.Handle
    try {
        if (-not $process.WaitForExit(60000)) { throw 'Settings probe fixture compilation timed out after 60 seconds.' }
        if ($process.ExitCode -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'Settings probe fixture compilation failed.' }
    } finally {
        if (-not $process.HasExited) {
            Stop-Process -Id $process.Id -Force
            if (-not $process.WaitForExit(5000)) { throw 'Settings probe fixture compiler did not stop.' }
        }
    }
}
function New-EveInterruptedFixture([string]$MainPath) {
    $MainPath=[IO.Path]::GetFullPath($MainPath)
    $directory=Split-Path -Parent $MainPath; $name=Split-Path -Leaf $MainPath
    $transaction=Join-Path $directory '.EVE-X-Preview-update'
    New-Item -ItemType Directory -Path (Join-Path $transaction 'backup') -Force | Out-Null
    $backup=Join-Path (Join-Path $transaction 'backup') $name
    Copy-Item -LiteralPath $MainPath -Destination $backup
    $journal=@{schemaVersion=2;installation=$directory;mainFile=$name;state='installing';operations=@(@{path=$name;existed=$true;sha256=(Get-FileHash -LiteralPath $backup).Hash})}
    $journal | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $transaction 'journal.json') -Encoding UTF8
    [IO.File]::WriteAllText($MainPath,'interrupted replacement')
}
if ($Mode -eq 'probe') { New-EveProbeFixture $Path $FixtureVersion $BaseExecutable }
elseif ($Mode -eq 'interruption') { New-EveInterruptedFixture $Path }
