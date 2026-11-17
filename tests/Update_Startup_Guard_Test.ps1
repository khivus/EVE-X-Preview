param([Parameter(Mandatory)][string]$OutputDirectory,
    [string]$AutoHotkeyDirectory = 'C:\Input\ahk-tools')
$ErrorActionPreference = 'Stop'
if ($env:USERNAME -ne 'WDAGUtilityAccount') { throw 'Run startup guard regressions inside Windows Sandbox.' }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$fixtureDirectory = Join-Path $env:TEMP ('evex-guard-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureDirectory | Out-Null
$fixture = Join-Path $fixtureDirectory 'guard fixture.exe'
$compiler = Start-Process (Join-Path $AutoHotkeyDirectory 'Compiler\Ahk2Exe.exe') -ArgumentList '/in',('"'+(Join-Path $PSScriptRoot 'fixtures\Update_Startup_Guard_Fixture.ahk')+'"'),'/out',('"'+$fixture+'"'),'/base',('"'+(Join-Path $AutoHotkeyDirectory 'v2\AutoHotkey64.exe')+'"'),'/silent' -WindowStyle Hidden -PassThru
$null = $compiler.Handle
try {
    if (-not $compiler.WaitForExit(60000) -or $compiler.ExitCode -ne 0 -or -not (Test-Path $fixture)) { throw 'Could not compile startup guard fixture.' }
} finally {
    if (-not $compiler.HasExited) {
        Stop-Process -Id $compiler.Id -Force
        if (-not $compiler.WaitForExit(5000)) { throw 'Startup guard fixture compiler did not stop.' }
    }
}
$results = @()
try {
    foreach ($case in @(
        @{mode='delayed-ready'; exitCode=0; message='PASS:'},
        @{mode='wrong-exit'; exitCode=1; message='unexpected exit code'},
        @{mode='hang-after-ack'; exitCode=1; message='did not exit'},
        @{mode='never-ready'; exitCode=1; message='not ready'},
        @{mode='no-notice'; exitCode=1; message='did not show'}
    )) {
        $name = $case.mode
        $env:EVEX_GUARD_FIXTURE_MODE = $name
        $env:EVEX_GUARD_FIXTURE_PID = Join-Path $OutputDirectory ($name + '.pid')
        $log = Join-Path $OutputDirectory ($name + '.log')
        Remove-Item -LiteralPath $log,$env:EVEX_GUARD_FIXTURE_PID -ErrorAction SilentlyContinue
        $helper = Start-Process C:\Runtime\AutoHotkey64.exe -ArgumentList '/ErrorStdOut',('"'+(Join-Path $PSScriptRoot 'Update_Startup_Guard.ahk')+'"'),('"'+$fixture+'"'),('"'+$log+'"') -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $OutputDirectory ($name+'-stdout.log')) -RedirectStandardError (Join-Path $OutputDirectory ($name+'-stderr.log'))
        $null = $helper.Handle
        $finished = $helper.WaitForExit(25000)
        if (-not $finished) { Stop-Process -Id $helper.Id -Force }
        $leftovers = @(Get-Process | Where-Object { try { $_.Path -eq $fixture } catch { $false } })
        $text = if (Test-Path $log) { Get-Content $log -Raw } else { '' }
        $fixturePid = 0
        $started = (Test-Path $env:EVEX_GUARD_FIXTURE_PID) -and [int]::TryParse((Get-Content $env:EVEX_GUARD_FIXTURE_PID -Raw), [ref]$fixturePid) -and $fixturePid -gt 0
        $passed = $started -and $finished -and $helper.ExitCode -eq $case.exitCode -and $text.Contains($case.message) -and $leftovers.Count -eq 0
        $results += [pscustomobject]@{case=$name; passed=$passed; exitCode=$helper.ExitCode; leftoverProcesses=$leftovers.Count}
        $leftovers | Stop-Process -Force
    }
} finally {
    Remove-Item Env:\EVEX_GUARD_FIXTURE_MODE, Env:\EVEX_GUARD_FIXTURE_PID -ErrorAction SilentlyContinue
    Get-Process | Where-Object { try { $_.Path -eq $fixture } catch { $false } } | Stop-Process -Force
}
$results | ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'results.json')
if (@($results | Where-Object { -not $_.passed }).Count) { throw 'Startup guard regressions failed; see results.json.' }
