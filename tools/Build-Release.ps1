[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputDirectory,
    [string]$AutoHotkeyDirectory = 'C:\Program Files\AutoHotkey'
)

$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$compiler = Join-Path $AutoHotkeyDirectory 'Compiler\Ahk2Exe.exe'
$runtime = Join-Path $AutoHotkeyDirectory 'v2\AutoHotkey64.exe'
foreach ($required in @($compiler, $runtime)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "Missing build tool: $required"
    }
}
$destination = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$packages = @(
    @{ Source = 'Main.ahk'; Name = 'EVE-X-Preview.exe' },
    @{ Source = 'Updater.ahk'; Name = 'EVE-X-Preview-Updater.exe' }
)
foreach ($package in $packages) {
    $source = Join-Path $repository $package.Source
    $output = Join-Path $destination $package.Name
    $match = [regex]::Match([IO.File]::ReadAllText($source), 'Ahk2Exe-Let U_version = ([\d.]+)')
    if (-not $match.Success) { throw "Missing release version in $source" }
    $expectedVersion = [version]$match.Groups[1].Value
    $arguments = '/in "{0}" /out "{1}" /base "{2}" /silent' -f $source, $output, $runtime
    $process = Start-Process -FilePath $compiler -ArgumentList $arguments -WorkingDirectory $repository -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $output -PathType Leaf)) {
        throw "Build failed for $($package.Name); exit code $($process.ExitCode)"
    }
    $info = [Diagnostics.FileVersionInfo]::GetVersionInfo($output)
    $actualVersion = [version]$info.FileVersion
    if ($actualVersion.Major -ne $expectedVersion.Major -or $actualVersion.Minor -ne $expectedVersion.Minor -or
        [Math]::Max(0, $actualVersion.Build) -ne [Math]::Max(0, $expectedVersion.Build) -or
        [Math]::Max(0, $actualVersion.Revision) -ne [Math]::Max(0, $expectedVersion.Revision)) {
        throw "Version mismatch for $($package.Name): $actualVersion instead of $expectedVersion"
    }
    [pscustomobject]@{ Package = $package.Name; Version = $info.FileVersion; Bytes = (Get-Item -LiteralPath $output).Length }
}
