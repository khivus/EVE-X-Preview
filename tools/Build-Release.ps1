[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputDirectory,
    [string]$AutoHotkeyDirectory = 'C:\Program Files\AutoHotkey',
    [switch]$BuiltInOnly,
    [string]$DotNetPath = 'dotnet',
    [switch]$NoRestore
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
if ((Test-Path -LiteralPath $destination) -and @(Get-ChildItem -LiteralPath $destination -Force).Count -ne 0) {
    throw 'Choose an empty release output directory so stale companion files cannot enter the package.'
}
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$packages = @(
    @{ Source = 'Main.ahk'; Name = 'EVE-X-Preview.exe' },
    @{ Source = 'Updater.ahk'; Name = 'EVE-X-Preview-Updater.exe' }
)
if (-not $BuiltInOnly) {
    $packages += @{ Source = 'SettingsHost.ahk'; Name = 'EVE-X-Preview-SettingsHost.exe' }
}
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

if (-not $BuiltInOnly) {
    # Keep SDK first-use and package caches local to the build workspace.
    $cache = Join-Path $repository '.vs\build'
    if (-not $env:DOTNET_CLI_HOME) { $env:DOTNET_CLI_HOME = Join-Path $cache 'home' }
    if (-not $env:NUGET_PACKAGES) { $env:NUGET_PACKAGES = Join-Path $cache 'packages' }
    if (-not $env:NUGET_HTTP_CACHE_PATH) { $env:NUGET_HTTP_CACHE_PATH = Join-Path $cache 'http-cache' }
    $env:DOTNET_GENERATE_ASPNET_CERTIFICATE = 'false'
    $env:DOTNET_ADD_GLOBAL_TOOLS_TO_PATH = '0'
    $env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
    $env:AVALONIA_TELEMETRY_OPTOUT = '1'
    $project = Join-Path $repository 'settings\EVE-X-Preview.Settings\EVE-X-Preview.Settings.csproj'
    $publish = @('publish', $project, '-c', 'Release', '-r', 'win-x64', '--self-contained', 'true',
        '-p:PublishTrimmed=false', '-p:UseSharedCompilation=false',
        ('-p:SettingsHostPath=' + (Join-Path $destination 'EVE-X-Preview-SettingsHost.exe')), '-o', $destination)
    if ($NoRestore) { $publish += '--no-restore' }
    & $DotNetPath @publish
    if ($LASTEXITCODE -ne 0) { throw "Settings companion publish failed: $LASTEXITCODE" }
    # Native NuGet packages can publish large PDB files as runtime assets.
    # Keep the portable distribution to executable/runtime dependencies.
    $destinationPrefix = $destination.TrimEnd('\') + '\'
    foreach ($symbol in Get-ChildItem -LiteralPath $destination -Filter '*.pdb' -File -Recurse) {
        if (-not $symbol.FullName.StartsWith($destinationPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Unexpected symbol path outside package directory: $($symbol.FullName)"
        }
        Remove-Item -LiteralPath $symbol.FullName
    }
    if (-not (Test-Path -LiteralPath (Join-Path $destination 'EVE-X-Preview-Settings.exe') -PathType Leaf)) {
        throw 'Single-file settings publish did not create its executable.'
    }
    # The compiled host is embedded in the settings executable, not a distribution dependency.
    $hostCache = Join-Path $cache ('hosts\' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $hostCache -Force | Out-Null
    Move-Item -LiteralPath (Join-Path $destination 'EVE-X-Preview-SettingsHost.exe') -Destination (Join-Path $hostCache 'EVE-X-Preview-SettingsHost.exe')
    $looseDependencies = @(Get-ChildItem -LiteralPath $destination -File | Where-Object { $_.Extension -ne '.exe' })
    if ($looseDependencies.Count) { throw ('Unexpected loose dependencies in single-file package: ' + ($looseDependencies.Name -join ', ')) }
    $assets = Get-ChildItem -LiteralPath $destination -File -Recurse
    [pscustomobject]@{
        Package = 'EVE-X-Preview-Settings.exe (single file)'
        Version = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $destination 'EVE-X-Preview-Settings.exe')).FileVersion
        Bytes = ($assets | Measure-Object -Property Length -Sum).Sum
    }
}
