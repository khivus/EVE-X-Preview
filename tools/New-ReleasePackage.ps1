[CmdletBinding()]
param([Parameter(Mandatory)][string]$Directory, [Parameter(Mandatory)][string]$Version,
    [Parameter(Mandatory)][string]$ArchivePath)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'src\PackageUpdate.ps1')
New-EveReleasePackage -Directory $Directory -Version $Version -ArchivePath $ArchivePath
