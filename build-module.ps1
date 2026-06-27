param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'dist')
)

$ErrorActionPreference = 'Stop'

$version = '3.1.0'
$moduleRoot = Join-Path $OutputPath 'DefenderShield'
$packageRoot = Join-Path $moduleRoot $version

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Recurse -Force
}

New-Item -Path $packageRoot -ItemType Directory -Force | Out-Null

foreach ($file in @('DefenderShield.ps1', 'DefenderShield.psm1', 'DefenderShield.psd1', 'README.md', 'LICENSE')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $file) -Destination $packageRoot -Force
}

Test-ModuleManifest -Path (Join-Path $packageRoot 'DefenderShield.psd1') | Out-Null

$zipPath = Join-Path $OutputPath "DefenderShield-v$version.zip"
Compress-Archive -Path (Join-Path $moduleRoot '*') -DestinationPath $zipPath -Force

Write-Host "Built $zipPath"
