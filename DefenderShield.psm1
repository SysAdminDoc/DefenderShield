function Invoke-DefenderShield {
    [CmdletBinding()]
    param(
        [ValidateSet('Defender', 'Firewall', 'Both', 'Status')]
        [string]$Mode = 'Status',

        [switch]$DryRun,

        [switch]$Silent,

        [switch]$Json,

        [ValidateSet('Services', 'Registry', 'Tasks', 'WMI', 'GroupPolicy', 'Features', 'SmartScreen', 'AppLocker', 'MDE', 'WindowsUpdate')]
        [string[]]$Only,

        [ValidateSet('Services', 'Registry', 'Tasks', 'WMI', 'GroupPolicy', 'Features', 'SmartScreen', 'AppLocker', 'MDE', 'WindowsUpdate')]
        [string[]]$Skip,

        [string]$SnapshotPath,

        [string]$CompareSnapshot,

        [switch]$InstallWatchdog,

        [switch]$RemoveWatchdog,

        [switch]$WatchdogCheck,

        [string[]]$ComputerName,

        [switch]$Portable
    )

    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'DefenderShield.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "DefenderShield.ps1 was not found next to the module."
    }

    $arguments = @{
        Mode = $Mode
    }
    if ($DryRun) { $arguments.DryRun = $true }
    if ($Silent) { $arguments.Silent = $true }
    if ($Json) { $arguments.Json = $true }
    if ($Only) { $arguments.Only = $Only }
    if ($Skip) { $arguments.Skip = $Skip }
    if ($SnapshotPath) { $arguments.SnapshotPath = $SnapshotPath }
    if ($CompareSnapshot) { $arguments.CompareSnapshot = $CompareSnapshot }
    if ($InstallWatchdog) { $arguments.InstallWatchdog = $true }
    if ($RemoveWatchdog) { $arguments.RemoveWatchdog = $true }
    if ($WatchdogCheck) { $arguments.WatchdogCheck = $true }
    if ($ComputerName) { $arguments.ComputerName = $ComputerName }
    if ($Portable) { $arguments.Portable = $true }

    & $scriptPath @arguments
}

Set-Alias -Name Repair-DefenderShield -Value Invoke-DefenderShield
Export-ModuleMember -Function Invoke-DefenderShield -Alias Repair-DefenderShield
