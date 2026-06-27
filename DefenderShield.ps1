<#
.SYNOPSIS
    DefenderShield - Windows Defender & Firewall Repair Tool
.DESCRIPTION
    Comprehensive repair tool for restoring Windows Defender and Windows Firewall
    after they've been disabled by privacy tools like privacy.sexy, O&O ShutUp10,
    Debloaters, or manual modifications.

    Features a GUI with pre-scan health dashboard, auto-select broken items,
    before/after comparison, and Quick Fix All. Also supports a first-class CLI
    for automation (PDQ, Intune, SCCM).

    CLI Examples:
      .\DefenderShield.ps1 -Mode Status
      .\DefenderShield.ps1 -Mode Both -DryRun
      .\DefenderShield.ps1 -Mode Defender -Silent
      .\DefenderShield.ps1 -Mode Both -Json
      .\DefenderShield.ps1 -Mode Status -SnapshotPath .\status.json
      .\DefenderShield.ps1 -Mode Status -CompareSnapshot .\status.json
      .\DefenderShield.ps1 -InstallWatchdog
      .\DefenderShield.ps1 -Mode Both -ComputerName PC-01,PC-02
.PARAMETER Mode
    CLI operation mode: Defender, Firewall, Both, or Status. Omit for GUI.
.PARAMETER DryRun
    Simulate all changes without writing. Shows what would be modified.
.PARAMETER Silent
    Suppress console output (exit code only).
.PARAMETER Json
    Output structured JSON instead of text.
.PARAMETER Only
    Run only specific phases: Services, Registry, Tasks, WMI, GroupPolicy, Features, SmartScreen
.PARAMETER Skip
    Skip specific phases (same values as -Only).
.PARAMETER SnapshotPath
    Save a Status mode snapshot to a JSON file.
.PARAMETER CompareSnapshot
    Compare current Status mode results against a saved snapshot JSON file.
.PARAMETER InstallWatchdog
    Install an opt-in scheduled task that repairs if Defender or Firewall drift off.
.PARAMETER RemoveWatchdog
    Remove the DefenderShield watchdog scheduled task.
.PARAMETER WatchdogCheck
    Internal watchdog entry point: repair only when status indicates drift.
.PARAMETER ComputerName
    Run the selected CLI mode against one or more WinRM targets.
.PARAMETER Portable
    Write logs, reports, and backups under .\Logs\ instead of Desktop.
.NOTES
    Author: Matt
    Requires: Administrator privileges
    Version: 3.1.0
#>
param(
    [ValidateSet('Defender', 'Firewall', 'Both', 'Status')]
    [string]$Mode,

    [switch]$DryRun,

    [switch]$Silent,

    [switch]$Json,

    [ValidateSet('Services', 'Registry', 'Tasks', 'WMI', 'GroupPolicy', 'Features', 'SmartScreen', 'AppLocker', 'MDE', 'WindowsUpdate')]
    [string[]]$Only,

    [ValidateSet('Services', 'Registry', 'Tasks', 'WMI', 'GroupPolicy', 'Features', 'SmartScreen', 'AppLocker', 'MDE', 'WindowsUpdate')]
    [string[]]$Skip,

    [switch]$Worker,

    [string]$WorkerLogPath,

    [string]$WorkerBackupPath,

    [string]$WorkerReportPath,

    [string]$SnapshotPath,

    [string]$CompareSnapshot,

    [switch]$InstallWatchdog,

    [switch]$RemoveWatchdog,

    [switch]$WatchdogCheck,

    [string[]]$ComputerName,

    [switch]$Portable
)

# ============================================================================
# CONSTANTS & MODE DETECTION
# ============================================================================

$Script:Version = '3.1.0'
$Script:CliMode = [bool]($Mode -or $InstallWatchdog -or $RemoveWatchdog -or $WatchdogCheck -or ($ComputerName -and $ComputerName.Count -gt 0))
$Script:IsDryRun = [bool]$DryRun
$Script:IsSilent = [bool]$Silent
$Script:IsJson = [bool]$Json
$Script:IsWorker = [bool]$Worker
$Script:IsPortable = [bool]$Portable
$Script:ExitCode = 0  # 0=success, 1=partial, 2=failed, 3=blocked

# Phase filtering
$Script:ActivePhases = if ($Only -and $Only.Count -gt 0) {
    $Only
} else {
    @('Services', 'Registry', 'Tasks', 'WMI', 'GroupPolicy', 'Features', 'SmartScreen', 'AppLocker', 'MDE', 'WindowsUpdate')
}
if ($Skip -and $Skip.Count -gt 0) {
    $Script:ActivePhases = $Script:ActivePhases | Where-Object { $_ -notin $Skip }
}

function Test-PhaseActive {
    param([string]$Phase)
    return ($Phase -in $Script:ActivePhases)
}

# JSON output collector for -Json mode
$Script:JsonOutput = @{
    version   = $Script:Version
    timestamp = (Get-Date -Format 'o')
    dryRun    = $Script:IsDryRun
    mode      = if ($Mode) { $Mode } else { 'GUI' }
    preScan   = $null
    postScan  = $null
    actions   = [System.Collections.ArrayList]::new()
    reports   = @{}
    undo      = $null
    exitCode  = 0
}

# Dry-run action log
$Script:DryRunActions = [System.Collections.ArrayList]::new()

# ============================================================================
# ELEVATION CHECK
# ============================================================================

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$hasFleetTargets = ($ComputerName -and $ComputerName.Count -gt 0)
$localRepairMode = ($Mode -in @('Defender', 'Firewall', 'Both')) -and -not $hasFleetTargets
$requiresAdmin = (-not $Script:IsWorker) -and ((-not $Script:CliMode) -or $localRepairMode -or $InstallWatchdog -or $RemoveWatchdog -or $WatchdogCheck)

if (-not $isAdmin -and $requiresAdmin) {
    if ($Script:CliMode) {
        if (-not $Script:IsSilent) {
            Write-Host 'ERROR: DefenderShield requires Administrator privileges.' -ForegroundColor Red
            Write-Host 'Re-run from an elevated PowerShell prompt.' -ForegroundColor Yellow
        }
        exit 2
    }
    try {
        $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        if ($Mode)    { $argList += " -Mode $Mode" }
        if ($DryRun)  { $argList += ' -DryRun' }
        if ($Silent)  { $argList += ' -Silent' }
        if ($Json)    { $argList += ' -Json' }
        if ($Only)    { $argList += " -Only $($Only -join ',')" }
        if ($Skip)    { $argList += " -Skip $($Skip -join ',')" }
        if ($Portable) { $argList += ' -Portable' }
        if ($SnapshotPath) { $argList += " -SnapshotPath `"$SnapshotPath`"" }
        if ($CompareSnapshot) { $argList += " -CompareSnapshot `"$CompareSnapshot`"" }
        if ($InstallWatchdog) { $argList += ' -InstallWatchdog' }
        if ($RemoveWatchdog) { $argList += ' -RemoveWatchdog' }
        if ($WatchdogCheck) { $argList += ' -WatchdogCheck' }
        if ($ComputerName) { $argList += " -ComputerName $($ComputerName -join ',')" }
        Start-Process powershell.exe -ArgumentList $argList -Verb RunAs -ErrorAction Stop
        exit
    }
    catch {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show("This tool requires Administrator privileges.`n`nPlease right-click and select 'Run as Administrator'.", "DefenderShield", "OK", "Error") | Out-Null
        exit 2
    }
}

# ============================================================================
# ASSEMBLIES
# ============================================================================

if (-not $Script:CliMode) {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms -ErrorAction SilentlyContinue
}

# ============================================================================
# CONFIGURATION
# ============================================================================

$Script:Config = @{
    LogPath    = "$env:USERPROFILE\Desktop\DefenderShield_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    BackupPath = "$env:USERPROFILE\Desktop\DefenderShield_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    ReportPath = "$env:USERPROFILE\Desktop\DefenderShield_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
}

if ($Script:IsPortable) {
    $portableRoot = Join-Path $PSScriptRoot 'Logs'
    $Script:Config.LogPath = Join-Path $portableRoot "DefenderShield_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $Script:Config.BackupPath = Join-Path $portableRoot "DefenderShield_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $Script:Config.ReportPath = Join-Path $portableRoot "DefenderShield_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
}

if ($WorkerLogPath) { $Script:Config.LogPath = $WorkerLogPath }
if ($WorkerBackupPath) { $Script:Config.BackupPath = $WorkerBackupPath }
if ($WorkerReportPath) { $Script:Config.ReportPath = $WorkerReportPath }

$Script:DefenderServices = @(
    @{ Name = 'WinDefend'; DisplayName = 'Microsoft Defender Antivirus Service'; StartType = 'Automatic' },
    @{ Name = 'WdNisSvc'; DisplayName = 'Microsoft Defender Antivirus Network Inspection Service'; StartType = 'Manual' },
    @{ Name = 'WdNisDrv'; DisplayName = 'Microsoft Defender Antivirus Network Inspection Driver'; StartType = 'Manual' },
    @{ Name = 'WdFilter'; DisplayName = 'Microsoft Defender Antivirus Mini-Filter Driver'; StartType = 'Boot' },
    @{ Name = 'WdBoot'; DisplayName = 'Microsoft Defender Antivirus Boot Driver'; StartType = 'Boot' },
    @{ Name = 'Sense'; DisplayName = 'Windows Defender Advanced Threat Protection Service'; StartType = 'Manual' },
    @{ Name = 'SecurityHealthService'; DisplayName = 'Windows Security Service'; StartType = 'Manual' }
)

$Script:FirewallServices = @(
    @{ Name = 'mpssvc'; DisplayName = 'Windows Defender Firewall'; StartType = 'Automatic' },
    @{ Name = 'BFE'; DisplayName = 'Base Filtering Engine'; StartType = 'Automatic' },
    @{ Name = 'IKEEXT'; DisplayName = 'IKE and AuthIP IPsec Keying Modules'; StartType = 'Manual' },
    @{ Name = 'PolicyAgent'; DisplayName = 'IPsec Policy Agent'; StartType = 'Manual' }
)

$Script:WindowsUpdateServices = @(
    @{ Name = 'wuauserv'; DisplayName = 'Windows Update'; StartType = 'Manual'; StartValue = 3 },
    @{ Name = 'UsoSvc'; DisplayName = 'Update Orchestrator Service'; StartType = 'Automatic'; StartValue = 2 },
    @{ Name = 'DoSvc'; DisplayName = 'Delivery Optimization'; StartType = 'AutomaticDelayed'; StartValue = 2 },
    @{ Name = 'BITS'; DisplayName = 'Background Intelligent Transfer Service'; StartType = 'Manual'; StartValue = 3 }
)

$Script:MdeServices = @(
    @{ Name = 'Sense'; DisplayName = 'Microsoft Defender for Endpoint Sensor'; StartType = 'Manual'; StartValue = 3 }
)

$Script:ThirdPartyAVGuidance = @(
    @{ Pattern = 'Norton|Symantec'; Name = 'Norton/Symantec'; Guidance = 'Use Norton Remove and Reinstall Tool, then reboot before repairing Defender.' },
    @{ Pattern = 'McAfee'; Name = 'McAfee'; Guidance = 'Use McAfee Consumer Product Removal (MCPR), then reboot before repairing Defender.' },
    @{ Pattern = 'Avast'; Name = 'Avast'; Guidance = 'Use Avast Uninstall Utility in Safe Mode when normal uninstall leaves providers registered.' },
    @{ Pattern = 'AVG'; Name = 'AVG'; Guidance = 'Use AVG Clear in Safe Mode when normal uninstall leaves providers registered.' },
    @{ Pattern = 'Bitdefender'; Name = 'Bitdefender'; Guidance = 'Use Bitdefender Uninstall Tool for the installed product family, then reboot.' },
    @{ Pattern = 'Kaspersky'; Name = 'Kaspersky'; Guidance = 'Use kavremover if normal uninstall does not remove the Security Center provider.' },
    @{ Pattern = 'ESET'; Name = 'ESET'; Guidance = 'Use ESET Uninstaller from Safe Mode if standard removal leaves services registered.' },
    @{ Pattern = 'Sophos'; Name = 'Sophos'; Guidance = 'Use SophosZap only after normal uninstall fails, then reboot.' },
    @{ Pattern = 'Trend Micro'; Name = 'Trend Micro'; Guidance = 'Use Trend Micro Diagnostic Toolkit uninstall cleanup, then reboot.' },
    @{ Pattern = 'Malwarebytes'; Name = 'Malwarebytes'; Guidance = 'Disable registered antivirus mode or uninstall with Malwarebytes Support Tool before repairing Defender.' }
)

# Stores pre-repair scan results for before/after comparison
$Script:PreRepairScan = $null
$Script:UndoManifest = $null
$Script:WmiRemovalReport = [System.Collections.ArrayList]::new()
$Script:ReportEntries = [System.Collections.ArrayList]::new()
$Script:LastReportPath = $null

# ============================================================================
# LOGGING
# ============================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR', 'SECTION')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"

    # Collect for JSON output
    if ($Script:IsJson) {
        [void]$Script:JsonOutput.actions.Add(@{ level = $Level; message = $Message; timestamp = $timestamp })
    }

    if ($Script:ReportEntries) {
        [void]$Script:ReportEntries.Add([ordered]@{
            timestamp = $timestamp
            level     = $Level
            message   = $Message
        })
    }

    try {
        $logDir = Split-Path -Parent $Script:Config.LogPath
        if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Add-Content -Path $Script:Config.LogPath -Value $logMessage -ErrorAction SilentlyContinue
    }
    catch { }

    return $logMessage
}

function Ensure-BackupDirectory {
    try {
        if (-not (Test-Path -LiteralPath $Script:Config.BackupPath)) {
            New-Item -Path $Script:Config.BackupPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        }
        return $true
    }
    catch {
        return $false
    }
}

function Initialize-UndoManifest {
    if ($Script:UndoManifest) { return }

    Ensure-BackupDirectory | Out-Null
    $Script:UndoManifest = [ordered]@{
        tool          = 'DefenderShield'
        version       = $Script:Version
        created       = (Get-Date -Format 'o')
        computerName  = $env:COMPUTERNAME
        logPath       = $Script:Config.LogPath
        backupPath    = $Script:Config.BackupPath
        dryRun        = $Script:IsDryRun
        changes       = [System.Collections.ArrayList]::new()
    }
}

function Add-UndoEntry {
    param(
        [string]$Type,
        [string]$Action,
        [string]$Target,
        [object]$Before,
        [object]$After,
        [object]$Rollback,
        [string]$Status = 'Planned'
    )

    Initialize-UndoManifest
    $entry = [ordered]@{
        id        = $Script:UndoManifest.changes.Count + 1
        timestamp = (Get-Date -Format 'o')
        type      = $Type
        action    = $Action
        target    = $Target
        before    = $Before
        after     = $After
        rollback  = $Rollback
        status    = $Status
    }
    [void]$Script:UndoManifest.changes.Add($entry)
    return $entry
}

function Save-UndoManifest {
    Initialize-UndoManifest
    try {
        Ensure-BackupDirectory | Out-Null
        $manifestPath = Join-Path $Script:Config.BackupPath 'undo-manifest.json'
        $Script:UndoManifest.completed = (Get-Date -Format 'o')
        $Script:UndoManifest.logPath = $Script:Config.LogPath
        $Script:UndoManifest.reportPath = $Script:Config.ReportPath
        $Script:UndoManifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
        $Script:JsonOutput.undo = $manifestPath
        Update-Status "Undo manifest: $manifestPath" -Level SUCCESS
        return $manifestPath
    }
    catch {
        Update-Status "Could not write undo manifest: $($_.Exception.Message)" -Level WARNING
        return $null
    }
}

function ConvertTo-HtmlSafe {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Export-RepairReport {
    try {
        Ensure-BackupDirectory | Out-Null
        $rows = New-Object System.Text.StringBuilder
        foreach ($entry in $Script:ReportEntries) {
            $class = $entry.level.ToLowerInvariant()
            [void]$rows.AppendLine("<tr class='$class'><td>$(ConvertTo-HtmlSafe $entry.timestamp)</td><td>$(ConvertTo-HtmlSafe $entry.level)</td><td>$(ConvertTo-HtmlSafe $entry.message)</td></tr>")
        }

        $changeRows = New-Object System.Text.StringBuilder
        if ($Script:UndoManifest -and $Script:UndoManifest.changes) {
            foreach ($change in $Script:UndoManifest.changes) {
                [void]$changeRows.AppendLine("<tr><td>$(ConvertTo-HtmlSafe $change.id)</td><td>$(ConvertTo-HtmlSafe $change.type)</td><td>$(ConvertTo-HtmlSafe $change.action)</td><td>$(ConvertTo-HtmlSafe $change.target)</td><td>$(ConvertTo-HtmlSafe $change.status)</td></tr>")
            }
        }

        $html = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>DefenderShield Repair Report</title>
<style>
body { background:#1e1e2e; color:#cdd6f4; font-family:Segoe UI,Arial,sans-serif; margin:24px; }
h1,h2 { color:#89b4fa; }
table { border-collapse:collapse; width:100%; margin:16px 0 28px; }
th,td { border:1px solid #45475a; padding:8px 10px; text-align:left; vertical-align:top; }
th { background:#181825; color:#bac2de; }
.success td { color:#a6e3a1; }
.warning td { color:#f9e2af; }
.error td { color:#f38ba8; }
.section td { color:#89b4fa; font-weight:600; }
.meta { color:#a6adc8; }
</style>
</head>
<body>
<h1>DefenderShield v$($Script:Version) Repair Report</h1>
<p class="meta">Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') on $(ConvertTo-HtmlSafe $env:COMPUTERNAME)</p>
<p class="meta">Log: $(ConvertTo-HtmlSafe $Script:Config.LogPath)</p>
<p class="meta">Backups: $(ConvertTo-HtmlSafe $Script:Config.BackupPath)</p>
<h2>Changes</h2>
<table>
<thead><tr><th>ID</th><th>Type</th><th>Action</th><th>Target</th><th>Status</th></tr></thead>
<tbody>
$changeRows
</tbody>
</table>
<h2>Run Log</h2>
<table>
<thead><tr><th>Time</th><th>Level</th><th>Message</th></tr></thead>
<tbody>
$rows
</tbody>
</table>
</body>
</html>
"@
        $html | Set-Content -LiteralPath $Script:Config.ReportPath -Encoding UTF8
        $Script:LastReportPath = $Script:Config.ReportPath
        $Script:JsonOutput.reports['html'] = $Script:Config.ReportPath
        Update-Status "HTML report: $($Script:Config.ReportPath)" -Level SUCCESS
        return $Script:Config.ReportPath
    }
    catch {
        Update-Status "Could not export HTML report: $($_.Exception.Message)" -Level WARNING
        return $null
    }
}

function Update-Status {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    Write-Log -Message $Message -Level $Level | Out-Null

    # CLI mode: write to console
    if ($Script:CliMode) {
        if (-not $Script:IsSilent) {
            $fgColor = switch ($Level) {
                'SUCCESS' { 'Green' }
                'WARNING' { 'Yellow' }
                'ERROR'   { 'Red' }
                'SECTION' { 'Cyan' }
                default   { 'White' }
            }
            if ($Script:IsDryRun -and $Message -and -not $Message.StartsWith('[DRY-RUN]') -and $Level -in @('SUCCESS','WARNING')) {
                # Dry-run messages are prefixed elsewhere
            }
            Write-Host $Message -ForegroundColor $fgColor
        }
        return
    }

    # GUI mode: write to RichTextBox
    if ($Script:StatusTextBox) {
        Add-GuiStatusLine -Message $Message -Level $Level
    }
}

function Add-GuiStatusLine {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    if (-not $Script:StatusTextBox) { return }

    try {
        $Script:StatusTextBox.Dispatcher.Invoke([action]{
            $color = switch ($Level) {
                'SUCCESS' { '#a6e3a1' }
                'WARNING' { '#f9e2af' }
                'ERROR'   { '#f38ba8' }
                'SECTION' { '#89b4fa' }
                default   { '#cdd6f4' }
            }

            $paragraph = New-Object System.Windows.Documents.Paragraph
            $run = New-Object System.Windows.Documents.Run($Message)
            $run.Foreground = $color
            $paragraph.Inlines.Add($run)
            $paragraph.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)
            $Script:StatusTextBox.Document.Blocks.Add($paragraph)
            $Script:StatusTextBox.ScrollToEnd()
        }, [System.Windows.Threading.DispatcherPriority]::Background)
    }
    catch { }
}

# ============================================================================
# HEALTH SCAN FUNCTIONS
# ============================================================================

function Get-ServiceHealthStatus {
    param([string]$ServiceName)

    try {
        $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if (-not $svc) { return 'Missing' }

        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
        $startType = $null
        if (Test-Path $regPath) {
            $startType = (Get-ItemProperty -Path $regPath -Name 'Start' -ErrorAction SilentlyContinue).Start
        }

        if ($startType -eq 4) { return 'Disabled' }
        if ($svc.Status -eq 'Running') { return 'Running' }
        return 'Stopped'
    }
    catch {
        return 'Missing'
    }
}

function Get-HealthScan {
    $scan = @{}

    # Service statuses
    $scan['WinDefend'] = Get-ServiceHealthStatus 'WinDefend'
    $scan['SecurityHealthService'] = Get-ServiceHealthStatus 'SecurityHealthService'
    $scan['wscsvc'] = Get-ServiceHealthStatus 'wscsvc'
    $scan['MpsSvc'] = Get-ServiceHealthStatus 'mpssvc'

    # Real-time protection
    try {
        $mpStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
        if ($mpStatus) {
            $scan['RealTimeProtection'] = if ($mpStatus.RealTimeProtectionEnabled) { 'ON' } else { 'OFF' }
            $scan['TamperProtection'] = if ($mpStatus.IsTamperProtected) { 'ON' } else { 'OFF' }

            # Definition age
            $defAge = (New-TimeSpan -Start $mpStatus.AntivirusSignatureLastUpdated -End (Get-Date)).Days
            $scan['DefinitionAge'] = $defAge
        }
        else {
            $scan['RealTimeProtection'] = 'OFF'
            $scan['TamperProtection'] = 'OFF'
            $scan['DefinitionAge'] = -1
        }
    }
    catch {
        $scan['RealTimeProtection'] = 'OFF'
        $scan['TamperProtection'] = 'OFF'
        $scan['DefinitionAge'] = -1
    }

    # Group Policy blocking
    $gpBlocking = $false
    $gpPaths = @(
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableAntiSpyware' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableAntiVirus' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableRealtimeMonitoring' }
    )
    foreach ($gp in $gpPaths) {
        try {
            if (Test-Path $gp.Path) {
                $val = Get-ItemProperty -Path $gp.Path -Name $gp.Name -ErrorAction SilentlyContinue
                if ($null -ne $val -and $val.$($gp.Name) -eq 1) {
                    $gpBlocking = $true
                    break
                }
            }
        }
        catch { }
    }
    $scan['GroupPolicyBlocking'] = if ($gpBlocking) { 'Yes' } else { 'No' }

    # Windows Security app
    try {
        $app = Get-AppxPackage -Name 'Microsoft.SecHealthUI' -ErrorAction SilentlyContinue
        $scan['WindowsSecurityApp'] = if ($app) { 'Registered' } else { 'Missing' }
    }
    catch {
        $scan['WindowsSecurityApp'] = 'Missing'
    }

    # SmartScreen status
    try {
        $ssVal = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer' -Name 'SmartScreenEnabled' -ErrorAction SilentlyContinue
        if ($ssVal -and $ssVal.SmartScreenEnabled -in @('Prompt', 'RequireAdmin')) {
            $scan['SmartScreen'] = 'ON'
        }
        elseif ($ssVal -and $ssVal.SmartScreenEnabled -eq 'Off') {
            $scan['SmartScreen'] = 'OFF'
        }
        else {
            # Check policy
            $ssPol = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Name 'EnableSmartScreen' -ErrorAction SilentlyContinue
            if ($ssPol -and $ssPol.EnableSmartScreen -eq 0) {
                $scan['SmartScreen'] = 'OFF'
            }
            else {
                $scan['SmartScreen'] = 'ON'
            }
        }
    }
    catch {
        $scan['SmartScreen'] = 'Unknown'
    }

    try {
        $avList = Get-ThirdPartyAV
        $scan['ThirdPartyAV'] = if ($avList) { ($avList | ForEach-Object { $_.Name }) -join ', ' } else { 'None' }
    }
    catch {
        $scan['ThirdPartyAV'] = 'Unknown'
    }

    try {
        $sense = Get-ServiceHealthStatus 'Sense'
        $scan['MDE'] = $sense
    }
    catch {
        $scan['MDE'] = 'Unknown'
    }

    return $scan
}

function Get-HealthColor {
    param([string]$Component, [string]$Value)

    # Catppuccin Mocha: Green=#a6e3a1  Yellow=#f9e2af  Red=#f38ba8  Text=#cdd6f4
    switch ($Component) {
        'WinDefend'            { if ($Value -eq 'Running') { '#a6e3a1' } elseif ($Value -eq 'Stopped') { '#f9e2af' } else { '#f38ba8' } }
        'SecurityHealthService' { if ($Value -eq 'Running') { '#a6e3a1' } elseif ($Value -eq 'Stopped') { '#f9e2af' } else { '#f38ba8' } }
        'wscsvc'               { if ($Value -eq 'Running') { '#a6e3a1' } elseif ($Value -eq 'Stopped') { '#f9e2af' } else { '#f38ba8' } }
        'MpsSvc'               { if ($Value -eq 'Running') { '#a6e3a1' } elseif ($Value -eq 'Stopped') { '#f9e2af' } else { '#f38ba8' } }
        'RealTimeProtection'   { if ($Value -eq 'ON') { '#a6e3a1' } else { '#f38ba8' } }
        'TamperProtection'     { if ($Value -eq 'ON') { '#a6e3a1' } else { '#f9e2af' } }
        'SmartScreen'          { if ($Value -eq 'ON') { '#a6e3a1' } elseif ($Value -eq 'OFF') { '#f38ba8' } else { '#f9e2af' } }
        'ThirdPartyAV'         { if ($Value -eq 'None') { '#a6e3a1' } elseif ($Value -eq 'Unknown') { '#f9e2af' } else { '#f38ba8' } }
        'MDE'                  { if ($Value -eq 'Running') { '#a6e3a1' } elseif ($Value -in @('Stopped', 'Missing')) { '#f9e2af' } else { '#f38ba8' } }
        'DefinitionAge'        {
            $days = [int]$Value
            if ($days -lt 0) { '#f38ba8' } elseif ($days -le 3) { '#a6e3a1' } elseif ($days -le 7) { '#f9e2af' } else { '#f38ba8' }
        }
        'GroupPolicyBlocking'  { if ($Value -eq 'No') { '#a6e3a1' } else { '#f38ba8' } }
        'WindowsSecurityApp'   { if ($Value -eq 'Registered') { '#a6e3a1' } else { '#f38ba8' } }
        default { '#cdd6f4' }
    }
}

function Update-HealthDashboard {
    param([hashtable]$Scan)

    if (-not $Script:DashboardLabels) { return }

    if ($Script:DashboardLabels['lblTileDefenderStatus']) {
        $defenderBroken = Test-DefenderBroken -Scan $Scan
        $firewallBroken = Test-FirewallBroken -Scan $Scan
        $tileData = @(
            @{ Name = 'lblTileDefenderStatus'; Value = if ($defenderBroken) { 'Repair needed' } else { 'Healthy' }; Component = 'RealTimeProtection'; ColorValue = if ($defenderBroken) { 'OFF' } else { 'ON' } },
            @{ Name = 'lblTileFirewallStatus'; Value = if ($firewallBroken) { 'Repair needed' } else { 'Healthy' }; Component = 'MpsSvc'; ColorValue = $Scan['MpsSvc'] },
            @{ Name = 'lblTileTamperStatus'; Value = $Scan['TamperProtection']; Component = 'TamperProtection'; ColorValue = $Scan['TamperProtection'] },
            @{ Name = 'lblTileSignatureStatus'; Value = if ($Scan['DefinitionAge'] -ge 0) { "$($Scan['DefinitionAge']) days" } else { 'Unknown' }; Component = 'DefinitionAge'; ColorValue = $Scan['DefinitionAge'] },
            @{ Name = 'lblTileAvStatus'; Value = $Scan['ThirdPartyAV']; Component = 'ThirdPartyAV'; ColorValue = $Scan['ThirdPartyAV'] }
        )

        foreach ($tile in $tileData) {
            $label = $Script:DashboardLabels[$tile.Name]
            if ($label) {
                $label.Text = $tile.Value
                $label.Foreground = Get-HealthColor -Component $tile.Component -Value ([string]$tile.ColorValue)
            }
        }
    }

    $labelMap = @{
        'WinDefend'            = @{ Label = 'lblWinDefend'; Value = $Scan['WinDefend'] }
        'SecurityHealthService' = @{ Label = 'lblSecHealth'; Value = $Scan['SecurityHealthService'] }
        'wscsvc'               = @{ Label = 'lblWscsvc'; Value = $Scan['wscsvc'] }
        'MpsSvc'               = @{ Label = 'lblMpsSvc'; Value = $Scan['MpsSvc'] }
        'RealTimeProtection'   = @{ Label = 'lblRTP'; Value = $Scan['RealTimeProtection'] }
        'TamperProtection'     = @{ Label = 'lblTamper'; Value = $Scan['TamperProtection'] }
        'SmartScreen'          = @{ Label = 'lblSmartScreen'; Value = $Scan['SmartScreen'] }
        'DefinitionAge'        = @{ Label = 'lblDefAge'; Value = if ($Scan['DefinitionAge'] -ge 0) { "$($Scan['DefinitionAge']) days" } else { 'Unknown' } }
        'GroupPolicyBlocking'  = @{ Label = 'lblGPBlock'; Value = $Scan['GroupPolicyBlocking'] }
        'WindowsSecurityApp'   = @{ Label = 'lblWinSecApp'; Value = $Scan['WindowsSecurityApp'] }
    }

    foreach ($key in $labelMap.Keys) {
        $info = $labelMap[$key]
        $label = $Script:DashboardLabels[$info.Label]
        if ($label) {
            $label.Text = $info.Value
            $color = Get-HealthColor -Component $key -Value $Scan[$key].ToString()
            $label.Foreground = $color
        }
    }
}

function Test-DefenderBroken {
    param([hashtable]$Scan)

    $broken = $false
    if ($Scan['WinDefend'] -ne 'Running') { $broken = $true }
    if ($Scan['RealTimeProtection'] -ne 'ON') { $broken = $true }
    if ($Scan['GroupPolicyBlocking'] -eq 'Yes') { $broken = $true }
    if ($Scan['SecurityHealthService'] -notin @('Running', 'Stopped')) { $broken = $true }
    if ($Scan['WindowsSecurityApp'] -ne 'Registered') { $broken = $true }
    $defAge = $Scan['DefinitionAge']
    if ($defAge -lt 0 -or $defAge -gt 7) { $broken = $true }
    return $broken
}

function Test-FirewallBroken {
    param([hashtable]$Scan)

    return ($Scan['MpsSvc'] -ne 'Running')
}

function Get-ComparisonReport {
    param(
        [hashtable]$Before,
        [hashtable]$After
    )

    $lines = @()
    $lines += ''
    $lines += '=== REPAIR RESULTS ==='

    $components = @(
        @{ Key = 'WinDefend'; Label = 'WinDefend Service' },
        @{ Key = 'SecurityHealthService'; Label = 'SecurityHealthService' },
        @{ Key = 'wscsvc'; Label = 'Security Center (wscsvc)' },
        @{ Key = 'MpsSvc'; Label = 'Firewall (MpsSvc)' },
        @{ Key = 'RealTimeProtection'; Label = 'Real-Time Protection' },
        @{ Key = 'TamperProtection'; Label = 'Tamper Protection' },
        @{ Key = 'SmartScreen'; Label = 'SmartScreen' },
        @{ Key = 'GroupPolicyBlocking'; Label = 'Group Policy Blocking' },
        @{ Key = 'WindowsSecurityApp'; Label = 'Windows Security App' }
    )

    foreach ($comp in $components) {
        $bVal = $Before[$comp.Key]
        $aVal = $After[$comp.Key]
        $padLabel = $comp.Label.PadRight(26)

        if ($bVal -ne $aVal) {
            $lines += "$padLabel $bVal -> $aVal [FIXED]"
        }
        else {
            $lines += "$padLabel $aVal -> $aVal [OK - No change needed]"
        }
    }

    # Definition age
    $bDef = $Before['DefinitionAge']
    $aDef = $After['DefinitionAge']
    $bDefStr = if ($bDef -ge 0) { "$bDef days" } else { 'Unknown' }
    $aDefStr = if ($aDef -ge 0) { "$aDef days" } else { 'Unknown' }
    $padLabel = 'Definition Age'.PadRight(26)
    if ($bDefStr -ne $aDefStr) {
        $lines += "$padLabel $bDefStr -> $aDefStr [UPDATED]"
    }
    else {
        $lines += "$padLabel $aDefStr -> $aDefStr [OK - No change needed]"
    }

    return $lines
}

# ============================================================================
# PRE-FLIGHT: THIRD-PARTY AV DETECTION
# ============================================================================

function Get-ThirdPartyAV {
    <#
    .SYNOPSIS
        Detects third-party antivirus registered as the active security provider.
        Returns $null if only Windows Defender is present, otherwise returns the AV info.
    #>
    $avProducts = @()
    try {
        $wmiAV = Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName 'AntiVirusProduct' -ErrorAction SilentlyContinue
        foreach ($av in $wmiAV) {
            if ($av.displayName -and $av.displayName -notmatch 'Windows Defender|Microsoft Defender') {
                $avProducts += @{
                    Name    = $av.displayName
                    State   = $av.productState
                    Path    = $av.pathToSignedProductExe
                }
            }
        }
    }
    catch { }

    if ($avProducts.Count -gt 0) { return $avProducts }
    return $null
}

function Get-ThirdPartyAVGuidance {
    param([array]$AVList)

    $guidance = @()
    foreach ($av in @($AVList)) {
        if (-not $av) { continue }
        $match = $Script:ThirdPartyAVGuidance | Where-Object { $av.Name -match $_.Pattern } | Select-Object -First 1
        if ($match) {
            $guidance += [ordered]@{
                product  = $av.Name
                family   = $match.Name
                guidance = $match.Guidance
            }
        }
        else {
            $guidance += [ordered]@{
                product  = $av.Name
                family   = 'Unknown'
                guidance = 'Use the vendor cleanup tool or fully uninstall this antivirus, reboot, then repair Defender.'
            }
        }
    }

    return $guidance
}

function Test-ThirdPartyAVBlocking {
    <#
    .SYNOPSIS
        Pre-flight check: if a 3rd-party AV is the active provider, warn and abort.
        Returns $true if repair should be blocked.
    #>
    $avList = Get-ThirdPartyAV
    if (-not $avList) { return $false }

    $avNames = ($avList | ForEach-Object { $_.Name }) -join ', '
    Update-Status "BLOCKED: Third-party antivirus detected: $avNames" -Level ERROR
    Update-Status "A third-party AV is registered as the active security provider." -Level WARNING
    Update-Status "Uninstall the third-party AV before running DefenderShield, or it will fight the repair." -Level WARNING
    foreach ($hint in Get-ThirdPartyAVGuidance -AVList $avList) {
        Update-Status "$($hint.product): $($hint.guidance)" -Level WARNING
    }

    $Script:ExitCode = 3
    return $true
}

# ============================================================================
# PRIVACY TOOL DETECTION
# ============================================================================

function Get-DetectedPrivacyTools {
    <#
    .SYNOPSIS
        Auto-detect privacy/debloater tools that have likely run on this system.
        Returns a list of detected tool signatures.
    #>
    $detected = @()

    # privacy.sexy signatures
    $privacySexySigs = @(
        "$env:APPDATA\privacy.sexy",
        "$env:LOCALAPPDATA\privacy.sexy",
        "$env:TEMP\privacy-sexy-*"
    )
    foreach ($sig in $privacySexySigs) {
        if (Test-Path $sig -ErrorAction SilentlyContinue) {
            $detected += @{ Tool = 'privacy.sexy'; Evidence = "Found: $sig" }
            break
        }
    }

    # O&O ShutUp10
    $ooSigs = @(
        "${env:ProgramFiles}\O&O\ShutUp10",
        "${env:ProgramFiles(x86)}\O&O\ShutUp10",
        "$env:APPDATA\O&O\ShutUp10"
    )
    foreach ($sig in $ooSigs) {
        if (Test-Path $sig -ErrorAction SilentlyContinue) {
            $detected += @{ Tool = 'O&O ShutUp10'; Evidence = "Found: $sig" }
            break
        }
    }
    # OOSU10.cfg on desktop or downloads
    $ooConfigs = @("$env:USERPROFILE\Desktop\OOSU10.cfg", "$env:USERPROFILE\Downloads\OOSU10.cfg")
    foreach ($cfg in $ooConfigs) {
        if ((Test-Path $cfg -ErrorAction SilentlyContinue) -and $detected.Tool -notcontains 'O&O ShutUp10') {
            $detected += @{ Tool = 'O&O ShutUp10'; Evidence = "Found config: $cfg" }
            break
        }
    }

    # Chris Titus WinUtil / winutil
    try {
        $winutilTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -match 'WinUtil|ChrisTitus' }
        if ($winutilTasks) {
            $detected += @{ Tool = 'Chris Titus WinUtil'; Evidence = "Scheduled task: $($winutilTasks[0].TaskName)" }
        }
    }
    catch { }

    # Windows10Debloater / Sycnex
    $debloaterSigs = @(
        "$env:TEMP\Windows10Debloater*",
        "$env:USERPROFILE\Desktop\Windows10Debloater*",
        "$env:USERPROFILE\Downloads\Windows10Debloater*"
    )
    foreach ($sig in $debloaterSigs) {
        if (Test-Path $sig -ErrorAction SilentlyContinue) {
            $detected += @{ Tool = 'Windows10Debloater (Sycnex)'; Evidence = "Found: $sig" }
            break
        }
    }

    # Sophia Script
    $sophiaSigs = @(
        "$env:TEMP\Sophia*",
        "$env:USERPROFILE\Desktop\Sophia*",
        "$env:USERPROFILE\Downloads\Sophia*"
    )
    foreach ($sig in $sophiaSigs) {
        if (Test-Path $sig -ErrorAction SilentlyContinue) {
            $detected += @{ Tool = 'Sophia Script'; Evidence = "Found: $sig" }
            break
        }
    }

    # Debloat-Win11 / Raphire
    try {
        $debloat11Tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -match 'Debloat' }
        if ($debloat11Tasks -and $detected.Tool -notcontains 'Windows10Debloater (Sycnex)') {
            $detected += @{ Tool = 'Debloat-Win11'; Evidence = "Scheduled task: $($debloat11Tasks[0].TaskName)" }
        }
    }
    catch { }

    # DefenderControl
    $dcSigs = @(
        "$env:TEMP\dControl*",
        "$env:USERPROFILE\Desktop\dControl*",
        "$env:USERPROFILE\Downloads\dControl*",
        "${env:ProgramFiles}\DefenderControl"
    )
    foreach ($sig in $dcSigs) {
        if (Test-Path $sig -ErrorAction SilentlyContinue) {
            $detected += @{ Tool = 'DefenderControl'; Evidence = "Found: $sig" }
            break
        }
    }

    return $detected
}

# ============================================================================
# POLICY SOURCE AUDIT
# ============================================================================

function Get-PolicySourceAudit {
    <#
    .SYNOPSIS
        Enumerates every active Defender/Firewall blocker (registry, WMI, scheduled task,
        GPO) and returns a structured table of "what's holding Defender/Firewall down".
    #>
    $blockers = @()

    # --- Registry blockers ---
    $regBlockers = @(
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableAntiSpyware'; Label = 'GP: Disable AntiSpyware' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableAntiVirus'; Label = 'GP: Disable AntiVirus' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableRealtimeMonitoring'; Label = 'GP: Disable RTP' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableBehaviorMonitoring'; Label = 'GP: Disable Behavior Monitoring' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableOnAccessProtection'; Label = 'GP: Disable On-Access' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableIOAVProtection'; Label = 'GP: Disable IOAV' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableScanOnRealtimeEnable'; Label = 'GP: Disable Scan on RTE' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender'; Name = 'DisableAntiSpyware'; Label = 'Direct: Disable AntiSpyware' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender'; Name = 'DisableAntiVirus'; Label = 'Direct: Disable AntiVirus' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableRealtimeMonitoring'; Label = 'Direct: Disable RTP' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile'; Name = 'EnableFirewall'; Label = 'GP: Firewall Domain Disabled'; BlockValue = 0 },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile'; Name = 'EnableFirewall'; Label = 'GP: Firewall Private Disabled'; BlockValue = 0 },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile'; Name = 'EnableFirewall'; Label = 'GP: Firewall Public Disabled'; BlockValue = 0 },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\StandardProfile'; Name = 'EnableFirewall'; Label = 'GP: Firewall Standard Disabled'; BlockValue = 0 }
    )

    foreach ($rb in $regBlockers) {
        try {
            if (Test-Path $rb.Path) {
                $val = Get-ItemProperty -Path $rb.Path -Name $rb.Name -ErrorAction SilentlyContinue
                if ($null -ne $val) {
                    $actualVal = $val.$($rb.Name)
                    $blockVal = if ($rb.ContainsKey('BlockValue')) { $rb.BlockValue } else { 1 }
                    if ($actualVal -eq $blockVal) {
                        $blockers += @{
                            Source = 'Registry'
                            Label  = $rb.Label
                            Detail = "$($rb.Path)\$($rb.Name) = $actualVal"
                        }
                    }
                }
            }
        }
        catch { }
    }

    # --- Service blockers (disabled services) ---
    $criticalServices = @('WinDefend', 'SecurityHealthService', 'wscsvc', 'mpssvc', 'BFE')
    foreach ($svcName in $criticalServices) {
        try {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$svcName"
            if (Test-Path $regPath) {
                $startType = (Get-ItemProperty -Path $regPath -Name 'Start' -ErrorAction SilentlyContinue).Start
                if ($startType -eq 4) {
                    $blockers += @{
                        Source = 'Service'
                        Label  = "$svcName disabled"
                        Detail = "Start type = 4 (Disabled)"
                    }
                }
            }
        }
        catch { }
    }

    foreach ($svc in @($Script:WindowsUpdateServices + $Script:MdeServices)) {
        try {
            $startType = Get-ServiceStartValue -ServiceName $svc.Name
            if ($startType -eq 4) {
                $blockers += @{
                    Source = 'Service'
                    Label  = "$($svc.Name) disabled"
                    Detail = "$($svc.DisplayName) Start type = 4 (Disabled)"
                }
            }
        }
        catch { }
    }

    try {
        $appLockerBlockers = @(Get-AppLockerMsMpEngBlockers)
        foreach ($blocker in $appLockerBlockers) {
            $blockers += @{
                Source = $blocker.Source
                Label  = "$($blocker.Source) blocks Defender"
                Detail = $blocker.Detail
            }
        }
    }
    catch { }

    # --- WMI subscription blockers ---
    try {
        $filters = Get-WmiObject -Query "SELECT * FROM __EventFilter WHERE Name LIKE '%Defender%' OR Name LIKE '%WinDefend%'" -Namespace 'root\subscription' -ErrorAction SilentlyContinue
        foreach ($f in $filters) {
            $blockers += @{
                Source = 'WMI Subscription'
                Label  = "WMI filter: $($f.Name)"
                Detail = $f.Query
            }
        }
    }
    catch { }

    # --- Scheduled task blockers ---
    try {
        $suspiciousTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
            $_.TaskName -match 'DisableDefender|DisableWinDefend|KillDefender|StopDefender'
        }
        foreach ($t in $suspiciousTasks) {
            $blockers += @{
                Source = 'Scheduled Task'
                Label  = "Malicious task: $($t.TaskName)"
                Detail = "State: $($t.State)"
            }
        }
    }
    catch { }

    # --- Disabled Defender scheduled tasks ---
    $defenderTasks = @('Windows Defender Cache Maintenance', 'Windows Defender Cleanup', 'Windows Defender Scheduled Scan', 'Windows Defender Verification')
    foreach ($taskName in $defenderTasks) {
        try {
            $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($task -and $task.State -eq 'Disabled') {
                $blockers += @{
                    Source = 'Scheduled Task'
                    Label  = "Defender task disabled: $taskName"
                    Detail = "State: Disabled"
                }
            }
        }
        catch { }
    }

    # --- Group Policy file ---
    $machinePolPath = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
    if (Test-Path $machinePolPath -ErrorAction SilentlyContinue) {
        try {
            $polContent = [System.IO.File]::ReadAllBytes($machinePolPath)
            $polText = [System.Text.Encoding]::Unicode.GetString($polContent)
            if ($polText -match 'Windows Defender|AntiSpyware|AntiVirus') {
                $blockers += @{
                    Source = 'Group Policy'
                    Label  = 'Machine Registry.pol contains Defender policy'
                    Detail = $machinePolPath
                }
            }
        }
        catch { }
    }

    return $blockers
}

function Show-PolicySourceAudit {
    param([array]$Blockers)

    if (-not $Blockers -or $Blockers.Count -eq 0) {
        Update-Status "No active blockers detected." -Level SUCCESS
        return
    }

    Update-Status "Found $($Blockers.Count) active blocker(s):" -Level WARNING
    Update-Status ("{0,-20} {1,-45} {2}" -f 'SOURCE', 'LABEL', 'DETAIL') -Level SECTION

    foreach ($b in $Blockers) {
        $line = "{0,-20} {1,-45} {2}" -f $b.Source, $b.Label, $b.Detail
        Update-Status $line -Level WARNING
    }
}

function New-StatusSnapshot {
    $avList = Get-ThirdPartyAV
    $thirdPartyAV = @()
    $avGuidance = @()
    if ($avList) {
        $thirdPartyAV = @($avList)
        $avGuidance = @(Get-ThirdPartyAVGuidance -AVList $avList)
    }

    return [ordered]@{
        version       = $Script:Version
        timestamp     = (Get-Date -Format 'o')
        computerName  = $env:COMPUTERNAME
        scan          = Get-HealthScan
        privacyTools  = @(Get-DetectedPrivacyTools)
        blockers      = @(Get-PolicySourceAudit)
        thirdPartyAV  = $thirdPartyAV
        avGuidance    = $avGuidance
    }
}

function Save-StatusSnapshot {
    param(
        [object]$Snapshot,
        [string]$Path
    )

    try {
        $dir = Split-Path -Parent $Path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -Path $dir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        }
        $Snapshot | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
        return $true
    }
    catch {
        Write-Warning "Could not write snapshot: $($_.Exception.Message)"
        return $false
    }
}

function Compare-StatusSnapshot {
    param(
        [object]$Previous,
        [object]$Current
    )

    $diff = [System.Collections.ArrayList]::new()
    $previousScan = @{}
    if ($Previous.scan) {
        foreach ($prop in $Previous.scan.PSObject.Properties) {
            $previousScan[$prop.Name] = $prop.Value
        }
    }

    foreach ($prop in $Current.scan.GetEnumerator()) {
        $old = if ($previousScan.ContainsKey($prop.Key)) { $previousScan[$prop.Key] } else { '<missing>' }
        $new = $prop.Value
        if ([string]$old -ne [string]$new) {
            [void]$diff.Add([ordered]@{
                component = $prop.Key
                before    = $old
                after     = $new
            })
        }
    }

    $oldBlockers = if ($Previous.blockers) { @($Previous.blockers).Count } else { 0 }
    $newBlockers = if ($Current.blockers) { @($Current.blockers).Count } else { 0 }
    if ($oldBlockers -ne $newBlockers) {
        [void]$diff.Add([ordered]@{
            component = 'ActiveBlockers'
            before    = $oldBlockers
            after     = $newBlockers
        })
    }

    return @($diff)
}

function Show-StatusDiff {
    param([array]$Diff)

    if (-not $Diff -or $Diff.Count -eq 0) {
        Write-Host 'No status drift detected.' -ForegroundColor Green
        return
    }

    Write-Host 'Status drift detected:' -ForegroundColor Yellow
    Write-Host ("{0,-24} {1,-24} {2}" -f 'COMPONENT', 'BEFORE', 'AFTER') -ForegroundColor Cyan
    foreach ($item in $Diff) {
        Write-Host ("{0,-24} {1,-24} {2}" -f $item.component, $item.before, $item.after) -ForegroundColor Yellow
    }
}

function Install-WatchdogTask {
    $taskName = 'DefenderShield Watchdog'
    $scriptPath = $PSCommandPath
    $argument = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -WatchdogCheck -Silent"
    if ($Script:IsPortable) { $argument += ' -Portable' }

    try {
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument
        $triggers = @(
            (New-ScheduledTaskTrigger -AtLogOn),
            (New-ScheduledTaskTrigger -Daily -At 9am)
        )
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 2)
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Principal $principal -Settings $settings -Force | Out-Null
        Write-Host "Installed scheduled task: $taskName" -ForegroundColor Green
        exit 0
    }
    catch {
        Write-Warning "Could not install watchdog: $($_.Exception.Message)"
        exit 2
    }
}

function Remove-WatchdogTask {
    $taskName = 'DefenderShield Watchdog'
    try {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
        Write-Host "Removed scheduled task: $taskName" -ForegroundColor Green
        exit 0
    }
    catch {
        Write-Warning "Could not remove watchdog: $($_.Exception.Message)"
        exit 2
    }
}

function Invoke-WatchdogCheck {
    $scan = Get-HealthScan
    $defBroken = Test-DefenderBroken -Scan $scan
    $fwBroken = Test-FirewallBroken -Scan $scan
    if (-not $defBroken -and -not $fwBroken) {
        if (-not $Script:IsSilent) { Write-Host 'DefenderShield watchdog: system healthy.' -ForegroundColor Green }
        exit 0
    }

    if (-not $Script:IsSilent) { Write-Host 'DefenderShield watchdog: repair needed.' -ForegroundColor Yellow }
    $Script:JsonOutput.preScan = $scan
    Start-Repair -RepairFirewall $fwBroken -RepairDefender $defBroken -CreateRestorePoint $false | Out-Null
    exit $Script:ExitCode
}

function Invoke-FleetRepair {
    param([string[]]$Targets)

    $scriptText = Get-Content -LiteralPath $PSCommandPath -Raw
    $results = @()
    foreach ($target in $Targets) {
        try {
            $remoteResult = Invoke-Command -ComputerName $target -ScriptBlock {
                param($ScriptText, $ModeValue, $DryRunValue, $SilentValue, $JsonValue, $OnlyValue, $SkipValue)

                $remotePath = Join-Path $env:TEMP 'DefenderShield-Remote.ps1'
                Set-Content -LiteralPath $remotePath -Value $ScriptText -Encoding UTF8
                $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $remotePath, '-Mode', $ModeValue)
                if ($DryRunValue) { $args += '-DryRun' }
                if ($SilentValue) { $args += '-Silent' }
                if ($JsonValue) { $args += '-Json' }
                if ($OnlyValue) { $args += @('-Only', ($OnlyValue -join ',')) }
                if ($SkipValue) { $args += @('-Skip', ($SkipValue -join ',')) }
                & powershell.exe @args
                $LASTEXITCODE
            } -ArgumentList $scriptText, $Mode, $Script:IsDryRun, $Script:IsSilent, $Script:IsJson, $Only, $Skip -ErrorAction Stop

            $exitCode = @($remoteResult)[-1]
            $results += [ordered]@{ computerName = $target; exitCode = $exitCode; status = 'Complete' }
            if (-not $Script:IsSilent) { Write-Host "$target : exit $exitCode" -ForegroundColor Cyan }
        }
        catch {
            $results += [ordered]@{ computerName = $target; exitCode = 2; status = $_.Exception.Message }
            if (-not $Script:IsSilent) { Write-Warning "$target : $($_.Exception.Message)" }
        }
    }

    if ($Script:IsJson) {
        Write-Output (@{ version = $Script:Version; fleet = $results } | ConvertTo-Json -Depth 5)
    }

    if (($results | Where-Object { $_.exitCode -ne 0 }).Count -gt 0) { exit 1 }
    exit 0
}

# ============================================================================
# SMARTSCREEN REPAIR
# ============================================================================

function Repair-SmartScreen {
    <#
    .SYNOPSIS
        Repairs Windows SmartScreen if privacy tools have disabled it.
    #>
    Update-Status "Repairing SmartScreen..." -Level SECTION

    $smartScreenPaths = @(
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'EnableSmartScreen'; Value = 1; Type = 'DWord' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer'; Name = 'SmartScreenEnabled'; Value = 'Prompt'; Type = 'String' },
        @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppHost'; Name = 'EnableWebContentEvaluation'; Value = 1; Type = 'DWord' },
        @{ Path = 'HKCU:\SOFTWARE\Microsoft\Edge\SmartScreenEnabled'; Name = '(Default)'; Value = 1; Type = 'DWord' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'ShellSmartScreenLevel'; Value = 'Warn'; Type = 'String' }
    )

    $repaired = 0
    foreach ($item in $smartScreenPaths) {
        try {
            if ($Script:IsDryRun) {
                [void]$Script:DryRunActions.Add("[DRY-RUN] Would set $($item.Path)\$($item.Name) = $($item.Value)")
                Update-Status "[DRY-RUN] Would set $($item.Path)\$($item.Name) = $($item.Value)" -Level INFO
                $repaired++
                continue
            }

            if (Set-RegistryValue -Path $item.Path -Name $item.Name -Value $item.Value -Type $item.Type) {
                $repaired++
            }
        }
        catch { }
    }

    # Remove SmartScreen-disabling policies
    $disablingPolicies = @(
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'EnableSmartScreen'; BlockValue = 0 },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\MicrosoftEdge\PhishingFilter'; Name = 'EnabledV9'; BlockValue = 0 }
    )

    foreach ($pol in $disablingPolicies) {
        try {
            if (Test-Path $pol.Path) {
                $val = Get-ItemProperty -Path $pol.Path -Name $pol.Name -ErrorAction SilentlyContinue
                if ($null -ne $val -and $val.$($pol.Name) -eq $pol.BlockValue) {
                    if ($Script:IsDryRun) {
                        [void]$Script:DryRunActions.Add("[DRY-RUN] Would remove blocking policy $($pol.Path)\$($pol.Name)")
                        Update-Status "[DRY-RUN] Would remove blocking policy $($pol.Path)\$($pol.Name)" -Level INFO
                    }
                    else {
                        Remove-RegistryValue -Path $pol.Path -Name $pol.Name | Out-Null
                    }
                    $repaired++
                }
            }
        }
        catch { }
    }

    if ($repaired -gt 0) {
        Update-Status "SmartScreen: $repaired setting(s) repaired" -Level SUCCESS
    }
    else {
        Update-Status "SmartScreen: Already properly configured" -Level SUCCESS
    }
}

# ============================================================================
# FIREWALL DEPENDENCY VALIDATION
# ============================================================================

function Test-FirewallServiceDependencies {
    <#
    .SYNOPSIS
        Validates BFE -> mpssvc -> IKEEXT -> PolicyAgent dependency tree.
        Returns $true if the dependency order is sane.
    #>
    Update-Status "Validating firewall service dependency tree..." -Level SECTION

    $dependencyOrder = @(
        @{ Name = 'BFE'; MustRunBefore = @('mpssvc', 'IKEEXT', 'PolicyAgent') },
        @{ Name = 'mpssvc'; MustRunBefore = @() }
    )

    $issues = @()
    foreach ($dep in $dependencyOrder) {
        $svc = Get-Service -Name $dep.Name -ErrorAction SilentlyContinue
        if (-not $svc) {
            $issues += "$($dep.Name) service is MISSING"
            continue
        }

        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$($dep.Name)"
        if (Test-Path $regPath) {
            $startType = (Get-ItemProperty -Path $regPath -Name 'Start' -ErrorAction SilentlyContinue).Start
            if ($startType -eq 4) {
                $issues += "$($dep.Name) is disabled (Start=4); dependents will fail"
            }
        }
    }

    # Check that BFE can start before mpssvc
    try {
        $bfe = Get-Service -Name 'BFE' -ErrorAction SilentlyContinue
        $mpssvc = Get-Service -Name 'mpssvc' -ErrorAction SilentlyContinue
        if ($bfe -and $mpssvc) {
            $bfeStart = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\BFE' -Name 'Start' -ErrorAction SilentlyContinue).Start
            $mpsStart = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\mpssvc' -Name 'Start' -ErrorAction SilentlyContinue).Start
            if ($bfeStart -gt $mpsStart -and $bfeStart -ne 3) {
                $issues += "BFE start type ($bfeStart) is later than mpssvc ($mpsStart); BFE must start first"
            }
        }
    }
    catch { }

    if ($issues.Count -eq 0) {
        Update-Status "Firewall dependency tree: OK" -Level SUCCESS
        return $true
    }

    foreach ($issue in $issues) {
        Update-Status "Dependency issue: $issue" -Level WARNING
    }
    return $false
}

# ============================================================================
# REPAIR VERIFICATION
# ============================================================================

function Get-RepairVerification {
    <#
    .SYNOPSIS
        Post-run verification using Get-MpComputerStatus and Get-NetFirewallProfile.
        Returns a structured result with pass/fail for each assertion.
    #>
    $results = @()

    # Defender verification
    try {
        $mpStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
        if ($mpStatus) {
            $results += @{ Component = 'RealTimeProtection'; Expected = $true; Actual = $mpStatus.RealTimeProtectionEnabled; Pass = $mpStatus.RealTimeProtectionEnabled }
            $results += @{ Component = 'AntivirusEnabled'; Expected = $true; Actual = $mpStatus.AntivirusEnabled; Pass = $mpStatus.AntivirusEnabled }
            $results += @{ Component = 'BehaviorMonitor'; Expected = $true; Actual = $mpStatus.BehaviorMonitorEnabled; Pass = $mpStatus.BehaviorMonitorEnabled }
            $results += @{ Component = 'OnAccessProtection'; Expected = $true; Actual = $mpStatus.OnAccessProtectionEnabled; Pass = $mpStatus.OnAccessProtectionEnabled }
            $results += @{ Component = 'IoavProtection'; Expected = $true; Actual = $mpStatus.IoavProtectionEnabled; Pass = $mpStatus.IoavProtectionEnabled }

            $defAge = (New-TimeSpan -Start $mpStatus.AntivirusSignatureLastUpdated -End (Get-Date)).Days
            $results += @{ Component = 'DefinitionsUpToDate'; Expected = '<=7 days'; Actual = "$defAge days"; Pass = ($defAge -le 7) }
        }
        else {
            $results += @{ Component = 'DefenderStatus'; Expected = 'Available'; Actual = 'Unavailable'; Pass = $false }
        }
    }
    catch {
        $results += @{ Component = 'DefenderStatus'; Expected = 'Available'; Actual = 'Error'; Pass = $false }
    }

    # Firewall verification
    try {
        $fwProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
        foreach ($profile in $fwProfiles) {
            $results += @{
                Component = "Firewall_$($profile.Name)"
                Expected  = $true
                Actual    = $profile.Enabled
                Pass      = [bool]$profile.Enabled
            }
        }
    }
    catch {
        $results += @{ Component = 'FirewallProfiles'; Expected = 'Available'; Actual = 'Error'; Pass = $false }
    }

    return $results
}

function Show-RepairVerification {
    param([array]$Results)

    Update-Status "" -Level INFO
    Update-Status "=== POST-REPAIR VERIFICATION ===" -Level SECTION
    Update-Status ("{0,-25} {1,-15} {2,-15} {3}" -f 'COMPONENT', 'EXPECTED', 'ACTUAL', 'STATUS') -Level SECTION

    $passCount = 0
    $failCount = 0
    foreach ($r in $Results) {
        $status = if ($r.Pass) { 'PASS'; $passCount++ } else { 'FAIL'; $failCount++ }
        $level = if ($r.Pass) { 'SUCCESS' } else { 'ERROR' }
        $line = "{0,-25} {1,-15} {2,-15} {3}" -f $r.Component, $r.Expected, $r.Actual, $status
        Update-Status $line -Level $level
    }

    Update-Status "" -Level INFO
    Update-Status "Verification: $passCount passed, $failCount failed out of $($Results.Count) checks" -Level $(if ($failCount -eq 0) { 'SUCCESS' } else { 'WARNING' })

    # Set exit code based on verification
    if ($failCount -eq 0) {
        $Script:ExitCode = 0  # Full restoration
    }
    elseif ($passCount -gt 0) {
        $Script:ExitCode = 1  # Partial restoration
    }
    else {
        $Script:ExitCode = 2  # Failed
    }
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Set-RegistryValue {
    param(
        [string]$Path,
        [string]$Name,
        [object]$Value,
        [string]$Type = 'DWord'
    )

    try {
        $before = Get-RegistryValueState -Path $Path -Name $Name
        Add-UndoEntry -Type 'Registry' -Action 'SetValue' -Target "$Path\$Name" -Before $before -After @{ exists = $true; value = $Value; type = $Type } -Rollback @{ action = if ($before.exists) { 'SetValue' } else { 'RemoveValue' }; value = $before.value; type = $before.type } -Status $(if ($Script:IsDryRun) { 'DryRun' } else { 'Applied' }) | Out-Null

        if ($Script:IsDryRun) {
            return $true
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null
        }
        if ($before.exists) {
            Set-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -Force -ErrorAction SilentlyContinue
        }
        else {
            New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType $Type -Force -ErrorAction SilentlyContinue | Out-Null
        }
        return $true
    }
    catch {
        return $false
    }
}

function Remove-RegistryValue {
    param(
        [string]$Path,
        [string]$Name
    )

    try {
        $before = Get-RegistryValueState -Path $Path -Name $Name
        if ($before.exists) {
            Add-UndoEntry -Type 'Registry' -Action 'RemoveValue' -Target "$Path\$Name" -Before $before -After @{ exists = $false } -Rollback @{ action = 'SetValue'; value = $before.value; type = $before.type } -Status $(if ($Script:IsDryRun) { 'DryRun' } else { 'Applied' }) | Out-Null
            if (-not $Script:IsDryRun) {
                Remove-ItemProperty -LiteralPath $Path -Name $Name -Force -ErrorAction SilentlyContinue
            }
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

function Backup-RegistryKey {
    param(
        [string]$KeyPath,
        [string]$BackupName
    )

    try {
        Ensure-BackupDirectory | Out-Null

        $exportPath = Join-Path $Script:Config.BackupPath "$BackupName.reg"
        $regPath = $KeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\' -replace '^HKCU:\\', 'HKEY_CURRENT_USER\'

        $null = reg export $regPath $exportPath /y 2>&1
        return $exportPath
    }
    catch {
        return $false
    }
}

function Get-RegistryValueState {
    param(
        [string]$Path,
        [string]$Name
    )

    $state = [ordered]@{
        exists = $false
        value  = $null
        type   = $null
    }

    try {
        if (Test-Path -LiteralPath $Path) {
            $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $item -and $item.PSObject.Properties.Name -contains $Name) {
                $state.exists = $true
                $state.value = $item.$Name
                if ($null -ne $state.value) {
                    $state.type = $state.value.GetType().Name
                }
            }
        }
    }
    catch { }

    return $state
}

function Remove-RegistryKeyTree {
    param(
        [string]$Path,
        [string]$BackupName,
        [string]$Label = $Path
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return $false }
        $backupPath = Backup-RegistryKey -KeyPath $Path -BackupName $BackupName
        Add-UndoEntry -Type 'Registry' -Action 'RemoveKeyTree' -Target $Path -Before @{ exists = $true; backup = $backupPath } -After @{ exists = $false } -Rollback @{ action = 'reg import'; file = $backupPath } -Status $(if ($Script:IsDryRun) { 'DryRun' } else { 'Applied' }) | Out-Null
        if ($Script:IsDryRun) { return $true }

        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
        Update-Status "Removed: $Label" -Level SUCCESS
        return $true
    }
    catch {
        Update-Status "Could not remove: $Label (continuing...)" -Level WARNING
        return $false
    }
}

function Get-ServiceStartValue {
    param([string]$ServiceName)
    try {
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
        if (Test-Path -LiteralPath $regPath) {
            return (Get-ItemProperty -LiteralPath $regPath -Name 'Start' -ErrorAction SilentlyContinue).Start
        }
    }
    catch { }
    return $null
}

function Set-ServiceStartValue {
    param(
        [string]$ServiceName,
        [string]$DisplayName,
        [int]$StartValue,
        [string]$StartType
    )

    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
    $before = Get-ServiceStartValue -ServiceName $ServiceName
    Add-UndoEntry -Type 'Service' -Action 'SetStartType' -Target $ServiceName -Before @{ start = $before } -After @{ start = $StartValue; startType = $StartType } -Rollback @{ action = 'SetStart'; value = $before } -Status $(if ($Script:IsDryRun) { 'DryRun' } else { 'Applied' }) | Out-Null

    if ($Script:IsDryRun) {
        Update-Status "[DRY-RUN] Would set $regPath\Start = $StartValue" -Level INFO
        [void]$Script:DryRunActions.Add("Set $regPath\Start = $StartValue")
        return $true
    }

    try {
        if (Test-Path -LiteralPath $regPath) {
            Set-ItemProperty -LiteralPath $regPath -Name 'Start' -Value $StartValue -Force -ErrorAction SilentlyContinue
            Update-Status "${DisplayName}: Registry repaired" -Level SUCCESS
        }

        $scStart = switch ($StartType) {
            'Automatic' { 'auto' }
            'AutomaticDelayed' { 'delayed-auto' }
            'Manual' { 'demand' }
            'Disabled' { 'disabled' }
            default { $null }
        }
        if ($scStart) {
            $null = sc.exe config $ServiceName start= $scStart 2>&1
        }
        return $true
    }
    catch {
        Update-Status "${DisplayName}: Could not repair (continuing...)" -Level WARNING
        return $false
    }
}

function Start-ServiceQuiet {
    param(
        [string]$ServiceName,
        [string]$DisplayName
    )

    try {
        $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if (-not $svc) {
            Update-Status "$ServiceName : Service missing" -Level WARNING
            return $false
        }
        if ($svc.Status -eq 'Running') {
            Update-Status "${DisplayName}: Already running" -Level SUCCESS
            return $true
        }

        Add-UndoEntry -Type 'Service' -Action 'Start' -Target $ServiceName -Before @{ status = [string]$svc.Status } -After @{ status = 'Running' } -Rollback @{ action = 'StopIfOriginallyStopped' } -Status $(if ($Script:IsDryRun) { 'DryRun' } else { 'Applied' }) | Out-Null

        if ($Script:IsDryRun) {
            Update-Status "[DRY-RUN] Would start service: $DisplayName" -Level INFO
            [void]$Script:DryRunActions.Add("Start service: $ServiceName")
            return $true
        }

        $null = sc.exe start $ServiceName 2>&1
        Start-Sleep -Milliseconds 500
        $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') {
            Update-Status "${DisplayName}: Started" -Level SUCCESS
            return $true
        }

        Update-Status "${DisplayName}: Could not start (may need reboot)" -Level WARNING
        return $false
    }
    catch {
        Update-Status "$ServiceName : Error starting (continuing...)" -Level WARNING
        return $false
    }
}

# ============================================================================
# FIREWALL REPAIR FUNCTIONS
# ============================================================================

function Repair-FirewallServices {
    if (-not (Test-PhaseActive 'Services')) { return }
    Update-Status "Repairing Firewall Services..." -Level SECTION

    foreach ($svc in $Script:FirewallServices) {
        $startValue = switch ($svc.StartType) {
            'Automatic' { 2 }
            'Manual' { 3 }
            'Disabled' { 4 }
            'Boot' { 0 }
            'System' { 1 }
            default { 2 }
        }
        Set-ServiceStartValue -ServiceName $svc.Name -DisplayName $svc.DisplayName -StartValue $startValue -StartType $svc.StartType | Out-Null
    }
}

function Repair-FirewallRegistry {
    if (-not (Test-PhaseActive 'Registry')) { return }
    Update-Status "Removing Firewall Blocking Policies..." -Level SECTION

    $policiesToRemove = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall',
        'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile',
        'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile',
        'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile',
        'HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\StandardProfile'
    )

    foreach ($policy in $policiesToRemove) {
        try {
            if (Test-Path $policy) {
                if ($Script:IsDryRun) {
                    Update-Status "[DRY-RUN] Would remove: $policy" -Level INFO
                    [void]$Script:DryRunActions.Add("Remove registry key: $policy")
                }
                else {
                    $backupName = "FirewallPolicy_$($policy -replace '[^A-Za-z0-9]+', '_')"
                    Remove-RegistryKeyTree -Path $policy -BackupName $backupName -Label $policy | Out-Null
                }
            }
        }
        catch {
            Update-Status "Could not remove: $policy (continuing...)" -Level WARNING
        }
    }

    # Reset profile settings
    $profilePaths = @(
        'HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile',
        'HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\PublicProfile',
        'HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile'
    )

    foreach ($profilePath in $profilePaths) {
        try {
            if (Test-Path $profilePath) {
                if ($Script:IsDryRun) {
                    Update-Status "[DRY-RUN] Would set $profilePath\EnableFirewall = 1" -Level INFO
                    [void]$Script:DryRunActions.Add("Set $profilePath\EnableFirewall = 1")
                }
                else {
                    Set-RegistryValue -Path $profilePath -Name 'EnableFirewall' -Value 1 -Type DWord | Out-Null
                }
            }
        }
        catch { }
    }

    Update-Status "Firewall registry cleanup complete" -Level SUCCESS
}

function Start-FirewallServices {
    if (-not (Test-PhaseActive 'Services')) { return }
    Update-Status "Starting Firewall Services..." -Level SECTION

    $startOrder = @('BFE', 'mpssvc', 'IKEEXT', 'PolicyAgent')

    foreach ($svcName in $startOrder) {
        $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        $display = if ($svc) { $svc.DisplayName } else { $svcName }
        Start-ServiceQuiet -ServiceName $svcName -DisplayName $display | Out-Null
    }
}

function Test-CustomFirewallRule {
    param([object]$Rule)

    if (-not $Rule) { return $false }
    if ($Rule.PolicyStoreSourceType -and $Rule.PolicyStoreSourceType -notin @('PersistentStore', 'None')) { return $false }
    if ($Rule.Group -and $Rule.Group -like '@FirewallAPI.dll,*') { return $false }
    if ($Rule.DisplayGroup -and $Rule.DisplayGroup -like '@FirewallAPI.dll,*') { return $false }
    return $true
}

function Export-CustomFirewallRules {
    if (-not (Test-PhaseActive 'Features')) { return @() }

    Update-Status "Preserving custom firewall rules..." -Level SECTION

    if ($Script:IsDryRun) {
        Update-Status "[DRY-RUN] Would export firewall policy and custom rules before reset" -Level INFO
        [void]$Script:DryRunActions.Add("Export firewall policy and custom rules")
        return @()
    }

    Ensure-BackupDirectory | Out-Null
    $policyPath = Join-Path $Script:Config.BackupPath 'firewall-policy-before-reset.wfw'
    $rulesPath = Join-Path $Script:Config.BackupPath 'custom-firewall-rules.json'
    $customRules = @()

    try {
        $null = netsh advfirewall export "$policyPath" 2>&1
        Add-UndoEntry -Type 'Firewall' -Action 'ExportPolicyBackup' -Target $policyPath -Before @{ exists = $false } -After @{ exists = (Test-Path -LiteralPath $policyPath) } -Rollback @{ action = 'netsh advfirewall import'; file = $policyPath } -Status 'Applied' | Out-Null
    }
    catch {
        Update-Status "Could not export full firewall policy backup" -Level WARNING
    }

    try {
        $rules = Get-NetFirewallRule -PolicyStore ActiveStore -ErrorAction SilentlyContinue | Where-Object { Test-CustomFirewallRule -Rule $_ }
        foreach ($rule in $rules) {
            try {
                $address = $rule | Get-NetFirewallAddressFilter -ErrorAction SilentlyContinue
                $port = $rule | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue
                $app = $rule | Get-NetFirewallApplicationFilter -ErrorAction SilentlyContinue
                $service = $rule | Get-NetFirewallServiceFilter -ErrorAction SilentlyContinue

                $customRules += [ordered]@{
                    Name          = $rule.Name
                    DisplayName   = $rule.DisplayName
                    Description   = $rule.Description
                    Direction     = [string]$rule.Direction
                    Action        = [string]$rule.Action
                    Enabled       = [string]$rule.Enabled
                    Profile       = [string]$rule.Profile
                    Program       = if ($app) { [string]$app.Program } else { $null }
                    Service       = if ($service) { [string]$service.Service } else { $null }
                    Protocol      = if ($port) { [string]$port.Protocol } else { 'Any' }
                    LocalPort     = if ($port) { [string]$port.LocalPort } else { 'Any' }
                    RemotePort    = if ($port) { [string]$port.RemotePort } else { 'Any' }
                    LocalAddress  = if ($address) { [string]$address.LocalAddress } else { 'Any' }
                    RemoteAddress = if ($address) { [string]$address.RemoteAddress } else { 'Any' }
                }
            }
            catch { }
        }

        @($customRules) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $rulesPath -Encoding UTF8
        Add-UndoEntry -Type 'Firewall' -Action 'ExportCustomRules' -Target $rulesPath -Before @{ count = 0 } -After @{ count = $customRules.Count } -Rollback @{ action = 'RecreateRulesFromJson'; file = $rulesPath } -Status 'Applied' | Out-Null
        Update-Status "Custom firewall rules exported: $($customRules.Count)" -Level SUCCESS
    }
    catch {
        Update-Status "Could not export custom firewall rules" -Level WARNING
    }

    return @($customRules)
}

function Restore-CustomFirewallRules {
    param([array]$Rules)

    if (-not $Rules -or $Rules.Count -eq 0) {
        Update-Status "No custom firewall rules to re-import" -Level INFO
        return
    }

    Update-Status "Re-importing custom firewall rules..." -Level SECTION
    $restored = 0
    foreach ($rule in $Rules) {
        try {
            $existing = Get-NetFirewallRule -DisplayName $rule.DisplayName -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($existing) {
                continue
            }

            $params = @{
                DisplayName = $rule.DisplayName
                Direction   = $rule.Direction
                Action      = $rule.Action
                Enabled     = $rule.Enabled
                Profile     = $rule.Profile
            }
            if ($rule.Description) { $params.Description = $rule.Description }
            if ($rule.Program -and $rule.Program -ne 'Any') { $params.Program = $rule.Program }
            if ($rule.Service -and $rule.Service -ne 'Any') { $params.Service = $rule.Service }
            if ($rule.Protocol -and $rule.Protocol -ne 'Any') { $params.Protocol = $rule.Protocol }
            if ($rule.LocalPort -and $rule.LocalPort -ne 'Any') { $params.LocalPort = $rule.LocalPort }
            if ($rule.RemotePort -and $rule.RemotePort -ne 'Any') { $params.RemotePort = $rule.RemotePort }
            if ($rule.LocalAddress -and $rule.LocalAddress -ne 'Any') { $params.LocalAddress = $rule.LocalAddress }
            if ($rule.RemoteAddress -and $rule.RemoteAddress -ne 'Any') { $params.RemoteAddress = $rule.RemoteAddress }

            New-NetFirewallRule @params -ErrorAction Stop | Out-Null
            Add-UndoEntry -Type 'Firewall' -Action 'RestoreCustomRule' -Target $rule.DisplayName -Before @{ exists = $false } -After @{ exists = $true } -Rollback @{ action = 'Remove-NetFirewallRule'; displayName = $rule.DisplayName } -Status 'Applied' | Out-Null
            $restored++
        }
        catch {
            Update-Status "Could not re-import firewall rule: $($rule.DisplayName)" -Level WARNING
        }
    }

    Update-Status "Custom firewall rules re-imported: $restored" -Level SUCCESS
}

function Enable-FirewallProfiles {
    if (-not (Test-PhaseActive 'Features')) { return }
    Update-Status "Enabling Firewall Profiles..." -Level SECTION

    if ($Script:IsDryRun) {
        Update-Status "[DRY-RUN] Would enable all firewall profiles (Domain, Public, Private)" -Level INFO
        Update-Status "[DRY-RUN] Would reset firewall to defaults" -Level INFO
        [void]$Script:DryRunActions.Add("Enable firewall profiles: Domain, Public, Private")
        [void]$Script:DryRunActions.Add("Reset firewall to defaults via netsh")
        return
    }

    $customRules = Export-CustomFirewallRules

    try {
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True -ErrorAction SilentlyContinue
        Add-UndoEntry -Type 'Firewall' -Action 'EnableProfiles' -Target 'Domain,Public,Private' -Before @{ enabled = 'Unknown' } -After @{ enabled = $true } -Rollback @{ action = 'ManualReview'; backup = (Join-Path $Script:Config.BackupPath 'firewall-policy-before-reset.wfw') } -Status 'Applied' | Out-Null
        Update-Status "All firewall profiles enabled" -Level SUCCESS
    }
    catch {
        # Fallback to netsh
        try {
            $null = netsh advfirewall set domainprofile state on 2>&1
            $null = netsh advfirewall set privateprofile state on 2>&1
            $null = netsh advfirewall set publicprofile state on 2>&1
            Update-Status "Firewall profiles enabled via netsh" -Level SUCCESS
        }
        catch {
            Update-Status "Could not enable profiles (may need reboot)" -Level WARNING
        }
    }

    # Reset to defaults
    try {
        Add-UndoEntry -Type 'Firewall' -Action 'ResetPolicy' -Target 'advfirewall' -Before @{ backup = (Join-Path $Script:Config.BackupPath 'firewall-policy-before-reset.wfw') } -After @{ reset = $true } -Rollback @{ action = 'netsh advfirewall import'; file = (Join-Path $Script:Config.BackupPath 'firewall-policy-before-reset.wfw') } -Status 'Applied' | Out-Null
        $null = netsh advfirewall reset 2>&1
        Update-Status "Firewall reset to defaults" -Level SUCCESS
        Restore-CustomFirewallRules -Rules $customRules
    }
    catch { }
}

# ============================================================================
# DEFENDER REPAIR FUNCTIONS
# ============================================================================

function Repair-DefenderServices {
    if (-not (Test-PhaseActive 'Services')) { return }
    Update-Status "Repairing Defender Services..." -Level SECTION

    foreach ($svc in $Script:DefenderServices) {
        $startValue = switch ($svc.StartType) {
            'Automatic' { 2 }
            'Manual' { 3 }
            'Disabled' { 4 }
            'Boot' { 0 }
            'System' { 1 }
            default { 3 }
        }
        Set-ServiceStartValue -ServiceName $svc.Name -DisplayName $svc.DisplayName -StartValue $startValue -StartType $svc.StartType | Out-Null
    }

    # Repair drivers
    $drivers = @(
        @{ Name = 'WdFilter'; Start = 0 },
        @{ Name = 'WdNisDrv'; Start = 3 },
        @{ Name = 'WdBoot'; Start = 0 }
    )

    foreach ($driver in $drivers) {
        try {
            Set-ServiceStartValue -ServiceName $driver.Name -DisplayName $driver.Name -StartValue $driver.Start -StartType $(if ($driver.Start -eq 0) { 'Boot' } else { 'Manual' }) | Out-Null
        }
        catch { }
    }
}

function Repair-DefenderRegistry {
    if (-not (Test-PhaseActive 'Registry')) { return }
    Update-Status "Removing Defender Blocking Policies..." -Level SECTION

    # Backup first (skip in dry-run)
    if (-not $Script:IsDryRun) {
        Backup-RegistryKey -KeyPath 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender' -BackupName 'Policies_WindowsDefender'
        Backup-RegistryKey -KeyPath 'HKLM:\SOFTWARE\Microsoft\Windows Defender' -BackupName 'WindowsDefender'
    }

    $disablingValues = @(
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableAntiSpyware' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableAntiVirus' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableRoutinelyTakingAction' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'ServiceKeepAlive' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableBehaviorMonitoring' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableOnAccessProtection' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableScanOnRealtimeEnable' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableRealtimeMonitoring' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableIOAVProtection' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet'; Name = 'SpynetReporting' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet'; Name = 'SubmitSamplesConsent' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\MpEngine'; Name = 'MpEnablePus' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender'; Name = 'DisableAntiSpyware' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender'; Name = 'DisableAntiVirus' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableRealtimeMonitoring' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableBehaviorMonitoring' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableOnAccessProtection' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableScanOnRealtimeEnable' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableIOAVProtection' }
    )

    $removed = 0
    foreach ($item in $disablingValues) {
        try {
            if (Test-Path $item.Path) {
                $value = Get-ItemProperty -Path $item.Path -Name $item.Name -ErrorAction SilentlyContinue
                if ($null -ne $value) {
                    if ($Script:IsDryRun) {
                        Update-Status "[DRY-RUN] Would remove $($item.Path)\$($item.Name)" -Level INFO
                        [void]$Script:DryRunActions.Add("Remove $($item.Path)\$($item.Name)")
                    }
                    else {
                        Remove-RegistryValue -Path $item.Path -Name $item.Name | Out-Null
                    }
                    $removed++
                }
            }
        }
        catch {
            if (-not $Script:IsDryRun) {
                # Try setting to 0 instead
                try {
                    Set-RegistryValue -Path $item.Path -Name $item.Name -Value 0 -Type DWord | Out-Null
                }
                catch { }
            }
        }
    }

    Update-Status "Removed/reset $removed blocking policies" -Level SUCCESS

    # Remove policy trees
    $policyTrees = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Policy Manager',
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\UX Configuration'
    )

    foreach ($tree in $policyTrees) {
        try {
            if (Test-Path $tree) {
                if ($Script:IsDryRun) {
                    Update-Status "[DRY-RUN] Would remove policy tree: $tree" -Level INFO
                    [void]$Script:DryRunActions.Add("Remove policy tree: $tree")
                }
                else {
                    $backupName = "DefenderPolicy_$($tree -replace '[^A-Za-z0-9]+', '_')"
                    Remove-RegistryKeyTree -Path $tree -BackupName $backupName -Label $tree | Out-Null
                }
            }
        }
        catch { }
    }
}

function Repair-DefenderScheduledTasks {
    if (-not (Test-PhaseActive 'Tasks')) { return }
    Update-Status "Repairing Defender Scheduled Tasks..." -Level SECTION

    $defenderTasks = @(
        'Windows Defender Cache Maintenance',
        'Windows Defender Cleanup',
        'Windows Defender Scheduled Scan',
        'Windows Defender Verification'
    )

    foreach ($taskName in $defenderTasks) {
        try {
            $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($task -and $task.State -eq 'Disabled') {
                if ($Script:IsDryRun) {
                    Update-Status "[DRY-RUN] Would enable: $taskName" -Level INFO
                    [void]$Script:DryRunActions.Add("Enable scheduled task: $taskName")
                }
                else {
                    Add-UndoEntry -Type 'ScheduledTask' -Action 'Enable' -Target $taskName -Before @{ state = 'Disabled' } -After @{ state = 'Enabled' } -Rollback @{ action = 'Disable-ScheduledTask'; taskName = $taskName } -Status 'Applied' | Out-Null
                    Enable-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
                    Update-Status "Enabled: $taskName" -Level SUCCESS
                }
            }
            elseif ($task) {
                Update-Status "Already enabled: $taskName" -Level SUCCESS
            }
        }
        catch {
            Update-Status "Could not enable: $taskName (continuing...)" -Level WARNING
        }
    }

    # Remove malicious tasks
    try {
        $suspiciousTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
            $_.TaskName -match 'DisableDefender|DisableWinDefend|KillDefender|StopDefender'
        }

        foreach ($task in $suspiciousTasks) {
            try {
                if ($Script:IsDryRun) {
                    Update-Status "[DRY-RUN] Would remove malicious task: $($task.TaskName)" -Level INFO
                    [void]$Script:DryRunActions.Add("Remove malicious task: $($task.TaskName)")
                }
                else {
                    Add-UndoEntry -Type 'ScheduledTask' -Action 'Remove' -Target $task.TaskName -Before @{ exists = $true; state = [string]$task.State; path = $task.TaskPath } -After @{ exists = $false } -Rollback @{ action = 'ManualRecreateRequired'; taskName = $task.TaskName; taskPath = $task.TaskPath } -Status 'Applied' | Out-Null
                    Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                    Update-Status "Removed malicious task: $($task.TaskName)" -Level SUCCESS
                }
            }
            catch { }
        }
    }
    catch { }
}

function Get-DefenderWmiSubscriptions {
    $subscriptions = @()

    try {
        $filters = Get-WmiObject -Query "SELECT * FROM __EventFilter WHERE Name LIKE '%Defender%' OR Name LIKE '%WinDefend%' OR Query LIKE '%MsMpEng%' OR Query LIKE '%WinDefend%'" -Namespace 'root\subscription' -ErrorAction SilentlyContinue
        foreach ($filter in $filters) {
            $filterRef = "__EventFilter.Name='$($filter.Name)'"
            $bindings = Get-WmiObject -Query "SELECT * FROM __FilterToConsumerBinding WHERE Filter=""$filterRef""" -Namespace 'root\subscription' -ErrorAction SilentlyContinue
            $subscriptions += [ordered]@{
                filter   = $filter
                bindings = @($bindings)
            }
        }
    }
    catch { }

    return $subscriptions
}

function Export-WmiRemovalReport {
    try {
        Ensure-BackupDirectory | Out-Null
        $reportPath = Join-Path $Script:Config.BackupPath 'wmi-subscription-report.json'
        @($Script:WmiRemovalReport) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding UTF8
        $Script:JsonOutput.reports['wmi'] = $reportPath
        Update-Status "WMI subscription report: $reportPath" -Level SUCCESS
        return $reportPath
    }
    catch {
        Update-Status "Could not write WMI report: $($_.Exception.Message)" -Level WARNING
        return $null
    }
}

function Get-DefenderControlUndoArtifacts {
    $roots = @(
        "$env:USERPROFILE\Desktop",
        "$env:USERPROFILE\Downloads",
        "$env:TEMP",
        "${env:ProgramFiles}\DefenderControl",
        "${env:ProgramFiles(x86)}\DefenderControl"
    )

    $artifacts = @()
    foreach ($root in $roots) {
        try {
            if (-not (Test-Path -LiteralPath $root)) { continue }

            $files = Get-ChildItem -LiteralPath $root -File -ErrorAction SilentlyContinue | Where-Object {
                $_.FullName -match 'dControl|DefenderControl|Defender' -and $_.Name -match 'enable|restore|backup|undo|defender'
            }
            foreach ($file in $files) {
                $artifacts += $file.FullName
            }

            $candidateDirs = Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object {
                $_.Name -match 'dControl|DefenderControl'
            } | Select-Object -First 5
            foreach ($dir in $candidateDirs) {
                $files = Get-ChildItem -LiteralPath $dir.FullName -Recurse -File -ErrorAction SilentlyContinue -Include '*.reg', '*.bak', '*.backup' | Where-Object {
                    $_.Name -match 'enable|restore|backup|undo|defender'
                } | Select-Object -First 20
                foreach ($file in $files) {
                    $artifacts += $file.FullName
                }
            }
        }
        catch { }
    }

    return @($artifacts | Select-Object -Unique)
}

function Invoke-DefenderControlUndo {
    if (-not (Test-PhaseActive 'Registry')) { return $false }

    $artifacts = @(Get-DefenderControlUndoArtifacts)
    if ($artifacts.Count -eq 0) { return $false }

    Update-Status "DefenderControl undo artifacts found: $($artifacts.Count)" -Level SECTION
    $replayed = 0
    foreach ($artifact in $artifacts) {
        try {
            if ($Script:IsDryRun) {
                Update-Status "[DRY-RUN] Would replay DefenderControl artifact: $artifact" -Level INFO
                [void]$Script:DryRunActions.Add("Replay DefenderControl artifact: $artifact")
                $replayed++
                continue
            }

            if ([System.IO.Path]::GetExtension($artifact) -ieq '.reg') {
                Add-UndoEntry -Type 'DefenderControl' -Action 'ReplayUndoArtifact' -Target $artifact -Before @{ replayed = $false } -After @{ replayed = $true } -Rollback @{ action = 'ManualReview'; source = $artifact } -Status 'Applied' | Out-Null
                $null = reg import "$artifact" 2>&1
                Update-Status "Replayed DefenderControl artifact: $artifact" -Level SUCCESS
                $replayed++
            }
        }
        catch {
            Update-Status "Could not replay DefenderControl artifact: $artifact" -Level WARNING
        }
    }

    return ($replayed -gt 0)
}

function Get-AppLockerMsMpEngBlockers {
    $blockers = @()
    $pattern = 'MsMpEng|WinDefend|Windows Defender|Microsoft Defender|Defender\\Platform'

    try {
        [xml]$policyXml = Get-AppLockerPolicy -Local -Xml -ErrorAction SilentlyContinue
        if ($policyXml) {
            $rules = $policyXml.SelectNodes("//*[local-name()='FilePathRule' or local-name()='FilePublisherRule' or local-name()='FileHashRule']")
            foreach ($rule in $rules) {
                if ($rule.Action -eq 'Deny' -and $rule.OuterXml -match $pattern) {
                    $blockers += [ordered]@{
                        Source = 'AppLocker'
                        Name   = $rule.Name
                        Id     = $rule.Id
                        Detail = $rule.OuterXml
                    }
                }
            }
        }
    }
    catch { }

    $srpRoots = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths',
        'HKCU:\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths'
    )
    foreach ($root in $srpRoots) {
        try {
            if (-not (Test-Path -LiteralPath $root)) { continue }
            $keys = Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue
            foreach ($key in $keys) {
                $item = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
                $itemData = [string]$item.ItemData
                if ($itemData -match $pattern) {
                    $providerPath = $key.Name -replace '^HKEY_LOCAL_MACHINE', 'HKLM:' -replace '^HKEY_CURRENT_USER', 'HKCU:'
                    $blockers += [ordered]@{
                        Source = 'SRP'
                        Name   = Split-Path -Leaf $key.Name
                        Id     = $providerPath
                        Detail = $itemData
                    }
                }
            }
        }
        catch { }
    }

    return $blockers
}

function Repair-AppLockerAndSRP {
    if (-not (Test-PhaseActive 'AppLocker')) { return }
    Update-Status "Checking AppLocker and SRP Defender blockers..." -Level SECTION

    $blockers = @(Get-AppLockerMsMpEngBlockers)
    if ($blockers.Count -eq 0) {
        Update-Status "No AppLocker/SRP Defender blockers found" -Level SUCCESS
        return
    }

    Update-Status "Found $($blockers.Count) AppLocker/SRP Defender blocker(s)" -Level WARNING

    try {
        [xml]$policyXml = Get-AppLockerPolicy -Local -Xml -ErrorAction SilentlyContinue
        if ($policyXml) {
            $pattern = 'MsMpEng|WinDefend|Windows Defender|Microsoft Defender|Defender\\Platform'
            $rules = @($policyXml.SelectNodes("//*[local-name()='FilePathRule' or local-name()='FilePublisherRule' or local-name()='FileHashRule']"))
            $removedRules = 0
            foreach ($rule in $rules) {
                if ($rule.Action -eq 'Deny' -and $rule.OuterXml -match $pattern) {
                    Update-Status "$(if ($Script:IsDryRun) { '[DRY-RUN] Would remove' } else { 'Removing' }) AppLocker deny rule: $($rule.Name)" -Level WARNING
                    if ($Script:IsDryRun) {
                        [void]$Script:DryRunActions.Add("Remove AppLocker deny rule: $($rule.Name)")
                    }
                    else {
                        [void]$rule.ParentNode.RemoveChild($rule)
                    }
                    $removedRules++
                }
            }

            if ($removedRules -gt 0 -and -not $Script:IsDryRun) {
                Ensure-BackupDirectory | Out-Null
                $backupXml = Join-Path $Script:Config.BackupPath 'applocker-policy-before-defender-repair.xml'
                $newXml = Join-Path $Script:Config.BackupPath 'applocker-policy-after-defender-repair.xml'
                [xml]$originalXml = Get-AppLockerPolicy -Local -Xml -ErrorAction SilentlyContinue
                $originalXml.Save($backupXml)
                $policyXml.Save($newXml)
                Add-UndoEntry -Type 'AppLocker' -Action 'RemoveDenyRules' -Target 'Local AppLocker Policy' -Before @{ backup = $backupXml; removed = $removedRules } -After @{ policy = $newXml } -Rollback @{ action = 'Set-AppLockerPolicy'; file = $backupXml } -Status 'Applied' | Out-Null
                Set-AppLockerPolicy -XMLPolicy $newXml -ErrorAction SilentlyContinue
                Update-Status "AppLocker deny rules removed: $removedRules" -Level SUCCESS
            }
        }
    }
    catch {
        Update-Status "Could not repair AppLocker policy (continuing...)" -Level WARNING
    }

    foreach ($blocker in $blockers | Where-Object { $_.Source -eq 'SRP' }) {
        try {
            if ($Script:IsDryRun) {
                Update-Status "[DRY-RUN] Would remove SRP rule: $($blocker.Detail)" -Level INFO
                [void]$Script:DryRunActions.Add("Remove SRP rule: $($blocker.Detail)")
            }
            else {
                $backupName = "SRP_$($blocker.Name -replace '[^A-Za-z0-9]+', '_')"
                Remove-RegistryKeyTree -Path $blocker.Id -BackupName $backupName -Label "SRP rule $($blocker.Detail)" | Out-Null
            }
        }
        catch { }
    }
}

function Repair-MDEComponents {
    if (-not (Test-PhaseActive 'MDE')) { return }
    Update-Status "Checking Microsoft Defender for Endpoint components..." -Level SECTION

    $isEnterprise = $false
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os.Caption -match 'Enterprise|Education') { $isEnterprise = $true }
    }
    catch { }

    foreach ($svc in $Script:MdeServices) {
        $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
        if (-not $service) {
            $level = if ($isEnterprise) { 'WARNING' } else { 'INFO' }
            Update-Status "$($svc.DisplayName): Service not present$(if ($isEnterprise) { ' - MDE onboarding may be missing' } else { '' })" -Level $level
            continue
        }

        $start = Get-ServiceStartValue -ServiceName $svc.Name
        if ($start -eq 4 -or $null -eq $start) {
            Set-ServiceStartValue -ServiceName $svc.Name -DisplayName $svc.DisplayName -StartValue $svc.StartValue -StartType $svc.StartType | Out-Null
        }
        else {
            Update-Status "$($svc.DisplayName): Start type is not disabled" -Level SUCCESS
        }
    }
}

function Repair-WindowsUpdateServices {
    if (-not (Test-PhaseActive 'WindowsUpdate')) { return }
    Update-Status "Repairing Windows Update services used by Defender signatures..." -Level SECTION

    foreach ($svc in $Script:WindowsUpdateServices) {
        $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
        if (-not $service) {
            Update-Status "$($svc.DisplayName): Service missing" -Level WARNING
            continue
        }

        $start = Get-ServiceStartValue -ServiceName $svc.Name
        if ($start -eq 4 -or $start -ne $svc.StartValue) {
            Set-ServiceStartValue -ServiceName $svc.Name -DisplayName $svc.DisplayName -StartValue $svc.StartValue -StartType $svc.StartType | Out-Null
        }
        else {
            Update-Status "$($svc.DisplayName): Start type OK" -Level SUCCESS
        }
    }

    foreach ($svcName in @('BITS', 'DoSvc')) {
        $service = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        if ($service) {
            Start-ServiceQuiet -ServiceName $svcName -DisplayName $service.DisplayName | Out-Null
        }
    }
}

function Repair-DefenderWMI {
    if (-not (Test-PhaseActive 'WMI')) { return }
    Update-Status "Checking WMI Subscriptions..." -Level SECTION

    try {
        $subscriptions = @(Get-DefenderWmiSubscriptions)

        if ($subscriptions.Count -gt 0) {
            foreach ($sub in $subscriptions) {
                $filter = $sub.filter
                [void]$Script:WmiRemovalReport.Add([ordered]@{
                    name     = $filter.Name
                    query    = $filter.Query
                    bindings = @($sub.bindings | ForEach-Object { $_.__RELPATH })
                    action   = if ($Script:IsDryRun) { 'WouldRemove' } else { 'Removed' }
                })

                try {
                    if ($Script:IsDryRun) {
                        Update-Status "[DRY-RUN] Would remove WMI subscription: $($filter.Name)" -Level INFO
                        [void]$Script:DryRunActions.Add("Remove WMI subscription: $($filter.Name)")
                    }
                    else {
                        Add-UndoEntry -Type 'WMI' -Action 'RemoveSubscription' -Target $filter.Name -Before @{ query = $filter.Query; relPath = $filter.__RELPATH; bindings = @($sub.bindings | ForEach-Object { $_.__RELPATH }) } -After @{ exists = $false } -Rollback @{ action = 'ManualRecreateRequired' } -Status 'Applied' | Out-Null
                        foreach ($binding in $sub.bindings) {
                            $binding | Remove-WmiObject -ErrorAction SilentlyContinue
                        }
                        $filter | Remove-WmiObject -ErrorAction SilentlyContinue
                        Update-Status "Removed WMI subscription: $($filter.Name)" -Level SUCCESS
                    }
                }
                catch { }
            }
        }
        else {
            Update-Status "No malicious WMI subscriptions found" -Level SUCCESS
        }

        Export-WmiRemovalReport | Out-Null
    }
    catch {
        Update-Status "Could not query WMI (continuing...)" -Level WARNING
    }
}

function Start-DefenderServices {
    if (-not (Test-PhaseActive 'Services')) { return }
    Update-Status "Starting Defender Services..." -Level SECTION

    # Start Security Center first
    try {
        $secCenter = Get-Service -Name 'wscsvc' -ErrorAction SilentlyContinue
        if ($secCenter -and $secCenter.Status -ne 'Running') {
            Set-ServiceStartValue -ServiceName 'wscsvc' -DisplayName 'Security Center' -StartValue 2 -StartType 'Automatic' | Out-Null
            Start-ServiceQuiet -ServiceName 'wscsvc' -DisplayName 'Security Center' | Out-Null
        }
    }
    catch { }

    foreach ($svc in $Script:DefenderServices) {
        Start-ServiceQuiet -ServiceName $svc.Name -DisplayName $svc.DisplayName | Out-Null
    }
}

function Enable-DefenderFeatures {
    if (-not (Test-PhaseActive 'Features')) { return }
    Update-Status "Enabling Defender Protection Features..." -Level SECTION

    if ($Script:IsDryRun) {
        $features = @(
            'DisableRealtimeMonitoring=false', 'DisableBehaviorMonitoring=false',
            'DisableBlockAtFirstSeen=false', 'DisableIOAVProtection=false',
            'DisablePrivacyMode=false', 'DisableScriptScanning=false',
            'DisableArchiveScanning=false', 'DisableIntrusionPreventionSystem=false',
            'MAPSReporting=Advanced', 'SubmitSamplesConsent=SendAllSamples',
            'PUAProtection=Enabled'
        )
        foreach ($f in $features) {
            Update-Status "[DRY-RUN] Would set Set-MpPreference -$f" -Level INFO
            [void]$Script:DryRunActions.Add("Set-MpPreference -$f")
        }
        Update-Status "[DRY-RUN] Would update virus definitions" -Level INFO
        [void]$Script:DryRunActions.Add("Update-MpSignature")
        return
    }

    try {
        $beforePrefs = $null
        try {
            $beforePrefs = Get-MpPreference -ErrorAction SilentlyContinue | Select-Object DisableRealtimeMonitoring, DisableBehaviorMonitoring, DisableBlockAtFirstSeen, DisableIOAVProtection, DisablePrivacyMode, DisableScriptScanning, DisableArchiveScanning, DisableIntrusionPreventionSystem, MAPSReporting, SubmitSamplesConsent, PUAProtection
        }
        catch { }
        Add-UndoEntry -Type 'DefenderPreference' -Action 'SetProtectionFeatures' -Target 'Set-MpPreference' -Before $beforePrefs -After @{ protectionEnabled = $true } -Rollback @{ action = 'Set-MpPreference'; values = $beforePrefs } -Status 'Applied' | Out-Null

        Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction SilentlyContinue
        Set-MpPreference -DisableBehaviorMonitoring $false -ErrorAction SilentlyContinue
        Set-MpPreference -DisableBlockAtFirstSeen $false -ErrorAction SilentlyContinue
        Set-MpPreference -DisableIOAVProtection $false -ErrorAction SilentlyContinue
        Set-MpPreference -DisablePrivacyMode $false -ErrorAction SilentlyContinue
        Set-MpPreference -DisableScriptScanning $false -ErrorAction SilentlyContinue
        Set-MpPreference -DisableArchiveScanning $false -ErrorAction SilentlyContinue
        Set-MpPreference -DisableIntrusionPreventionSystem $false -ErrorAction SilentlyContinue
        Set-MpPreference -MAPSReporting Advanced -ErrorAction SilentlyContinue
        Set-MpPreference -SubmitSamplesConsent SendAllSamples -ErrorAction SilentlyContinue
        Set-MpPreference -PUAProtection Enabled -ErrorAction SilentlyContinue

        Update-Status "Protection features enabled" -Level SUCCESS
    }
    catch {
        Update-Status "Some features could not be enabled (continuing...)" -Level WARNING
    }

    # Update signatures
    try {
        Update-Status "Updating virus definitions..."
        Update-MpSignature -ErrorAction SilentlyContinue
        Update-Status "Definitions updated" -Level SUCCESS
    }
    catch {
        Update-Status "Could not update definitions (try manually later)" -Level WARNING
    }
}

function Reset-GroupPolicy {
    if (-not (Test-PhaseActive 'GroupPolicy')) { return }
    Update-Status "Resetting Group Policy..." -Level SECTION

    try {
        $machinePolPath = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"

        if (Test-Path $machinePolPath) {
            if ($Script:IsDryRun) {
                Update-Status "[DRY-RUN] Would backup and remove $machinePolPath" -Level INFO
                Update-Status "[DRY-RUN] Would run gpupdate /force" -Level INFO
                [void]$Script:DryRunActions.Add("Remove $machinePolPath")
                [void]$Script:DryRunActions.Add("gpupdate /force")
                return
            }
            $backupPol = Join-Path $Script:Config.BackupPath "Registry.pol"
            Ensure-BackupDirectory | Out-Null
            Copy-Item -Path $machinePolPath -Destination $backupPol -Force -ErrorAction SilentlyContinue
            Add-UndoEntry -Type 'GroupPolicy' -Action 'RemoveRegistryPol' -Target $machinePolPath -Before @{ exists = $true; backup = $backupPol } -After @{ exists = $false } -Rollback @{ action = 'CopyBack'; file = $backupPol } -Status 'Applied' | Out-Null
            Remove-Item -Path $machinePolPath -Force -ErrorAction SilentlyContinue
            Update-Status "Machine policy file removed" -Level SUCCESS
        }

        if (-not $Script:IsDryRun) {
            $null = gpupdate /force 2>&1
            Update-Status "Group Policy refreshed" -Level SUCCESS
        }
    }
    catch {
        Update-Status "Could not reset Group Policy (continuing...)" -Level WARNING
    }
}

function Repair-WindowsSecurity {
    if (-not (Test-PhaseActive 'Features')) { return }
    Update-Status "Repairing Windows Security App..." -Level SECTION

    if ($Script:IsDryRun) {
        Update-Status "[DRY-RUN] Would re-register Microsoft.SecHealthUI AppxPackage" -Level INFO
        [void]$Script:DryRunActions.Add("Re-register Microsoft.SecHealthUI AppxPackage")
        return
    }

    try {
        $pkg = Get-AppxPackage -Name 'Microsoft.SecHealthUI' -ErrorAction SilentlyContinue
        if ($pkg) {
            Add-UndoEntry -Type 'Appx' -Action 'RegisterPackage' -Target 'Microsoft.SecHealthUI' -Before @{ installLocation = $pkg.InstallLocation } -After @{ registered = $true } -Rollback @{ action = 'ManualReview' } -Status 'Applied' | Out-Null
            Add-AppxPackage -DisableDevelopmentMode -Register "$($pkg.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue
            Update-Status "Windows Security app re-registered" -Level SUCCESS
        }
        else {
            # Try AllUsers reprovisioning
            $pkg = Get-AppxPackage -AllUsers -Name 'Microsoft.SecHealthUI' -ErrorAction SilentlyContinue
            if ($pkg) {
                Add-UndoEntry -Type 'Appx' -Action 'RegisterPackage' -Target 'Microsoft.SecHealthUI AllUsers' -Before @{ installLocation = $pkg.InstallLocation } -After @{ registered = $true } -Rollback @{ action = 'ManualReview' } -Status 'Applied' | Out-Null
                Add-AppxPackage -DisableDevelopmentMode -Register "$($pkg.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue
                Update-Status "Windows Security app re-registered (AllUsers)" -Level SUCCESS
            }
            else {
                Update-Status "Windows Security app package not found - may need manual reinstall" -Level WARNING
            }
        }
    }
    catch {
        Update-Status "Could not re-register app (continuing...)" -Level WARNING
    }
}

# ============================================================================
# MAIN REPAIR FUNCTION
# ============================================================================

function Start-Repair {
    param(
        [bool]$RepairFirewall,
        [bool]$RepairDefender,
        [bool]$CreateRestorePoint
    )

    $startTime = Get-Date
    Initialize-UndoManifest

    # Capture pre-repair scan
    $Script:PreRepairScan = Get-HealthScan

    $modeLabel = if ($Script:IsDryRun) { ' [DRY-RUN]' } else { '' }
    Update-Status "DefenderShield v$($Script:Version) Repair Started$modeLabel" -Level SECTION
    Update-Status "Log file: $($Script:Config.LogPath)"

    # Pre-flight: third-party AV check
    if ($RepairDefender) {
        if (Test-ThirdPartyAVBlocking) {
            Save-UndoManifest | Out-Null
            Export-RepairReport | Out-Null
            return $false
        }
    }

    # Privacy tool detection
    Update-Status "" -Level INFO
    Update-Status "========== PRIVACY TOOL DETECTION ==========" -Level SECTION
    $privacyTools = Get-DetectedPrivacyTools
    if ($privacyTools -and $privacyTools.Count -gt 0) {
        Update-Status "Detected $($privacyTools.Count) privacy tool(s) that may have caused this breakage:" -Level WARNING
        foreach ($pt in $privacyTools) {
            Update-Status "  - $($pt.Tool): $($pt.Evidence)" -Level WARNING
        }
    }
    else {
        Update-Status "No known privacy tool signatures detected." -Level SUCCESS
    }

    # Policy source audit
    Update-Status "" -Level INFO
    Update-Status "========== POLICY SOURCE AUDIT ==========" -Level SECTION
    $blockers = Get-PolicySourceAudit
    Show-PolicySourceAudit -Blockers $blockers

    # Create restore point
    if ($CreateRestorePoint -and -not $Script:IsDryRun) {
        Update-Status "" -Level INFO
        Update-Status "Creating System Restore Point..." -Level SECTION
        try {
            Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
            Checkpoint-Computer -Description "DefenderShield - Before Repair" -RestorePointType MODIFY_SETTINGS -ErrorAction SilentlyContinue
            Update-Status "Restore point created" -Level SUCCESS
        }
        catch {
            Update-Status "Could not create restore point (continuing...)" -Level WARNING
        }
    }
    elseif ($CreateRestorePoint -and $Script:IsDryRun) {
        Update-Status "[DRY-RUN] Would create System Restore Point" -Level INFO
    }

    # Repair Firewall
    if ($RepairFirewall) {
        Update-Status ""
        Update-Status "========== FIREWALL REPAIR ==========" -Level SECTION

        # Validate dependency tree first
        Test-FirewallServiceDependencies | Out-Null

        Repair-FirewallServices
        Repair-FirewallRegistry
        Start-FirewallServices
        Enable-FirewallProfiles
    }

    # Repair Defender
    if ($RepairDefender) {
        Update-Status ""
        Update-Status "========== DEFENDER REPAIR ==========" -Level SECTION
        Repair-DefenderServices
        Invoke-DefenderControlUndo | Out-Null
        Repair-DefenderRegistry
        Repair-AppLockerAndSRP
        Repair-MDEComponents
        Repair-WindowsUpdateServices
        Repair-DefenderScheduledTasks
        Repair-DefenderWMI
        Reset-GroupPolicy
        Start-DefenderServices
        Enable-DefenderFeatures
        Repair-WindowsSecurity
    }

    # SmartScreen repair (runs for both Defender and Firewall modes)
    if (($RepairDefender -or $RepairFirewall) -and (Test-PhaseActive 'SmartScreen')) {
        Update-Status ""
        Update-Status "========== SMARTSCREEN REPAIR ==========" -Level SECTION
        Repair-SmartScreen
    }

    # Dry-run summary
    if ($Script:IsDryRun) {
        Update-Status "" -Level INFO
        Update-Status "========== DRY-RUN SUMMARY ==========" -Level SECTION
        Update-Status "$($Script:DryRunActions.Count) action(s) would be performed." -Level INFO
        Update-Status "No changes were made to the system." -Level SUCCESS
        $Script:ExitCode = 0
    }
    else {
        # Post-repair scan and comparison
        Start-Sleep -Seconds 2
        $postScan = Get-HealthScan

        # Update dashboard with new state (GUI only)
        if (-not $Script:CliMode) {
            Update-HealthDashboard -Scan $postScan
        }

        # Show before/after comparison
        $comparisonLines = Get-ComparisonReport -Before $Script:PreRepairScan -After $postScan

        Update-Status ""
        foreach ($line in $comparisonLines) {
            if ($line -match '\[FIXED\]') {
                Update-Status $line -Level SUCCESS
            }
            elseif ($line -match '\[UPDATED\]') {
                Update-Status $line -Level SUCCESS
            }
            elseif ($line -match '=== REPAIR RESULTS ===') {
                Update-Status $line -Level SECTION
            }
            else {
                Update-Status $line -Level INFO
            }
        }

        # Post-repair verification
        $verifyResults = Get-RepairVerification
        Show-RepairVerification -Results $verifyResults

        # Store post-scan for JSON
        $Script:JsonOutput.postScan = $postScan
    }

    # Complete
    $duration = (Get-Date) - $startTime

    Update-Status ""
    Update-Status "========== REPAIR COMPLETE ==========" -Level SECTION
    Update-Status "Duration: $([math]::Round($duration.TotalSeconds, 1)) seconds" -Level SUCCESS

    if (-not $Script:IsDryRun) {
        Update-Status ""
        Update-Status "A RESTART is strongly recommended!" -Level WARNING
        Update-Status "Check Windows Security after restart." -Level INFO
    }

    Save-UndoManifest | Out-Null
    Export-RepairReport | Out-Null

    return $true
}

# ============================================================================
# CLI ENTRY POINT
# ============================================================================

if ($Script:IsWorker) {
    return
}

function Invoke-CliStatus {
    <#
    .SYNOPSIS
        CLI Status mode: show health scan, detected privacy tools, and policy audit.
    #>
    $snapshot = New-StatusSnapshot
    $scan = $snapshot.scan
    $diff = $null
    if ($CompareSnapshot) {
        try {
            $previous = Get-Content -LiteralPath $CompareSnapshot -Raw -ErrorAction Stop | ConvertFrom-Json
            $diff = Compare-StatusSnapshot -Previous $previous -Current $snapshot
        }
        catch {
            Write-Warning "Could not compare snapshot: $($_.Exception.Message)"
        }
    }
    if ($SnapshotPath) {
        $snapshotSaved = Save-StatusSnapshot -Snapshot $snapshot -Path $SnapshotPath
    }

    if ($Script:IsJson) {
        $Script:JsonOutput.preScan = $snapshot.scan
        $Script:JsonOutput.privacyTools = $snapshot.privacyTools
        $Script:JsonOutput.blockers = $snapshot.blockers
        $Script:JsonOutput.thirdPartyAV = $snapshot.thirdPartyAV
        $Script:JsonOutput.avGuidance = $snapshot.avGuidance
        $Script:JsonOutput.snapshotPath = $SnapshotPath
        $Script:JsonOutput.diff = $diff
        $Script:JsonOutput.exitCode = 0
        Write-Output ($Script:JsonOutput | ConvertTo-Json -Depth 5)
        exit 0
    }

    Write-Host ''
    Write-Host "DefenderShield v$($Script:Version) - System Status" -ForegroundColor Cyan
    Write-Host ('=' * 50) -ForegroundColor Cyan
    Write-Host ''

    $components = @(
        @{ Label = 'WinDefend Service'; Value = $scan['WinDefend'] },
        @{ Label = 'SecurityHealthService'; Value = $scan['SecurityHealthService'] },
        @{ Label = 'Security Center (wscsvc)'; Value = $scan['wscsvc'] },
        @{ Label = 'Firewall (MpsSvc)'; Value = $scan['MpsSvc'] },
        @{ Label = 'Real-Time Protection'; Value = $scan['RealTimeProtection'] },
        @{ Label = 'Tamper Protection'; Value = $scan['TamperProtection'] },
        @{ Label = 'SmartScreen'; Value = $scan['SmartScreen'] },
        @{ Label = 'Definition Age'; Value = if ($scan['DefinitionAge'] -ge 0) { "$($scan['DefinitionAge']) days" } else { 'Unknown' } },
        @{ Label = 'Group Policy Blocking'; Value = $scan['GroupPolicyBlocking'] },
        @{ Label = 'Windows Security App'; Value = $scan['WindowsSecurityApp'] },
        @{ Label = 'MDE Sense'; Value = $scan['MDE'] },
        @{ Label = 'Third-Party AV'; Value = $scan['ThirdPartyAV'] }
    )

    foreach ($c in $components) {
        $color = 'Green'
        if ($c.Label -eq 'Third-Party AV' -and $c.Value -ne 'None') { $color = 'Red' }
        elseif ($c.Value -in @('OFF', 'Missing', 'Disabled', 'Yes', 'Unknown')) { $color = 'Red' }
        elseif ($c.Value -in @('Stopped')) { $color = 'Yellow' }
        Write-Host ("{0,-26} " -f $c.Label) -NoNewline
        Write-Host $c.Value -ForegroundColor $color
    }

    # Third-party AV
    Write-Host ''
    $avList = $snapshot.thirdPartyAV
    if ($avList) {
        Write-Host 'Third-Party AV Detected:' -ForegroundColor Yellow
        foreach ($av in $avList) {
            Write-Host "  - $($av.Name)" -ForegroundColor Yellow
        }
        foreach ($hint in $snapshot.avGuidance) {
            Write-Host "    $($hint.guidance)" -ForegroundColor Yellow
        }
    }

    # Privacy tools
    Write-Host ''
    $privacyTools = $snapshot.privacyTools
    if ($privacyTools -and $privacyTools.Count -gt 0) {
        Write-Host 'Detected Privacy Tools:' -ForegroundColor Yellow
        foreach ($pt in $privacyTools) {
            Write-Host "  - $($pt.Tool): $($pt.Evidence)" -ForegroundColor Yellow
        }
    }

    # Policy audit
    Write-Host ''
    $blockers = $snapshot.blockers
    if ($blockers -and $blockers.Count -gt 0) {
        Write-Host "Active Blockers ($($blockers.Count)):" -ForegroundColor Red
        Write-Host ("{0,-20} {1,-45} {2}" -f 'SOURCE', 'LABEL', 'DETAIL') -ForegroundColor Cyan
        foreach ($b in $blockers) {
            Write-Host ("{0,-20} {1,-45} {2}" -f $b.Source, $b.Label, $b.Detail) -ForegroundColor Yellow
        }
    }
    else {
        Write-Host 'No active blockers detected.' -ForegroundColor Green
    }

    Write-Host ''
    if ($SnapshotPath) {
        if ($snapshotSaved -and (Test-Path -LiteralPath $SnapshotPath)) {
            Write-Host "Snapshot saved: $SnapshotPath" -ForegroundColor Green
        }
        else {
            Write-Warning "Snapshot was not saved: $SnapshotPath"
        }
    }
    if ($diff) {
        Show-StatusDiff -Diff $diff
        Write-Host ''
    }

    $defBroken = Test-DefenderBroken -Scan $scan
    $fwBroken = Test-FirewallBroken -Scan $scan
    if ($defBroken -or $fwBroken) {
        Write-Host 'Issues detected. Run with -Mode Both to repair.' -ForegroundColor Yellow
        exit 1
    }
    else {
        Write-Host 'All components appear healthy.' -ForegroundColor Green
        exit 0
    }
}

# CLI dispatch
if ($Script:CliMode) {
    if ($InstallWatchdog) {
        Install-WatchdogTask
    }
    elseif ($RemoveWatchdog) {
        Remove-WatchdogTask
    }
    elseif ($WatchdogCheck) {
        Invoke-WatchdogCheck
    }
    elseif ($ComputerName -and $ComputerName.Count -gt 0) {
        if (-not $Mode) { $Mode = 'Status' }
        Invoke-FleetRepair -Targets $ComputerName
    }
    elseif ($Mode -eq 'Status') {
        Invoke-CliStatus
    }
    else {
        $repairFW = $Mode -in @('Firewall', 'Both')
        $repairDef = $Mode -in @('Defender', 'Both')

        $Script:JsonOutput.preScan = (Get-HealthScan)

        $result = Start-Repair -RepairFirewall $repairFW -RepairDefender $repairDef -CreateRestorePoint (-not $Script:IsDryRun)

        if ($Script:IsJson) {
            $Script:JsonOutput.exitCode = $Script:ExitCode
            $Script:JsonOutput.dryRunActions = @($Script:DryRunActions)
            Write-Output ($Script:JsonOutput | ConvertTo-Json -Depth 5)
        }

        exit $Script:ExitCode
    }
}

# ============================================================================
# GUI (Catppuccin Mocha Theme)
# ============================================================================

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="DefenderShield v3.1.0 - Windows Security Repair Tool"
    Height="940" Width="860"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    Background="#1e1e2e">

    <Window.Resources>
        <!-- Catppuccin Mocha palette -->
        <!-- Base: #1e1e2e  Mantle: #181825  Crust: #11111b  Surface0: #313244  Surface1: #45475a  Surface2: #585b70 -->
        <!-- Text: #cdd6f4  Subtext0: #a6adc8  Subtext1: #bac2de  Overlay0: #6c7086 -->
        <!-- Red: #f38ba8  Green: #a6e3a1  Yellow: #f9e2af  Blue: #89b4fa  Mauve: #cba6f7  Teal: #94e2d5  Peach: #fab387 -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#cdd6f4"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Margin" Value="0,8,0,8"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#89b4fa"/>
            <Setter Property="Foreground" Value="#1e1e2e"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="20,12"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="DefenderShield" FontSize="28" FontWeight="Bold" Foreground="#f38ba8" HorizontalAlignment="Center"/>
            <TextBlock Text="Windows Defender and Firewall Repair Tool  v3.1.0" FontSize="14" Foreground="#a6adc8" HorizontalAlignment="Center" Margin="0,5,0,0"/>
        </StackPanel>

        <!-- Health Dashboard -->
        <StackPanel Grid.Row="1" Margin="0,0,0,10">
            <TextBlock Text="System Health Dashboard" FontSize="16" FontWeight="SemiBold" Foreground="#cdd6f4" Margin="0,0,0,10"/>
            <UniformGrid Columns="5">
                <Border Background="#181825" CornerRadius="8" Padding="12" Margin="0,0,8,0" MinHeight="82">
                    <StackPanel>
                        <TextBlock Text="Defender" Foreground="#a6adc8" FontSize="12"/>
                        <TextBlock x:Name="lblTileDefenderStatus" Text="Scanning..." Foreground="#6c7086" FontSize="18" FontWeight="Bold" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>
                <Border Background="#181825" CornerRadius="8" Padding="12" Margin="0,0,8,0" MinHeight="82">
                    <StackPanel>
                        <TextBlock Text="Firewall" Foreground="#a6adc8" FontSize="12"/>
                        <TextBlock x:Name="lblTileFirewallStatus" Text="Scanning..." Foreground="#6c7086" FontSize="18" FontWeight="Bold" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>
                <Border Background="#181825" CornerRadius="8" Padding="12" Margin="0,0,8,0" MinHeight="82">
                    <StackPanel>
                        <TextBlock Text="Tamper" Foreground="#a6adc8" FontSize="12"/>
                        <TextBlock x:Name="lblTileTamperStatus" Text="Scanning..." Foreground="#6c7086" FontSize="18" FontWeight="Bold" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>
                <Border Background="#181825" CornerRadius="8" Padding="12" Margin="0,0,8,0" MinHeight="82">
                    <StackPanel>
                        <TextBlock Text="Signature Age" Foreground="#a6adc8" FontSize="12"/>
                        <TextBlock x:Name="lblTileSignatureStatus" Text="Scanning..." Foreground="#6c7086" FontSize="18" FontWeight="Bold" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>
                <Border Background="#181825" CornerRadius="8" Padding="12" MinHeight="82">
                    <StackPanel>
                        <TextBlock Text="Third-Party AV" Foreground="#a6adc8" FontSize="12"/>
                        <TextBlock x:Name="lblTileAvStatus" Text="Scanning..." Foreground="#6c7086" FontSize="16" FontWeight="Bold" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>
            </UniformGrid>
        </StackPanel>

        <!-- Options Panel -->
        <Border Grid.Row="2" Background="#181825" CornerRadius="8" Padding="20" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="Select Components to Repair:" FontSize="16" FontWeight="SemiBold" Foreground="#cdd6f4" Margin="0,0,0,15"/>

                <CheckBox x:Name="chkFirewall" IsChecked="True">
                    <StackPanel>
                        <TextBlock Text="Windows Firewall" FontWeight="SemiBold" Foreground="#cdd6f4"/>
                        <TextBlock Text="Repairs services, removes blocking policies, enables all profiles" FontSize="11" Foreground="#6c7086"/>
                    </StackPanel>
                </CheckBox>

                <CheckBox x:Name="chkDefender" IsChecked="True">
                    <StackPanel>
                        <TextBlock Text="Windows Defender Antivirus" FontWeight="SemiBold" Foreground="#cdd6f4"/>
                        <TextBlock Text="Repairs services, registry, scheduled tasks, WMI, SmartScreen, enables protection" FontSize="11" Foreground="#6c7086"/>
                    </StackPanel>
                </CheckBox>

                <Rectangle Height="1" Fill="#45475a" Margin="0,10"/>

                <CheckBox x:Name="chkRestorePoint" IsChecked="True">
                    <StackPanel>
                        <TextBlock Text="Create System Restore Point" FontWeight="SemiBold" Foreground="#cdd6f4"/>
                        <TextBlock Text="Recommended - allows you to undo changes if needed" FontSize="11" Foreground="#6c7086"/>
                    </StackPanel>
                </CheckBox>
            </StackPanel>
        </Border>

        <!-- Tamper Protection Helper -->
        <Border Grid.Row="3" Background="#302030" CornerRadius="8" Padding="12" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" TextWrapping="Wrap" Foreground="#f9e2af" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold">Tamper Protection:</Run>
                    <Run>If Defender won't start after repair, open Windows Security, turn OFF Tamper Protection, repair again, then turn it back ON.</Run>
                </TextBlock>
                <StackPanel Grid.Column="1" Orientation="Vertical" Margin="10,0,0,0">
                    <Button x:Name="btnTamperHelper" Content="Open Settings" Width="120" FontSize="12" Padding="10,8" Background="#cba6f7" Foreground="#1e1e2e" Margin="0,0,0,5"/>
                    <Button x:Name="btnTamperRecheck" Content="Re-check" Width="120" FontSize="12" Padding="10,8" Background="#45475a" Foreground="#cdd6f4"/>
                </StackPanel>
            </Grid>
        </Border>

        <ProgressBar x:Name="prgRepair" Grid.Row="4" Height="8" Margin="0,0,0,10" IsIndeterminate="True" Visibility="Collapsed" Background="#313244" Foreground="#89b4fa"/>

        <!-- Status Output -->
        <TabControl Grid.Row="5" Background="#11111b" BorderBrush="#45475a" Foreground="#cdd6f4">
            <TabItem Header="Live Log" Background="#181825" Foreground="#cdd6f4">
                <Border Background="#11111b" Padding="10">
                    <RichTextBox x:Name="txtStatus"
                                 Background="Transparent"
                                 Foreground="#cdd6f4"
                                 FontFamily="Consolas"
                                 FontSize="12"
                                 IsReadOnly="True"
                                 BorderThickness="0"
                                 VerticalScrollBarVisibility="Auto">
                        <FlowDocument>
                            <Paragraph>
                                <Run Foreground="#6c7086">Ready. Select options above and click Start Repair.</Run>
                            </Paragraph>
                        </FlowDocument>
                    </RichTextBox>
                </Border>
            </TabItem>
            <TabItem Header="Report" Background="#181825" Foreground="#cdd6f4">
                <Border Background="#11111b" Padding="10">
                    <TextBox x:Name="txtReport"
                             Text="No report has been generated yet."
                             Background="Transparent"
                             Foreground="#cdd6f4"
                             FontFamily="Consolas"
                             FontSize="12"
                             IsReadOnly="True"
                             BorderThickness="0"
                             TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto"/>
                </Border>
            </TabItem>
        </TabControl>

        <!-- Buttons -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnQuickFix" Content="Quick Fix All" Margin="0,0,10,0" Width="150" Background="#f38ba8" Foreground="#1e1e2e"/>
            <Button x:Name="btnStart" Content="Start Repair" Margin="0,0,10,0" Width="150"/>
            <Button x:Name="btnExportReport" Content="Open Report" Margin="0,0,10,0" Width="140" Background="#94e2d5" Foreground="#1e1e2e" IsEnabled="False"/>
            <Button x:Name="btnRestart" Content="Restart PC" Width="120" Background="#45475a" Foreground="#cdd6f4" IsEnabled="False"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Parse XAML
$reader = New-Object System.Xml.XmlNodeReader($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Get controls
$chkFirewall = $window.FindName('chkFirewall')
$chkDefender = $window.FindName('chkDefender')
$chkRestorePoint = $window.FindName('chkRestorePoint')
$txtStatus = $window.FindName('txtStatus')
$txtReport = $window.FindName('txtReport')
$prgRepair = $window.FindName('prgRepair')
$btnStart = $window.FindName('btnStart')
$btnRestart = $window.FindName('btnRestart')
$btnQuickFix = $window.FindName('btnQuickFix')
$btnExportReport = $window.FindName('btnExportReport')
$btnTamperHelper = $window.FindName('btnTamperHelper')
$btnTamperRecheck = $window.FindName('btnTamperRecheck')

# Dashboard labels
$Script:DashboardLabels = @{
    'lblTileDefenderStatus'  = $window.FindName('lblTileDefenderStatus')
    'lblTileFirewallStatus'  = $window.FindName('lblTileFirewallStatus')
    'lblTileTamperStatus'    = $window.FindName('lblTileTamperStatus')
    'lblTileSignatureStatus' = $window.FindName('lblTileSignatureStatus')
    'lblTileAvStatus'        = $window.FindName('lblTileAvStatus')
    'lblWinDefend'           = $window.FindName('lblWinDefend')
    'lblSecHealth'           = $window.FindName('lblSecHealth')
    'lblWscsvc'              = $window.FindName('lblWscsvc')
    'lblMpsSvc'              = $window.FindName('lblMpsSvc')
    'lblRTP'                 = $window.FindName('lblRTP')
    'lblTamper'              = $window.FindName('lblTamper')
    'lblSmartScreen'         = $window.FindName('lblSmartScreen')
    'lblDefAge'              = $window.FindName('lblDefAge')
    'lblGPBlock'             = $window.FindName('lblGPBlock')
    'lblWinSecApp'           = $window.FindName('lblWinSecApp')
}

# Store reference for logging
$Script:StatusTextBox = $txtStatus
$Script:RepairPowerShell = $null
$Script:RepairAsync = $null
$Script:RepairTimer = $null
$Script:RepairLogPosition = 0
$Script:ActiveRepairLogPath = $null
$Script:ActiveRepairReportPath = $null
$Script:ActiveRepairBackupPath = $null

function Read-GuiRepairLog {
    if (-not $Script:ActiveRepairLogPath) { return }
    if (-not (Test-Path -LiteralPath $Script:ActiveRepairLogPath)) { return }

    try {
        $fs = [System.IO.File]::Open($Script:ActiveRepairLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $null = $fs.Seek($Script:RepairLogPosition, [System.IO.SeekOrigin]::Begin)
            $reader = New-Object System.IO.StreamReader($fs)
            $text = $reader.ReadToEnd()
            $Script:RepairLogPosition = $fs.Position
        }
        finally {
            $fs.Close()
        }

        foreach ($line in ($text -split "`r?`n")) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $level = 'INFO'
            $message = $line
            if ($line -match '^\[[^\]]+\]\s+\[([^\]]+)\]\s+(.*)$') {
                $level = $Matches[1]
                $message = $Matches[2]
            }
            Add-GuiStatusLine -Message $message -Level $level
        }
    }
    catch { }
}

function Complete-GuiRepair {
    Read-GuiRepairLog

    try {
        if ($Script:RepairPowerShell -and $Script:RepairAsync) {
            $null = $Script:RepairPowerShell.EndInvoke($Script:RepairAsync)
            if ($Script:RepairPowerShell.Streams.Error.Count -gt 0) {
                foreach ($err in $Script:RepairPowerShell.Streams.Error) {
                    Add-GuiStatusLine -Message "Worker error: $($err.Exception.Message)" -Level ERROR
                }
            }
        }
    }
    catch {
        Add-GuiStatusLine -Message "Repair worker failed: $($_.Exception.Message)" -Level ERROR
    }
    finally {
        if ($Script:RepairPowerShell) {
            $Script:RepairPowerShell.Dispose()
        }
        $Script:RepairPowerShell = $null
        $Script:RepairAsync = $null
    }

    $prgRepair.Visibility = 'Collapsed'
    $btnStart.Content = "Start Repair"
    $btnQuickFix.Content = "Quick Fix All"
    $btnStart.IsEnabled = $true
    $btnQuickFix.IsEnabled = $true
    $btnRestart.IsEnabled = $true
    $chkFirewall.IsEnabled = $true
    $chkDefender.IsEnabled = $true
    $chkRestorePoint.IsEnabled = $true

    try {
        $scan = Get-HealthScan
        Update-HealthDashboard -Scan $scan
    }
    catch { }

    $manifestPath = Join-Path $Script:ActiveRepairBackupPath 'undo-manifest.json'
    if (Test-Path -LiteralPath $Script:ActiveRepairReportPath) {
        $btnExportReport.IsEnabled = $true
        $txtReport.Text = "Report generated:`r`n$($Script:ActiveRepairReportPath)`r`n`r`nUndo manifest:`r`n$manifestPath`r`n`r`nLog:`r`n$($Script:ActiveRepairLogPath)"
    }
    else {
        $txtReport.Text = "Repair completed, but no HTML report was generated.`r`nLog:`r`n$($Script:ActiveRepairLogPath)"
    }
}

# Helper to run repair with current checkbox state
$Script:RunRepair = {
    if (-not $chkFirewall.IsChecked -and -not $chkDefender.IsChecked) {
        Update-Status "Select at least one component to repair." -Level WARNING
        return
    }

    # Clear status
    $txtStatus.Document.Blocks.Clear()

    # Disable controls during repair
    $btnStart.IsEnabled = $false
    $btnQuickFix.IsEnabled = $false
    $btnExportReport.IsEnabled = $false
    $chkFirewall.IsEnabled = $false
    $chkDefender.IsEnabled = $false
    $chkRestorePoint.IsEnabled = $false
    $btnStart.Content = "Repairing..."
    $btnQuickFix.Content = "Repairing..."
    $prgRepair.Visibility = 'Visible'

    # Store selections
    $repairFW = [bool]$chkFirewall.IsChecked
    $repairDef = [bool]$chkDefender.IsChecked
    $createRP = [bool]$chkRestorePoint.IsChecked

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $Script:ActiveRepairLogPath = "$env:USERPROFILE\Desktop\DefenderShield_$stamp.log"
    $Script:ActiveRepairBackupPath = "$env:USERPROFILE\Desktop\DefenderShield_Backup_$stamp"
    $Script:ActiveRepairReportPath = "$env:USERPROFILE\Desktop\DefenderShield_Report_$stamp.html"
    $Script:RepairLogPosition = 0
    $txtReport.Text = "Repair running...`r`nLog:`r`n$($Script:ActiveRepairLogPath)"

    $workerScript = @'
param($scriptPath, $repairFW, $repairDef, $createRP, $logPath, $backupPath, $reportPath)
. $scriptPath -Worker -WorkerLogPath $logPath -WorkerBackupPath $backupPath -WorkerReportPath $reportPath
Start-Repair -RepairFirewall $repairFW -RepairDefender $repairDef -CreateRestorePoint $createRP | Out-Null
$Script:ExitCode
'@

    try {
        $Script:RepairPowerShell = [PowerShell]::Create()
        $workerCommand = $Script:RepairPowerShell.AddScript($workerScript)
        [void]$workerCommand.AddArgument($PSCommandPath)
        [void]$workerCommand.AddArgument($repairFW)
        [void]$workerCommand.AddArgument($repairDef)
        [void]$workerCommand.AddArgument($createRP)
        [void]$workerCommand.AddArgument($Script:ActiveRepairLogPath)
        [void]$workerCommand.AddArgument($Script:ActiveRepairBackupPath)
        [void]$workerCommand.AddArgument($Script:ActiveRepairReportPath)
        $Script:RepairAsync = $Script:RepairPowerShell.BeginInvoke()

        $Script:RepairTimer = New-Object System.Windows.Threading.DispatcherTimer
        $Script:RepairTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $Script:RepairTimer.Add_Tick({
            Read-GuiRepairLog
            if ($Script:RepairAsync -and $Script:RepairAsync.IsCompleted) {
                $Script:RepairTimer.Stop()
                Complete-GuiRepair
            }
        })
        $Script:RepairTimer.Start()
    }
    catch {
        Add-GuiStatusLine -Message "Could not start repair worker: $($_.Exception.Message)" -Level ERROR
        Complete-GuiRepair
    }
}

# Start button click
$btnStart.Add_Click({
    & $Script:RunRepair
})

# Quick Fix All button click
$btnQuickFix.Add_Click({
    # Select all broken items
    $scan = Get-HealthScan
    $chkFirewall.IsChecked = Test-FirewallBroken -Scan $scan
    $chkDefender.IsChecked = Test-DefenderBroken -Scan $scan

    # If nothing is broken, check both anyway
    if (-not $chkFirewall.IsChecked -and -not $chkDefender.IsChecked) {
        $chkFirewall.IsChecked = $true
        $chkDefender.IsChecked = $true
    }

    $chkRestorePoint.IsChecked = $true

    & $Script:RunRepair
})

# Restart button click
$btnRestart.Add_Click({
    Update-Status "Restart requested. Restarting now..." -Level WARNING
    Restart-Computer -Force
})

# Open report button click
$btnExportReport.Add_Click({
    if ($Script:ActiveRepairReportPath -and (Test-Path -LiteralPath $Script:ActiveRepairReportPath)) {
        Start-Process $Script:ActiveRepairReportPath
    }
    else {
        Update-Status "No report is available yet." -Level WARNING
    }
})

# Tamper Protection helper: open Windows Security to the right page
$btnTamperHelper.Add_Click({
    try {
        Start-Process 'windowsdefender://threatsettings'
    }
    catch {
        try {
            Start-Process 'ms-settings:windowsdefender'
        }
        catch {
            [System.Windows.MessageBox]::Show("Could not open Windows Security.`n`nManually open: Windows Security > Virus & threat protection > Manage settings > Tamper Protection", "DefenderShield", "OK", "Warning") | Out-Null
        }
    }
})

# Tamper Protection re-check: rescan and update dashboard
$btnTamperRecheck.Add_Click({
    $scan = Get-HealthScan
    Update-HealthDashboard -Scan $scan

    $tamperVal = $scan['TamperProtection']
    if ($tamperVal -eq 'OFF') {
        Update-Status "Tamper Protection is OFF - you can now run repair." -Level WARNING
    }
    else {
        Update-Status "Tamper Protection is ON - good, it is protecting Defender." -Level SUCCESS
    }
})

# Run initial health scan on window load
$window.Add_ContentRendered({
    $scan = Get-HealthScan
    Update-HealthDashboard -Scan $scan

    # Auto-select broken items
    $defenderBroken = Test-DefenderBroken -Scan $scan
    $firewallBroken = Test-FirewallBroken -Scan $scan

    $chkDefender.IsChecked = $defenderBroken
    $chkFirewall.IsChecked = $firewallBroken

    # Update status with scan summary
    $txtStatus.Document.Blocks.Clear()

    $paragraph = New-Object System.Windows.Documents.Paragraph
    $run = New-Object System.Windows.Documents.Run("Health scan complete. ")
    $run.Foreground = '#6c7086'
    $paragraph.Inlines.Add($run)

    if ($defenderBroken -or $firewallBroken) {
        $issueRun = New-Object System.Windows.Documents.Run("Issues detected - broken items auto-selected.")
        $issueRun.Foreground = '#f9e2af'
        $paragraph.Inlines.Add($issueRun)
    }
    else {
        $okRun = New-Object System.Windows.Documents.Run("All components appear healthy.")
        $okRun.Foreground = '#a6e3a1'
        $paragraph.Inlines.Add($okRun)
    }

    $paragraph.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)
    $txtStatus.Document.Blocks.Add($paragraph)

    # Check for third-party AV and show in status
    $avList = Get-ThirdPartyAV
    if ($avList) {
        $avNames = ($avList | ForEach-Object { $_.Name }) -join ', '
        $avParagraph = New-Object System.Windows.Documents.Paragraph
        $avRun = New-Object System.Windows.Documents.Run("Third-party AV detected: $avNames - Defender repair may be blocked.")
        $avRun.Foreground = '#f9e2af'
        $avParagraph.Inlines.Add($avRun)
        $avParagraph.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)
        $txtStatus.Document.Blocks.Add($avParagraph)
    }
})

# Show window
$null = $window.ShowDialog()
