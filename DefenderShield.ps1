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
.NOTES
    Author: Matt
    Requires: Administrator privileges
    Version: 3.0.0
#>
param(
    [ValidateSet('Defender', 'Firewall', 'Both', 'Status')]
    [string]$Mode,

    [switch]$DryRun,

    [switch]$Silent,

    [switch]$Json,

    [ValidateSet('Services', 'Registry', 'Tasks', 'WMI', 'GroupPolicy', 'Features', 'SmartScreen')]
    [string[]]$Only,

    [ValidateSet('Services', 'Registry', 'Tasks', 'WMI', 'GroupPolicy', 'Features', 'SmartScreen')]
    [string[]]$Skip
)

# ============================================================================
# CONSTANTS & MODE DETECTION
# ============================================================================

$Script:Version = '3.0.0'
$Script:CliMode = [bool]$Mode
$Script:IsDryRun = [bool]$DryRun
$Script:IsSilent = [bool]$Silent
$Script:IsJson = [bool]$Json
$Script:ExitCode = 0  # 0=success, 1=partial, 2=failed, 3=blocked

# Phase filtering
$Script:ActivePhases = if ($Only -and $Only.Count -gt 0) {
    $Only
} else {
    @('Services', 'Registry', 'Tasks', 'WMI', 'GroupPolicy', 'Features', 'SmartScreen')
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
    exitCode  = 0
}

# Dry-run action log
$Script:DryRunActions = [System.Collections.ArrayList]::new()

# ============================================================================
# ELEVATION CHECK
# ============================================================================

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
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
}

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

# Stores pre-repair scan results for before/after comparison
$Script:PreRepairScan = $null

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

    try {
        Add-Content -Path $Script:Config.LogPath -Value $logMessage -ErrorAction SilentlyContinue
    }
    catch { }

    return $logMessage
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
        try {
            $Script:StatusTextBox.Dispatcher.Invoke([action]{
                # Catppuccin Mocha colors
                $color = switch ($Level) {
                    'SUCCESS' { '#a6e3a1' }   # Green
                    'WARNING' { '#f9e2af' }   # Yellow
                    'ERROR'   { '#f38ba8' }   # Red
                    'SECTION' { '#89b4fa' }   # Blue
                    default   { '#cdd6f4' }   # Text
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

            if (-not (Test-Path $item.Path)) {
                New-Item -Path $item.Path -Force -ErrorAction SilentlyContinue | Out-Null
            }

            if ($item.Name -eq '(Default)') {
                Set-ItemProperty -Path $item.Path -Name '(Default)' -Value $item.Value -Force -ErrorAction SilentlyContinue
            }
            else {
                Set-ItemProperty -Path $item.Path -Name $item.Name -Value $item.Value -Type $item.Type -Force -ErrorAction SilentlyContinue
            }
            $repaired++
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
                        Remove-ItemProperty -Path $pol.Path -Name $pol.Name -Force -ErrorAction SilentlyContinue
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
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force -ErrorAction SilentlyContinue
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
        if (Test-Path $Path) {
            Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction SilentlyContinue
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
        if (-not (Test-Path $Script:Config.BackupPath)) {
            New-Item -Path $Script:Config.BackupPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        }

        $exportPath = Join-Path $Script:Config.BackupPath "$BackupName.reg"
        $regPath = $KeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\' -replace '^HKCU:\\', 'HKEY_CURRENT_USER\'

        $null = reg export $regPath $exportPath /y 2>&1
        return $true
    }
    catch {
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
        try {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Name)"

            $startValue = switch ($svc.StartType) {
                'Automatic' { 2 }
                'Manual' { 3 }
                'Disabled' { 4 }
                'Boot' { 0 }
                'System' { 1 }
                default { 2 }
            }

            if ($Script:IsDryRun) {
                Update-Status "[DRY-RUN] Would set $regPath\Start = $startValue" -Level INFO
                [void]$Script:DryRunActions.Add("Set $regPath\Start = $startValue")
                continue
            }

            if (Test-Path $regPath) {
                Set-ItemProperty -Path $regPath -Name 'Start' -Value $startValue -Type DWord -Force -ErrorAction SilentlyContinue
                Update-Status "$($svc.DisplayName): Registry repaired" -Level SUCCESS
            }

            # Also try sc.exe
            $null = sc.exe config $svc.Name start= $svc.StartType.ToLower() 2>&1
        }
        catch {
            Update-Status "$($svc.DisplayName): Could not repair (continuing...)" -Level WARNING
        }
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
                    Remove-Item -Path $policy -Recurse -Force -ErrorAction SilentlyContinue
                    Update-Status "Removed: $policy" -Level SUCCESS
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
                    Set-ItemProperty -Path $profilePath -Name 'EnableFirewall' -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
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
        try {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -ne 'Running') {
                if ($Script:IsDryRun) {
                    Update-Status "[DRY-RUN] Would start service: $($svc.DisplayName)" -Level INFO
                    [void]$Script:DryRunActions.Add("Start service: $svcName")
                    continue
                }
                Start-Service -Name $svcName -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 500

                $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
                if ($svc.Status -eq 'Running') {
                    Update-Status "$($svc.DisplayName): Started" -Level SUCCESS
                }
                else {
                    Update-Status "$($svc.DisplayName): Could not start (may need reboot)" -Level WARNING
                }
            }
            elseif ($svc) {
                Update-Status "$($svc.DisplayName): Already running" -Level SUCCESS
            }
        }
        catch {
            Update-Status "$svcName : Error starting (continuing...)" -Level WARNING
        }
    }
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

    try {
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True -ErrorAction SilentlyContinue
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
        $null = netsh advfirewall reset 2>&1
        Update-Status "Firewall reset to defaults" -Level SUCCESS
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
        try {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Name)"

            $startValue = switch ($svc.StartType) {
                'Automatic' { 2 }
                'Manual' { 3 }
                'Disabled' { 4 }
                'Boot' { 0 }
                'System' { 1 }
                default { 3 }
            }

            if ($Script:IsDryRun) {
                Update-Status "[DRY-RUN] Would set $regPath\Start = $startValue" -Level INFO
                [void]$Script:DryRunActions.Add("Set $regPath\Start = $startValue")
                continue
            }

            if (Test-Path $regPath) {
                Set-ItemProperty -Path $regPath -Name 'Start' -Value $startValue -Type DWord -Force -ErrorAction SilentlyContinue
                Update-Status "$($svc.DisplayName): Registry repaired" -Level SUCCESS
            }
        }
        catch {
            Update-Status "$($svc.DisplayName): Could not repair (Tamper Protection?)" -Level WARNING
        }
    }

    # Repair drivers
    $drivers = @(
        @{ Name = 'WdFilter'; Start = 0 },
        @{ Name = 'WdNisDrv'; Start = 3 },
        @{ Name = 'WdBoot'; Start = 0 }
    )

    foreach ($driver in $drivers) {
        try {
            $drvPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$($driver.Name)"
            if ($Script:IsDryRun) {
                Update-Status "[DRY-RUN] Would set $drvPath\Start = $($driver.Start)" -Level INFO
                [void]$Script:DryRunActions.Add("Set $drvPath\Start = $($driver.Start)")
                continue
            }
            if (Test-Path $drvPath) {
                Set-ItemProperty -Path $drvPath -Name 'Start' -Value $driver.Start -Type DWord -Force -ErrorAction SilentlyContinue
            }
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
                        Remove-ItemProperty -Path $item.Path -Name $item.Name -Force -ErrorAction SilentlyContinue
                    }
                    $removed++
                }
            }
        }
        catch {
            if (-not $Script:IsDryRun) {
                # Try setting to 0 instead
                try {
                    Set-ItemProperty -Path $item.Path -Name $item.Name -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
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
                    Remove-Item -Path $tree -Recurse -Force -ErrorAction SilentlyContinue
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
                    Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                    Update-Status "Removed malicious task: $($task.TaskName)" -Level SUCCESS
                }
            }
            catch { }
        }
    }
    catch { }
}

function Repair-DefenderWMI {
    if (-not (Test-PhaseActive 'WMI')) { return }
    Update-Status "Checking WMI Subscriptions..." -Level SECTION

    try {
        $filters = Get-WmiObject -Query "SELECT * FROM __EventFilter WHERE Name LIKE '%Defender%' OR Name LIKE '%WinDefend%'" -Namespace 'root\subscription' -ErrorAction SilentlyContinue

        if ($filters) {
            foreach ($filter in $filters) {
                try {
                    if ($Script:IsDryRun) {
                        Update-Status "[DRY-RUN] Would remove WMI subscription: $($filter.Name)" -Level INFO
                        [void]$Script:DryRunActions.Add("Remove WMI subscription: $($filter.Name)")
                    }
                    else {
                        $bindingQuery = "SELECT * FROM __FilterToConsumerBinding WHERE Filter=""__EventFilter.Name='$($filter.Name)'"""
                        Get-WmiObject -Query $bindingQuery -Namespace 'root\subscription' -ErrorAction SilentlyContinue | Remove-WmiObject -ErrorAction SilentlyContinue
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
            if ($Script:IsDryRun) {
                Update-Status "[DRY-RUN] Would start Security Center (wscsvc)" -Level INFO
                [void]$Script:DryRunActions.Add("Start service: wscsvc")
            }
            else {
                Set-Service -Name 'wscsvc' -StartupType Automatic -ErrorAction SilentlyContinue
                Start-Service -Name 'wscsvc' -ErrorAction SilentlyContinue
            }
        }
    }
    catch { }

    foreach ($svc in $Script:DefenderServices) {
        try {
            $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
            if ($service -and $service.Status -ne 'Running') {
                if ($Script:IsDryRun) {
                    Update-Status "[DRY-RUN] Would start $($svc.DisplayName)" -Level INFO
                    [void]$Script:DryRunActions.Add("Start service: $($svc.Name)")
                    continue
                }
                Start-Service -Name $svc.Name -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 300

                $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
                if ($service.Status -eq 'Running') {
                    Update-Status "$($svc.DisplayName): Started" -Level SUCCESS
                }
                else {
                    Update-Status "$($svc.DisplayName): Could not start (may need reboot)" -Level WARNING
                }
            }
            elseif ($service) {
                Update-Status "$($svc.DisplayName): Already running" -Level SUCCESS
            }
        }
        catch {
            Update-Status "$($svc.DisplayName): Error (continuing...)" -Level WARNING
        }
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
            Copy-Item -Path $machinePolPath -Destination $backupPol -Force -ErrorAction SilentlyContinue
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
            Add-AppxPackage -DisableDevelopmentMode -Register "$($pkg.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue
            Update-Status "Windows Security app re-registered" -Level SUCCESS
        }
        else {
            # Try AllUsers reprovisioning
            $pkg = Get-AppxPackage -AllUsers -Name 'Microsoft.SecHealthUI' -ErrorAction SilentlyContinue
            if ($pkg) {
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

    # Capture pre-repair scan
    $Script:PreRepairScan = Get-HealthScan

    $modeLabel = if ($Script:IsDryRun) { ' [DRY-RUN]' } else { '' }
    Update-Status "DefenderShield v$($Script:Version) Repair Started$modeLabel" -Level SECTION
    Update-Status "Log file: $($Script:Config.LogPath)"

    # Pre-flight: third-party AV check
    if ($RepairDefender) {
        if (Test-ThirdPartyAVBlocking) {
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
        Test-FirewallServiceDependencies

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
        Repair-DefenderRegistry
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

    return $true
}

# ============================================================================
# CLI ENTRY POINT
# ============================================================================

function Invoke-CliStatus {
    <#
    .SYNOPSIS
        CLI Status mode: show health scan, detected privacy tools, and policy audit.
    #>
    $scan = Get-HealthScan

    if ($Script:IsJson) {
        $Script:JsonOutput.preScan = $scan
        $Script:JsonOutput.privacyTools = @(Get-DetectedPrivacyTools)
        $Script:JsonOutput.blockers = @(Get-PolicySourceAudit)
        $Script:JsonOutput.thirdPartyAV = Get-ThirdPartyAV
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
        @{ Label = 'Definition Age'; Value = if ($scan['DefinitionAge'] -ge 0) { "$($scan['DefinitionAge']) days" } else { 'Unknown' } },
        @{ Label = 'Group Policy Blocking'; Value = $scan['GroupPolicyBlocking'] },
        @{ Label = 'Windows Security App'; Value = $scan['WindowsSecurityApp'] }
    )

    foreach ($c in $components) {
        $color = 'Green'
        if ($c.Value -in @('OFF', 'Missing', 'Disabled', 'Yes', 'Unknown')) { $color = 'Red' }
        elseif ($c.Value -in @('Stopped')) { $color = 'Yellow' }
        Write-Host ("{0,-26} " -f $c.Label) -NoNewline
        Write-Host $c.Value -ForegroundColor $color
    }

    # Third-party AV
    Write-Host ''
    $avList = Get-ThirdPartyAV
    if ($avList) {
        Write-Host 'Third-Party AV Detected:' -ForegroundColor Yellow
        foreach ($av in $avList) {
            Write-Host "  - $($av.Name)" -ForegroundColor Yellow
        }
    }

    # Privacy tools
    Write-Host ''
    $privacyTools = Get-DetectedPrivacyTools
    if ($privacyTools -and $privacyTools.Count -gt 0) {
        Write-Host 'Detected Privacy Tools:' -ForegroundColor Yellow
        foreach ($pt in $privacyTools) {
            Write-Host "  - $($pt.Tool): $($pt.Evidence)" -ForegroundColor Yellow
        }
    }

    # Policy audit
    Write-Host ''
    $blockers = Get-PolicySourceAudit
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
    if ($Mode -eq 'Status') {
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
    Title="DefenderShield v3.0.0 - Windows Security Repair Tool"
    Height="900" Width="740"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanMinimize"
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="DefenderShield" FontSize="28" FontWeight="Bold" Foreground="#f38ba8" HorizontalAlignment="Center"/>
            <TextBlock Text="Windows Defender and Firewall Repair Tool  v3.0.0" FontSize="14" Foreground="#a6adc8" HorizontalAlignment="Center" Margin="0,5,0,0"/>
        </StackPanel>

        <!-- Health Dashboard -->
        <Border Grid.Row="1" Background="#181825" CornerRadius="8" Padding="15" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="System Health Dashboard" FontSize="16" FontWeight="SemiBold" Foreground="#cdd6f4" Margin="0,0,0,10"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Left column -->
                    <StackPanel Grid.Row="0" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="WinDefend:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblWinDefend" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="1" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="SecurityHealth:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblSecHealth" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="2" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Security Center:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblWscsvc" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="3" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Firewall (MpsSvc):  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblMpsSvc" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="4" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Win Security App:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblWinSecApp" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>

                    <!-- Right column -->
                    <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Real-Time Protection:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblRTP" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="1" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Tamper Protection:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblTamper" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="SmartScreen:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblSmartScreen" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="3" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Definition Age:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblDefAge" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="4" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="GP Blocking:  " Foreground="#a6adc8" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblGPBlock" Text="Scanning..." Foreground="#6c7086" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                </Grid>
            </StackPanel>
        </Border>

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

        <!-- Status Output -->
        <Border Grid.Row="4" Background="#11111b" CornerRadius="8" Padding="10">
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

        <!-- Buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnQuickFix" Content="Quick Fix All" Margin="0,0,10,0" Width="150" Background="#f38ba8" Foreground="#1e1e2e"/>
            <Button x:Name="btnStart" Content="Start Repair" Margin="0,0,10,0" Width="150"/>
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
$btnStart = $window.FindName('btnStart')
$btnRestart = $window.FindName('btnRestart')
$btnQuickFix = $window.FindName('btnQuickFix')
$btnTamperHelper = $window.FindName('btnTamperHelper')
$btnTamperRecheck = $window.FindName('btnTamperRecheck')

# Dashboard labels
$Script:DashboardLabels = @{
    'lblWinDefend'   = $window.FindName('lblWinDefend')
    'lblSecHealth'   = $window.FindName('lblSecHealth')
    'lblWscsvc'      = $window.FindName('lblWscsvc')
    'lblMpsSvc'      = $window.FindName('lblMpsSvc')
    'lblRTP'         = $window.FindName('lblRTP')
    'lblTamper'      = $window.FindName('lblTamper')
    'lblSmartScreen' = $window.FindName('lblSmartScreen')
    'lblDefAge'      = $window.FindName('lblDefAge')
    'lblGPBlock'     = $window.FindName('lblGPBlock')
    'lblWinSecApp'   = $window.FindName('lblWinSecApp')
}

# Store reference for logging
$Script:StatusTextBox = $txtStatus

# Helper to run repair with current checkbox state
$Script:RunRepair = {
    if (-not $chkFirewall.IsChecked -and -not $chkDefender.IsChecked) {
        [System.Windows.MessageBox]::Show("Please select at least one component to repair.", "DefenderShield", "OK", "Warning") | Out-Null
        return
    }

    # Clear status
    $txtStatus.Document.Blocks.Clear()

    # Disable controls during repair
    $btnStart.IsEnabled = $false
    $btnQuickFix.IsEnabled = $false
    $chkFirewall.IsEnabled = $false
    $chkDefender.IsEnabled = $false
    $chkRestorePoint.IsEnabled = $false
    $btnStart.Content = "Repairing..."
    $btnQuickFix.Content = "Repairing..."

    # Store selections
    $repairFW = $chkFirewall.IsChecked
    $repairDef = $chkDefender.IsChecked
    $createRP = $chkRestorePoint.IsChecked

    # Process UI events then run repair
    [System.Windows.Forms.Application]::DoEvents()

    # Run repair
    Start-Repair -RepairFirewall $repairFW -RepairDefender $repairDef -CreateRestorePoint $createRP

    # Re-enable controls
    $btnStart.Content = "Start Repair"
    $btnQuickFix.Content = "Quick Fix All"
    $btnStart.IsEnabled = $true
    $btnQuickFix.IsEnabled = $true
    $btnRestart.IsEnabled = $true
    $chkFirewall.IsEnabled = $true
    $chkDefender.IsEnabled = $true
    $chkRestorePoint.IsEnabled = $true
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
    $result = [System.Windows.MessageBox]::Show("Are you sure you want to restart your computer now?", "Restart Computer", "YesNo", "Question")
    if ($result -eq 'Yes') {
        Restart-Computer -Force
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
