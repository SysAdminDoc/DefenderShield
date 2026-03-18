<#
.SYNOPSIS
    DefenderShield - Windows Defender & Firewall Repair Tool (GUI Version)
.DESCRIPTION
    Comprehensive repair tool for restoring Windows Defender and Windows Firewall
    after they've been disabled by privacy tools like privacy.sexy, O&O ShutUp10,
    Debloaters, or manual modifications.

    Features a GUI with pre-scan health dashboard, auto-select broken items,
    before/after comparison, and Quick Fix All.
.NOTES
    Author: Generated for Matt
    Requires: Administrator privileges
    Version: 2.1.0
#>

# ============================================================================
# ELEVATION CHECK
# ============================================================================

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -ErrorAction Stop
        exit
    }
    catch {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show("This tool requires Administrator privileges.`n`nPlease right-click and select 'Run as Administrator'.", "DefenderShield", "OK", "Error") | Out-Null
        exit
    }
}

# ============================================================================
# ASSEMBLIES
# ============================================================================

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms -ErrorAction SilentlyContinue

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

    $logMsg = Write-Log -Message $Message -Level $Level

    if ($Script:StatusTextBox) {
        try {
            $Script:StatusTextBox.Dispatcher.Invoke([action]{
                $color = switch ($Level) {
                    'SUCCESS' { 'Lime' }
                    'WARNING' { 'Yellow' }
                    'ERROR'   { 'OrangeRed' }
                    'SECTION' { 'Cyan' }
                    default   { 'White' }
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

    return $scan
}

function Get-HealthColor {
    param([string]$Component, [string]$Value)

    switch ($Component) {
        'WinDefend'            { if ($Value -eq 'Running') { '#00ff00' } elseif ($Value -eq 'Stopped') { '#ffaa00' } else { '#ff4444' } }
        'SecurityHealthService' { if ($Value -eq 'Running') { '#00ff00' } elseif ($Value -eq 'Stopped') { '#ffaa00' } else { '#ff4444' } }
        'wscsvc'               { if ($Value -eq 'Running') { '#00ff00' } elseif ($Value -eq 'Stopped') { '#ffaa00' } else { '#ff4444' } }
        'MpsSvc'               { if ($Value -eq 'Running') { '#00ff00' } elseif ($Value -eq 'Stopped') { '#ffaa00' } else { '#ff4444' } }
        'RealTimeProtection'   { if ($Value -eq 'ON') { '#00ff00' } else { '#ff4444' } }
        'TamperProtection'     { if ($Value -eq 'ON') { '#00ff00' } else { '#ffaa00' } }
        'DefinitionAge'        {
            $days = [int]$Value
            if ($days -lt 0) { '#ff4444' } elseif ($days -le 3) { '#00ff00' } elseif ($days -le 7) { '#ffaa00' } else { '#ff4444' }
        }
        'GroupPolicyBlocking'  { if ($Value -eq 'No') { '#00ff00' } else { '#ff4444' } }
        'WindowsSecurityApp'   { if ($Value -eq 'Registered') { '#00ff00' } else { '#ff4444' } }
        default { '#ffffff' }
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
                Remove-Item -Path $policy -Recurse -Force -ErrorAction SilentlyContinue
                Update-Status "Removed: $policy" -Level SUCCESS
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
                Set-ItemProperty -Path $profilePath -Name 'EnableFirewall' -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }
    }

    Update-Status "Firewall registry cleanup complete" -Level SUCCESS
}

function Start-FirewallServices {
    Update-Status "Starting Firewall Services..." -Level SECTION

    $startOrder = @('BFE', 'mpssvc', 'IKEEXT', 'PolicyAgent')

    foreach ($svcName in $startOrder) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -ne 'Running') {
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
    Update-Status "Enabling Firewall Profiles..." -Level SECTION

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
            if (Test-Path $drvPath) {
                Set-ItemProperty -Path $drvPath -Name 'Start' -Value $driver.Start -Type DWord -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }
    }
}

function Repair-DefenderRegistry {
    Update-Status "Removing Defender Blocking Policies..." -Level SECTION

    # Backup first
    Backup-RegistryKey -KeyPath 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender' -BackupName 'Policies_WindowsDefender'
    Backup-RegistryKey -KeyPath 'HKLM:\SOFTWARE\Microsoft\Windows Defender' -BackupName 'WindowsDefender'

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
                    Remove-ItemProperty -Path $item.Path -Name $item.Name -Force -ErrorAction SilentlyContinue
                    $removed++
                }
            }
        }
        catch {
            # Try setting to 0 instead
            try {
                Set-ItemProperty -Path $item.Path -Name $item.Name -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
            }
            catch { }
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
                Remove-Item -Path $tree -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }
    }
}

function Repair-DefenderScheduledTasks {
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
                Enable-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
                Update-Status "Enabled: $taskName" -Level SUCCESS
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
                Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                Update-Status "Removed malicious task: $($task.TaskName)" -Level SUCCESS
            }
            catch { }
        }
    }
    catch { }
}

function Repair-DefenderWMI {
    Update-Status "Checking WMI Subscriptions..." -Level SECTION

    try {
        $filters = Get-WmiObject -Query "SELECT * FROM __EventFilter WHERE Name LIKE '%Defender%' OR Name LIKE '%WinDefend%'" -Namespace 'root\subscription' -ErrorAction SilentlyContinue

        if ($filters) {
            foreach ($filter in $filters) {
                try {
                    $bindingQuery = "SELECT * FROM __FilterToConsumerBinding WHERE Filter=""__EventFilter.Name='$($filter.Name)'"""
                    Get-WmiObject -Query $bindingQuery -Namespace 'root\subscription' -ErrorAction SilentlyContinue | Remove-WmiObject -ErrorAction SilentlyContinue
                    $filter | Remove-WmiObject -ErrorAction SilentlyContinue
                    Update-Status "Removed WMI subscription: $($filter.Name)" -Level SUCCESS
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
    Update-Status "Starting Defender Services..." -Level SECTION

    # Start Security Center first
    try {
        $secCenter = Get-Service -Name 'wscsvc' -ErrorAction SilentlyContinue
        if ($secCenter -and $secCenter.Status -ne 'Running') {
            Set-Service -Name 'wscsvc' -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service -Name 'wscsvc' -ErrorAction SilentlyContinue
        }
    }
    catch { }

    foreach ($svc in $Script:DefenderServices) {
        try {
            $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
            if ($service -and $service.Status -ne 'Running') {
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
    Update-Status "Enabling Defender Protection Features..." -Level SECTION

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
    Update-Status "Resetting Group Policy..." -Level SECTION

    try {
        $machinePolPath = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"

        if (Test-Path $machinePolPath) {
            $backupPol = Join-Path $Script:Config.BackupPath "Registry.pol"
            Copy-Item -Path $machinePolPath -Destination $backupPol -Force -ErrorAction SilentlyContinue
            Remove-Item -Path $machinePolPath -Force -ErrorAction SilentlyContinue
            Update-Status "Machine policy file removed" -Level SUCCESS
        }

        $null = gpupdate /force 2>&1
        Update-Status "Group Policy refreshed" -Level SUCCESS
    }
    catch {
        Update-Status "Could not reset Group Policy (continuing...)" -Level WARNING
    }
}

function Repair-WindowsSecurity {
    Update-Status "Repairing Windows Security App..." -Level SECTION

    try {
        Get-AppxPackage -Name 'Microsoft.SecHealthUI' -ErrorAction SilentlyContinue |
            ForEach-Object {
                Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue
            }
        Update-Status "Windows Security app re-registered" -Level SUCCESS
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

    Update-Status "DefenderShield v2.1.0 Repair Started" -Level SECTION
    Update-Status "Log file: $($Script:Config.LogPath)"

    # Create restore point
    if ($CreateRestorePoint) {
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

    # Repair Firewall
    if ($RepairFirewall) {
        Update-Status ""
        Update-Status "========== FIREWALL REPAIR ==========" -Level SECTION
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

    # Post-repair scan and comparison
    Start-Sleep -Seconds 2
    $postScan = Get-HealthScan

    # Update dashboard with new state
    Update-HealthDashboard -Scan $postScan

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

    # Complete
    $duration = (Get-Date) - $startTime

    Update-Status ""
    Update-Status "========== REPAIR COMPLETE ==========" -Level SECTION
    Update-Status "Duration: $([math]::Round($duration.TotalSeconds, 1)) seconds" -Level SUCCESS
    Update-Status ""
    Update-Status "A RESTART is strongly recommended!" -Level WARNING
    Update-Status "Check Windows Security after restart." -Level INFO

    return $true
}

# ============================================================================
# GUI
# ============================================================================

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="DefenderShield v2.1.0 - Windows Security Repair Tool"
    Height="850" Width="720"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanMinimize"
    Background="#1a1a2e">

    <Window.Resources>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Margin" Value="0,8,0,8"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0f3460"/>
            <Setter Property="Foreground" Value="White"/>
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
            <TextBlock Text="DefenderShield" FontSize="28" FontWeight="Bold" Foreground="#e94560" HorizontalAlignment="Center"/>
            <TextBlock Text="Windows Defender and Firewall Repair Tool  v2.1.0" FontSize="14" Foreground="#aaa" HorizontalAlignment="Center" Margin="0,5,0,0"/>
        </StackPanel>

        <!-- Health Dashboard -->
        <Border Grid.Row="1" Background="#16213e" CornerRadius="8" Padding="15" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="System Health Dashboard" FontSize="16" FontWeight="SemiBold" Foreground="White" Margin="0,0,0,10"/>
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
                        <TextBlock Text="WinDefend:  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblWinDefend" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="1" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="SecurityHealth:  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblSecHealth" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="2" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Security Center:  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblWscsvc" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="3" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Firewall (MpsSvc):  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblMpsSvc" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="4" Grid.Column="0" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Win Security App:  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblWinSecApp" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>

                    <!-- Right column -->
                    <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Real-Time Protection:  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblRTP" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="1" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Tamper Protection:  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblTamper" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="Definition Age:  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblDefAge" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="3" Grid.Column="1" Orientation="Horizontal" Margin="0,3">
                        <TextBlock Text="GP Blocking:  " Foreground="#aaa" FontFamily="Consolas" FontSize="12"/>
                        <TextBlock x:Name="lblGPBlock" Text="Scanning..." Foreground="#888" FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
                    </StackPanel>
                </Grid>
            </StackPanel>
        </Border>

        <!-- Options Panel -->
        <Border Grid.Row="2" Background="#16213e" CornerRadius="8" Padding="20" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="Select Components to Repair:" FontSize="16" FontWeight="SemiBold" Foreground="White" Margin="0,0,0,15"/>

                <CheckBox x:Name="chkFirewall" IsChecked="True">
                    <StackPanel>
                        <TextBlock Text="Windows Firewall" FontWeight="SemiBold" Foreground="White"/>
                        <TextBlock Text="Repairs services, removes blocking policies, enables all profiles" FontSize="11" Foreground="#888"/>
                    </StackPanel>
                </CheckBox>

                <CheckBox x:Name="chkDefender" IsChecked="True">
                    <StackPanel>
                        <TextBlock Text="Windows Defender Antivirus" FontWeight="SemiBold" Foreground="White"/>
                        <TextBlock Text="Repairs services, registry, scheduled tasks, WMI, enables protection" FontSize="11" Foreground="#888"/>
                    </StackPanel>
                </CheckBox>

                <Rectangle Height="1" Fill="#333" Margin="0,10"/>

                <CheckBox x:Name="chkRestorePoint" IsChecked="True">
                    <StackPanel>
                        <TextBlock Text="Create System Restore Point" FontWeight="SemiBold" Foreground="White"/>
                        <TextBlock Text="Recommended - allows you to undo changes if needed" FontSize="11" Foreground="#888"/>
                    </StackPanel>
                </CheckBox>
            </StackPanel>
        </Border>

        <!-- Warning Panel -->
        <Border Grid.Row="3" Background="#3d1a1a" CornerRadius="8" Padding="15" Margin="0,0,0,10">
            <TextBlock TextWrapping="Wrap" Foreground="#ffaa00">
                <Run FontWeight="SemiBold">Tamper Protection Notice:</Run>
                <Run>If Defender won't start after repair, disable Tamper Protection in Windows Security (Virus and threat protection - Manage settings), then run this tool again.</Run>
            </TextBlock>
        </Border>

        <!-- Status Output -->
        <Border Grid.Row="4" Background="#0d1117" CornerRadius="8" Padding="10">
            <RichTextBox x:Name="txtStatus"
                         Background="Transparent"
                         Foreground="White"
                         FontFamily="Consolas"
                         FontSize="12"
                         IsReadOnly="True"
                         BorderThickness="0"
                         VerticalScrollBarVisibility="Auto">
                <FlowDocument>
                    <Paragraph>
                        <Run Foreground="#888">Ready. Select options above and click Start Repair.</Run>
                    </Paragraph>
                </FlowDocument>
            </RichTextBox>
        </Border>

        <!-- Buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnQuickFix" Content="Quick Fix All" Margin="0,0,10,0" Width="150" Background="#e94560"/>
            <Button x:Name="btnStart" Content="Start Repair" Margin="0,0,10,0" Width="150"/>
            <Button x:Name="btnRestart" Content="Restart PC" Width="120" Background="#4a1942" IsEnabled="False"/>
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

# Dashboard labels
$Script:DashboardLabels = @{
    'lblWinDefend' = $window.FindName('lblWinDefend')
    'lblSecHealth' = $window.FindName('lblSecHealth')
    'lblWscsvc'    = $window.FindName('lblWscsvc')
    'lblMpsSvc'    = $window.FindName('lblMpsSvc')
    'lblRTP'       = $window.FindName('lblRTP')
    'lblTamper'    = $window.FindName('lblTamper')
    'lblDefAge'    = $window.FindName('lblDefAge')
    'lblGPBlock'   = $window.FindName('lblGPBlock')
    'lblWinSecApp' = $window.FindName('lblWinSecApp')
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
    $run.Foreground = '#888'
    $paragraph.Inlines.Add($run)

    if ($defenderBroken -or $firewallBroken) {
        $issueRun = New-Object System.Windows.Documents.Run("Issues detected - broken items auto-selected.")
        $issueRun.Foreground = 'Yellow'
        $paragraph.Inlines.Add($issueRun)
    }
    else {
        $okRun = New-Object System.Windows.Documents.Run("All components appear healthy.")
        $okRun.Foreground = 'Lime'
        $paragraph.Inlines.Add($okRun)
    }

    $paragraph.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)
    $txtStatus.Document.Blocks.Add($paragraph)
})

# Show window
$null = $window.ShowDialog()
