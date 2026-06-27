@{
    RootModule        = 'DefenderShield.psm1'
    ModuleVersion     = '3.1.0'
    GUID              = '9bfe1fe0-f6d8-49d5-8f86-5c06501f23da'
    Author            = 'Matt'
    CompanyName       = 'SysAdminDoc'
    Copyright         = '(c) 2026 Matt. All rights reserved.'
    Description       = 'Windows Defender and Firewall repair tool for reversing privacy-tool breakage.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Invoke-DefenderShield')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @('Repair-DefenderShield')
    FileList          = @('DefenderShield.ps1', 'DefenderShield.psm1', 'DefenderShield.psd1', 'README.md', 'LICENSE')
    PrivateData       = @{
        PSData = @{
            Tags         = @('Defender', 'Firewall', 'WindowsSecurity', 'Repair', 'PowerShell')
            LicenseUri   = 'https://github.com/SysAdminDoc/DefenderShield/blob/main/LICENSE'
            ProjectUri   = 'https://github.com/SysAdminDoc/DefenderShield'
            ReleaseNotes = 'v3.1.0 adds undo manifests, HTML repair reports, async GUI repair, dashboard tiles, firewall rule preservation, AppLocker/SRP repair, MDE repair, Windows Update repair, module packaging, status snapshots, watchdog tasks, fleet mode, DefenderControl replay, portable logs, and third-party AV guidance.'
        }
    }
}
