# CLAUDE.md - DefenderShield

## Overview
Repair and restore Windows Defender and Firewall after debloaters (privacy.sexy, O&O ShutUp10, etc.) have broken them. The inverse of DefenderControl. v2.1.0.

## Tech Stack
- PowerShell 5.1, WPF GUI with checkboxes for selective repair

## Key Details
- Single-file PowerShell script
- **Health Dashboard**: Auto-scans on launch, shows color-coded status of 9 components (WinDefend, SecurityHealthService, wscsvc, MpsSvc, RTP, Tamper Protection, definition age, GP blocking, Windows Security app)
- **Auto-select**: Broken items are auto-checked, healthy items unchecked
- **Before/After comparison**: Post-repair re-scan with formatted comparison in log panel
- **Quick Fix All**: One-click button selects all broken items and runs repair
- Repairs: service registry, blocking policy keys, scheduled tasks, WMI subscriptions, Group Policy reset, Set-MpPreference re-enablement, Windows Security app re-registration
- Backs up registry keys before modifications
- Logs to Desktop (`DefenderShield_*.log`)
- Backups to Desktop (`DefenderShield_Backup_*`)
- "Restart PC" button activates after repair

## Build/Run
```powershell
.\DefenderShield.ps1
```

## Version History
- 2.1.0 - Health dashboard, auto-select broken items, before/after comparison, Quick Fix All button
- 2.0.0 - Initial GUI version with selective component repair
