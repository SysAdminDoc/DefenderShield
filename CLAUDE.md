# CLAUDE.md - DefenderShield

## Overview
Repair and restore Windows Defender and Firewall after debloaters (privacy.sexy, O&O ShutUp10, etc.) have broken them. The inverse of DefenderControl. v2.0.0.

## Tech Stack
- PowerShell 5.1, WPF GUI with checkboxes for selective repair

## Key Details
- ~848 lines, single-file
- Repairs: service registry, blocking policy keys, scheduled tasks, WMI subscriptions, Group Policy reset, Set-MpPreference re-enablement, Windows Security app re-registration
- Backs up registry keys before modifications
- Logs to Desktop (`DefenderShield_*.log`)
- Backups to Desktop (`DefenderShield_Backup_*`)
- "Restart PC" button activates after repair

## Build/Run
```powershell
.\DefenderShield.ps1
```

## Version
2.0.0
