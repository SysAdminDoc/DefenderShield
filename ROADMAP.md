# DefenderShield Roadmap

Repair tool for restoring Windows Defender and Firewall after privacy scripts have disabled them. Tracks work after current release.

## Planned Features

### Core Repair
- Policy source audit: enumerate every active Defender/Firewall blocker (registry, WMI subscription, scheduled task, AppLocker rule, GPO) and render a single "what's holding this down" table before repair
- JSON undo manifest of every change (registry / service / task) so a repair can be rolled back if it breaks something
- Dry-Run mode: simulate without writing (match Debloat-Win11's pattern)
- `-Only` / `-Skip` phase flags so the user can rerun just the Services phase after a failed run
- Repair verification: post-run `Get-MpComputerStatus` + `Get-NetFirewallProfile` assertion; exit code reflects partial vs full restoration
- Auto-detect privacy tools that ran on this box (signatures for privacy.sexy, O&O ShutUp10, Debloater scripts) and surface them in the report

### UI
- Restyle WPF to Catppuccin Mocha (match other project GUIs)
- Dashboard tiles: Defender status, Firewall status, Tamper Protection, Signature age, 3rd-party AV present
- Progress bar + streaming log (async runspace) so the UI never freezes
- Report pane: exportable HTML with every change, color-coded for SUCCESS / WARN / ERROR
- Tamper Protection helper: one-click open Windows Security with inline instructions and a "re-check" button

### CLI
- First-class CLI alongside the GUI: `-Mode Firewall|Defender|Both|Status`, `-DryRun`, `-Silent`
- Structured JSON output via `-Json`
- Exit codes for PDQ / Intune / SCCM detection
- `Install-Module DefenderShield` publish path

### Coverage Expansion
- Detect and remove AppLocker/SRP rules that block MsMpEng
- Detect and repair broken MsSense / MDE components on Windows 11 Enterprise
- Repair Windows Security UI if SecHealthUI was deprovisioned (`Get-AppxPackage -AllUsers Microsoft.SecHealthUI | Add-AppxPackage -Register`)
- Repair SmartScreen independently (often flipped off by the same privacy tools)
- Repair Windows Update surfaces that privacy tools flip: `wuauserv`, `UsoSvc`, `DoSvc`, `BITS` start-type
- Handle WMI event subscription removal with a dedicated phase + report

### Safety
- Firewall service order sanity: validate `BFE → mpssvc → IKEEXT → PolicyAgent` dependency tree before starting services
- Pre-flight: abort and warn if a third-party AV is registered as the active provider (don't fight it)
- System Restore point creation via `Checkpoint-Computer` with throttle awareness
- Preserve user-defined firewall rules: export all custom rules before reset, re-import after

### Packaging
- Authenticode signing + SHA256SUMS per release
- PSGallery module + chocolatey package
- GitHub Actions workflow producing signed `.ps1`, `.nupkg`, and `.zip` artifacts on tag

## Competitive Research

- **ReSetX / Windows Repair Toolbox** — Closed-source repair shops' kits; validates demand but poor auditability. DefenderShield's open-source + structured log is the differentiator.
- **Windows-Defender-Remover `-restore` flag** — Single-purpose reverse of the same author's disabler; narrower scope, useful reference for Appx reprovisioning logic.
- **privacy.sexy** — The tool most often responsible for the breakage DefenderShield fixes. Keeping parity with their published keys means a weekly Action syncing their scripts to a `known-blockers.json`.
- **O&O ShutUp10++** — Closed-source privacy tool with a built-in "reset defaults" option. Mention as a first-try option in docs so users don't default to DefenderShield for trivial reversals.

## Nice-to-Haves

- "What changed?" diff: compare two Status snapshots taken weeks apart to spot when Defender drifted off
- Watchdog scheduled task (opt-in) that re-runs Repair if it sees Defender/Firewall flipped off
- Fleet mode (WinRM) with explicit opt-in to repair multiple machines in a session
- Built-in AV market share lookup — when a 3rd-party AV is detected, show its uninstall instructions rather than fighting it
- Integration with DefenderControl: if DefenderControl's undo manifest is on disk, DefenderShield prefers replaying it instead of generic repair
- Minimal "portable" mode that runs entirely from a USB stick with writable logs to `.\Logs\`

## Open-Source Research (Round 2)

### Related OSS Projects
- **zoicware/DefenderProTools** — https://github.com/zoicware/DefenderProTools — Has a restore-path that repopulates Defender registry keys, services, and scheduled tasks; closest peer for "undo" logic.
- **AndyFul/ConfigureDefender** — https://github.com/AndyFul/ConfigureDefender — Exposes the full set of Defender registry toggles that are not surfaced in the Security Center UI; useful reference for which values to validate post-repair.
- **metablaster/WindowsFirewallRuleset** — https://github.com/metablaster/WindowsFirewallRuleset — Ships `Reset-Firewall.ps1` plus a log of every service whose startup mode changed; model for reversible firewall ops.
- **TairikuOokami/Windows** (Microsoft Defender Enable.bat) — https://github.com/TairikuOokami/Windows/blob/main/Microsoft%20Defender%20Enable.bat — Concrete re-enable batch for autologgers (DefenderApiLogger/DefenderAuditLogger) and scheduled tasks (Cache Maintenance, Cleanup, Scheduled Scan, Verification).
- **axelmierczuk Windows Defender Link Fix gist** — https://gist.github.com/axelmierczuk/c18496eb984091b287d95ccbcf1bd4d9 — Reinstalls `Microsoft.SecHealthUI` AppxPackage + Microsoft.UI.Xaml.2.7; useful when Settings-app Defender pane is broken.
- **builtbybel/privatezilla** — https://github.com/builtbybel/privatezilla — Privacy audit that can report the delta between "secure" and "default" states; inverse tool worth linking from DefenderShield's README.
- **modzero/fix-windows-privacy** — https://github.com/modzero/fix-windows-privacy — ~130 XML-defined privacy rules with built-in "Restore Backup"; backup/restore pattern transferable.
- **ElliotKillick/LdrLockLiberator** — https://github.com/ElliotKillick/LdrLockLiberator — Tangential; useful reference for SYSTEM-level privilege acquisition patterns.

### Features to Borrow
- SecHealthUI AppX re-register step when the Settings Defender page is blank/missing — borrow from `axelmierczuk` gist.
- Restore known-good scheduled tasks list (Cache Maintenance / Cleanup / Scheduled Scan / Verification / autologgers) — borrow from `TairikuOokami` enable batch.
- Post-repair validation matrix: verify every setting is ≥ "default" using the same table `ConfigureDefender` displays — borrow from `ConfigureDefender` UI.
- Snapshot-backup + `Restore Backup` UI so users can undo DefenderShield itself — borrow from `modzero/fix-windows-privacy`.
- Changelog of every service whose StartType was touched — borrow from `WindowsFirewallRuleset` logging.
- "Repair from privacy.sexy preset X" one-click fix profiles targeting specific breakage patterns — borrow from `privatezilla` community rule-sets model.
- Detect/fix missing `WinDefend` / `mpssvc` / `BFE` services by restoring from `.reg` templates — borrow from Tenforums Safe-Mode procedure codified in scripts.

### Patterns & Architectures Worth Studying
- `WindowsFirewallRuleset` two-phase approach: capture pre-change state to a log, then reset; post-run diff shows exactly what changed. Good model for DefenderShield's "what did I repair?" report.
- `ConfigureDefender`'s table of (RegKey, Name, ExpectedValue, Status) — reuse as DefenderShield's repair-verification checklist so every phase ends with a pass/fail matrix.
- `modzero`'s XML-rule DSL for privacy ops — if DefenderShield reaches >30 repair ops, externalize them to a similar declarative file for easier community contribution.
