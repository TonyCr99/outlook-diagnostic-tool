# OUTLOOK Diagnostic Tool

A cyberpunk-styled PowerShell CLI for diagnosing the most common Outlook/Exchange Online support tickets: mail not arriving/sending, sync issues, and a slow-loading desktop client — including the classic ">100K items in a folder" hang/loop.

Available in **Spanish, English, and Portuguese**. Runs on Windows PowerShell 5.1+ or PowerShell 7+.

```
   ███  █   █ █████ █      ███   ███  █   █
  █   █ █   █   █   █     █   █ █   █ █  █
  █   █ █   █   █   █     █   █ █   █ ███
  █   █ █   █   █   █     █   █ █   █ █  █
   ███   ███    █   █████  ███   ███  █   █

        DIAGNOSTIC TOOL
        Performance & Sync Diagnostics
```

## What it does

Combines two diagnostic passes in one tool:

- **Remote** — connects to Exchange Online PowerShell and inspects the mailbox: quota usage, folder item counts (flags folders over the 100K-item Microsoft-recommended limit — the classic cause of client hangs/loops), inbox rules, mailbox-level forwarding, mobile device sync status, and a 48h message trace for inbound/outbound delivery failures.
- **Local** — inspects the Outlook client on the machine it runs on: RAM/disk, OST/PST file size, autocomplete cache, active add-ins, running processes, and TCP connectivity to the core Microsoft 365 endpoints.

Every check produces a "finding" with a severity (`OK` / `INFO` / `WARN` / `CRIT`), a detail block, and — for warnings/criticals — a recommendation. A running summary and an exportable, dark-themed HTML report tie it all together.

## Remediation

A guided remediation menu turns the most common findings into safe, auditable fixes:

| Action | What it does |
|---|---|
| Enable online archive | `Enable-Mailbox -Archive` (+ optional auto-expanding archive) |
| Move old items to archive | Assigns a retention policy and runs the Managed Folder Assistant — the recommended fix for the >100K-items case (moves mail, doesn't delete it) |
| Age-guided deletion | Content Search–based **soft-delete only** (recoverable), with a dry-run estimate before anything runs |
| Client repair runbook | Prints the manual steps for `/safe`, `/cleanviews`, `/resetnavpane`, OST regeneration, and reducing the cached-mode sync window |

Every remediation action requires **typed confirmation** of an exact phrase before it runs, and is logged to a local audit file. Before attempting a write action, the tool checks whether your Exchange Online session actually has the required cmdlet (a common RBAC/role gap) and tells you which admin role to check for instead of surfacing a raw PowerShell error.

## Requirements

- Windows PowerShell 5.1+ (built into Windows) or PowerShell 7+
- [`ExchangeOnlineManagement`](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) module, for remote checks:
  ```powershell
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
  ```
- An Exchange Online account with at least read access for diagnostics; remediation actions need write roles (the tool tells you which one is missing if you hit a gap)

## Usage

```powershell
# Interactive menu (language picker shown first)
.\Diagnose-OutlookPerf.ps1

# Skip the picker, force a language
.\Diagnose-OutlookPerf.ps1 -Language en

# Diagnose a specific mailbox directly
.\Diagnose-OutlookPerf.ps1 -Mailbox user@domain.com

# Unattended full diagnostic + HTML report
.\Diagnose-OutlookPerf.ps1 -Mailbox user@domain.com -Auto -ReportPath .\report.html

# Local client checks only (run on the affected user's machine)
.\Diagnose-OutlookPerf.ps1 -LocalOnly

# Mailbox checks only, skip local
.\Diagnose-OutlookPerf.ps1 -RemoteOnly -Mailbox user@domain.com
```

| Parameter | Description |
|---|---|
| `-Mailbox <upn>` | Target mailbox for remote checks |
| `-LocalOnly` | Only run the local client diagnostic |
| `-RemoteOnly` | Only run the mailbox diagnostic |
| `-Auto` | Run everything unattended and exit (no menu) |
| `-ReportPath <path>` | HTML report output path (used with `-Auto`) |
| `-Language es\|en\|pt` | Interface language (default: `es`) |

## Building a standalone .exe

`Build-NetrunnerExe.ps1` compiles a small native launcher (~5 KB) that locates and runs `Diagnose-OutlookPerf.ps1` from its own folder, so it can be double-clicked without typing PowerShell flags:

```powershell
.\Build-NetrunnerExe.ps1 -OutputPath .\OutlookDiagnostic.exe
```

Requires the .NET Framework C# compiler (`csc.exe`, included in Windows). Copy the generated `.exe` next to `Diagnose-OutlookPerf.ps1` — it still needs the script alongside it and PowerShell 5.1+ installed on the target machine.

## Notes

- All remediation deletions are **soft-delete only** — items go to Recoverable Items and can be restored during the retention window. There is no permanent-purge option.
- Adjustable thresholds (folder item limits, quota %, OST size, RAM, etc.) live in the `$Script:Thresholds` hashtable near the top of the script.
- An audit log (`OutlookDiag_Remediation_YYYYMMDD.log`) is written next to the script for every remediation action taken.

## License

MIT — see [LICENSE](LICENSE).
