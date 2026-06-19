# Windows Update Scripts

## Winget + Microsoft Store updater

Use [`run-winget-ms-store-update-all.cmd`](/E:/Sandbox/danmade-playground/windows-update-scripts/run-winget-ms-store-update-all.cmd) to launch [`winget-ms-store-update-all.ps1`](/E:/Sandbox/danmade-playground/windows-update-scripts/winget-ms-store-update-all.ps1).

Why the `.cmd` launcher exists:
- The PowerShell script now requires PowerShell 7 or later.
- If `pwsh` is not installed, the launcher can offer to install the latest PowerShell with `winget`.
- After PowerShell 7 is available, the launcher runs the script for you.

If you run the `.ps1` file directly from Windows PowerShell 5.1 or another unsupported host, it exits with an advisory to use the `.cmd` launcher or install PowerShell 7 first.

## Package selection UX

When `Out-GridView` is available, the script opens a grid picker for package selection.

Grid selection tips:
- `Ctrl+Click` selects nonconsecutive rows.
- `Shift+Click` selects a range.
- `Ctrl+A` selects all rows.
- Choose `OK` to continue or close/cancel the window to skip changes.

If `Out-GridView` is unavailable in the current session, the script falls back to the text-based numbered picker.

## PowerShell 7 special handling

If `Microsoft.PowerShell` is selected for upgrade, the updater does not try to replace PowerShell 7 from the active `pwsh` session.

Instead, it opens a separate Windows PowerShell helper window that:

- waits for all `pwsh.exe` processes to exit
- reminds you to close lingering PowerShell 7 terminals and check Task Manager if needed
- runs the deferred `winget` upgrade after PowerShell 7 is no longer in use

To exercise that path even when no real PowerShell update is currently available, run the main script with:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\windows-update-scripts\winget-ms-store-update-all.ps1 -TestModeAddPowerShell
```

That injects a test-only `Microsoft.PowerShell` entry into the picker so you can verify the deferred helper flow from the main updater.

## Remote winget updater

Use [`winget-ms-store-update-remote.ps1`](/E:/Sandbox/danmade-playground/windows-update-scripts/winget-ms-store-update-remote.ps1) when you need to query or run winget upgrades on another Windows computer over PowerShell remoting.

This remote variant is intentionally separate from the local interactive updater:
- no `Out-GridView`, `Read-Host`, or local UAC popup flow
- no end-of-run pause prompt
- defaults to listing available upgrades unless you explicitly pass `-All` or `-PackageId`
- returns structured objects so the calling session can filter, export, or review results

Examples:

```powershell
# List available upgrades on a remote computer
.\winget-ms-store-update-remote.ps1 -ComputerName SERVER01

# Upgrade all available winget packages on a remote computer
.\winget-ms-store-update-remote.ps1 -ComputerName SERVER01 -All

# Upgrade specific packages on a remote computer
.\winget-ms-store-update-remote.ps1 -ComputerName SERVER01 -PackageId Git.Git, Microsoft.PowerShell
```

Notes:
- PowerShell remoting/WinRM must already be enabled on the target computer.
- The remote session must already have enough rights to run `winget`; this script does not try to open a UAC prompt remotely.
- Some Microsoft Store or per-user installs may still behave differently in a remote admin session than they do in a local interactive desktop session.

## Danmade Patch Agent

Use [`danmade-patch-agent.ps1`](/E:/Sandbox/danmade-playground/windows-update-scripts/danmade-patch-agent.ps1) for unattended domain-managed patching through Group Policy scheduled tasks.

The patch agent is intentionally separate from the interactive updater:
- no prompts, `Out-GridView`, UAC relaunches, or end-of-run pause
- policy-driven package allow/block lists and retry limits
- Wazuh-friendly Windows Event Log and JSONL reporting
- designed to be signed with a Domain CA code-signing certificate and run under `AllSigned`

Start with [`danmade-patch-agent.policy.sample.json`](/E:/Sandbox/danmade-playground/windows-update-scripts/danmade-patch-agent.policy.sample.json), then follow the deployment and signing guide in [`docs/danmade-patch-agent/admin-guide.md`](/E:/Sandbox/danmade-playground/docs/danmade-patch-agent/admin-guide.md).

## Portable package cleanup

If you want to remove installed `winget` portable packages and delete their package directories through `winget uninstall --purge`, run:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\windows-update-scripts\winget-ms-store-update-all.ps1 -CleanupPortablePackages
```

This cleanup mode is explicit on purpose:

- it lists installed portable packages registered with `winget uninstall --product-code`
- it lets you choose which ones to remove
- it does not automatically remove a different major version just because a newer branch is installed

For example, `PHP.PHP.8.2` and `PHP.PHP.8.3` are separate packages. Cleanup mode lets you remove one or both intentionally.

## Validation

After changing [`winget-ms-store-update-all.ps1`](/E:/Sandbox/danmade-playground/windows-update-scripts/winget-ms-store-update-all.ps1), run:

```powershell
powershell -NoProfile -File .agents/skills/winget-parser-maintainer/scripts/test-winget-fallback-parser.ps1
```
