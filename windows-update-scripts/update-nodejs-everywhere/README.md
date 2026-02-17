# Update Node Everywhere (Windows)

This folder contains a resilient PowerShell script to update Node.js on Windows servers in both common install models:

- NVM-managed Node.js versions
- System-installed Node.js (winget first, MSI fallback)

The script is designed to:

- log each action to console and file
- continue to later steps if one step fails
- summarize success/fail status for each step
- pause at the end so the operator can review results

## Files

- `Update-NodeEverywhere.ps1`
- `Update-NodeEverywhere.cmd`
- `logs/` (created automatically at runtime)

## Usage

Run from an elevated PowerShell window:

```powershell
.\Update-NodeEverywhere.ps1
```

Run from Command Prompt (or by double-clicking the `.cmd` file):

```cmd
Update-NodeEverywhere.cmd
```

Dry run:

```powershell
.\Update-NodeEverywhere.ps1 -WhatIf
```

Update and remove older NVM patch versions per major branch:

```powershell
.\Update-NodeEverywhere.ps1 -CleanupOldNvmVersions
```

Update only system install using current (non-LTS) channel:

```powershell
.\Update-NodeEverywhere.ps1 -UpdateNvm:$false -Channel Current
```

## Notes

- `nvm` must be in `PATH` for NVM update steps.
- `winget` is used when available; otherwise the script downloads the official MSI from `nodejs.org`.
- Script exits after user presses Enter (pause prompt at the end).
