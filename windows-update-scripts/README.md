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

## Validation

After changing [`winget-ms-store-update-all.ps1`](/E:/Sandbox/danmade-playground/windows-update-scripts/winget-ms-store-update-all.ps1), run:

```powershell
powershell -NoProfile -File .agents/skills/winget-parser-maintainer/scripts/test-winget-fallback-parser.ps1
```
