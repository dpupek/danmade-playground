# Winget Python Runtime Diagnostics

## Summary

- Improved `windows-update-scripts/winget-ms-store-update-all.ps1` so the winget fallback parser handles upgrade rows whose version columns contain spaces, including parenthetical build numbers and comparator-prefixed versions.
- Added Python Install Manager diagnostics for managed runtimes that are detected but not selected for upgrade.
- Broadened Python runtime online matching across runtime IDs, tags, `install-for`, and `run-for` aliases.
- Added a cleaned-up `py install --update --dry-run` summary when a managed Python runtime is skipped.

## Local Result

- The managed Python runtime at `C:\Users\dan.pupek\AppData\Local\Python\pythoncore-3.14-64\python.exe` is now reported explicitly as already current when Python Install Manager sees installed `3.14.5` and online `3.14.5`.
- The live upgrade helper check still reports the available Zoom Store upgrade normally.

## Validation

```powershell
powershell -NoProfile -File .agents/skills/winget-parser-maintainer/scripts/test-winget-fallback-parser.ps1
```

Result: pass.

## Files

- `windows-update-scripts/winget-ms-store-update-all.ps1`
- `.agents/skills/winget-parser-maintainer/scripts/test-winget-fallback-parser.ps1`
- `docs/progress/2026-06-03-winget-python-runtime-diagnostics.md`
