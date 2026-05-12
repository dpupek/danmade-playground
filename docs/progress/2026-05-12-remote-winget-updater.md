Progress
- Added a separate remote-safe winget updater at `windows-update-scripts/winget-ms-store-update-remote.ps1`.
- Kept the remote flow non-interactive so it works over PowerShell remoting without `Out-GridView`, `Read-Host`, local UAC relaunch, or pause prompts.
- Reused the winget upgrade-list parsing approach from the local updater so remote enumeration still works without depending on UI flow.
- Added per-package remote log capture under `C:\ProgramData\winget-update-script-logs` and structured result objects for caller-side review.
- Updated `windows-update-scripts/README.md` with remote usage examples and caveats.
- Created GitHub issue `#1` to track the requested follow-up: change the remote default from preview/list-only to silent upgrade-all while keeping graceful failures and logging.

Decisions
- Kept the remote implementation in a separate script instead of forcing the interactive local updater through remoting, because the local script depends on prompts, grid selection, and local elevation behavior that do not translate cleanly to WinRM sessions.
- Made the first remote version conservative by defaulting to list-only until the desired default behavior was clarified.

Validation
- `pwsh -NoProfile -File .\windows-update-scripts\winget-ms-store-update-remote.ps1 -ComputerName localhost -WhatIf` -> wrapper path resolves and advertises list behavior.
- `pwsh -NoProfile -File .\windows-update-scripts\winget-ms-store-update-remote.ps1 -ComputerName localhost -PackageId Git.Git -WhatIf` -> wrapper path resolves and advertises targeted upgrade behavior.
- Live remote execution against another machine was not run, so target-environment behavior is still unverified.

Files
- `windows-update-scripts/winget-ms-store-update-remote.ps1`
- `windows-update-scripts/README.md`
- `docs/progress/2026-05-12-remote-winget-updater.md`

Next steps
- Change the remote default mode to silent upgrade-all.
- Add an explicit preview/list-only switch instead of making preview the default.
- Expand remote failure hints where useful to match more of the local updater behavior.
- Validate the remote script against at least one real WinRM target.
- If GitHub issue comments are needed from this environment, add a supported path for posting issue updates.
