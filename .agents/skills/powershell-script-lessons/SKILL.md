---
name: PowerShell Script Lessons (Winget + Admin Flow)
description: Use this skill when writing or fixing interactive PowerShell automation that runs external CLI tools (especially `winget`) and needs resilient UX, diagnostics, and elevation flow.

---
## When To Use

- Interactive update/install scripts with per-item processing.
- Scripts that need to retry failures with elevation.
- Scripts where users need readable failure diagnostics.
- Any script that should always pause before exit so users can review output.

## Core Principles

1. Do not parse human-oriented tables with naive whitespace splitting when machine formats exist.
2. Process all selected items; never stop on first failure.
3. Separate "run UX" from "diagnostics capture".
4. Preserve shell parity for elevation (`pwsh` should relaunch `pwsh`).
5. Guarantee an end-of-run pause with a `finally` block.

## Patterns

### 1) Robust CLI Listing Parse

- Prefer `--output json` when supported.
- Support schema variants (`PackageIdentifier` vs `Id`, `InstalledVersion` vs `Version`, etc.).
- Deduplicate by stable key (`Id`).
- Only fall back to text parsing when JSON is unavailable.

### 2) Single-Item Parameter Safety

PowerShell can pass a single object differently than arrays.

- Accept `[object[]]` for collection parameters.
- Iterate with `foreach ($x in @($Items))`.

### 3) Progress Bar Rendering

If you redirect output to files and replay it, progress bars print as many lines.

- For clean in-place progress, invoke CLI directly: `& winget @args`.
- Capture diagnostics with tool-native logging (`winget --log <path>`) instead of stream redirection.
- Do not pipe the normal interactive `winget` run through `Tee-Object` or other PowerShell pipeline stages when the operator needs the native progress bar or spinner.

### 4) Better Error Messages

Store and report:

- Tool exit code in decimal and hex (HRESULT-like values).
- Installer exit code (e.g., MSI `1603`) parsed from log file.
- Human-readable hint and concrete retry suggestion.
- When the wrapper log does not contain the real cause, correlate to winget's own CLI diagnostic log under `AppData\Local\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\DiagOutputDir`.
- Prefer exact correlation using the command's `--log` path before reading CLI diagnostic text; do not guess from a package-id-only match if an exact log-path match is missing.
- Treat installer exit `3010` as restart-required state, not a normal hard failure that should automatically enter the retry ladder.

### 5) Elevation Fidelity

Do not hardcode `powershell`.

- Use current host executable: `(Get-Process -Id $PID).Path`.
- Fallback based on edition (`pwsh` for Core, `powershell` for Desktop).

### 6) Retry Failed Items In Elevated Window

- Collect failed package IDs.
- Prompt user once to retry only failures in an elevated session.
- Keep successful packages untouched.

### 7) Always Pause At End

Use structured flow control to ensure pause executes on every path.

- Wrap control flow in `try/catch/finally`.
- Signal intentional exits through a known sentinel exception.
- Read final input in `finally`, then `exit` once with tracked code.

## Suggested Implementation Skeleton

```powershell
$script:RequestedExitCode = 0
$script:PausePrompt = "Press Enter to close this window"

function Exit-WithPause {
  param([int]$Code = 0, [string]$Prompt = "Press Enter to close this window")
  $script:RequestedExitCode = $Code
  $script:PausePrompt = $Prompt
  throw [System.OperationCanceledException]::new("__SCRIPT_EXIT__")
}

try {
  # main flow
  # run all selected items
  # summarize failures
  # optionally retry failures elevated
  Exit-WithPause
}
catch [System.OperationCanceledException] {
  if ($_.Exception.Message -ne "__SCRIPT_EXIT__") {
    $script:RequestedExitCode = 1
    throw
  }
}
catch {
  $script:RequestedExitCode = 1
  Write-Warning "Unexpected error: $($_.Exception.Message)"
}
finally {
  Read-Host $script:PausePrompt | Out-Null
}

exit $script:RequestedExitCode
```

## Anti-Patterns To Avoid

- `-split '\s{2,}'` table parsing when columns may contain spaces.
- `[System.Collections.IEnumerable]` parameters for user selections without array normalization.
- `Start-Process ... -RedirectStandardOutput` for tools with animated progress output.
- `2>&1 | Tee-Object` around the normal interactive `winget` execution path when native progress is part of the intended UX.
- Falling back to the newest winget CLI diag log that mentions a package id when the exact command log path cannot be matched; this can bind retries to stale failure evidence.
- Hardcoded `Start-Process "PowerShell" -Verb RunAs`.
- Multiple `exit` points that bypass user pause.

## Validation Checklist

- Script parses with PowerShell parser API.
- Single package selection works.
- Multiple package selection works.
- Failed package summary includes actionable guidance.
- Restart-required results (for example installer `3010`) are reported as reboot-needed rather than generic failures.
- Diagnostic correlation prefers exact command/log matches and does not reuse stale retry logs.
- Retry-elevated prompt appears when failures occur.
- Final pause appears for success, failure, skip, and unexpected error paths.
