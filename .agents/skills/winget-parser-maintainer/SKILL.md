---
name: winget-parser-maintainer
description: Maintain winget parsing in windows-update-scripts/winget-ms-store-update-all.ps1. Use when changing package list parsing, fallback behavior, ID normalization, or upgrade selection display. Always run the bundled fallback-parser regression tests before finalizing changes.
---

# Winget Parser Maintainer

Run this workflow when editing `windows-update-scripts/winget-ms-store-update-all.ps1`.

## Workflow

1. Keep JSON parsing as primary path. Treat text parsing as compatibility fallback.
2. Preserve support for both dotted and non-dotted IDs.
3. Preserve support for single-space and multi-space table separators.
4. Preserve normalization that strips leading mojibake artifacts from IDs.
5. When restoring native winget progress, keep the listing/parser logic separate from installer diagnostics; do not reintroduce pipeline capture on the interactive execution path just to inspect output.
6. If you add diagnostic-log correlation, prefer exact `--log` path matches over package-id-only heuristics to avoid stale retry evidence.
7. Run regression tests before and after parser edits.

## Required Validation

Run:

```powershell
powershell -NoProfile -File .agents/skills/winget-parser-maintainer/scripts/test-winget-fallback-parser.ps1
```

Do not finalize parser changes unless tests pass.

## Guardrails

- Do not require a dot in package IDs.
- Do not use fixed column index parsing for fallback table rows.
- Do not duplicate parser logic in tests; test against loaded functions from the target script.
- Do not weaken the interactive winget UX just to make parser or diagnostic helpers easier to implement.
- Do not let a missing wrapper log silently fall back to the newest diag log for the same package id when that could belong to a different retry attempt.
