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
5. Run regression tests before and after parser edits.

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

