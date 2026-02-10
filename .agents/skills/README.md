# Local Skills Index

This directory contains repo-local skills for Codex agents.

## Available Skills

### `powershell-script-lessons`
- Path: `.agents/skills/powershell-script-lessons/SKILL.md`
- Purpose: Practical patterns for writing resilient interactive PowerShell automation (especially `winget` flows).
- Use when: You are building/fixing PowerShell scripts with package upgrades, progress rendering, elevation retry logic, and graceful failure summaries.

## Usage Notes

- Open the skill's `SKILL.md` and follow it directly.
- Prefer adapting existing script patterns from this repo instead of rewriting from scratch.
- Keep edits minimal and validate with a PowerShell parser check before finalizing.
