# Baseline

## Current Repository Shape

- Source repo root currently contains:
  - `.agents`
  - `windows-update-scripts`
  - `pitfall-clone`
- The source repo working tree was clean before the split work started.

## Game Self-Containment

`pitfall-clone` already contains its own:

- `package.json` and `package-lock.json`
- `tsconfig.json`
- `vite.config.ts`
- game source under `src/`
- docs under `docs/`
- asset-generation script under `scripts/`
- tests under `src/**/*.test.ts`
- static assets under `public/`

## Naming Drift

The game currently uses mixed names:

- folder and package name: `pitfall-clone`
- README and HTML title: `Temple Runaway`
- runtime config title: `Dan's Dungeon`

## Split Constraints

- Preserve history for the extracted game path.
- New local repo target: `E:\Sandbox\dungeon-dan`
- No remote bootstrap or CI work in this iteration.
- Remove `pitfall-clone` from the source repo after validation.
