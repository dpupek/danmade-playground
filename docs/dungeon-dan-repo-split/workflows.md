# Dungeon Dan Repo Split Workflow

## Personas

- Maintainer: owns the `danmade-playground` repo and wants the game isolated from unrelated utility scripts.
- Game contributor: needs a focused repository with accurate naming, setup steps, and preserved history.
- Future collaborator: needs a clean repo root that only contains the game and its supporting docs.

## Motivation

The current repository mixes a standalone Phaser game with unrelated Windows maintenance scripts. Splitting the game reduces repository noise, makes onboarding clearer, and preserves the game's evolution separately from script work.

## Workflow

1. Document the current repo shape and split decisions in `docs/dungeon-dan-repo-split/`.
2. Extract `pitfall-clone` into its own repository rooted at `E:\Sandbox\dungeon-dan` while preserving git history.
3. Normalize the extracted repo identity around Dan's Dungeon without changing gameplay behavior.
4. Validate install, build, tests, and browser shell identity in the new repo.
5. Remove `pitfall-clone` from `danmade-playground` once the extracted repo is verified.

## Expected Outcome

- `E:\Sandbox\dungeon-dan` is a standalone git repository containing the game at repo root.
- Public-facing names consistently describe the game as Dan's Dungeon.
- `E:\Sandbox\danmade-playground` no longer contains the game folder.
