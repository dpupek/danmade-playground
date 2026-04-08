# Roadmap

## Phase 1: Baseline and decisions

- [x] Confirm the game lives under `pitfall-clone`.
- [x] Confirm the target repo path is `E:\Sandbox\dungeon-dan`.
- [x] Decide to preserve history during extraction.
- [x] Decide the canonical identity is Dan's Dungeon.
- [x] Decide to remove the game from the source repo after validation.

## Phase 2: Extraction and normalization

- [x] Create the shaping docs in the source repo.
- [x] Extract `pitfall-clone` into a standalone repository rooted at `E:\Sandbox\dungeon-dan`.
- [x] Normalize package, docs, and browser-facing names to Dan's Dungeon.
- [x] Update any hard-coded local setup paths to the new repo root.

## Phase 3: Validation and cutover

- [x] Run install, build, and test in the extracted repo.
- [x] Verify the browser shell title reflects Dan's Dungeon.
- [x] Remove `pitfall-clone` from `danmade-playground`.
- [x] Confirm the source repo only contains the intentional split changes.
